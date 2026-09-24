"""Garante que os endpoints de debug mostram exatamente o que o extrator
recebe. Motivo: uma fixture escrita a partir de um trecho colado a mao ja
divergiu do texto real que o extrator recebia (ver CLAUDE.md, "Layout
real de uma DANFE") -- entao a ponte entre "o que eu inspeciono" e "o que
o extrator ve" precisa ser garantida por teste, nao por leitura de codigo.

httpx (TestClient) nao esta nas dependencias, entao chama as funcoes dos
endpoints direto com um UploadFile montado na mao.
"""
import asyncio
import json
from io import BytesIO
from pathlib import Path

import pymupdf
from starlette.datastructures import Headers, UploadFile

from app import basic_extractor, main
from app.extractors.base import Palavra
from app.extractors.danfe_tabela import montar_tabela_itens

FIXTURE_PALAVRAS = Path(__file__).parent / "fixtures" / "danfe_palavras_anonimizado.json"


def _pdf_duas_paginas() -> bytes:
    doc = pymupdf.open()
    pagina1 = doc.new_page()
    pagina1.insert_text((50, 60), "FOLHA 1/2", fontsize=10)
    pagina1.insert_text((50, 80), "DANFE Documento Auxiliar", fontsize=10)
    pagina1.insert_text((50, 100), "Emitente: EMPRESA EXEMPLO LTDA", fontsize=10)
    pagina2 = doc.new_page()
    pagina2.insert_text((50, 60), "FOLHA 2/2", fontsize=10)
    pagina2.insert_text((50, 80), "Valor Total da Nota 10,00", fontsize=10)
    dados = doc.tobytes()
    doc.close()
    return dados


def _upload(conteudo: bytes) -> UploadFile:
    return UploadFile(
        file=BytesIO(conteudo),
        filename="teste.pdf",
        headers=Headers({"content-type": "application/pdf"}),
    )


def _rodar(corrotina):
    return asyncio.run(corrotina)


def _sem_numeracao(linhas_numeradas: list[str]) -> list[str]:
    return [linha.split(": ", 1)[1] if ": " in linha else "" for linha in linhas_numeradas]


def test_todos_os_caminhos_veem_as_mesmas_linhas():
    pdf = _pdf_duas_paginas()

    # Caminho de /extract-document: _ler_pdf -> montar_contexto.
    lido = _rodar(main._ler_pdf(_upload(pdf)))
    linhas_do_extrator = basic_extractor.montar_contexto(lido.texto, lido.paginas_palavras).linhas

    # /debug/extract-text
    debug_texto = _rodar(main.debug_extract_text(_upload(pdf)))

    # /debug/extractor-input
    debug_entrada = main.montar_debug_entrada_do_extrator(lido)
    linhas_exibidas = [
        linha
        for pagina in debug_entrada["paginas"]
        for linha in _sem_numeracao(pagina["linhas"])
    ]

    assert debug_texto["linhas"] == linhas_do_extrator
    assert debug_entrada["exibicao_fiel_ao_texto_do_extrator"] is True
    # A exibicao por pagina, sem as linhas em branco das pontas, e o mesmo
    # texto que o extrator recebe.
    assert [linha for linha in linhas_exibidas if linha] == [
        linha for linha in linhas_do_extrator if linha
    ]


def test_debug_entrada_separa_paginas_e_mostra_extrator_escolhido():
    lido = _rodar(main._ler_pdf(_upload(_pdf_duas_paginas())))
    debug = main.montar_debug_entrada_do_extrator(lido)

    assert debug["numero_paginas"] == 2
    assert [p["pagina"] for p in debug["paginas"]] == [1, 2]
    assert any("FOLHA 1/2" in linha for linha in debug["paginas"][0]["linhas"])
    assert any("FOLHA 2/2" in linha for linha in debug["paginas"][1]["linhas"])
    assert debug["extrator_escolhido"] == "DanfeExtractor"  # tem o marcador "danfe"
    assert set(debug["pontuacoes"]) == {"DanfeExtractor", "BoletoExtractor", "GenericExtractor"}
    assert "documento" in debug["resultado"]
    assert "confiancas" in debug["resultado"]


def test_extract_document_e_debug_extraem_o_mesmo_documento():
    """O resultado mostrado pelo debug e o mesmo que /extract-document
    devolve no modo basico (a IA fica de fora de proposito)."""
    lido = _rodar(main._ler_pdf(_upload(_pdf_duas_paginas())))
    debug = main.montar_debug_entrada_do_extrator(lido)
    direto = basic_extractor.extrair_com_metadados(lido.texto, lido.paginas_palavras)

    assert debug["resultado"]["documento"] == direto.documento.model_dump()
    assert debug["resultado"]["confiancas"] == direto.confiancas


def _paginas_palavras_fixture() -> list[list[Palavra]]:
    dados = json.loads(FIXTURE_PALAVRAS.read_text(encoding="utf-8"))
    return [
        [Palavra(texto=p["texto"], x0=p["x0"], x1=p["x1"], top=p["top"], bottom=p["bottom"]) for p in pagina]
        for pagina in dados["paginas"]
    ]


def test_diagnostico_da_tabela_nao_muda_o_resultado_e_explica_a_decisao():
    paginas = _paginas_palavras_fixture()
    sem_diagnostico = montar_tabela_itens(paginas)

    diagnostico: list[dict] = []
    com_diagnostico = montar_tabela_itens(paginas, diagnostico=diagnostico)

    assert com_diagnostico == sem_diagnostico

    assert len(diagnostico) == 1
    pagina = diagnostico[0]
    assert pagina["cabecalho_encontrado"] is True
    assert pagina["ancoras"][0][0] == "codigo"
    assert [linha["decisao"] for linha in pagina["linhas"]] == ["item"]
    assert pagina["linhas"][0]["colunas"]["codigo"] == "460-BCZS"


def test_diagnostico_mostra_a_linha_que_encerrou_a_tabela():
    paginas = _paginas_palavras_fixture()
    paginas[0].extend(
        [
            Palavra(texto="INFORMAÇÕES", x0=40, x1=100, top=491, bottom=498),
            Palavra(texto="COMPLEMENTARES", x0=110, x1=190, top=491, bottom=498),
        ]
    )
    diagnostico: list[dict] = []
    montar_tabela_itens(paginas, diagnostico=diagnostico)

    ultima = diagnostico[0]["linhas"][-1]
    assert ultima["decisao"].startswith("fim: marcador de secao")
    assert "INFORMAÇÕES COMPLEMENTARES" in ultima["texto"]


def test_diagnostico_quando_nao_acha_cabecalho():
    paginas = [[Palavra(texto="Nada", x0=10, x1=30, top=10, bottom=17)]]
    diagnostico: list[dict] = []
    tabela = montar_tabela_itens(paginas, diagnostico=diagnostico)

    assert tabela.itens == []
    assert diagnostico[0]["cabecalho_encontrado"] is False
    assert diagnostico[0]["minimo_exigido"] > diagnostico[0]["melhor_pontuacao_cabecalho"]
