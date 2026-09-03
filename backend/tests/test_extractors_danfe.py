"""Testes do DanfeExtractor e da reconstrucao de tabela por coordenadas.

Fixtures construidas a partir de dumps reais (/debug/extract-words e
/debug/extract-text) de uma DANFE real (nota fiscal da Dell) -- a ORDEM
das linhas do texto e exatamente a que o pdfplumber produziu (nao uma
linearizacao "rotulo: valor" inventada). Dados do destinatario (pessoa
fisica de terceiro) sao ficticios ("FULANO DE TAL"/CPF zerado) porque a
propria pessoa nao deu consentimento -- o usuario anonimizou antes de
repassar. Emitente/CNPJ/chave de acesso/tabela de itens sao reais e
publicos. Ver CLAUDE.md para o historico de por que essa arquitetura
exige material real em vez de suposicao de layout: a primeira versao
desta fixture era uma linearizacao escrita a mao (nao uma extracao real),
e nao reproduzia bugs reais de ordem de leitura (emissor saindo "DANFE",
destinatario/data_emissao/valor_total vazios) que so apareceram testando
o PDF real no navegador.
"""
import json
from pathlib import Path

from app.extractors.base import ContextoExtracao, Palavra
from app.extractors.danfe import DanfeExtractor
from app.extractors.danfe_tabela import _eh_texto_vertical, montar_tabela_itens

FIXTURE_TEXTO = Path(__file__).parent / "fixtures" / "danfe_real_anonimizado.txt"
FIXTURE_PALAVRAS = Path(__file__).parent / "fixtures" / "danfe_palavras_anonimizado.json"


def _carregar_paginas_palavras() -> list[list[Palavra]]:
    dados = json.loads(FIXTURE_PALAVRAS.read_text(encoding="utf-8"))
    return [
        [Palavra(texto=p["texto"], x0=p["x0"], x1=p["x1"], top=p["top"], bottom=p["bottom"]) for p in pagina]
        for pagina in dados["paginas"]
    ]


def _contexto() -> ContextoExtracao:
    texto = FIXTURE_TEXTO.read_text(encoding="utf-8")
    return ContextoExtracao(
        texto=texto, linhas=texto.splitlines(), paginas_palavras=_carregar_paginas_palavras()
    )


def test_deteccao_danfe():
    contexto = _contexto()
    assert DanfeExtractor().pontuacao_deteccao(contexto) > 0.5


def test_extrair_emissor_pula_titulo_da_danfe():
    """Regressao do bug real: buscar 'Emitente' como rotulo generico
    pegava a legenda "Identificacao do Emitente" e a "proxima linha" caia
    numa caixa vizinha, devolvendo "DANFE" (o titulo do documento) como
    emissor. _extrair_emissor usa a primeira linha do documento, mas pula
    explicitamente marcadores de titulo conhecidos (MARCADORES_DETECCAO)
    -- este teste cobre esse pulo mesmo que, na fixture real completa, o
    nome ja seja a primeira linha (sem titulo antes)."""
    texto = "\n".join(
        [
            "DANFE",
            "Documento Auxiliar da Nota Fiscal Eletronica",
            "EMPRESA EXEMPLO LTDA",
            "CNPJ: 11.222.333/0001-44",
        ]
    )
    contexto = ContextoExtracao(texto=texto, linhas=texto.splitlines(), paginas_palavras=None)
    emissor = DanfeExtractor()._extrair_emissor(contexto)
    assert emissor is not None
    assert "EMPRESA EXEMPLO LTDA" in emissor
    assert "DANFE" not in emissor


def test_danfe_completa():
    resultado = DanfeExtractor().extrair(_contexto())
    doc = resultado.documento

    print(doc.model_dump_json(indent=2))
    print("confiancas:", resultado.confiancas)
    print("avisos:", resultado.avisos)

    assert doc.tipo_documento == "nota_fiscal"

    campos = {c.campo: c.valor for c in doc.campos_adicionais}
    # Chave de acesso real (44 digitos, DV valido -- conferido a mao com o
    # mesmo algoritmo de test_chave_acesso.py).
    assert campos.get("Chave de Acesso") == "35260472381189001001550010000123451123456786"
    assert resultado.confiancas.get("Chave de Acesso") == "alta"

    # emissor: o nome aparece antes de qualquer rotulo no documento real
    # (ver DanfeExtractor._extrair_emissor) -- buscar por rotulo generico
    # "Emitente" e o que causava o bug real (emissor saindo "DANFE").
    assert doc.emissor is not None and "DELL COMPUTADORES DO BRASIL LTDA" in doc.emissor
    assert "72.381.189/0010-01" in doc.emissor

    # destinatario: rotulo real e "NOME/RAZÃO SOCIAL", nao "Destinatário"
    # (que o GenericExtractor buscava e nunca encontrava).
    assert doc.destinatario is not None and "FULANO DE TAL" in doc.destinatario
    assert resultado.confiancas.get("destinatario") == "alta"

    # data_emissao: rotulo real e "DATA DA EMISSÃO" ("da", nao "de") e o
    # formato tem mes sem zero a esquerda ("15/4/2026") -- os dois motivos
    # reais de o campo vir vazio antes do fix.
    assert doc.data_emissao == "15/4/2026"
    assert resultado.confiancas.get("data_emissao") == "alta"

    # valor_total: no documento real o valor vem ANTES do rotulo "VALOR
    # TOTAL DA NOTA" (legenda por baixo da caixa), nao "rotulo: valor".
    assert doc.valor_total == 229.0
    assert resultado.confiancas.get("valor_total") == "alta"

    # --- Tabela de itens por coordenadas ---
    assert len(doc.itens) == 1
    item = doc.itens[0]
    assert item.descricao == "Mochila Dell Gaming Backpack 17, GM1720PM"
    assert item.quantidade == 1.0
    assert item.valor_unitario == 215.03
    assert item.valor_total == 215.03
    assert resultado.confiancas.get("itens") == "media"

    assert campos.get("CFOP") == "5102"

    # Soma dos itens (215,03) bate com "Valor Total dos Produtos" (215,03)
    # na fixture -- nao deve gerar aviso de divergencia.
    assert resultado.avisos == []


def test_soma_divergente_gera_aviso():
    # No documento real o valor vem na linha ANTES do rotulo (ver
    # comentario de valor_total em test_danfe_completa) -- troca o valor
    # que precede "VALOR TOTAL DOS PRODUTOS", nao o rotulo em si.
    texto = FIXTURE_TEXTO.read_text(encoding="utf-8").replace(
        "215,03\nVALOR TOTAL DOS PRODUTOS", "999,99\nVALOR TOTAL DOS PRODUTOS"
    )
    contexto = ContextoExtracao(
        texto=texto, linhas=texto.splitlines(), paginas_palavras=_carregar_paginas_palavras()
    )
    resultado = DanfeExtractor().extrair(contexto)
    assert any("nao bate" in aviso for aviso in resultado.avisos)


def test_montar_tabela_itens_descarta_texto_vertical():
    paginas = _carregar_paginas_palavras()
    tabela = montar_tabela_itens(paginas)

    assert len(tabela.itens) == 1
    # Nenhum dos rotulos verticais (SOTUDORP, LISARB, ...) deve aparecer
    # em nenhum campo do item.
    item_texto = tabela.itens[0].descricao
    for ruido in ["SOTUDORP", "LISARB", "ADARITER", "SERODATUPMOC", "OIRÁTANITSED"]:
        assert ruido not in item_texto


def test_montar_tabela_itens_para_em_informacoes_complementares():
    """Bug real: com o boleto de teste em navegador, o bloco 'INFORMACOES
    COMPLEMENTARES' que vem depois da tabela de itens (secao padronizada
    nacionalmente, nao especifica de um emissor) caia por coincidencia de
    x na coluna CODIGO e virava uma sequencia de itens fantasmas. As
    coordenadas exatas do bloco na DANFE real do usuario nao foram
    fornecidas (so o teste em navegador reportou top aproximado
    491-517) -- aqui adicionamos linhas representativas dessa secao por
    cima da fixture real (mesma pagina, coordenadas plausiveis, top maior
    que o item real) so pra validar que o corte por conteudo funciona,
    independente da posicao vertical exata."""
    paginas = _carregar_paginas_palavras()
    pagina = paginas[0]

    def w(texto, x0, top):
        return Palavra(texto=texto, x0=x0, x1=x0 + len(texto) * 6, top=top, bottom=top + 7)

    pagina.extend(
        [
            w("INFORMAÇÕES", 40, 491),
            w("COMPLEMENTARES", 110, 491),
            w("Reservado", 40, 505),
            w("ao", 100, 505),
            w("Fisco", 115, 505),
        ]
    )

    tabela = montar_tabela_itens(paginas)
    assert len(tabela.itens) == 1
    assert "Reservado" not in tabela.itens[0].descricao


def test_montar_tabela_itens_separa_valores_colados():
    """O token real "13,9718,00" (VALOR I.P.I. + ALIQUOTA ICMS colados)
    nao pode vazar pro valor_total do item nem corromper o valor_unitario
    -- eles vem de colunas antes do token colado, entao devem sair
    intactos."""
    tabela = montar_tabela_itens(_carregar_paginas_palavras())
    item = tabela.itens[0]
    assert item.valor_unitario == 215.03 or item.valor_unitario == 215.0300  # mesma coisa
    assert item.valor_total == 215.03


def test_eh_texto_vertical():
    vertical = Palavra(texto="SOTUDORP", x0=20, x1=28, top=150, bottom=210)
    horizontal = Palavra(texto="460-BCZS", x0=85, x1=100, top=422.1, bottom=429.4)
    assert _eh_texto_vertical(vertical) is True
    assert _eh_texto_vertical(horizontal) is False


def test_eh_texto_vertical_fora_da_faixa_x_nao_e_descartado():
    """A regra so vale na faixa de x0 15-85 (margem lateral) -- uma
    palavra alta e estreita fora dessa faixa nao deveria ser confundida
    com um rotulo rotacionado."""
    fora_da_faixa = Palavra(texto="ALGO", x0=300, x1=308, top=150, bottom=210)
    assert _eh_texto_vertical(fora_da_faixa) is False
