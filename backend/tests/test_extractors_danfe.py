"""Testes do DanfeExtractor e da reconstrucao de tabela por coordenadas.

Fixtures construidas a partir de um dump real via /debug/extract-words de
uma DANFE real (nota fiscal da Dell) -- coordenadas do cabecalho da
tabela e da linha de item sao reais, confirmadas pelo usuario. Dados do
destinatario (pessoa fisica de terceiro) sao ficticios porque nao foram
fornecidos: o dump real so cobriu emitente + tabela de itens. Ver
CLAUDE.md para o historico de por que essa arquitetura exige material
real em vez de suposicao de layout.
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


def test_danfe_completa():
    resultado = DanfeExtractor().extrair(_contexto())
    doc = resultado.documento

    print(doc.model_dump_json(indent=2))
    print("confiancas:", resultado.confiancas)
    print("avisos:", resultado.avisos)

    assert doc.tipo_documento == "nota_fiscal"

    campos = {c.campo: c.valor for c in doc.campos_adicionais}
    assert campos.get("Chave de Acesso") == "1" * 43 + "2"
    assert resultado.confiancas.get("Chave de Acesso") == "alta"

    assert doc.emissor is not None and "DELL COMPUTADORES DO BRASIL LTDA" in doc.emissor
    assert "72.381.189/0010-01" in doc.emissor
    assert doc.destinatario is not None and "Cliente Exemplo Ltda" in doc.destinatario
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
    texto = FIXTURE_TEXTO.read_text(encoding="utf-8").replace(
        "Valor Total dos Produtos 215,03", "Valor Total dos Produtos 999,99"
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
