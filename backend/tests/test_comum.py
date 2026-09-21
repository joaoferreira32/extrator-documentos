"""Testes de helpers compartilhados de extractors/comum.py."""
import pytest

from app.extractors import comum


@pytest.mark.parametrize(
    "entrada, esperado",
    [
        ("15/4/2026", "15/04/2026"),  # DANFE real imprime o mes sem zero
        ("5/4/2026", "05/04/2026"),
        ("15/04/2026", "15/04/2026"),  # ja normalizada: nao muda (boleto)
        ("15-4-2026", "15-04-2026"),  # mantem o separador original
        ("nao e data", "nao e data"),  # nunca levanta, devolve como veio
    ],
)
def test_normalizar_data(entrada, esperado):
    assert comum.normalizar_data(entrada) == esperado


def test_extrair_data_devolve_normalizada():
    linhas = ["DATA DA EMISSÃO", "15/4/2026"]
    assert comum.extrair_data(linhas, ["Data da Emissão"], ["Data da Emissão"]) == "15/04/2026"


def test_extrair_valor_rotulo_por_padrao_so_olha_a_mesma_linha():
    """O default NAO le a linha seguinte: numa DANFE toda linha vizinha de
    um rotulo tem valores de OUTROS campos. (A DANFE le por posicao na
    grade de totais e nao liga aceitar_linha_seguinte.)"""
    texto = "Valor Documento (-) desconto\n1.000,00"
    assert comum.extrair_valor_rotulo(texto, ["Valor Documento"]) is None


def test_extrair_valor_rotulo_aceitar_linha_seguinte_e_opt_in():
    """Layout de cabecalho de tabela do boleto: rotulo numa linha, valor na
    proxima."""
    texto = "Valor Documento (-) desconto\n1.000,00"
    assert comum.extrair_valor_rotulo(texto, ["Valor Documento"], aceitar_linha_seguinte=True) == "1.000,00"


def test_aceitar_linha_seguinte_prefere_o_valor_da_mesma_linha():
    texto = "Valor Documento 2.000,00\n1.000,00"
    assert comum.extrair_valor_rotulo(texto, ["Valor Documento"], aceitar_linha_seguinte=True) == "2.000,00"


def test_aceitar_linha_seguinte_nao_olha_a_linha_anterior():
    texto = "1.000,00\nValor Documento (-) desconto"
    assert comum.extrair_valor_rotulo(texto, ["Valor Documento"], aceitar_linha_seguinte=True) is None
