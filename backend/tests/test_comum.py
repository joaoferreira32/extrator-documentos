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
