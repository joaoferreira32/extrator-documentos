"""Registro dos extratores disponiveis (padrao Strategy).

Para adicionar um tipo de documento novo: crie `extractors/novo_tipo.py`
implementando `ExtratorDocumento` (ver `base.py`) e adicione a classe na
lista `_EXTRATORES` abaixo. Nao e preciso tocar em nenhum extrator
existente nem em `selecionar_extrator`.
"""
from app.extractors.base import ContextoExtracao, ExtratorDocumento, Palavra, ResultadoExtracao
from app.extractors.boleto import BoletoExtractor
from app.extractors.danfe import DanfeExtractor
from app.extractors.generico import GenericExtractor

_EXTRATORES: list[ExtratorDocumento] = [
    DanfeExtractor(),
    BoletoExtractor(),
    GenericExtractor(),  # sempre por ultimo: pontuacao de fallback, so vence se nada mais reconhecer
]


def selecionar_extrator(contexto: ContextoExtracao) -> ExtratorDocumento:
    return max(_EXTRATORES, key=lambda extrator: extrator.pontuacao_deteccao(contexto))


def pontuacoes(contexto: ContextoExtracao) -> dict[str, float]:
    """Pontuacao de deteccao de cada extrator registrado (so pra debug)."""
    return {type(extrator).__name__: extrator.pontuacao_deteccao(contexto) for extrator in _EXTRATORES}


__all__ = [
    "ContextoExtracao",
    "ExtratorDocumento",
    "Palavra",
    "ResultadoExtracao",
    "pontuacoes",
    "selecionar_extrator",
]
