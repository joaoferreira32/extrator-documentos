"""Testa `extractors.selecionar_extrator()`: cada tipo de texto deve ser
roteado pro extrator certo."""
from pathlib import Path

from app import extractors
from app.extractors.base import ContextoExtracao
from app.extractors.boleto import BoletoExtractor
from app.extractors.danfe import DanfeExtractor
from app.extractors.generico import GenericExtractor

FIXTURE_BOLETO = Path(__file__).parent / "fixtures" / "boleto_real_anonimizado.txt"


def _contexto(texto: str) -> ContextoExtracao:
    return ContextoExtracao(texto=texto, linhas=texto.splitlines())


def test_boleto_real_roteia_para_boleto_extractor():
    texto = FIXTURE_BOLETO.read_text(encoding="utf-8")
    extrator = extractors.selecionar_extrator(_contexto(texto))
    assert isinstance(extrator, BoletoExtractor)


def test_marcadores_danfe_roteiam_para_danfe_extractor():
    texto = "DANFE - Documento Auxiliar da Nota Fiscal Eletronica\nChave de Acesso: 1234"
    extrator = extractors.selecionar_extrator(_contexto(texto))
    assert isinstance(extrator, DanfeExtractor)


def test_texto_generico_roteia_para_generic_extractor():
    texto = "Relatorio mensal de vendas\nTotal: R$ 1.000,00"
    extrator = extractors.selecionar_extrator(_contexto(texto))
    assert isinstance(extrator, GenericExtractor)


def test_texto_vazio_roteia_para_generic_extractor():
    extrator = extractors.selecionar_extrator(_contexto(""))
    assert isinstance(extrator, GenericExtractor)


def test_nota_fiscal_sem_marcador_danfe_nao_vira_danfe():
    """Uma nota fiscal de servico generica (sem os marcadores especificos
    de DANFE) nao deve ser tratada como DANFE -- so GenericExtractor."""
    texto = "Nota Fiscal de Servicos\nPrestador: Empresa Exemplo\nValor Total: R$ 500,00"
    extrator = extractors.selecionar_extrator(_contexto(texto))
    assert isinstance(extrator, GenericExtractor)
