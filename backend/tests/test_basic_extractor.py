"""Teste de compatibilidade: `basic_extractor.extrair()` continua com a
mesma assinatura publica (recebe texto, devolve DocumentoExtraido) depois
da refatoracao pra padrao Strategy (`app/extractors/`).

Os testes detalhados de cada tipo de documento vivem em
`test_extractors_boleto.py` / `test_extractors_danfe.py`, testando as
classes diretamente. Este arquivo so garante que a camada fina de
compatibilidade (`basic_extractor.py`) continua funcionando pra quem
chama do jeito antigo.
"""
from pathlib import Path

from app import basic_extractor

FIXTURE = Path(__file__).parent / "fixtures" / "boleto_real_anonimizado.txt"


def test_extrair_delega_para_o_extrator_certo():
    texto = FIXTURE.read_text(encoding="utf-8")
    doc = basic_extractor.extrair(texto)

    assert doc.tipo_documento == "boleto"
    assert doc.numero_documento == "1234567890"


def test_extrair_com_metadados_expoe_confianca():
    texto = FIXTURE.read_text(encoding="utf-8")
    resultado = basic_extractor.extrair_com_metadados(texto)

    assert resultado.documento.numero_documento == "1234567890"
    assert resultado.confiancas.get("numero_documento") == "alta"


def test_extrair_com_texto_vazio_nao_quebra():
    doc = basic_extractor.extrair("")
    assert doc.tipo_documento == "desconhecido"
