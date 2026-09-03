"""Teste de regressao com o texto bruto de um boleto real (anonimizado).

Este fixture existe porque cenarios sinteticos escritos a mao nao
reproduziam o layout real de um boleto bancario -- eles passavam mas o
extrator continuava errando no documento de verdade (o mesmo rotulo
aparece varias vezes em contextos diferentes: cabecalho do recibo,
tabela-resumo no rodape, campos de verdade). Ver CLAUDE.md.

Movido de test_basic_extractor.py quando a extracao de boleto virou
`BoletoExtractor` (padrao Strategy) -- mesmo comportamento, testado
diretamente na classe em vez de pela funcao de compatibilidade.
"""
import logging
from pathlib import Path

from app.extractors.base import ContextoExtracao
from app.extractors.boleto import BoletoExtractor
from app.extractors import comum

FIXTURE = Path(__file__).parent / "fixtures" / "boleto_real_anonimizado.txt"


def _extrair(texto: str):
    contexto = ContextoExtracao(texto=texto, linhas=texto.splitlines())
    return BoletoExtractor().extrair(contexto)


def test_boleto_real_anonimizado(caplog):
    texto = FIXTURE.read_text(encoding="utf-8")

    with caplog.at_level(logging.DEBUG, logger="app.extractors.comum"):
        resultado = _extrair(texto)
    doc = resultado.documento

    print("\n--- log de extracao (rotulo/linha escolhido por campo) ---")
    for registro in caplog.records:
        print(registro.getMessage())

    print("\n--- resultado ---")
    print(doc.model_dump_json(indent=2))
    print("confiancas:", resultado.confiancas)

    campos = {c.campo: c.valor for c in doc.campos_adicionais}

    assert doc.destinatario is not None and "Maria" in doc.destinatario, (
        f"destinatario deveria conter o nome do pagador, veio {doc.destinatario!r}"
    )
    assert doc.numero_documento == "1234567890", (
        f"numero_documento errado: {doc.numero_documento!r}"
    )
    assert campos.get("Nosso Número") == "10200000001-9", (
        f"nosso_numero errado: {campos.get('Nosso Número')!r}"
    )
    # Campos vindos de rotulo explicito devem ter confianca alta.
    assert resultado.confiancas.get("numero_documento") == "alta"
    assert resultado.confiancas.get("destinatario") == "alta"


def test_deteccao_boleto():
    texto = FIXTURE.read_text(encoding="utf-8")
    contexto = ContextoExtracao(texto=texto, linhas=texto.splitlines())
    assert BoletoExtractor().pontuacao_deteccao(contexto) > 0.5


def test_corrige_confusao_ocr_em_cnpj():
    """OCR troca com frequencia O<->0, I<->1, S<->5 em campos numericos."""
    texto = "Beneficiário: Empresa Teste Ltda\nCNPJ: 12.34S.678/OOO1-9I"
    doc = _extrair(texto).documento
    assert doc.emissor == "Empresa Teste Ltda (CNPJ 12.345.678/0001-91)"


def test_corrige_confusao_ocr_em_cpf():
    texto = "Pagador Fulano de Tal CPF: I11.222.333-44"
    doc = _extrair(texto).documento
    assert doc.destinatario == "Fulano de Tal (CPF 111.222.333-44)"


def test_corrige_confusao_ocr_em_valor():
    from app.extractors.boleto import ROTULOS_VALOR_A_PAGAR

    valor = comum.extrair_valor_rotulo("Valor a Pagar = R$ 9OO,OO", ROTULOS_VALOR_A_PAGAR)
    assert valor == "900,00"


def test_numero_documento_curto_sem_rotulo_e_rejeitado():
    """Regressao do bug original: numero solto curto (ex: '02') nao pode
    virar numero_documento sem vir de um rotulo confiavel."""
    texto = "BANCO EXEMPLO S.A.\nAceite N  02  Especie Doc DM\nCedente: Empresa ABC Ltda"
    doc = _extrair(texto).documento
    assert doc.numero_documento != "02"
