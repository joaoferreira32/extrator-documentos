"""Teste de regressao com o texto bruto de um boleto real (anonimizado).

Este fixture existe porque cenarios sinteticos escritos a mao nao
reproduziam o layout real de um boleto bancario -- eles passavam mas o
extrator continuava errando no documento de verdade (o mesmo rotulo
aparece varias vezes em contextos diferentes: cabecalho do recibo,
tabela-resumo no rodape, campos de verdade). Ver CLAUDE.md.
"""
import logging
from pathlib import Path

from app import basic_extractor

FIXTURE = Path(__file__).parent / "fixtures" / "boleto_real_anonimizado.txt"


def test_boleto_real_anonimizado(caplog):
    texto = FIXTURE.read_text(encoding="utf-8")

    with caplog.at_level(logging.DEBUG, logger="app.basic_extractor"):
        doc = basic_extractor.extrair(texto)

    print("\n--- log de extracao (rotulo/linha escolhido por campo) ---")
    for registro in caplog.records:
        print(registro.getMessage())

    print("\n--- resultado ---")
    print(doc.model_dump_json(indent=2))

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
