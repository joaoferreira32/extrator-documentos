"""`avisos` (lista, um por item) convive com `aviso` (string juntada, por
compatibilidade). O frontend lista `avisos` no banner do resultado."""
import asyncio
from io import BytesIO
from types import SimpleNamespace

import pymupdf
from starlette.datastructures import Headers, UploadFile

from app import main
from app.extractors.base import ResultadoExtracao
from app.schemas import DocumentoExtraido


def _texto_lido(origem="digital"):
    return SimpleNamespace(origem=origem)


def test_resultado_basico_expoe_cada_aviso_separado_e_o_aviso_juntado():
    extracao = ResultadoExtracao(
        documento=DocumentoExtraido(),
        avisos=["A soma dos itens nao bate.", "Os totais nao fecham."],
    )
    resultado = main._resultado_basico(_texto_lido(), extracao)

    assert resultado.avisos == ["A soma dos itens nao bate.", "Os totais nao fecham."]
    assert resultado.aviso == "A soma dos itens nao bate. Os totais nao fecham."


def test_aviso_extra_do_main_vem_primeiro():
    extracao = ResultadoExtracao(documento=DocumentoExtraido(), avisos=["Soma nao bate."])
    resultado = main._resultado_basico(_texto_lido(), extracao, aviso_extra="IA indisponivel.")

    assert resultado.avisos == ["IA indisponivel.", "Soma nao bate."]
    assert resultado.aviso == "IA indisponivel. Soma nao bate."


def test_sem_avisos_a_lista_e_vazia_e_aviso_e_none():
    resultado = main._resultado_basico(
        _texto_lido(), ResultadoExtracao(documento=DocumentoExtraido())
    )
    assert resultado.avisos == []
    assert resultado.aviso is None


def test_pdf_escaneado_tem_um_aviso_na_lista_igual_ao_aviso():
    doc = pymupdf.open()
    doc.new_page()  # pagina em branco: sem texto extraivel
    upload = UploadFile(
        file=BytesIO(doc.tobytes()),
        filename="branco.pdf",
        headers=Headers({"content-type": "application/pdf"}),
    )
    resultado = asyncio.run(main.extract_document(upload))

    assert len(resultado.avisos) == 1
    assert resultado.aviso == resultado.avisos[0]
