"""Limites contra abuso de um servico publico (item 2 do diagnostico de
maturidade): tamanho de upload nos 3 endpoints de debug (que ate aqui nao
tinham limite nenhum, diferente de /extract-document) e teto de itens/
campos/avisos/documentos em /export-excel (que aceitava um payload
fabricado com 100 mil itens -- 12,9 MB de JSON -- sem rejeitar nada, e
levava 58s pra montar o .xlsx, bloqueando o servidor inteiro nesse tempo;
ver tests/test_concorrencia.py sobre o bloqueio em si).

httpx (TestClient) nao esta nas dependencias -- chama `main._ler_pdf` e os
endpoints de debug direto, como test_debug_endpoints.py ja faz.
"""
import asyncio
import os
import time
from io import BytesIO

import pymupdf
import pytest
from fastapi import HTTPException
from pydantic import ValidationError
from starlette.datastructures import Headers, UploadFile

from app import main
from app.schemas import (
    CampoAdicional,
    DocumentoExtraido,
    DocumentoParaExportar,
    ExportarExcelRequest,
    ExtractionResult,
    ItemDocumento,
    TETO_CAMPOS_ADICIONAIS,
    TETO_ITENS,
    TETO_ITENS_TOTAL_DO_LOTE,
)


def _upload(tamanho_bytes: int) -> UploadFile:
    # nao precisa ser um PDF de verdade -- o teto de tamanho e checado ANTES
    # de tentar ler o conteudo como PDF
    return UploadFile(
        file=BytesIO(b"%PDF-1.4\n" + b"0" * tamanho_bytes),
        filename="grande.pdf",
        headers=Headers({"content-type": "application/pdf"}),
    )


def _pdf_valido_do_tamanho(tamanho_bytes: int) -> UploadFile:
    """PDF estruturalmente valido (nao lixo) do tamanho pedido -- padding
    via anexo binario, igual ao usado no diagnostico de PDF grande. Existe
    porque um PDF de lixo desse tamanho faz o pdfplumber gastar ~20s
    tentando (sem sucesso) recuperar a estrutura antes de desistir -- o que
    testaria outra coisa (tempo de parsing de lixo), nao o teto de tamanho."""
    doc = pymupdf.open()
    doc.new_page().insert_text((50, 50), "PDF valido", fontsize=12)
    doc.embfile_add("padding.bin", os.urandom(max(tamanho_bytes - 3000, 0)))
    conteudo = doc.tobytes()
    doc.close()
    return UploadFile(
        file=BytesIO(conteudo), filename="grande.pdf", headers=Headers({"content-type": "application/pdf"})
    )


def _rodar(corrotina):
    return asyncio.run(corrotina)


# ---------- tamanho de upload (os 4 endpoints que leem PDF) ----------


def test_ler_pdf_rejeita_acima_de_20mb():
    with pytest.raises(HTTPException) as exc:
        _rodar(main._ler_pdf(_upload(21 * 1024 * 1024)))
    assert exc.value.status_code == 400
    assert "20 MB" in exc.value.detail


def test_ler_pdf_nao_rejeita_por_tamanho_ate_20mb():
    lido = _rodar(main._ler_pdf(_pdf_valido_do_tamanho(20 * 1024 * 1024 - 1000)))
    assert lido.numero_paginas == 1  # nao rejeitou -- leu o PDF normalmente


@pytest.mark.parametrize("nome_endpoint", ["debug_extract_text", "debug_extract_words", "debug_extractor_input"])
def test_endpoints_de_debug_tambem_rejeitam_acima_de_20mb(nome_endpoint):
    """Bug real: os 3 endpoints de debug chamavam _ler_pdf SEM o limite que
    /extract-document ja tinha -- um PDF de 55 MB era processado
    normalmente neles. Sao endpoints publicos (qualquer um acessa via
    /docs), entao isso era uma porta aberta sem limite nenhum."""
    endpoint = getattr(main, nome_endpoint)
    with pytest.raises(HTTPException) as exc:
        _rodar(endpoint(_upload(21 * 1024 * 1024)))
    assert exc.value.status_code == 400
    assert "20 MB" in exc.value.detail


# ---------- teto de itens/campos/avisos/documentos (/export-excel) ----------


def _documento(**kwargs):
    kwargs.setdefault("tipo_documento", "nota_fiscal")
    return DocumentoExtraido(**kwargs)


def test_tetos_tem_valores_absolutos_razoaveis():
    """Achado por checagem de mutacao real: os testes abaixo (ex:
    test_itens_alem_do_teto_e_rejeitado) testam TETO_ITENS+1 contra
    TETO_ITENS -- ou seja, sao relativos ao PROPRIO valor da constante.
    Mudar TETO_ITENS pra 1 milhao (esqueceram de ajustar, por exemplo)
    continua passando neles, porque o teste "acompanha" a mudanca. So um
    numero absoluto, independente da constante, pega isso."""
    assert TETO_ITENS <= 2000
    assert TETO_CAMPOS_ADICIONAIS <= 500
    assert TETO_ITENS_TOTAL_DO_LOTE <= 10_000


def test_itens_alem_do_teto_e_rejeitado():
    with pytest.raises(ValidationError):
        _documento(itens=[ItemDocumento(descricao="x") for _ in range(TETO_ITENS + 1)])
    _documento(itens=[ItemDocumento(descricao="x") for _ in range(TETO_ITENS)])  # no teto: aceita


def test_campos_adicionais_alem_do_teto_e_rejeitado():
    with pytest.raises(ValidationError):
        _documento(campos_adicionais=[CampoAdicional(campo="C", valor="1") for _ in range(TETO_CAMPOS_ADICIONAIS + 1)])
    _documento(campos_adicionais=[CampoAdicional(campo="C", valor="1") for _ in range(TETO_CAMPOS_ADICIONAIS)])


def test_avisos_alem_do_teto_e_rejeitado():
    with pytest.raises(ValidationError):
        ExtractionResult(modo_extracao="basico", avisos=["a"] * 51, documento=_documento())
    ExtractionResult(modo_extracao="basico", avisos=["a"] * 50, documento=_documento())


def test_lote_de_documentos_alem_do_teto_e_rejeitado():
    def doc():
        return DocumentoParaExportar(resultado=ExtractionResult(modo_extracao="basico", documento=_documento()))

    with pytest.raises(ValidationError):
        ExportarExcelRequest(documentos=[doc() for _ in range(101)])
    ExportarExcelRequest(documentos=[doc() for _ in range(100)])  # no teto: aceita


def test_teto_combinado_conta_o_total_de_itens_de_todos_os_documentos():
    """Sem o teto combinado, TETO_ITENS (por documento) e o max_length de
    `documentos` SOZINHOS ainda permitiriam 100 documentos x 1000 itens =
    100 mil itens no total -- o mesmo tamanho que levava 58s pra montar o
    .xlsx. Usa numeros fixos (nao TETO_ITENS) de proposito: um teste que
    usasse TETO_ITENS pra montar 100 documentos ficaria lento demais (ou
    trava) se alguem aumentasse essa constante -- o de baixo
    (test_tetos_tem_valores_absolutos_razoaveis) e quem garante que a
    constante em si nao vira algo absurdo."""
    def doc_com(n_itens):
        documento = _documento(itens=[ItemDocumento(descricao="x") for _ in range(n_itens)])
        return DocumentoParaExportar(resultado=ExtractionResult(modo_extracao="basico", documento=documento))

    # 5 documentos de 999 itens = 4995, dentro do teto de 5000
    ExportarExcelRequest(documentos=[doc_com(999) for _ in range(5)])
    # 5 documentos de 1000 itens = 5000, ainda dentro (no teto exato)
    ExportarExcelRequest(documentos=[doc_com(1000) for _ in range(5)])
    assert 5 * 1000 == TETO_ITENS_TOTAL_DO_LOTE
    # 6 documentos de 1000 = 6000, acima do teto
    with pytest.raises(ValidationError):
        ExportarExcelRequest(documentos=[doc_com(1000) for _ in range(6)])


def test_payload_fabricado_de_100_mil_itens_e_rejeitado_rapido_em_vez_de_travar():
    """Reproduz o ataque medido antes da correcao (100 mil itens fabricados,
    58s pra montar o .xlsx, bloqueando o servidor inteiro nesse tempo -- ver
    tests/test_concorrencia.py) e confirma que agora e rejeitado na
    VALIDACAO, em milissegundos, antes de qualquer tentativa de montar o
    Excel."""
    t0 = time.perf_counter()
    with pytest.raises(ValidationError):
        _documento(itens=[ItemDocumento(descricao=f"Item {i}") for i in range(100_000)])
    duracao = time.perf_counter() - t0
    assert duracao < 5.0, f"rejeitar deveria ser quase instantaneo, levou {duracao:.1f}s"
