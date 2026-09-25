"""Item 5 do diagnostico de maturidade: mensagens de erro diferenciadas por
causa em `_ler_pdf`, em vez de um "corrompido" generico pra tudo.

Bugs reais, medidos gerando os PDFs de verdade (nao supondo o texto do
erro): um PDF protegido por senha e um PDF sem nenhuma pagina caiam os
dois na mesma mensagem generica de "corrompido"/"parece escaneado" --
diagnostico errado nos dois casos, o arquivo esta perfeito.

httpx (TestClient) nao esta nas dependencias -- chama `main._ler_pdf` e
`main.extract_document` direto, como test_debug_endpoints.py ja faz.
"""
import asyncio
from io import BytesIO

import pymupdf
import pytest
from fastapi import HTTPException
from starlette.datastructures import Headers, UploadFile

from app import main, pdf_extractor


def _upload(conteudo: bytes, nome="teste.pdf") -> UploadFile:
    return UploadFile(file=BytesIO(conteudo), filename=nome, headers=Headers({"content-type": "application/pdf"}))


def _rodar(corrotina):
    return asyncio.run(corrotina)


def _pdf_protegido_por_senha() -> bytes:
    doc = pymupdf.open()
    doc.new_page().insert_text((50, 50), "conteudo", fontsize=12)
    buf = BytesIO()
    doc.save(buf, encryption=pymupdf.PDF_ENCRYPT_AES_256, owner_pw="dono123", user_pw="senha123")
    doc.close()
    return buf.getvalue()


def _pdf_zero_paginas() -> bytes:
    """PyMuPDF recusa salvar um documento com 0 paginas ("cannot save with
    zero pages") -- escreve um PDF minimo, valido, com /Kids vazio na mao.
    E sintaticamente correto, so que sem nenhuma pagina (confirmado que o
    pdfplumber abre e devolve `len(pdf.pages) == 0` sem erro)."""
    objetos = [b"<< /Type /Catalog /Pages 2 0 R >>", b"<< /Type /Pages /Kids [] /Count 0 >>"]
    partes = [b"%PDF-1.4\n"]
    offsets = []
    for i, corpo in enumerate(objetos, start=1):
        offsets.append(sum(len(p) for p in partes))
        partes.append(f"{i} 0 obj\n".encode() + corpo + b"\nendobj\n")
    offset_xref = sum(len(p) for p in partes)
    n = len(objetos) + 1
    xref = [f"xref\n0 {n}\n".encode(), b"0000000000 65535 f \n"]
    xref += [f"{off:010d} 00000 n \n".encode() for off in offsets]
    partes += xref
    partes.append(f"trailer\n<< /Size {n} /Root 1 0 R >>\nstartxref\n{offset_xref}\n%%EOF".encode())
    return b"".join(partes)


# ---------- PDF protegido por senha ----------


def test_pdfplumber_realmente_levanta_pdfpasswordincorrect_embrulhada():
    """Confirma a premissa (nao supoe): o que o pdfplumber levanta pra um
    PDF com senha e `PdfminerException` com a `PDFPasswordIncorrect` real
    guardada em `.args[0]` -- e o que `extrair_texto` precisa reconhecer."""
    import pdfplumber
    from pdfminer.pdfdocument import PDFPasswordIncorrect
    from pdfplumber.utils.exceptions import PdfminerException

    with pytest.raises(PdfminerException) as exc:
        with pdfplumber.open(BytesIO(_pdf_protegido_por_senha())):
            pass
    assert exc.value.args and isinstance(exc.value.args[0], PDFPasswordIncorrect)


def test_extrair_texto_levanta_excecao_propria_pra_pdf_com_senha():
    with pytest.raises(pdf_extractor.PDFProtegidoPorSenha):
        pdf_extractor.extrair_texto(_pdf_protegido_por_senha())


def test_ler_pdf_devolve_mensagem_especifica_pra_pdf_com_senha_nao_generica():
    with pytest.raises(HTTPException) as exc:
        _rodar(main._ler_pdf(_upload(_pdf_protegido_por_senha())))
    assert exc.value.status_code == 400
    assert "senha" in exc.value.detail.lower()
    assert "corrompido" not in exc.value.detail.lower()


def test_pdf_genuinamente_corrompido_continua_com_a_mensagem_generica():
    """Nao pode virar falso positivo: lixo aleatorio (nao tem NADA a ver
    com senha) tem que continuar caindo no "corrompido" de sempre."""
    with pytest.raises(HTTPException) as exc:
        _rodar(main._ler_pdf(_upload(b"isto nao e um pdf de jeito nenhum")))
    assert "corrompido" in exc.value.detail.lower()


# ---------- PDF sem nenhuma pagina ----------


def test_pdf_zero_paginas_abre_normalmente_com_0_paginas():
    """Confirma a premissa: o pdfplumber abre esse PDF sem erro nenhum --
    o problema NAO e de leitura (por isso o bug nao aparecia em
    `_ler_pdf`, e sim na interpretacao que /extract-document fazia do
    resultado)."""
    resultado = pdf_extractor.extrair_texto(_pdf_zero_paginas())
    assert resultado.numero_paginas == 0
    assert resultado.parece_escaneado is True  # texto vazio -- mas por um motivo DIFERENTE de "e uma imagem"


def test_extract_document_avisa_pdf_sem_paginas_em_vez_de_dizer_que_e_imagem_escaneada():
    resultado = _rodar(main.extract_document(_upload(_pdf_zero_paginas())))
    assert "nenhuma pagina" in resultado.aviso.lower() or "nenhuma página" in resultado.aviso.lower()
    assert "escaneada" not in resultado.aviso.lower()
    assert "imagem" not in resultado.aviso.lower()


def test_extract_document_pdf_escaneado_de_verdade_continua_com_o_aviso_de_imagem():
    """Nao pode virar falso positivo: um PDF com 1+ pagina mas sem texto
    nenhum (imagem escaneada de verdade) tem que continuar com o aviso de
    sempre, nao o de "sem pagina"."""
    doc = pymupdf.open()
    doc.new_page()  # 1 pagina, sem nenhum texto (simula escaneado sem OCR)
    buf = BytesIO()
    doc.save(buf)
    doc.close()
    resultado = _rodar(main.extract_document(_upload(buf.getvalue())))
    assert "escaneada" in resultado.aviso.lower()
    assert "nenhuma pagina" not in resultado.aviso.lower() and "nenhuma página" not in resultado.aviso.lower()
