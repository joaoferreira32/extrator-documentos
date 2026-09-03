"""Testes do fallback digital -> OCR em pdf_extractor.

A orquestracao (quando acionar o OCR, o que fazer se ele nao estiver
disponivel) e testada com mocks, sem depender de o Tesseract estar
instalado na maquina que roda os testes. Ha tambem um teste "de verdade"
com OCR real, que so roda se o Tesseract e o pacote de idioma pt estiverem
de fato disponiveis -- caso contrario e pulado, nunca falha por isso.
"""
import pymupdf
import pytest

from app import pdf_extractor


def _pdf_com_texto(texto: str) -> bytes:
    doc = pymupdf.open()
    pagina = doc.new_page()
    pagina.insert_text((50, 72), texto)
    return doc.tobytes()


def _pdf_em_branco() -> bytes:
    """Uma pagina sem nenhum texto -- simula um documento escaneado."""
    doc = pymupdf.open()
    doc.new_page()
    return doc.tobytes()


def _ocr_real_disponivel() -> bool:
    if not pdf_extractor._OCR_IMPORTADO:
        return False
    texto, disponivel = pdf_extractor._tentar_ocr(_pdf_em_branco())
    return disponivel


def test_texto_digital_suficiente_nao_aciona_ocr(monkeypatch):
    chamado = []
    monkeypatch.setattr(
        pdf_extractor, "_tentar_ocr", lambda conteudo: chamado.append(1) or ("nunca", True)
    )

    resultado = pdf_extractor.extrair_texto(_pdf_com_texto("Um texto digital bem mais longo que vinte caracteres."))

    assert resultado.origem == "digital"
    assert not chamado, "OCR nao deveria ser chamado quando ja ha texto digital suficiente"


def test_ocr_indisponivel_nao_quebra(monkeypatch):
    monkeypatch.setattr(pdf_extractor, "_OCR_IMPORTADO", False)

    resultado = pdf_extractor.extrair_texto(_pdf_em_branco())

    assert resultado.parece_escaneado
    assert resultado.ocr_disponivel is False
    assert resultado.origem == "digital"


def test_ocr_usado_quando_texto_digital_insuficiente(monkeypatch):
    monkeypatch.setattr(
        pdf_extractor, "_tentar_ocr", lambda conteudo: ("Texto vindo do OCR simulado", True)
    )

    resultado = pdf_extractor.extrair_texto(_pdf_em_branco())

    assert resultado.origem == "ocr"
    assert resultado.texto == "Texto vindo do OCR simulado"
    assert not resultado.parece_escaneado


def test_falha_no_ocr_cai_para_texto_digital_sem_quebrar(monkeypatch):
    def _tentar_ocr_com_erro(conteudo):
        return "", False

    monkeypatch.setattr(pdf_extractor, "_tentar_ocr", _tentar_ocr_com_erro)

    resultado = pdf_extractor.extrair_texto(_pdf_em_branco())

    assert resultado.parece_escaneado
    assert resultado.ocr_disponivel is False


@pytest.mark.skipif(not _ocr_real_disponivel(), reason="Tesseract ou pacote de idioma 'por' nao disponivel")
def test_ocr_real_extrai_texto_de_pdf_so_imagem(tmp_path):
    from PIL import Image, ImageDraw

    imagem = Image.new("RGB", (600, 100), "white")
    ImageDraw.Draw(imagem).text((10, 30), "TESTE DE OCR REAL", fill="black")
    caminho_png = tmp_path / "pagina.png"
    imagem.save(caminho_png)

    doc = pymupdf.open()
    pagina = doc.new_page(width=600, height=100)
    pagina.insert_image(pagina.rect, filename=str(caminho_png))
    conteudo_pdf = doc.tobytes()
    doc.close()

    resultado = pdf_extractor.extrair_texto(conteudo_pdf)

    assert resultado.origem == "ocr"
    assert "TESTE" in resultado.texto.upper()
