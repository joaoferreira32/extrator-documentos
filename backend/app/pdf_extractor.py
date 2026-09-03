"""Extracao de texto de PDFs: pdfplumber primeiro, OCR como fallback.

Fluxo: pdfplumber tenta extrair o texto (camada digital do PDF). Se vier
vazio ou curto demais para ser confiavel, tenta OCR (Tesseract via
pytesseract) sobre uma imagem renderizada de cada pagina (PyMuPDF, sem
depender de nenhum binario externo alem do proprio Tesseract).

OCR e uma dependencia de sistema externa (nao um pacote Python) -- se o
Tesseract nao estiver instalado na maquina, isso NUNCA deve quebrar o
app: `extrair_texto` continua funcionando, so devolve `ocr_disponivel =
False` para quem chama decidir a mensagem certa.
"""
import logging
import shutil
import sys
from dataclasses import dataclass
from io import BytesIO
from pathlib import Path

import pdfplumber

logger = logging.getLogger(__name__)

try:
    import pymupdf
    import pytesseract
    from PIL import Image

    _OCR_IMPORTADO = True
except ImportError:
    _OCR_IMPORTADO = False

# O instalador oficial do Tesseract para Windows nao adiciona o binario ao
# PATH automaticamente. Sem isso, pytesseract levanta TesseractNotFoundError
# mesmo com o Tesseract instalado -- confunde "nao instalado" com "instalado
# mas nao configurado". Se nao achar no PATH, tenta o local padrao de
# instalacao antes de desistir.
if _OCR_IMPORTADO and sys.platform == "win32" and shutil.which("tesseract") is None:
    _CAMINHO_PADRAO_WINDOWS = Path(r"C:\Program Files\Tesseract-OCR\tesseract.exe")
    if _CAMINHO_PADRAO_WINDOWS.is_file():
        pytesseract.pytesseract.tesseract_cmd = str(_CAMINHO_PADRAO_WINDOWS)

# Abaixo desse tamanho, o texto digital e tratado como "vazio demais para
# confiar" e o OCR e acionado -- PDFs escaneados as vezes tem uma ou duas
# palavras de texto real (ex: um carimbo) misturadas com a imagem.
TAMANHO_MINIMO_TEXTO_DIGITAL = 20

RESOLUCAO_OCR_DPI = 200
IDIOMA_OCR = "por"


@dataclass
class TextoExtraido:
    texto: str
    numero_paginas: int
    origem: str = "digital"  # "digital" ou "ocr"
    ocr_disponivel: bool = True
    """False quando o Tesseract (ou as bibliotecas de OCR) nao estao
    instalados no sistema -- diferente de "OCR tentou e nao achou nada"."""

    @property
    def parece_escaneado(self) -> bool:
        """True quando nao ha texto extraivel nem pelo pdfplumber nem
        (quando disponivel) pelo OCR."""
        return not self.texto.strip()


def extrair_texto(conteudo_pdf: bytes) -> TextoExtraido:
    """Le todas as paginas de um PDF e devolve o texto concatenado.

    Nao levanta excecao quando o PDF nao tem texto extraivel nem quando o
    OCR nao esta disponivel; quem chama deve checar `parece_escaneado` e
    `ocr_disponivel` e decidir a mensagem.
    """
    with pdfplumber.open(BytesIO(conteudo_pdf)) as pdf:
        paginas_texto = [pagina.extract_text() or "" for pagina in pdf.pages]
        texto_digital = "\n".join(paginas_texto).strip()
        numero_paginas = len(pdf.pages)

    if len(texto_digital) >= TAMANHO_MINIMO_TEXTO_DIGITAL:
        return TextoExtraido(texto=texto_digital, numero_paginas=numero_paginas, origem="digital")

    texto_ocr, ocr_disponivel = _tentar_ocr(conteudo_pdf)
    if len(texto_ocr) > len(texto_digital):
        return TextoExtraido(
            texto=texto_ocr,
            numero_paginas=numero_paginas,
            origem="ocr",
            ocr_disponivel=ocr_disponivel,
        )

    return TextoExtraido(
        texto=texto_digital,
        numero_paginas=numero_paginas,
        origem="digital",
        ocr_disponivel=ocr_disponivel,
    )


def _tentar_ocr(conteudo_pdf: bytes) -> tuple[str, bool]:
    """Renderiza cada pagina como imagem e roda Tesseract em cima.

    Devolve (texto, ocr_disponivel). Nunca levanta excecao: qualquer falha
    (bibliotecas ausentes, Tesseract nao instalado, pacote de idioma
    faltando, PDF corrompido) resulta em ("", False) e quem chama cai de
    volta para o texto digital (mesmo que vazio).
    """
    if not _OCR_IMPORTADO:
        return "", False

    try:
        paginas_texto = []
        with pymupdf.open(stream=conteudo_pdf, filetype="pdf") as documento:
            for pagina in documento:
                pixmap = pagina.get_pixmap(dpi=RESOLUCAO_OCR_DPI)
                imagem = Image.open(BytesIO(pixmap.tobytes("png")))
                paginas_texto.append(pytesseract.image_to_string(imagem, lang=IDIOMA_OCR))
        return "\n".join(paginas_texto).strip(), True
    except pytesseract.TesseractNotFoundError:
        logger.warning("Tesseract nao encontrado no sistema -- OCR indisponivel.")
        return "", False
    except Exception:
        logger.exception("Falha ao rodar OCR neste PDF.")
        return "", False
