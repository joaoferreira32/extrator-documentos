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
from dataclasses import dataclass, field
from io import BytesIO
from pathlib import Path
from typing import Optional

import pdfplumber

from app.extractors.base import Palavra

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
    paginas_palavras: Optional[list[list[Palavra]]] = None
    """Palavras posicionadas (x0/x1/top/bottom) de cada pagina, na mesma
    passada do pdfplumber que extraiu o texto digital. None quando a
    origem e "ocr" -- imagem escaneada nao tem posicao de palavra real,
    so o texto que o Tesseract reconheceu. Extratores que reconstroem
    tabela por coordenada (ex: DANFE) dependem disso e devem degradar
    graciosamente quando for None."""
    paginas_texto: list[str] = field(default_factory=list)
    """Texto de cada pagina separado (so na origem "digital"; vazio no
    OCR). `texto` e a juncao disso com "\\n" -- guardado so pra o endpoint
    de debug mostrar onde cada pagina comeca, ja que essa fronteira some
    na juncao."""
    caracteres_girados_descartados: int = 0
    """Quantos caracteres de texto girado (canhoto, rotulos laterais) foram
    descartados na leitura -- ver `_sem_texto_girado`."""

    @property
    def parece_escaneado(self) -> bool:
        """True quando nao ha texto extraivel nem pelo pdfplumber nem
        (quando disponivel) pelo OCR."""
        return not self.texto.strip()


def extrair_texto(conteudo_pdf: bytes) -> TextoExtraido:
    """Le todas as paginas de um PDF e devolve o texto concatenado (e, se
    a origem for digital, as palavras posicionadas de cada pagina).

    Nao levanta excecao quando o PDF nao tem texto extraivel nem quando o
    OCR nao esta disponivel; quem chama deve checar `parece_escaneado` e
    `ocr_disponivel` e decidir a mensagem.
    """
    with pdfplumber.open(BytesIO(conteudo_pdf)) as pdf:
        paginas_texto = []
        paginas_palavras = []
        girados_descartados = 0
        for pagina in pdf.pages:
            pagina, descartados = _sem_texto_girado(pagina)
            girados_descartados += descartados
            paginas_texto.append(pagina.extract_text() or "")
            paginas_palavras.append(
                [
                    Palavra(
                        texto=palavra["text"],
                        x0=palavra["x0"],
                        x1=palavra["x1"],
                        top=palavra["top"],
                        bottom=palavra["bottom"],
                    )
                    for palavra in pagina.extract_words()
                ]
            )
        texto_digital = "\n".join(paginas_texto).strip()
        numero_paginas = len(pdf.pages)

    if len(texto_digital) >= TAMANHO_MINIMO_TEXTO_DIGITAL:
        return TextoExtraido(
            texto=texto_digital,
            numero_paginas=numero_paginas,
            origem="digital",
            paginas_palavras=paginas_palavras,
            paginas_texto=paginas_texto,
            caracteres_girados_descartados=girados_descartados,
        )

    texto_ocr, ocr_disponivel = _tentar_ocr(conteudo_pdf)
    if len(texto_ocr) > len(texto_digital):
        return TextoExtraido(
            texto=texto_ocr,
            numero_paginas=numero_paginas,
            origem="ocr",
            ocr_disponivel=ocr_disponivel,
            paginas_palavras=None,
        )

    return TextoExtraido(
        texto=texto_digital,
        numero_paginas=numero_paginas,
        origem="digital",
        ocr_disponivel=ocr_disponivel,
        paginas_palavras=paginas_palavras,
        paginas_texto=paginas_texto,
        caracteres_girados_descartados=girados_descartados,
    )


def _sem_texto_girado(pagina):
    """Devolve (pagina, quantos_caracteres_descartados) sem os caracteres
    girados (`upright == False`, i.e. texto a 90/270 graus).

    Bug real (DANFE): o canhoto e os rotulos laterais sao impressos girados,
    e o pdfplumber os extrai como linhas de letras soltas/invertidas
    ("SERODATUPMOC", "e-FN"...) ANTES do conteudo de verdade -- poluiam o
    texto de qualquer extrator, e a primeira linha do documento virava
    "FOLHA 1/". Descartar na leitura resolve na raiz, pra todos os
    extratores, em vez de cada um se defender do lixo.

    Salvaguarda: so filtra quando os girados sao MINORIA. Se a maior parte
    do texto da pagina esta "girada" (pagina inteira em paisagem via
    /Rotate, por exemplo), filtrar apagaria o documento -- nesse caso
    mantem tudo. Limitacao conhecida: texto de cabeca pra baixo (180 graus)
    continua `upright == True` e nao e descartado."""
    caracteres = pagina.chars
    girados = sum(1 for c in caracteres if not c.get("upright", True))
    if girados == 0 or girados * 2 >= len(caracteres):
        return pagina, 0
    filtrada = pagina.filter(lambda obj: obj.get("object_type") != "char" or obj.get("upright", True))
    return filtrada, girados


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
