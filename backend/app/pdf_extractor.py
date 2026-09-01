"""Extracao de texto de PDFs com pdfplumber."""
from dataclasses import dataclass
from io import BytesIO

import pdfplumber


@dataclass
class TextoExtraido:
    texto: str
    numero_paginas: int

    @property
    def parece_escaneado(self) -> bool:
        """True quando o PDF nao tem texto extraivel (provavelmente e uma
        imagem escaneada, sem camada de texto)."""
        return not self.texto.strip()


def extrair_texto(conteudo_pdf: bytes) -> TextoExtraido:
    """Le todas as paginas de um PDF e devolve o texto concatenado.

    Nao levanta excecao quando o PDF nao tem texto extraivel (PDF escaneado);
    quem chama deve checar `parece_escaneado` e decidir o que fazer.

    TODO: adicionar OCR (ex: pytesseract) como fallback para PDFs
    escaneados. Fora do escopo deste MVP.
    """
    with pdfplumber.open(BytesIO(conteudo_pdf)) as pdf:
        paginas_texto = [pagina.extract_text() or "" for pagina in pdf.pages]
        return TextoExtraido(
            texto="\n".join(paginas_texto).strip(),
            numero_paginas=len(pdf.pages),
        )
