"""Extracao heuristica por regex, sem depender de LLM.

Modo "basico": sempre disponivel, nao precisa de chave de API. Bom para
capturar datas, valores monetarios, numero do documento e emissor/
destinatario quando o texto do PDF traz rotulos razoavelmente previsiveis.
Nao tenta reconstruir a tabela de itens -- texto corrido de PDF sem
estrutura nao da pra confiar nisso via regex. `itens` fica vazio; o modo IA
existe justamente para cobrir esse caso melhor.
"""
import re
from typing import Optional

from app.schemas import DocumentoExtraido

DATA_RE = re.compile(r"\b(\d{2}[/-]\d{2}[/-]\d{4})\b")
VALOR_RE = re.compile(r"R\$\s*([\d.]+,\d{2})")
NUM_DOCUMENTO_RE = re.compile(
    r"(?:N[º°o.]{1,3}|N[uú]mero)\s*[:\-]?\s*(\d[\w\-./]*)", re.IGNORECASE
)
EMISSOR_RE = re.compile(r"(?:Emitente|Fornecedor)\s*:?\s*(.+)", re.IGNORECASE)
DESTINATARIO_RE = re.compile(r"(?:Destinat[aá]rio|Cliente)\s*:?\s*(.+)", re.IGNORECASE)

PALAVRAS_CHAVE_TIPO = [
    ("nota fiscal", "nota_fiscal"),
    ("pedido de compra", "pedido_compra"),
    ("pedido", "pedido_compra"),
    ("relat", "relatorio"),  # cobre "relatorio"/"relatório"
]


def para_numero(bruto: str) -> float | str:
    """Converte um valor no formato BR ("1.234,56") para float; devolve a
    string original quando a conversao falha."""
    normalizado = bruto.strip().replace(".", "").replace(",", ".")
    try:
        return float(normalizado)
    except ValueError:
        return bruto.strip()


def _detectar_tipo_documento(texto_lower: str) -> str:
    for palavra, tipo in PALAVRAS_CHAVE_TIPO:
        if palavra in texto_lower:
            return tipo
    return "desconhecido"


def _primeiro_grupo(regex: re.Pattern, texto: str) -> Optional[str]:
    m = regex.search(texto)
    return m.group(1).strip() if m else None


def _valor_total(texto: str) -> Optional[str]:
    """Prioriza uma linha que mencione "total"; senao usa a ultima
    ocorrencia de valor monetario no texto."""
    for linha in texto.splitlines():
        if "total" in linha.lower():
            m = VALOR_RE.search(linha)
            if m:
                return m.group(1)
    ocorrencias = VALOR_RE.findall(texto)
    return ocorrencias[-1] if ocorrencias else None


def extrair(texto: str) -> DocumentoExtraido:
    texto_lower = texto.lower()

    valor_total_bruto = _valor_total(texto)

    return DocumentoExtraido(
        tipo_documento=_detectar_tipo_documento(texto_lower),
        numero_documento=_primeiro_grupo(NUM_DOCUMENTO_RE, texto),
        data_emissao=_primeiro_grupo(DATA_RE, texto),
        emissor=_primeiro_grupo(EMISSOR_RE, texto),
        destinatario=_primeiro_grupo(DESTINATARIO_RE, texto),
        valor_total=para_numero(valor_total_bruto) if valor_total_bruto else None,
        itens=[],
        campos_adicionais=[],
    )
