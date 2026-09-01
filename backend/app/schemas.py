"""Modelos de dados compartilhados pelo modo basico e pelo modo IA.

DocumentoExtraido e o mesmo contrato de dados nos dois modos de extracao
(regex/heuristica ou LLM) -- e isso que permite a tabela e a exportacao
Excel funcionarem igual independente de como o documento foi processado.
"""
from typing import List, Optional, Union

from pydantic import BaseModel

Numero = Union[float, str]
"""float quando a conversao numerica funciona, string original como fallback
quando o valor extraido nao pode ser convertido com confianca."""


class ItemDocumento(BaseModel):
    descricao: str
    quantidade: Optional[Numero] = None
    valor_unitario: Optional[Numero] = None
    valor_total: Optional[Numero] = None


class CampoAdicional(BaseModel):
    campo: str
    valor: str


class DocumentoExtraido(BaseModel):
    tipo_documento: str = "desconhecido"
    numero_documento: Optional[str] = None
    data_emissao: Optional[str] = None
    emissor: Optional[str] = None
    destinatario: Optional[str] = None
    valor_total: Optional[Numero] = None
    itens: List[ItemDocumento] = []
    campos_adicionais: List[CampoAdicional] = []


class ExtractionResult(BaseModel):
    modo_extracao: str  # "basico" ou "ia"
    aviso: Optional[str] = None  # informativo, nao e erro (ex: PDF escaneado, fallback IA -> basico)
    documento: DocumentoExtraido
