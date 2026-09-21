"""Modelos de dados compartilhados pelo modo basico e pelo modo IA.

DocumentoExtraido e o mesmo contrato de dados nos dois modos de extracao
(regex/heuristica ou LLM) -- e isso que permite a tabela e a exportacao
Excel funcionarem igual independente de como o documento foi processado.
"""
from typing import Dict, List, Optional, Union

from pydantic import BaseModel, Field

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
    data_vencimento: Optional[str] = None
    emissor: Optional[str] = None
    destinatario: Optional[str] = None
    valor_total: Optional[Numero] = None
    itens: List[ItemDocumento] = []
    campos_adicionais: List[CampoAdicional] = []


class ExtractionResult(BaseModel):
    modo_extracao: str  # "basico" ou "ia"
    origem_texto: str = "digital"  # "digital" (pdfplumber) ou "ocr" (Tesseract)
    confiancas: Dict[str, str] = {}
    """Confianca por campo, so preenchida no modo basico: "alta" (rotulo
    explicito), "media" (heuristica posicional, ex: tabela por
    coordenada), "baixa" (fallback sem rotulo). Chave e o nome do atributo
    em DocumentoExtraido, ou, para itens de campos_adicionais, o texto
    exato do campo. Fica fora de DocumentoExtraido de proposito: esse
    schema tambem e o output_format do modo IA (Structured Outputs da
    Anthropic), e a LLM nao tem uma nocao natural de confianca por
    campo."""
    avisos: List[str] = []
    """Os avisos, UM POR ITEM (ex: "soma dos itens nao bate", "fallback IA ->
    basico"). E o que o frontend lista no banner. `aviso` (abaixo) e o mesmo
    conteudo juntado numa string so, mantido por compatibilidade."""
    aviso: Optional[str] = None  # informativo, nao e erro (ex: PDF escaneado, fallback IA -> basico)
    documento: DocumentoExtraido


class DocumentoParaExportar(BaseModel):
    """Um documento na requisicao de exportacao para Excel."""

    arquivo: Optional[str] = None
    """Nome do arquivo PDF de origem (vai pra coluna Arquivo do Resumo)."""
    resultado: ExtractionResult
    corrigidos: Dict[str, Union[float, str, None]] = {}
    """Campos que o usuario CORRIGIU na tela: {chave: valor ORIGINAL extraido}.
    A chave usa o mesmo espaco de nomes de `confiancas` (atributo do documento,
    ex. "emissor", ou o texto exato de um campo adicional, ex. "Chave de
    Acesso"). O backend so recebe o valor atual em `resultado`; sem isto ele
    nao teria como marcar o que foi corrigido nem mostrar o valor original."""


class ExportarExcelRequest(BaseModel):
    """Corpo de POST /export-excel. Uma LISTA de documentos (hoje a tela envia
    1; a estrutura ja comporta lote): cada documento vira uma linha no Resumo
    e as linhas dele nas outras abas, ligadas pela coluna ID."""

    documentos: List[DocumentoParaExportar] = Field(min_length=1)

