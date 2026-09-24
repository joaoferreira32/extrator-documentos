"""Modelos de dados compartilhados pelo modo basico e pelo modo IA.

DocumentoExtraido e o mesmo contrato de dados nos dois modos de extracao
(regex/heuristica ou LLM) -- e isso que permite a tabela e a exportacao
Excel funcionarem igual independente de como o documento foi processado.
"""
from typing import Dict, List, Optional, Union

from pydantic import BaseModel, Field, model_validator

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


# Tetos generosos (nenhum documento real chega perto): nenhum item de fisco
# tem centenas de itens de tabela nem dezenas de campos adicionais. Existem
# pra fechar um DoS real e medido -- /export-excel aceitava um DocumentoExtraido
# com 100 mil itens fabricados (12,9 MB de JSON) sem rejeitar nada, e montar
# o .xlsx com isso demorou 58s bloqueando o servidor inteiro (ver
# tests/test_concorrencia.py). `/extract-document` nunca produziria um
# documento assim (vem de PDF real), entao o teto nunca afeta uso legitimo.
TETO_ITENS = 1000
TETO_CAMPOS_ADICIONAIS = 100


class DocumentoExtraido(BaseModel):
    tipo_documento: str = "desconhecido"
    numero_documento: Optional[str] = None
    data_emissao: Optional[str] = None
    data_vencimento: Optional[str] = None
    emissor: Optional[str] = None
    destinatario: Optional[str] = None
    valor_total: Optional[Numero] = None
    itens: List[ItemDocumento] = Field(default_factory=list, max_length=TETO_ITENS)
    campos_adicionais: List[CampoAdicional] = Field(default_factory=list, max_length=TETO_CAMPOS_ADICIONAIS)


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
    avisos: List[str] = Field(default_factory=list, max_length=50)
    """Os avisos, UM POR ITEM (ex: "soma dos itens nao bate", "fallback IA ->
    basico"). E o que o frontend lista no banner. `aviso` (abaixo) e o mesmo
    conteudo juntado numa string so, mantido por compatibilidade. Teto de 50
    pela mesma razao de TETO_ITENS/TETO_CAMPOS_ADICIONAIS acima: nenhuma
    extracao real gera isso, so serve pra fechar o DoS via /export-excel."""
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


# Teto do TOTAL de itens somados de todos os documentos do lote. Sozinhos,
# TETO_ITENS (por documento) e o max_length de `documentos` abaixo NAO
# fecham o DoS: 100 documentos x 1000 itens cada ainda da 100 mil itens no
# total -- exatamente o tamanho medido levando 58s. Este teto e o que
# realmente limita o custo de UM lote inteiro.
TETO_ITENS_TOTAL_DO_LOTE = 5000


class ExportarExcelRequest(BaseModel):
    """Corpo de POST /export-excel. Uma LISTA de documentos (hoje a tela envia
    1; a estrutura ja comporta lote): cada documento vira uma linha no Resumo
    e as linhas dele nas outras abas, ligadas pela coluna ID."""

    # max_length: teto generoso pra uma futura interface de lote, ainda
    # fechando o DoS (ver TETO_ITENS acima) -- sem isso, o teto por
    # documento nao adianta contra "muitos documentos pequenos".
    documentos: List[DocumentoParaExportar] = Field(min_length=1, max_length=100)

    @model_validator(mode="after")
    def _limitar_total_de_itens_do_lote(self) -> "ExportarExcelRequest":
        total = sum(len(doc.resultado.documento.itens) for doc in self.documentos)
        if total > TETO_ITENS_TOTAL_DO_LOTE:
            raise ValueError(
                f"O lote tem {total} itens no total (limite: {TETO_ITENS_TOTAL_DO_LOTE}). "
                "Divida em lotes menores."
            )
        return self

