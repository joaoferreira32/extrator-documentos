"""Interface comum dos extratores por tipo de documento (padrao Strategy).

Cada tipo de documento (boleto, DANFE, generico) implementa
`ExtratorDocumento`. `extractors/__init__.py` mantem o registro dos
extratores disponiveis e escolhe qual usar via `pontuacao_deteccao()` --
adicionar um tipo novo nao exige tocar nos extratores existentes, so criar
o arquivo novo e registra-lo na lista.
"""
from abc import ABC, abstractmethod
from dataclasses import dataclass, field
from typing import Optional

from app.schemas import DocumentoExtraido


@dataclass
class Palavra:
    """Uma palavra extraida de um PDF digital, com sua posicao na pagina
    (unidades: pontos PDF, origem no canto superior esquerdo -- mesmo
    sistema de `pdfplumber.page.extract_words()`)."""

    texto: str
    x0: float
    x1: float
    top: float
    bottom: float


@dataclass
class ContextoExtracao:
    """Tudo que um extrator pode precisar para processar um documento."""

    texto: str
    linhas: list[str]
    paginas_palavras: Optional[list[list[Palavra]]] = None
    """Palavras posicionadas por pagina, na ordem em que aparecem no PDF.
    None quando o texto veio de OCR (sem posicao confiavel) -- extratores
    que dependem de coordenadas (ex: tabela de itens da DANFE) devem
    degradar graciosamente quando isso for None, nunca quebrar."""


@dataclass
class ResultadoExtracao:
    documento: DocumentoExtraido
    confiancas: dict[str, str] = field(default_factory=dict)
    """Confianca por campo do DocumentoExtraido: "alta" (rotulo explicito),
    "media" (heuristica posicional) ou "baixa" (fallback sem rotulo).
    Campos ausentes daqui nao tem confianca atribuida (ex: nada foi
    encontrado, ou o campo nao se aplica)."""
    avisos: list[str] = field(default_factory=list)
    """Avisos especificos deste extrator (ex: soma dos itens nao bate com
    o valor total). Quem monta a resposta final concatena com outros
    avisos (OCR, fallback de IA) no campo `aviso` unico da API."""


class ExtratorDocumento(ABC):
    """Estrategia de extracao para um tipo de documento especifico."""

    tipo: str

    @abstractmethod
    def pontuacao_deteccao(self, contexto: ContextoExtracao) -> float:
        """0.0 a 1.0: o quanto este extrator acredita que o texto e do seu
        tipo. `selecionar_extrator()` usa a maior pontuacao entre todos os
        extratores registrados."""

    @abstractmethod
    def extrair(self, contexto: ContextoExtracao) -> ResultadoExtracao:
        """Extrai os campos do documento. Nunca deve levantar excecao por
        falta de dado -- campos nao encontrados ficam None/vazios."""
