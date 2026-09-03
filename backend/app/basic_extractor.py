"""Ponto de entrada do modo "basico" (sem LLM): delega para o extrator
certo via padrao Strategy (`app/extractors/`).

Esta era uma extracao monolitica por regex; virou uma camada fina que
monta o `ContextoExtracao` e delega pro `ExtratorDocumento` escolhido por
`extractors.selecionar_extrator()`. Ver `app/extractors/base.py` e
`app/extractors/__init__.py` para a arquitetura, e CLAUDE.md para o
historico de por que cada heuristica existe.

`extrair()` mantem a assinatura publica de sempre (devolve so
`DocumentoExtraido`) para nao quebrar quem ja chama assim. `main.py` usa
`extrair_com_metadados()` pra tambem ter acesso a confianca por campo e a
avisos especificos do extrator (ex: soma da tabela de itens nao bate).
"""
from app import extractors
from app.extractors.base import ContextoExtracao, ResultadoExtracao
from app.schemas import DocumentoExtraido


def extrair_com_metadados(texto: str, paginas_palavras=None) -> ResultadoExtracao:
    contexto = ContextoExtracao(
        texto=texto, linhas=texto.splitlines(), paginas_palavras=paginas_palavras
    )
    extrator = extractors.selecionar_extrator(contexto)
    return extrator.extrair(contexto)


def extrair(texto: str) -> DocumentoExtraido:
    return extrair_com_metadados(texto).documento
