"""Log estruturado simples, sem infraestrutura pesada (sem Sentry, sem ELK,
sem servico pago) -- so `logging` da biblioteca padrao com timestamp e um
request-id propagado automaticamente.

Achado no diagnostico de maturidade do projeto: nao existia log nenhum alem
de 2 chamadas pontuais em `pdf_extractor.py` (falha de OCR), sem timestamp,
sem formato, sem jeito de juntar varias linhas da MESMA requisicao. Isso
sozinho ja e o suficiente pra responder "quanto tempo levou", "qual
extrator foi escolhido" e "quais campos mais ficam vazios ou com confianca
baixa" sem precisar de outra ferramenta -- ver `main._logar_extracao`.

Request-id: gerado uma vez por requisicao (`main.RequestIdMiddleware`) e
guardado num `contextvars.ContextVar` -- por isso QUALQUER logger do
projeto (`app.pdf_extractor`, `app.extractors.comum`, ...) ja devolve o id
certo sem precisar receber-lo como parametro em toda chamada. O mesmo id
volta no cabecalho `X-Request-ID` da resposta, pra correlacionar um erro
que o usuario relatar com a linha exata do log do servidor.
"""
import contextvars
import logging
import os

_request_id_var: contextvars.ContextVar[str] = contextvars.ContextVar("request_id", default="-")


def id_da_requisicao_atual() -> str:
    return _request_id_var.get()


def definir_id_da_requisicao(valor: str) -> None:
    _request_id_var.set(valor)


class _FiltroRequestId(logging.Filter):
    """Injeta `record.request_id` em toda linha, mesmo vinda de um logger
    diferente do que criou o contextvar (ex: app.pdf_extractor dentro de
    uma requisicao iniciada em main.py)."""

    def filter(self, record: logging.LogRecord) -> bool:
        record.request_id = _request_id_var.get()
        return True


def configurar_logging() -> None:
    """Chamado uma vez, na inicializacao do app (main.py). So configura o
    logger "app" (o prefixo de todo `logging.getLogger(__name__)` deste
    projeto) -- nao mexe no logger raiz nem nos loggers do proprio uvicorn
    (uvicorn.error/uvicorn.access), que ja tem a propria configuracao.
    `propagate = False` evita duplicar a linha no root."""
    logger_app = logging.getLogger("app")
    if logger_app.handlers:
        return  # ja configurado (ex: uvicorn --reload reimporta o modulo)

    handler = logging.StreamHandler()
    handler.addFilter(_FiltroRequestId())
    handler.setFormatter(logging.Formatter(
        "%(asctime)s %(levelname)s [%(request_id)s] %(name)s: %(message)s"
    ))
    logger_app.addHandler(handler)
    logger_app.setLevel(os.environ.get("LOG_LEVEL", "INFO").upper())
    logger_app.propagate = False
