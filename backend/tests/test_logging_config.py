"""app/logging_config.py: request-id por contextvar + filtro de log, e a
configuracao em si (idempotente, nao mexe no logger raiz).

Achado real verificado aqui: `logger_app.propagate = False` (pra nao
duplicar linha no root, que o uvicorn tambem configura) poderia ter
quebrado o `caplog` que test_extractors_boleto.py ja usava em
"app.extractors.comum" (um logger FILHO de "app"). Nao quebra: o pytest
(`caplog.at_level(nivel, logger=X)`) anexa o handler de captura DIRETO no
logger X pedido, nao depende de propagacao ate a raiz -- confirmado abaixo
com o mesmo padrao, num logger fabricado so pra este teste.
"""
import logging

from app import logging_config


def test_id_da_requisicao_comeca_e_muda_por_definir():
    logging_config.definir_id_da_requisicao("abc123")
    assert logging_config.id_da_requisicao_atual() == "abc123"
    logging_config.definir_id_da_requisicao("outro456")
    assert logging_config.id_da_requisicao_atual() == "outro456"


def test_filtro_injeta_o_request_id_atual_no_record():
    logging_config.definir_id_da_requisicao("req-teste-1")
    record = logging.LogRecord("qualquer", logging.INFO, __file__, 1, "msg", None, None)
    assert logging_config._FiltroRequestId().filter(record) is True
    assert record.request_id == "req-teste-1"


def test_configurar_logging_e_idempotente_nao_duplica_handler():
    logging_config.configurar_logging()  # garante que ja tem 1 handler antes de medir
    n_com_1_chamada = len(logging.getLogger("app").handlers)
    assert n_com_1_chamada >= 1
    logging_config.configurar_logging()
    logging_config.configurar_logging()
    assert len(logging.getLogger("app").handlers) == n_com_1_chamada


def test_configurar_logging_nao_propaga_pro_logger_raiz():
    """Motivo: uvicorn ja configura o logger raiz -- deixar "app" propagar
    tambem duplicaria toda linha (uma formatada por mim, outra pelo
    uvicorn)."""
    assert logging.getLogger("app").propagate is False


def test_propagate_false_do_logger_app_nao_quebra_caplog_de_um_logger_filho(caplog):
    """Prova empirica (nao so leitura de codigo) de que test_extractors_
    boleto.py (caplog em "app.extractors.comum", um FILHO de "app")
    continua funcionando: caplog.at_level(logger=X) anexa o handler direto
    em X, entao "app" ter propagate=False nao afeta nada."""
    logger_filho = logging.getLogger("app.um_filho_qualquer_pro_teste")
    with caplog.at_level(logging.INFO, logger="app.um_filho_qualquer_pro_teste"):
        logger_filho.info("mensagem de teste")
    assert any("mensagem de teste" in r.getMessage() for r in caplog.records)
