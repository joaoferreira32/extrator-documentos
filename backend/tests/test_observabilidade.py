"""Item 4 do diagnostico de maturidade: log estruturado por extracao +
request-id. Duas partes:

- `main._logar_extracao`: chamada direta com `caplog` (mesmo padrao de
  test_extractors_boleto.py) -- confere o CONTEUDO da linha de log (JSON).
- `RequestIdMiddleware`: sobe o servidor de verdade (helpers.iniciar_
  servidor, que descarta stdout/stderr -- por isso aqui so confere o
  cabecalho da RESPOSTA, nao o log do lado do servidor) e confere que o
  cabecalho X-Request-ID existe, parece um id de verdade, e MUDA a cada
  requisicao (nao e um valor fixo/cacheado).
"""
import json
import logging
import sys
import urllib.request
from pathlib import Path

from app import main
from app.pdf_extractor import TextoExtraido
from app.schemas import DocumentoExtraido, ExtractionResult

sys.path.insert(0, str(Path(__file__).resolve().parent / "e2e"))
import helpers  # noqa: E402


def _texto_extraido(**kwargs):
    kwargs.setdefault("texto", "algum texto")
    kwargs.setdefault("numero_paginas", 3)
    kwargs.setdefault("origem", "digital")
    return TextoExtraido(**kwargs)


def _resultado(**kwargs):
    kwargs.setdefault("modo_extracao", "basico")
    kwargs.setdefault("origem_texto", "digital")
    kwargs.setdefault("documento", DocumentoExtraido(tipo_documento="nota_fiscal"))
    return ExtractionResult(**kwargs)


def _unica_linha_json(caplog):
    linhas = [r.getMessage() for r in caplog.records if r.name == "app.main"]
    assert len(linhas) == 1, f"esperava 1 linha de log, veio {len(linhas)}: {linhas}"
    return json.loads(linhas[0])


def test_logar_extracao_registra_duracao_paginas_modo_e_origem(caplog):
    resultado_texto = _texto_extraido(numero_paginas=7, origem="digital")
    resultado = _resultado(modo_extracao="ia", origem_texto="digital")
    with caplog.at_level(logging.INFO, logger="app.main"):
        main._logar_extracao(resultado_texto, resultado, duracao_ms=123.456, arquivo="nota.pdf")
    dados = _unica_linha_json(caplog)
    assert dados["arquivo"] == "nota.pdf"
    assert dados["duracao_ms"] == 123.5  # arredondado
    assert dados["numero_paginas"] == 7
    assert dados["origem_texto"] == "digital"
    assert dados["modo_extracao"] == "ia"


def test_logar_extracao_lista_campos_media_ou_baixa_ordenados(caplog):
    resultado = _resultado(confiancas={"emissor": "alta", "valor_total": "baixa", "data_emissao": "media"})
    with caplog.at_level(logging.INFO, logger="app.main"):
        main._logar_extracao(_texto_extraido(), resultado, duracao_ms=1.0, arquivo="a.pdf")
    dados = _unica_linha_json(caplog)
    assert dados["campos_media_ou_baixa"] == ["data_emissao", "valor_total"]  # alfabetica, sem "emissor" (alta)


def test_logar_extracao_sem_confiancas_nao_lista_nada_modo_ia(caplog):
    resultado = _resultado(modo_extracao="ia", confiancas={})
    with caplog.at_level(logging.INFO, logger="app.main"):
        main._logar_extracao(_texto_extraido(), resultado, duracao_ms=1.0, arquivo="a.pdf")
    dados = _unica_linha_json(caplog)
    assert "campos_media_ou_baixa" not in dados


def test_logar_extracao_lista_campos_obrigatorios_vazios_pelo_tipo(caplog):
    doc_boleto_incompleto = DocumentoExtraido(tipo_documento="boleto")  # nada preenchido
    resultado = _resultado(documento=doc_boleto_incompleto)
    with caplog.at_level(logging.INFO, logger="app.main"):
        main._logar_extracao(_texto_extraido(), resultado, duracao_ms=1.0, arquivo="a.pdf")
    dados = _unica_linha_json(caplog)
    # boleto exige emissor/destinatario/numero/data_emissao/valor_total/data_vencimento -- todos vazios
    assert set(dados["campos_obrigatorios_vazios"]) == {
        "emissor", "destinatario", "numero_documento", "data_emissao", "valor_total", "data_vencimento",
    }


def test_logar_extracao_documento_completo_nao_lista_campo_obrigatorio_vazio(caplog):
    doc_completo = DocumentoExtraido(
        tipo_documento="nota_fiscal", numero_documento="1", data_emissao="15/04/2026",
        emissor="E", destinatario="D", valor_total=10.0,
    )
    resultado = _resultado(documento=doc_completo)
    with caplog.at_level(logging.INFO, logger="app.main"):
        main._logar_extracao(_texto_extraido(), resultado, duracao_ms=1.0, arquivo="a.pdf")
    dados = _unica_linha_json(caplog)
    assert "campos_obrigatorios_vazios" not in dados


def test_logar_extracao_extra_e_mesclado_no_json(caplog):
    with caplog.at_level(logging.INFO, logger="app.main"):
        main._logar_extracao(_texto_extraido(), _resultado(), duracao_ms=1.0, arquivo="a.pdf",
                              extra={"extrator": "DanfeExtractor", "fallback_ia": "TimeoutError"})
    dados = _unica_linha_json(caplog)
    assert dados["extrator"] == "DanfeExtractor" and dados["fallback_ia"] == "TimeoutError"


# ---------- X-Request-ID (servidor de verdade) ----------


def _pedir(url: str):
    # HTTPResponse.getheader() e case-INSENSITIVE (via email.message.Message);
    # dict(resp.getheaders()) nao e -- guarda os nomes como o urllib devolve
    # (minusculo), entao um dict literal com "X-Request-ID" nunca bateria.
    with urllib.request.urlopen(url, timeout=10) as resp:
        return resp.getheader("X-Request-ID")


def test_x_request_id_aparece_no_cabecalho_e_muda_a_cada_requisicao():
    processo, url = helpers.iniciar_servidor()
    try:
        rid_1 = _pedir(url + "health")
        rid_2 = _pedir(url + "health")
        assert rid_1 and rid_2, "cabecalho X-Request-ID ausente"
        assert len(rid_1) == 8 and all(c in "0123456789abcdef" for c in rid_1)
        assert rid_1 != rid_2, "cada requisicao deveria ter um id proprio, nao um valor fixo"
    finally:
        helpers.parar_servidor(processo)
