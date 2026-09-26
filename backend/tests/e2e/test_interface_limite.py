"""Testes e2e (Chromium real) do limite de uso da demonstracao publica: quando o
backend responde 429 a tela mostra a mensagem dele, clara, e nao um erro
generico ("Erro 429" / "Falha ao extrair: ...").

Rode com: `pytest tests/e2e --e2e -v`. PDFs FICTICIOS (ver helpers.py).
"""
import json

import pytest

pytest.importorskip("playwright.sync_api", reason="instale requirements-dev.txt")

import helpers  # noqa: E402
from conftest import Tela  # noqa: E402

pytestmark = pytest.mark.e2e

MENSAGEM = (
    "Você atingiu o limite de 10 extrações por minuto da demonstração pública. "
    "Aguarde 42 segundos e envie o documento de novo."
)


def _responder_429(page, rota, detail):
    page.route(
        rota,
        lambda route: route.fulfill(
            status=429,
            headers={"Content-Type": "application/json", "Retry-After": "42"},
            body=json.dumps({"detail": detail}),
        ),
    )


def _texto_do_erro(page):
    page.wait_for_selector("#error-card:not([hidden])")
    return page.locator("#error-text").inner_text()


def test_429_na_extracao_mostra_a_mensagem_do_limite_sem_prefixo_de_falha(tela):
    tela.status_http_esperados.add(429)
    _responder_429(tela.page, "**/extract-document", MENSAGEM)
    tela.page.goto(tela.url)
    tela.page.set_input_files("#file-input", str(tela.pdfs["danfe_ok"]))
    tela.page.click("#btn-extrair")

    assert _texto_do_erro(tela.page) == MENSAGEM  # exatamente o que o backend disse
    assert "Falha" not in MENSAGEM and "429" not in _texto_do_erro(tela.page)
    assert tela.page.locator("#error-card").get_attribute("role") == "alert"  # leitor de tela anuncia
    assert tela.page.locator("#resultado").is_hidden()
    assert tela.page.locator("#btn-extrair").is_enabled()  # dá pra tentar de novo depois de esperar


def test_429_na_exportacao_mostra_a_mensagem_do_limite_e_nao_so_erro_429(tela):
    tela.status_http_esperados.add(429)
    tela.extrair("danfe_ok")
    mensagem = "Você atingiu o limite de 10 exportações por minuto da demonstração pública. Aguarde 42 segundos e baixe o Excel de novo."
    _responder_429(tela.page, "**/export-excel", mensagem)
    tela.page.click("#btn-excel")

    assert _texto_do_erro(tela.page) == mensagem
    assert tela.page.locator("#btn-excel").is_enabled()


def test_erro_comum_continua_com_o_prefixo_de_falha(tela):
    """O 429 e a excecao: um erro de verdade (ex: 500 sem `detail`) segue como falha."""
    tela.status_http_esperados.add(500)
    tela.page.route("**/extract-document", lambda route: route.fulfill(status=500, body="boom"))
    tela.page.goto(tela.url)
    tela.page.set_input_files("#file-input", str(tela.pdfs["danfe_ok"]))
    tela.page.click("#btn-extrair")
    assert _texto_do_erro(tela.page) == "Falha ao extrair: Erro 500"


@pytest.fixture(scope="module")
def servidor_limitado():
    processo, url = helpers.iniciar_servidor({"RATE_LIMIT_POR_MINUTO": "2"})
    yield url
    helpers.parar_servidor(processo)


def test_fluxo_real_a_terceira_extracao_no_minuto_e_barrada_com_mensagem_clara(navegador, servidor_limitado, pdfs, tmp_path):
    contexto = navegador.new_context(viewport={"width": 1280, "height": 900})
    page = contexto.new_page()
    erros = []
    page.on("console", lambda m: erros.append(m.text) if m.type == "error" and "429" not in m.text else None)
    page.on("pageerror", lambda e: erros.append(str(e)))
    try:
        tela = Tela(page, servidor_limitado, pdfs, tmp_path)
        tela.extrair("danfe_ok")  # 1a
        tela.extrair("boleto")  # 2a: esgota a cota
        page.goto(servidor_limitado)
        page.set_input_files("#file-input", str(pdfs["generico"]))
        page.click("#btn-extrair")  # 3a: barrada

        texto = _texto_do_erro(page)
        assert "limite de 2 extrações por minuto" in texto and "demonstração pública" in texto
        assert "Aguarde" in texto and "segundo" in texto
        assert not texto.startswith("Falha")
        assert page.locator("#resultado").is_hidden()
        assert not erros, erros
    finally:
        contexto.close()
