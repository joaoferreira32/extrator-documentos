"""Fixtures dos testes e2e (Playwright). Opt-in: `pytest tests/e2e --e2e`.

Sem `--e2e` tudo aqui e pulado ANTES de subir servidor ou navegador, entao o
`pytest tests` normal continua rapido e nao exige Playwright.
"""
import json

import pytest

import helpers


@pytest.fixture(scope="session", autouse=True)
def _exige_flag_e2e(request):
    if not request.config.getoption("--e2e"):
        pytest.skip("e2e: rode `pytest tests/e2e --e2e` (requer requirements-dev.txt e `playwright install chromium`)")


@pytest.fixture(scope="session")
def servidor():
    processo, url = helpers.iniciar_servidor()
    yield url
    helpers.parar_servidor(processo)


@pytest.fixture(scope="session")
def pdfs(tmp_path_factory):
    return helpers.gerar_pdfs(tmp_path_factory.mktemp("pdfs_ficticios"))


@pytest.fixture(scope="session")
def navegador():
    sync_api = pytest.importorskip("playwright.sync_api", reason="instale requirements-dev.txt")
    playwright = sync_api.sync_playwright().start()
    try:
        browser = playwright.chromium.launch()
    except Exception as erro:  # Chromium nao baixado
        playwright.stop()
        pytest.skip(f"Chromium indisponivel ({str(erro).splitlines()[0]}); rode `playwright install chromium`")
    yield browser
    browser.close()
    playwright.stop()


class Tela:
    """Atalhos sobre a pagina: extrair um PDF, ler o resumo, baixar o Excel."""

    def __init__(self, page, url, pdfs, tmp_path):
        self.page = page
        self.url = url
        self.pdfs = pdfs
        self.tmp_path = tmp_path

    def extrair(self, nome_pdf):
        """Envia o PDF e devolve o JSON que o backend respondeu."""
        self.page.goto(self.url)
        self.page.set_input_files("#file-input", str(self.pdfs[nome_pdf]))
        with self.page.expect_response("**/extract-document") as resposta:
            self.page.click("#btn-extrair")
        self.page.wait_for_selector("#resultado:not([hidden])")
        return resposta.value.json()

    def resumo(self):
        return self.page.locator("#resumo").inner_text().replace("\n", " ")

    def json_bruto(self):
        # o <pre> fica dentro de um <details> fechado: inner_text() voltaria vazio
        return json.loads(self.page.locator("#json-bruto").text_content())

    def baixar_excel(self):
        """Clica em Baixar Excel e devolve {campo: valor} da aba Resumo."""
        import openpyxl

        with self.page.expect_download() as download:
            self.page.click("#btn-excel")
        destino = self.tmp_path / "exportado.xlsx"
        download.value.save_as(destino)
        planilha = openpyxl.load_workbook(destino)
        return {linha[0]: linha[1] for linha in planilha["Resumo"].iter_rows(min_row=2, values_only=True)}

    def campo(self, id_campo):
        return self.page.locator(f'[data-campo="{id_campo}"]')


@pytest.fixture
def tela(navegador, servidor, pdfs, tmp_path):
    contexto = navegador.new_context(viewport={"width": 1280, "height": 900}, accept_downloads=True)
    page = contexto.new_page()
    erros = []
    page.on("console", lambda m: erros.append(m.text) if m.type == "error" else None)
    page.on("pageerror", lambda e: erros.append(str(e)))
    yield Tela(page, servidor, pdfs, tmp_path)
    contexto.close()
    # Erro de JS no meio de um handler pode deixar a tela "quase certa" (ja
    # aconteceu: o bug do blur reentrante) -- entao todo teste vigia o console.
    assert not erros, f"erros de console/JS durante o teste: {erros}"
