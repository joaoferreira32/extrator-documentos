"""Regenera docs/screenshot.png (usado no README) com dados FICTICIOS.

Uso, a partir de backend/:   python tests/e2e/gerar_screenshot.py

Sempre parte dos PDFs sinteticos de helpers.py (nunca de documento real):
screenshot de tela com documento real ja foi um risco de vazamento de dados
pessoais neste projeto. Mostra a DANFE "com problemas" porque exibe quase tudo
da interface: banner de avisos, chips de confianca (alta/media), o resumo e
um campo obrigatorio vazio ja aberto pra preencher.
"""
import sys
import tempfile
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent))

import helpers  # noqa: E402
from playwright.sync_api import sync_playwright  # noqa: E402

DESTINO = helpers.BACKEND.parent / "docs" / "screenshot.png"


def main():
    with tempfile.TemporaryDirectory() as pasta:
        pdfs = helpers.gerar_pdfs(Path(pasta))
        conteudo = pdfs["danfe_problemas"].read_bytes()
        processo, url = helpers.iniciar_servidor()
        try:
            with sync_playwright() as playwright:
                navegador = playwright.chromium.launch()
                contexto = navegador.new_context(viewport={"width": 1280, "height": 900}, device_scale_factor=2)
                pagina = contexto.new_page()
                pagina.goto(url)
                # nome de arquivo amigavel na tela (o PDF de teste tem nome tecnico)
                pagina.set_input_files(
                    "#file-input",
                    files=[{"name": "danfe_exemplo.pdf", "mimeType": "application/pdf", "buffer": conteudo}],
                )
                pagina.click("#btn-extrair")
                pagina.wait_for_selector("#resultado:not([hidden])")
                pagina.screenshot(path=str(DESTINO), full_page=True)
                navegador.close()
        finally:
            helpers.parar_servidor(processo)
    print(f"gerado: {DESTINO}")


if __name__ == "__main__":
    main()
