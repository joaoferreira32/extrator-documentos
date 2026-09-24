import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent.parent))


def pytest_addoption(parser):
    # Definida aqui (e nao em tests/e2e/conftest.py) porque pytest_addoption so
    # vale em conftests carregados no inicio da sessao.
    parser.addoption(
        "--e2e",
        action="store_true",
        default=False,
        help="roda os testes e2e (Playwright + Chromium; ver requirements-dev.txt)",
    )
    parser.addoption(
        "--ooxml-sdk",
        action="store_true",
        default=False,
        help="valida o .xlsx com o validador oficial da Microsoft (Open XML SDK; so Windows; baixa o pacote na 1a vez)",
    )
