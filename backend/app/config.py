"""Configuracao do app. A chave de API e opcional -- sem ela o sistema
continua funcionando normalmente em modo basico (ver main.py)."""
import os

from dotenv import load_dotenv

load_dotenv()

ANTHROPIC_API_KEY = os.environ.get("ANTHROPIC_API_KEY", "").strip() or None


def ia_disponivel() -> bool:
    return ANTHROPIC_API_KEY is not None
