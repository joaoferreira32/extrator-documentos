"""Configuracao do app. A chave de API e opcional -- sem ela o sistema
continua funcionando normalmente em modo basico (ver main.py)."""
import os

from dotenv import load_dotenv

load_dotenv()

ANTHROPIC_API_KEY = os.environ.get("ANTHROPIC_API_KEY", "").strip() or None

# Requisicoes por minuto, por IP e por grupo de endpoint (extracao/exportacao).
# Criterio do valor: ver o docstring de app/rate_limit.py.
LIMITE_REQUISICOES_POR_MINUTO_PADRAO = 10


def ia_disponivel() -> bool:
    return ANTHROPIC_API_KEY is not None


def limite_requisicoes_por_minuto() -> int:
    """RATE_LIMIT_POR_MINUTO: 0 desliga o limite. Valor vazio ou invalido volta
    ao padrao -- uma variavel mal escrita nao pode derrubar nem escancarar o app."""
    bruto = os.environ.get("RATE_LIMIT_POR_MINUTO", "").strip()
    if not bruto:
        return LIMITE_REQUISICOES_POR_MINUTO_PADRAO
    try:
        return max(0, int(bruto))
    except ValueError:
        return LIMITE_REQUISICOES_POR_MINUTO_PADRAO


_VERDADEIRO = {"1", "true", "yes", "sim"}
_FALSO = {"0", "false", "no", "nao", "não"}


def confiar_x_forwarded_for() -> bool:
    """Se o IP do cliente deve vir do X-Forwarded-For (ver app/rate_limit.py).

    Com CONFIAR_X_FORWARDED_FOR explicita, vale ela. Sem ela, o padrao depende de
    onde o app roda: LIGADO no Render (que sempre define RENDER=true e poe todo
    trafego atras de um proxy -- desligado ali, todos os visitantes dividiriam a
    cota do IP do proxy) e DESLIGADO no resto (servidor exposto direto: o
    cabecalho e forjavel). Nao depende do render.yaml: um servico criado pelo
    painel, sem Blueprint, ignora os envVars do arquivo."""
    bruto = os.environ.get("CONFIAR_X_FORWARDED_FOR", "").strip().lower()
    if bruto in _VERDADEIRO:
        return True
    if bruto in _FALSO:
        return False
    return os.environ.get("RENDER", "").strip().lower() == "true"
