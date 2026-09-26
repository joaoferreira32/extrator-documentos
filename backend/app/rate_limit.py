"""Limite simples de requisicoes por IP, pensado para a demonstracao publica.

Escopo deliberadamente pequeno: janela deslizante EM MEMORIA, por processo.
Serve pra impedir que uma pessoa (ou um script) sature a instancia gratuita
mandando dezenas de PDFs por minuto; NAO e defesa contra ataque distribuido
(isso e papel do Cloudflare/Render na frente) e zera quando o processo
reinicia ou hiberna. Com mais de uma instancia cada uma contaria por conta
propria -- fora de escopo enquanto a demo roda numa so.

Valor padrao (config.LIMITE_REQUISICOES_POR_MINUTO_PADRAO = 10 por minuto, por
IP, POR GRUPO de endpoint) -- criterio: um humano usando a tela faz no maximo
~3-4 extracoes por minuto (enviar, conferir, exportar); 10 da folga de ~2,5x
pra quem grava um video enviando varios documentos em sequencia, e ainda limita
um script a 600 por hora por IP. Grupos separados ("extracao" e "exportacao")
pra exportar nao consumir a cota de upload.

Por que middleware ASGI e nao `Depends`: uma dependencia so roda DEPOIS de o
FastAPI ler e interpretar o corpo multipart; barrar aqui poupa o trabalho de
receber um upload de ate 20 MB que vai ser recusado de qualquer jeito.
"""
import ipaddress
import json
import logging
import math
import time
from collections import OrderedDict, deque
from dataclasses import dataclass
from typing import Callable, Optional

from starlette.responses import JSONResponse

logger = logging.getLogger(__name__)

# Todo endpoint POST do app esta aqui (tests/test_rate_limit.py confere): quem
# acrescentar um endpoint de upload novo e esquecer de listar aqui e avisado
# pelo teste, em vez de descobrir com a instancia saturada.
GRUPOS_POR_ROTA = {
    "/extract-document": "extracao",
    "/debug/extract-text": "extracao",
    "/debug/extract-words": "extracao",
    "/debug/extractor-input": "extracao",
    "/export-excel": "exportacao",
}

JANELA_SEGUNDOS = 60.0
MAX_CHAVES = 10_000  # teto de IPs rastreados: o dicionario nao pode crescer sem limite


@dataclass(frozen=True)
class Decisao:
    permitido: bool
    restantes: int
    espera_s: int  # segundos ate poder tentar de novo; 0 quando permitido


class LimitadorPorJanela:
    """Janela deslizante: guarda o instante das ultimas requisicoes ACEITAS de
    cada chave. Requisicao recusada NAO e registrada -- insistir nao estende o
    bloqueio, e o "aguarde N segundos" da resposta fica exato.

    Sincrono e sem `await`: dentro do event loop do asyncio cada chamada roda
    inteira antes de outra comecar, entao nao precisa de lock.

    Memoria limitada a `max_chaves` chaves (LRU). Sob um ataque que varia o IP,
    os IPs mais antigos saem primeiro: quem perde o registro ganha uma janela
    nova (mais permissivo, nunca mais restritivo) e o processo nao estoura."""

    def __init__(self, limite: int, janela_s: float, agora: Callable[[], float] = time.monotonic, max_chaves: int = MAX_CHAVES):
        if limite < 1 or max_chaves < 1:
            raise ValueError("limite e max_chaves precisam ser >= 1")
        self._limite = limite
        self._janela = janela_s
        self._agora = agora
        self._max_chaves = max_chaves
        self._por_chave: "OrderedDict[str, deque[float]]" = OrderedDict()

    def __len__(self) -> int:
        return len(self._por_chave)

    def tentar(self, chave: str) -> Decisao:
        agora = self._agora()
        instantes = self._por_chave.get(chave)
        if instantes is None:
            instantes = deque()
            self._por_chave[chave] = instantes
            if len(self._por_chave) > self._max_chaves:
                self._por_chave.popitem(last=False)
        else:
            self._por_chave.move_to_end(chave)

        corte = agora - self._janela
        while instantes and instantes[0] <= corte:
            instantes.popleft()

        if len(instantes) >= self._limite:
            espera = math.ceil(instantes[0] + self._janela - agora)
            return Decisao(False, 0, max(espera, 1))
        instantes.append(agora)
        return Decisao(True, self._limite - len(instantes), 0)


def _normalizar_ip(texto: str) -> Optional[str]:
    """Chave canonica de um IP, ou None se nao for um IP valido. Aceita porta
    ("1.2.3.4:5678", "[::1]:443") e IPv4 embrulhado em IPv6 ("::ffff:1.2.3.4").
    IPv6 vira o prefixo /64: uma pessoa com IPv6 recebe uma rede inteira, e
    contar endereco por endereco deixaria o limite trivial de contornar."""
    t = texto.strip()
    if t.startswith("["):
        fim = t.find("]")
        if fim == -1:
            return None
        t = t[1:fim]
    elif t.count(":") == 1:
        t = t.split(":")[0]
    try:
        ip = ipaddress.ip_address(t)
    except ValueError:
        return None
    if isinstance(ip, ipaddress.IPv6Address):
        if ip.ipv4_mapped:
            return str(ip.ipv4_mapped)
        return str(ipaddress.ip_network(f"{ip}/64", strict=False))
    return str(ip)


def ip_do_cliente(scope: dict, confiar_x_forwarded_for: bool) -> str:
    """IP usado como chave do limite.

    Sem `confiar_x_forwarded_for` (padrao): o IP da conexao, e o cabecalho e
    ignorado -- quem fala direto com o servidor nao consegue forjar outro IP.
    Com ele (ligado no Render, onde toda requisicao chega por um proxy): o
    PRIMEIRO IP de X-Forwarded-For, que e o cliente quando o proxy o preenche.
    LIMITACAO CONHECIDA: o proxy do Render so ANEXA ao cabecalho que o cliente
    mandou, entao quem envia um X-Forwarded-For forjado escolhe a propria chave
    e contorna o limite. Aceitavel pra um limite de demo (ver docstring do
    modulo); quem precisar de mais que isso deve limitar na borda."""
    cliente = scope.get("client")
    conexao = cliente[0] if cliente else "desconhecido"
    if not confiar_x_forwarded_for:
        return conexao

    valores = [v.decode("latin-1") for nome, v in scope.get("headers", []) if nome == b"x-forwarded-for"]
    if valores:
        primeiro = ",".join(valores).split(",")[0]
        ip = _normalizar_ip(primeiro)
        if ip:
            return ip
    return _normalizar_ip(conexao) or conexao


def _mensagem(grupo: str, limite: int, espera_s: int) -> str:
    segundos = "1 segundo" if espera_s == 1 else f"{espera_s} segundos"
    if grupo == "extracao":
        return (
            f"Você atingiu o limite de {limite} extrações por minuto da demonstração pública. "
            f"Aguarde {segundos} e envie o documento de novo."
        )
    return (
        f"Você atingiu o limite de {limite} exportações por minuto da demonstração pública. "
        f"Aguarde {segundos} e baixe o Excel de novo."
    )


class LimiteDeRequisicoesMiddleware:
    """Middleware ASGI puro. `limite_por_minuto` 0 desliga (nada e limitado)."""

    def __init__(
        self,
        app,
        limite_por_minuto: int,
        confiar_x_forwarded_for: bool = False,
        agora: Callable[[], float] = time.monotonic,
    ):
        self.app = app
        self.limite = limite_por_minuto
        self.confiar_x_forwarded_for = confiar_x_forwarded_for
        self.limitadores = (
            {grupo: LimitadorPorJanela(limite_por_minuto, JANELA_SEGUNDOS, agora) for grupo in set(GRUPOS_POR_ROTA.values())}
            if limite_por_minuto > 0
            else {}
        )

    async def __call__(self, scope, receive, send):
        if not self.limitadores or scope["type"] != "http" or scope["method"] != "POST":
            await self.app(scope, receive, send)
            return
        grupo = GRUPOS_POR_ROTA.get(scope["path"])
        if grupo is None:
            await self.app(scope, receive, send)
            return

        decisao = self.limitadores[grupo].tentar(ip_do_cliente(scope, self.confiar_x_forwarded_for))
        cabecalhos = {"X-RateLimit-Limit": str(self.limite)}

        if not decisao.permitido:
            # sem o IP no log: e dado pessoal e o grupo/espera ja bastam pra ver o padrao
            logger.warning(json.dumps({"evento": "limite_atingido", "grupo": grupo, "espera_s": decisao.espera_s}))
            resposta = JSONResponse(
                status_code=429,
                content={"detail": _mensagem(grupo, self.limite, decisao.espera_s)},
                headers={**cabecalhos, "X-RateLimit-Remaining": "0", "Retry-After": str(decisao.espera_s)},
            )
            await resposta(scope, receive, send)
            return

        cabecalhos["X-RateLimit-Remaining"] = str(decisao.restantes)

        async def send_com_cabecalhos(mensagem):
            if mensagem["type"] == "http.response.start":
                extras = [(k.lower().encode("latin-1"), v.encode("latin-1")) for k, v in cabecalhos.items()]
                mensagem = {**mensagem, "headers": [*mensagem.get("headers", []), *extras]}
            await send(mensagem)

        await self.app(scope, receive, send_com_cabecalhos)
