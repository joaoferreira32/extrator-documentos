"""Limite de requisicoes por IP (app/rate_limit.py) -- item 7 do diagnostico.

Quatro camadas: o limitador puro (relogio falso, deterministico), a escolha do
IP, o middleware ASGI isolado, e o app real (ASGI direto e um uvicorn de
verdade, este ultimo pra provar o comportamento pelo caminho HTTP inteiro).
`httpx` (TestClient) nao esta nas dependencias -- o app e chamado como ASGI puro.
"""
import asyncio
import json
import sys
import urllib.error
import urllib.request
from pathlib import Path

import pytest

sys.path.insert(0, str(Path(__file__).resolve().parent / "e2e"))
import helpers  # noqa: E402

from app import config, main  # noqa: E402
from app.rate_limit import (  # noqa: E402
    GRUPOS_POR_ROTA,
    JANELA_SEGUNDOS,
    LimiteDeRequisicoesMiddleware,
    LimitadorPorJanela,
    ip_do_cliente,
)


class Relogio:
    def __init__(self):
        self.t = 1000.0

    def __call__(self):
        return self.t

    def avancar(self, segundos):
        self.t += segundos


# ---------- limitador puro ----------


def test_permite_ate_o_limite_e_recusa_a_seguinte():
    r = Relogio()
    lim = LimitadorPorJanela(3, 60, r)
    assert [(d.permitido, d.restantes) for d in (lim.tentar("a") for _ in range(3))] == [(True, 2), (True, 1), (True, 0)]
    d = lim.tentar("a")
    assert d.permitido is False and d.restantes == 0 and d.espera_s == 60


def test_espera_e_exata_e_insistir_nao_estende_o_bloqueio():
    r = Relogio()
    lim = LimitadorPorJanela(2, 60, r)
    lim.tentar("a"), lim.tentar("a")  # t=0: as duas vagas
    r.avancar(10)
    assert lim.tentar("a").espera_s == 50
    r.avancar(10)
    assert lim.tentar("a").espera_s == 40  # a recusa em t=10 nao foi registrada: a espera so encolhe
    r.avancar(40)  # t=60: a 1a requisicao completa a janela
    assert lim.tentar("a").permitido is True


def test_a_janela_desliza_em_vez_de_zerar_de_uma_vez():
    r = Relogio()
    lim = LimitadorPorJanela(3, 60, r)
    lim.tentar("a")  # t=0
    r.avancar(30)
    lim.tentar("a")  # t=30
    r.avancar(29)
    lim.tentar("a")  # t=59: cheio
    assert lim.tentar("a").permitido is False
    r.avancar(2)  # t=61: so a de t=0 saiu da janela
    assert lim.tentar("a").permitido is True  # abriu UMA vaga...
    assert lim.tentar("a").permitido is False  # ...nao tres


def test_espera_arredonda_pra_cima_e_nunca_e_zero():
    r = Relogio()
    lim = LimitadorPorJanela(1, 60, r)
    lim.tentar("a")
    r.avancar(59.2)
    assert lim.tentar("a").espera_s == 1  # faltam 0,8 s: "aguarde 1 segundo", nao "0"


def test_chaves_sao_independentes():
    lim = LimitadorPorJanela(1, 60, Relogio())
    assert lim.tentar("a").permitido and not lim.tentar("a").permitido
    assert lim.tentar("b").permitido


def test_memoria_limitada_a_max_chaves_descartando_as_mais_antigas():
    r = Relogio()
    lim = LimitadorPorJanela(1, 60, r, max_chaves=3)
    for i in range(50):
        lim.tentar(f"ip{i}")
    assert len(lim) == 3
    # a chave mais antiga foi descartada: ganha janela nova (mais permissivo, nunca mais restritivo)
    assert lim.tentar("ip0").permitido is True


def test_o_uso_recente_protege_a_chave_do_descarte():
    lim = LimitadorPorJanela(5, 60, Relogio(), max_chaves=2)
    lim.tentar("a")
    lim.tentar("b")
    lim.tentar("a")  # "a" volta pro fim da fila (LRU)
    lim.tentar("c")  # estoura: sai "b", nao "a"
    assert lim.tentar("a").restantes == 5 - 3  # ainda lembra das 2 anteriores


@pytest.mark.parametrize("limite, max_chaves", [(0, 10), (-1, 10), (1, 0)])
def test_limite_ou_teto_invalido_e_recusado(limite, max_chaves):
    with pytest.raises(ValueError):
        LimitadorPorJanela(limite, 60, max_chaves=max_chaves)


# ---------- IP do cliente ----------


def _scope(cliente=("10.0.0.5", 4000), xff=None):
    headers = [(b"x-forwarded-for", v.encode()) for v in ([xff] if isinstance(xff, str) else (xff or []))]
    scope = {"type": "http", "headers": headers}
    if cliente is not None:
        scope["client"] = cliente
    return scope


def test_sem_confiar_no_proxy_o_cabecalho_forjado_e_ignorado():
    assert ip_do_cliente(_scope(("203.0.113.9", 1), xff="1.2.3.4"), confiar_x_forwarded_for=False) == "203.0.113.9"


def test_confiando_no_proxy_vale_o_primeiro_ip_da_lista():
    scope = _scope(xff="198.51.100.7, 172.68.1.1, 10.0.0.9")
    assert ip_do_cliente(scope, confiar_x_forwarded_for=True) == "198.51.100.7"


def test_varias_linhas_do_cabecalho_contam_como_uma_lista_so():
    scope = _scope(xff=["198.51.100.7", "172.68.1.1"])
    assert ip_do_cliente(scope, confiar_x_forwarded_for=True) == "198.51.100.7"


@pytest.mark.parametrize(
    "bruto, esperado",
    [
        ("198.51.100.7:5678", "198.51.100.7"),  # com porta
        ("::ffff:198.51.100.7", "198.51.100.7"),  # IPv4 embrulhado em IPv6
        ("  198.51.100.7  ", "198.51.100.7"),
    ],
)
def test_formas_do_mesmo_ip_viram_a_mesma_chave(bruto, esperado):
    assert ip_do_cliente(_scope(xff=bruto), confiar_x_forwarded_for=True) == esperado


def test_ipv6_conta_pela_rede_64_e_nao_pelo_endereco():
    a = ip_do_cliente(_scope(xff="2001:db8:1:2::1"), True)
    b = ip_do_cliente(_scope(xff="2001:db8:1:2:ffff:ffff:ffff:ffff"), True)
    c = ip_do_cliente(_scope(xff="2001:db8:1:3::1"), True)
    d = ip_do_cliente(_scope(xff="[2001:db8:1:2::9]:443"), True)
    assert a == b == d == "2001:db8:1:2::/64"
    assert c != a


@pytest.mark.parametrize("lixo", ["nao-e-ip", "", "999.1.1.1", "[::1", "a,b,c"])
def test_cabecalho_com_lixo_cai_no_ip_da_conexao(lixo):
    assert ip_do_cliente(_scope(("10.0.0.5", 1), xff=lixo), True) == "10.0.0.5"


def test_sem_cliente_no_scope_nao_quebra():
    assert ip_do_cliente(_scope(cliente=None), False) == "desconhecido"
    assert ip_do_cliente(_scope(cliente=None), True) == "desconhecido"


# ---------- middleware ASGI isolado ----------


async def _chamar(app, metodo="POST", caminho="/extract-document", ip="203.0.113.1", cabecalhos=None, corpo=b""):
    scope = {
        "type": "http", "asgi": {"version": "3.0"}, "http_version": "1.1", "method": metodo, "path": caminho,
        "raw_path": caminho.encode(), "query_string": b"", "scheme": "http", "server": ("testserver", 80),
        "client": (ip, 5000),
        "headers": [(k.lower().encode(), v.encode()) for k, v in (cabecalhos or {}).items()],
    }
    enviado, respostas = False, []

    async def receive():
        nonlocal enviado
        if not enviado:
            enviado = True
            return {"type": "http.request", "body": corpo, "more_body": False}
        return {"type": "http.disconnect"}

    async def send(m):
        respostas.append(m)

    await app(scope, receive, send)
    inicio = next(m for m in respostas if m["type"] == "http.response.start")
    headers = {k.decode(): v.decode() for k, v in inicio["headers"]}
    corpo_resposta = b"".join(m.get("body", b"") for m in respostas if m["type"] == "http.response.body")
    return inicio["status"], headers, corpo_resposta


def _app_de_mentira():
    chamadas = []

    async def app(scope, receive, send):
        chamadas.append((scope["method"], scope["path"]))
        await send({"type": "http.response.start", "status": 200, "headers": []})
        await send({"type": "http.response.body", "body": b"ok"})

    return app, chamadas


def _mw(limite=3, relogio=None, confiar=False):
    app, chamadas = _app_de_mentira()
    return LimiteDeRequisicoesMiddleware(app, limite, confiar, relogio or Relogio()), chamadas


def _rodar(corrotina):
    return asyncio.run(corrotina)


def test_as_primeiras_passam_a_seguinte_recebe_429_e_nao_chega_ao_app():
    mw, chamadas = _mw(3)

    async def cenario():
        return [await _chamar(mw) for _ in range(4)]

    respostas = _rodar(cenario())
    assert [r[0] for r in respostas] == [200, 200, 200, 429]
    assert len(chamadas) == 3  # a recusada nao alcancou o app: o corpo do upload nem foi lido


def test_resposta_429_tem_mensagem_clara_em_json_e_cabecalhos_padrao():
    mw, _ = _mw(2)

    async def cenario():
        for _ in range(2):
            await _chamar(mw)
        return await _chamar(mw)

    status, headers, corpo = _rodar(cenario())
    detail = json.loads(corpo)["detail"]
    assert status == 429 and headers["content-type"].startswith("application/json")
    assert "limite de 2 extrações por minuto" in detail and "demonstração pública" in detail
    assert "Aguarde 60 segundos" in detail and "envie o documento de novo" in detail
    assert headers["retry-after"] == "60"
    assert headers["x-ratelimit-limit"] == "2" and headers["x-ratelimit-remaining"] == "0"


def test_mensagem_de_exportacao_e_a_de_1_segundo_no_singular():
    r = Relogio()
    mw, _ = _mw(1, r)

    async def cenario():
        await _chamar(mw, caminho="/export-excel")
        r.avancar(59)
        return await _chamar(mw, caminho="/export-excel")

    status, headers, corpo = _rodar(cenario())
    detail = json.loads(corpo)["detail"]
    assert status == 429 and "1 exportações por minuto" in detail  # limite 1: o texto usa o numero real
    assert "Aguarde 1 segundo e baixe o Excel de novo" in detail and headers["retry-after"] == "1"


def test_resposta_aceita_informa_quantas_restam():
    mw, _ = _mw(3)

    async def cenario():
        return [(await _chamar(mw))[1]["x-ratelimit-remaining"] for _ in range(3)]

    assert _rodar(cenario()) == ["2", "1", "0"]


def test_depois_de_esperar_a_janela_libera():
    r = Relogio()
    mw, _ = _mw(1, r)

    async def cenario():
        a = (await _chamar(mw))[0]
        b = (await _chamar(mw))[0]
        r.avancar(JANELA_SEGUNDOS)
        c = (await _chamar(mw))[0]
        return a, b, c

    assert _rodar(cenario()) == (200, 429, 200)


def test_ips_diferentes_nao_se_afetam():
    mw, _ = _mw(1)

    async def cenario():
        return [(await _chamar(mw, ip=ip))[0] for ip in ("203.0.113.1", "203.0.113.1", "203.0.113.2")]

    assert _rodar(cenario()) == [200, 429, 200]


def test_extracao_e_exportacao_tem_cotas_separadas_e_o_debug_divide_a_da_extracao():
    mw, _ = _mw(1)

    async def cenario():
        extracao = (await _chamar(mw, caminho="/extract-document"))[0]
        debug = (await _chamar(mw, caminho="/debug/extract-text"))[0]  # mesma cota da extracao: esgotada
        exportacao = (await _chamar(mw, caminho="/export-excel"))[0]  # cota propria: livre
        return extracao, debug, exportacao

    assert _rodar(cenario()) == (200, 429, 200)


def test_so_post_em_rota_listada_conta():
    mw, chamadas = _mw(1)

    async def cenario():
        for _ in range(5):
            await _chamar(mw, metodo="GET", caminho="/extract-document")
            await _chamar(mw, metodo="GET", caminho="/health")
            await _chamar(mw, metodo="POST", caminho="/uma-rota-que-nao-e-de-upload")
        return [(await _chamar(mw))[0] for _ in range(2)]

    assert _rodar(cenario()) == [200, 429]  # nada acima gastou a cota
    assert len(chamadas) == 5 * 3 + 1


def test_limite_zero_desliga():
    mw, chamadas = _mw(0)

    async def cenario():
        return {(await _chamar(mw))[0] for _ in range(50)}

    assert _rodar(cenario()) == {200} and len(chamadas) == 50


def test_x_forwarded_for_so_decide_o_ip_quando_se_confia_no_proxy():
    async def cenario(confiar):
        mw, _ = _mw(1, confiar=confiar)
        forjado = [{"X-Forwarded-For": f"198.51.100.{i}"} for i in (1, 2)]
        return [(await _chamar(mw, ip="10.0.0.5", cabecalhos=h))[0] for h in forjado]

    assert _rodar(cenario(False)) == [200, 429]  # ignorado: os dois vieram da mesma conexao
    assert _rodar(cenario(True)) == [200, 200]  # atras de proxy: cada IP do cabecalho tem a sua cota


# ---------- cobertura de rotas e configuracao ----------


def test_todo_endpoint_post_do_app_esta_sob_o_limite():
    """Se alguem criar um endpoint de upload novo e esquecer de listar em
    GRUPOS_POR_ROTA, este teste falha -- em vez de a instancia saturar."""
    posts = {r.path for r in main.app.routes if "POST" in (getattr(r, "methods", None) or set())}
    assert posts == set(GRUPOS_POR_ROTA)


def test_o_padrao_e_10_por_minuto_e_o_app_usa_a_configuracao(monkeypatch):
    monkeypatch.delenv("RATE_LIMIT_POR_MINUTO", raising=False)
    assert config.LIMITE_REQUISICOES_POR_MINUTO_PADRAO == 10 and config.limite_requisicoes_por_minuto() == 10
    assert JANELA_SEGUNDOS == 60  # as mensagens dizem "por minuto"


@pytest.mark.parametrize("bruto, esperado", [("25", 25), ("0", 0), ("-5", 0), ("abc", 10), ("", 10), (" 7 ", 7)])
def test_variavel_de_ambiente_do_limite(monkeypatch, bruto, esperado):
    monkeypatch.setenv("RATE_LIMIT_POR_MINUTO", bruto)
    assert config.limite_requisicoes_por_minuto() == esperado  # invalido volta ao padrao, nunca derruba o app


@pytest.mark.parametrize("bruto, esperado", [("true", True), ("1", True), ("SIM", True), ("", False), ("false", False), ("talvez", False)])
def test_confiar_no_proxy_e_desligado_por_padrao_fora_do_render(monkeypatch, bruto, esperado):
    monkeypatch.delenv("RENDER", raising=False)
    monkeypatch.setenv("CONFIAR_X_FORWARDED_FOR", bruto)
    assert config.confiar_x_forwarded_for() is esperado


@pytest.mark.parametrize("bruto, esperado", [("", True), ("talvez", True), ("false", False), ("0", False), ("true", True)])
def test_no_render_confiar_no_proxy_e_o_padrao_mas_da_pra_desligar(monkeypatch, bruto, esperado):
    """O Render sempre define RENDER=true. Sem isso, um servico criado pelo painel
    (sem Blueprint, que ignora os envVars do render.yaml) poria todos os
    visitantes na cota do IP do proxy."""
    monkeypatch.setenv("RENDER", "true")
    monkeypatch.setenv("CONFIAR_X_FORWARDED_FOR", bruto)
    assert config.confiar_x_forwarded_for() is esperado


# ---------- app real (ASGI direto) ----------


def _multipart_nao_pdf():
    boundary = "----limite"
    corpo = (
        f'--{boundary}\r\nContent-Disposition: form-data; name="file"; filename="x.txt"\r\n'
        f"Content-Type: text/plain\r\n\r\nnao e pdf\r\n--{boundary}--\r\n"
    ).encode()
    return corpo, {"Content-Type": f"multipart/form-data; boundary={boundary}"}


def test_app_real_barra_a_extracao_alem_do_limite_antes_de_validar_o_upload():
    limite = config.limite_requisicoes_por_minuto()
    assert limite > 0, "RATE_LIMIT_POR_MINUTO=0 no ambiente dos testes: o limite esta desligado"
    corpo, cabecalhos = _multipart_nao_pdf()

    async def cenario():
        return [await _chamar(main.app, ip="198.51.100.77", cabecalhos=cabecalhos, corpo=corpo) for _ in range(limite + 1)]

    respostas = _rodar(cenario())
    assert [r[0] for r in respostas[:limite]] == [400] * limite  # passaram pelo limite e caíram na validação
    status, headers, corpo_429 = respostas[-1]
    assert status == 429 and "extrações por minuto" in json.loads(corpo_429)["detail"]
    assert headers["retry-after"] and headers["x-request-id"]  # o 429 tambem sai com o id de correlacao


def test_health_e_a_pagina_nunca_sao_limitados_no_app_real():
    async def cenario():
        return {(await _chamar(main.app, metodo="GET", caminho="/health", ip="198.51.100.78"))[0] for _ in range(40)}

    assert _rodar(cenario()) == {200}


# ---------- uvicorn de verdade ----------


# O uvicorn, POR PADRAO, confia em quem chega de 127.0.0.1 (o proprio teste) e
# reescreve o IP do cliente a partir do X-Forwarded-For ANTES do nosso codigo
# ver a requisicao (`--forwarded-allow-ips`). Pra testar a logica do limite e
# nao a do uvicorn, os testes com o cabecalho o apontam pra outro endereco.
SEM_PROXY_DO_UVICORN = {"FORWARDED_ALLOW_IPS": "192.0.2.1"}


def _post(url, corpo, cabecalhos):
    req = urllib.request.Request(url + "extract-document", data=corpo, headers=cabecalhos, method="POST")
    try:
        with urllib.request.urlopen(req, timeout=10) as r:
            return r.status, {k.lower(): v for k, v in r.headers.items()}, r.read()
    except urllib.error.HTTPError as e:
        return e.code, {k.lower(): v for k, v in e.headers.items()}, e.read()


def test_servidor_real_devolve_429_json_e_o_health_continua_de_pe():
    processo, url = helpers.iniciar_servidor({"RATE_LIMIT_POR_MINUTO": "3"})
    try:
        corpo, cabecalhos = _multipart_nao_pdf()
        codigos = [_post(url, corpo, cabecalhos)[0] for _ in range(3)]
        status, headers, resposta = _post(url, corpo, cabecalhos)
        assert codigos == [400, 400, 400] and status == 429
        assert "limite de 3 extrações por minuto" in json.loads(resposta)["detail"]
        assert int(headers["retry-after"]) >= 1
        assert urllib.request.urlopen(url + "health", timeout=5).status == 200
    finally:
        helpers.parar_servidor(processo)


def test_servidor_real_ignora_x_forwarded_for_forjado_por_padrao():
    """Seguro por padrao: com o servidor exposto direto, inventar o cabecalho
    nao ganha uma cota nova."""
    processo, url = helpers.iniciar_servidor({"RATE_LIMIT_POR_MINUTO": "2", **SEM_PROXY_DO_UVICORN})
    try:
        corpo, cabecalhos = _multipart_nao_pdf()
        codigos = [_post(url, corpo, {**cabecalhos, "X-Forwarded-For": f"198.51.100.{i}"})[0] for i in range(1, 4)]
        assert codigos == [400, 400, 429]
    finally:
        helpers.parar_servidor(processo)


def test_servidor_real_atras_de_proxy_conta_por_ip_do_cabecalho():
    """Com CONFIAR_X_FORWARDED_FOR (Render), cada IP do cabecalho tem a propria
    cota. E a mesma propriedade que permite forjar o cabecalho -- limitacao
    documentada em app/rate_limit.py."""
    processo, url = helpers.iniciar_servidor({"RATE_LIMIT_POR_MINUTO": "1", "CONFIAR_X_FORWARDED_FOR": "true", **SEM_PROXY_DO_UVICORN})
    try:
        corpo, cabecalhos = _multipart_nao_pdf()
        por_ip = [_post(url, corpo, {**cabecalhos, "X-Forwarded-For": ip})[0] for ip in ("198.51.100.1", "198.51.100.1", "198.51.100.2")]
        assert por_ip == [400, 429, 400]
    finally:
        helpers.parar_servidor(processo)
