"""Bug real corrigido: os endpoints que fazem trabalho pesado (leitura de
PDF, extracao, geracao de Excel) eram `async def` mas chamavam codigo
sincrono diretamente, sem `run_in_threadpool` -- isso bloqueia o event loop
do asyncio inteiro, travando o processo INTEIRO (inclusive o /health, que um
monitor de uptime usaria) enquanto uma extracao pesada roda.

Medido antes da correcao: uma extracao de DANFE de 300 paginas (~9s) rodando
ao mesmo tempo que chamadas repetidas a /health -- uma chamada deu timeout
total (>5s), outra levou 3.817ms. Sem um teste que reproduza isso, alguem no
futuro reintroduz um `def` sincrono (ou tira o run_in_threadpool) sem
perceber que isso trava o servidor pra todo mundo durante uma extracao.

Sobe um servidor de verdade (mesmo helper dos testes e2e, tests/e2e/
helpers.py) porque o bug e sobre concorrencia real entre requisicoes --
chamar a funcao do endpoint direto (como test_debug_endpoints.py faz) nao
exercita o event loop nem o threadpool de verdade.
"""
import sys
import threading
import time
import urllib.request
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent / "e2e"))
import helpers  # noqa: E402
import pymupdf  # noqa: E402

TIMEOUT_HEALTH = 2.0
LIMITE_ACEITAVEL = 0.5  # bem acima do normal (poucos ms), bem abaixo do que o bug antigo dava


def _pdf_pesado(n_paginas: int) -> bytes:
    """DANFE (o caminho mais pesado: reconstrucao de tabela por coordenadas
    via extract_words) repetida N vezes, so pra garantir uma extracao que
    demore o suficiente pra o teste ter tempo de medir /health durante ela."""
    tmp = Path(helpers.__file__).resolve().parent / "_pdf_pesado_teste.pdf"
    helpers._danfe(tmp, com_numero=True, produtos="215,03")
    with pymupdf.open(tmp) as base, pymupdf.open() as combinado:
        for _ in range(n_paginas):
            combinado.insert_pdf(base)
        conteudo = combinado.tobytes()
    tmp.unlink()  # so depois de fechar os dois Document (Windows trava arquivo aberto)
    return conteudo


def _multipart(conteudo: bytes, nome: str) -> tuple[bytes, str]:
    boundary = "----teste-concorrencia"
    corpo = (
        f"--{boundary}\r\nContent-Disposition: form-data; name=\"file\"; filename=\"{nome}\"\r\n"
        f"Content-Type: application/pdf\r\n\r\n"
    ).encode() + conteudo + f"\r\n--{boundary}--\r\n".encode()
    return corpo, f"multipart/form-data; boundary={boundary}"


def _medir_health_durante(url_health: str, disparar_trabalho_pesado) -> list[float]:
    """Bate /health repetidamente enquanto `disparar_trabalho_pesado()` roda
    (bloqueante) numa thread separada. Devolve a latencia de cada chamada;
    uma chamada que falhar (timeout, conexao recusada) conta como
    LIMITE_ACEITAVEL*10 -- nunca "some" silenciosamente do resultado."""
    latencias: list[float] = []
    parar = threading.Event()

    def bater_health():
        while not parar.is_set():
            t0 = time.perf_counter()
            try:
                urllib.request.urlopen(url_health, timeout=TIMEOUT_HEALTH).read()
                latencias.append(time.perf_counter() - t0)
            except Exception:
                latencias.append(LIMITE_ACEITAVEL * 10)
            time.sleep(0.05)

    t_health = threading.Thread(target=bater_health, daemon=True)
    t_health.start()
    time.sleep(0.15)  # umas leituras de baseline antes do trabalho pesado comecar

    disparar_trabalho_pesado()

    time.sleep(0.15)
    parar.set()
    t_health.join(timeout=TIMEOUT_HEALTH + 1)
    return latencias


def test_extracao_pesada_nao_bloqueia_o_health_check():
    processo, url = helpers.iniciar_servidor()
    try:
        conteudo = _pdf_pesado(150)  # ~4s de processamento (medido: ~28ms/pagina de DANFE)
        corpo, content_type = _multipart(conteudo, "pesado.pdf")

        def extrair():
            req = urllib.request.Request(
                url + "extract-document", data=corpo, headers={"Content-Type": content_type}, method="POST"
            )
            urllib.request.urlopen(req, timeout=30).read()

        latencias = _medir_health_durante(url + "health", extrair)

        assert len(latencias) >= 3, "poucas leituras de /health -- o teste nao rodou tempo suficiente pra valer"
        assert max(latencias) < LIMITE_ACEITAVEL, (
            f"/health ficou lento durante a extracao pesada (latencias: {latencias}) -- "
            "isso indica que algum endpoint voltou a bloquear o event loop com trabalho "
            "sincrono sem run_in_threadpool."
        )
    finally:
        helpers.parar_servidor(processo)


def test_export_excel_grande_nao_bloqueia_o_health_check():
    """Mesmo bug, endpoint diferente: /export-excel tambem faz trabalho
    pesado (montar o .xlsx) de forma sincrona. Usa o MAIOR lote que os
    tetos de app/schemas.py (item 2) ainda aceitam -- 5 documentos de 1000
    itens cada, o teto de TETO_ITENS_TOTAL_DO_LOTE -- que ja demora ~3s
    (medido), o bastante pra observar /health durante o processamento."""
    processo, url = helpers.iniciar_servidor()
    try:
        def documento(id_doc):
            return {
                "arquivo": f"grande_{id_doc}.pdf",
                "resultado": {
                    "modo_extracao": "basico",
                    "confiancas": {},
                    "avisos": [],
                    "documento": {
                        "tipo_documento": "nota_fiscal",
                        "numero_documento": str(id_doc),
                        "itens": [
                            {"descricao": f"Item {i}", "quantidade": 1.0, "valor_unitario": 1.0, "valor_total": 1.0}
                            for i in range(1000)
                        ],
                    },
                },
                "corrigidos": {},
            }

        import json

        payload = json.dumps({"documentos": [documento(i) for i in range(5)]}).encode()  # 5x1000 = 5000

        def exportar():
            req = urllib.request.Request(
                url + "export-excel", data=payload, headers={"Content-Type": "application/json"}, method="POST"
            )
            urllib.request.urlopen(req, timeout=60).read()

        latencias = _medir_health_durante(url + "health", exportar)

        assert len(latencias) >= 3, "poucas leituras de /health -- o teste nao rodou tempo suficiente pra valer"
        assert max(latencias) < LIMITE_ACEITAVEL, (
            f"/health ficou lento durante a exportacao grande (latencias: {latencias})"
        )
    finally:
        helpers.parar_servidor(processo)
