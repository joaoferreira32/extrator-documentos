"""Apoio dos testes e2e: servidor uvicorn temporario e PDFs FICTICIOS.

Nenhum documento real entra aqui: os PDFs sao gerados por PyMuPDF com dados
inventados (a unica coisa "real" e publica e o CNPJ/nome da Dell, ja usados
nas fixtures; a chave de acesso e sintetica, com digito verificador valido).
"""
import os
import re
import socket
import subprocess
import sys
import time
import urllib.request
from pathlib import Path

import pymupdf

BACKEND = Path(__file__).resolve().parents[2]
FIXTURE_BOLETO = BACKEND / "tests" / "fixtures" / "boleto_real_anonimizado.txt"

CHAVE = "3526 0472 3811 8900 1001 5500 1000 0123 4511 2345 6786"  # sintetica
CHAVE_SEM_ESPACOS = CHAVE.replace(" ", "")
_T = 7  # tamanho da fonte da DANFE


# ---------- servidor ----------


def iniciar_servidor(env_extra: dict[str, str] | None = None):
    """Sobe o app numa porta livre e devolve (processo, url_base).

    `env_extra` sobrescreve variaveis de ambiente do servidor (ex: o limite de
    requisicoes). O limite por IP vem DESLIGADO por padrao aqui: as suites e2e e
    de concorrencia fazem dezenas de requisicoes por minuto do mesmo IP local, e
    seriam barradas. Os testes do proprio limite ligam de proposito."""
    with socket.socket() as s:
        s.bind(("127.0.0.1", 0))
        porta = s.getsockname()[1]

    # Chave vazia: o teste tem que ser sempre modo BASICO, mesmo que a maquina
    # tenha ANTHROPIC_API_KEY (load_dotenv nao sobrescreve variavel ja definida).
    env = {
        **os.environ,
        "ANTHROPIC_API_KEY": "",
        "PYTHONIOENCODING": "utf-8",
        "RATE_LIMIT_POR_MINUTO": "0",
        "CONFIAR_X_FORWARDED_FOR": "",
        "RENDER": "",  # rodando os testes numa maquina do Render, o padrao mudaria
        **(env_extra or {}),
    }
    processo = subprocess.Popen(
        [sys.executable, "-m", "uvicorn", "app.main:app", "--port", str(porta), "--log-level", "warning"],
        cwd=BACKEND,
        env=env,
        stdout=subprocess.DEVNULL,
        stderr=subprocess.DEVNULL,
    )
    url = f"http://127.0.0.1:{porta}/"
    limite = time.time() + 30
    while time.time() < limite:
        if processo.poll() is not None:
            raise RuntimeError("o uvicorn encerrou ao iniciar (rode `uvicorn app.main:app` pra ver o erro)")
        try:
            urllib.request.urlopen(url + "health", timeout=1)
            return processo, url
        except OSError:
            time.sleep(0.25)
    processo.kill()
    raise RuntimeError("o uvicorn nao respondeu em 30s")


def parar_servidor(processo):
    processo.terminate()
    try:
        processo.wait(timeout=10)
    except subprocess.TimeoutExpired:
        processo.kill()


# ---------- PDFs fictícios ----------


def _danfe(destino: Path, com_numero: bool, produtos: str):
    doc = pymupdf.open()
    p = doc.new_page(width=842, height=595)

    def h(txt, x, y):
        p.insert_text((x, y + _T), txt, fontsize=_T, fontname="helv")

    def girado(txt, x, y):
        p.insert_text((x, y), txt, fontsize=_T, fontname="helv", rotate=90)

    girado("TRANSPORTADOS", 20, 300)  # canhoto/rotulos laterais girados, como numa DANFE real
    girado("COMPUTADORES", 32, 300)
    h("FOLHA 1/", 40, 4)
    h("Identificacao do emitente", 40, 20)
    h("DANFE", 300, 20)
    h("DELL COMPUTADORES DO BRASIL LTDA", 40, 30)
    h("Documento Auxiliar da", 300, 30)
    h("Nota Fiscal Eletronica", 40, 40)
    linhas = ["AV EXEMPLO, 5000", "BAIRRO EXEMPLO", "Cidade Exemplo, SP", CHAVE]
    if com_numero:
        linhas.append("Nº000012345")
    linhas += ["linha a", "linha b"]
    for i, texto in enumerate(linhas):
        h(texto, 40, 50 + i * 10)
    h("INSCRICAO ESTADUAL INSCR. ESTADUAL DO SUBST. TRIBUT. CNPJ", 40, 120)
    h("748241245113 72.381.189/0010-01", 40, 130)
    h("NOME/RAZAO SOCIAL CNPJ/CPF DATA DA EMISSAO", 40, 150)
    h("FULANO DE TAL SILVA 000.000.000-00 15/4/2026", 300, 150)
    h("BASE DE CALCULO DO ICMS VALOR DO ICMS BASE DE CALCULO ICMS ST VALOR DO ICMS SUBSTITUICAO VALOR TOTAL DOS PRODUTOS", 40, 200)
    h(f"229,00 41,22 0,00 0,00 {produtos}", 40, 210)
    h("VALOR DO FRETE VALOR DO SEGURO DESCONTO OUTRAS DESPESAS ACESSORIAS VALOR TOTAL DO I.P.I. VALOR TOTAL DA NOTA", 40, 230)
    h("0,00 0,00 0,00 0,00 13,97 229,00", 40, 240)
    cabecalho = [
        ("CÓDIGO", 98.7), ("DESCRIÇÃO", 233.6), ("NCM/SH", 412.6), ("CST", 444.5), ("CFOP", 463.5),
        ("UNID.", 493.2), ("QUANT.", 530.5), ("UNITÁRIO", 562.9), ("TOTAL", 614.3),
        ("B.CALC.ICMS", 663.0), ("ICMS", 705.1), ("I.P.I.", 749.7),
    ]
    for texto, x in cabecalho:
        h(texto, x, 413.6)
    for texto, x in [("IC", 705.0), ("M", 714.0), ("S", 721.0), ("IP", 749.7), ("I", 760.0)]:
        h(texto, x, 416.8)  # fragmentos do cabecalho de varias linhas
    for texto, x in [
        ("460-BCZS", 85), ("Mochila", 169), ("42029200", 408), ("100", 445), ("5102", 463), ("UN", 507.8),
        ("1,0000", 537.4), ("215,0300", 577.2), ("215,03", 635.4),
    ]:
        h(texto, x, 422.1)
    h("INFORMAÇÕES COMPLEMENTARES", 40, 491)
    h("Reservado ao Fisco", 40, 505)
    doc.save(destino)


def _texto(destino: Path, linhas: list[str], fonte=9, passo=12):
    doc = pymupdf.open()
    p = doc.new_page(width=842, height=595)
    for i, linha in enumerate(linhas):
        p.insert_text((20, 30 + i * passo), linha, fontsize=fonte, fontname="helv")
    doc.save(destino)


def gerar_pdfs(pasta: Path) -> dict[str, Path]:
    """Gera os 4 PDFs de cenario em `pasta` e devolve {nome: caminho}."""
    pasta.mkdir(parents=True, exist_ok=True)
    caminhos = {nome: pasta / f"{nome}.pdf" for nome in ("danfe_ok", "danfe_problemas", "generico", "boleto")}

    _danfe(caminhos["danfe_ok"], com_numero=True, produtos="215,03")  # tudo certo
    _danfe(caminhos["danfe_problemas"], com_numero=False, produtos="999,99")  # sem Nº + soma nao fecha
    _texto(  # nota generica: numero e total "baixa", campos obrigatorios vazios
        caminhos["generico"],
        [
            "NOTA FISCAL",
            "Emitente: Empresa Exemplo Ltda CNPJ: 11.222.333/0001-81",
            "Nº 123456",
            "Total: R$ 1.234,56",
        ],
    )
    _texto(  # boleto: texto (ja ficticio) da fixture
        caminhos["boleto"], FIXTURE_BOLETO.read_text(encoding="utf-8").splitlines(), fonte=6.5, passo=9
    )
    return caminhos


# ---------- contraste (WCAG) ----------


def rgb(css):
    return tuple(int(x) for x in re.findall(r"\d+", css)[:3])


def contraste(cor_a, cor_b):
    """Razao de contraste WCAG entre duas cores (r, g, b)."""

    def luminancia(cor):
        def canal(c):
            c /= 255
            return c / 12.92 if c <= 0.03928 else ((c + 0.055) / 1.055) ** 2.4

        r, g, b = cor
        return 0.2126 * canal(r) + 0.7152 * canal(g) + 0.0722 * canal(b)

    claro, escuro = sorted((luminancia(cor_a), luminancia(cor_b)), reverse=True)
    return (claro + 0.05) / (escuro + 0.05)
