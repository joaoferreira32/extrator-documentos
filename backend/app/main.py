import json
import logging
import time
import uuid
from pathlib import Path

from fastapi import FastAPI, HTTPException, UploadFile
from fastapi.concurrency import run_in_threadpool
from fastapi.responses import StreamingResponse
from fastapi.staticfiles import StaticFiles
from starlette.middleware.base import BaseHTTPMiddleware
from starlette.requests import Request

from app import basic_extractor, config, excel_exporter, extractors, llm_extractor, logging_config, pdf_extractor
from app.confianca import CAMPOS_OBRIGATORIOS, vazio
from app.extractors.danfe import DanfeExtractor
from app.extractors.danfe_tabela import montar_tabela_itens
from app.rate_limit import LimiteDeRequisicoesMiddleware
from app.schemas import ExportarExcelRequest, ExtractionResult

BASE_DIR = Path(__file__).resolve().parent.parent.parent
FRONTEND_DIR = BASE_DIR / "frontend"

TAMANHO_MAXIMO_BYTES = 20 * 1024 * 1024  # 20 MB

logging_config.configurar_logging()
logger = logging.getLogger(__name__)

app = FastAPI(title="Extrator Inteligente de Documentos")


class RequestIdMiddleware(BaseHTTPMiddleware):
    """Um id curto por requisicao, disponivel pra QUALQUER logger do projeto
    via contextvar (ver app/logging_config.py) e devolvido no cabecalho
    X-Request-ID -- pra correlacionar um erro que o usuario relatar com a
    linha exata do log do servidor, sem precisar passar o id explicitamente
    em cada chamada de funcao."""

    async def dispatch(self, request: Request, call_next):
        rid = uuid.uuid4().hex[:8]
        logging_config.definir_id_da_requisicao(rid)
        resposta = await call_next(request)
        resposta.headers["X-Request-ID"] = rid
        return resposta


# Ordem importa: o ULTIMO adicionado fica por fora. O request-id envolve o
# limite, entao a resposta 429 tambem sai com X-Request-ID.
app.add_middleware(
    LimiteDeRequisicoesMiddleware,
    limite_por_minuto=config.limite_requisicoes_por_minuto(),
    confiar_x_forwarded_for=config.confiar_x_forwarded_for(),
)
app.add_middleware(RequestIdMiddleware)


def _resultado_basico(resultado_texto, resultado_extracao, aviso_extra: str | None = None) -> ExtractionResult:
    """Monta a resposta do modo basico a partir de um ResultadoExtracao,
    juntando os avisos especificos do extrator (ex: soma da tabela nao
    bate) com um aviso extra do proprio main.py (ex: fallback de IA), se
    houver."""
    avisos = list(resultado_extracao.avisos)
    if aviso_extra:
        avisos.insert(0, aviso_extra)
    return ExtractionResult(
        modo_extracao="basico",
        origem_texto=resultado_texto.origem,
        confiancas=resultado_extracao.confiancas,
        avisos=avisos,
        aviso=" ".join(avisos) if avisos else None,
        documento=resultado_extracao.documento,
    )


async def _ler_pdf(file: UploadFile) -> pdf_extractor.TextoExtraido:
    """Valida o upload e le o PDF. TODOS os endpoints que leem PDF passam
    por aqui (extract-document e os de debug), entao nao existe "outro
    jeito de ler" -- o que os endpoints de debug mostram e exatamente o
    que o extrator recebe.

    O limite de tamanho vale pros 4 (nao so /extract-document): os 3
    endpoints de debug sao publicos e ficaram SEM limite ate aqui -- um
    PDF de 55 MB era processado normalmente neles mesmo depois do limite
    ja existir em /extract-document. Bug real, medido."""
    if file.content_type != "application/pdf":
        raise HTTPException(status_code=400, detail="Este arquivo não é um PDF. Escolha um arquivo .pdf.")

    conteudo = await file.read()
    if not conteudo:
        raise HTTPException(status_code=400, detail="O arquivo está vazio.")
    if len(conteudo) > TAMANHO_MAXIMO_BYTES:
        tamanho_mb = f"{len(conteudo) / (1024 * 1024):.1f}".replace(".", ",")
        raise HTTPException(
            status_code=400,
            detail=f"O arquivo tem {tamanho_mb} MB, e o limite é {TAMANHO_MAXIMO_BYTES // (1024 * 1024)} MB. Envie um PDF menor.",
        )

    try:
        # run_in_threadpool: extrair_texto e sincrono e pode demorar segundos
        # (PDF grande/muitas paginas). Chamar direto aqui bloquearia o event
        # loop do asyncio INTEIRO -- travando ate o /health -- enquanto uma
        # unica extracao roda. Bug real, medido com dois usuarios
        # simultaneos (ver tests/test_concorrencia.py).
        return await run_in_threadpool(pdf_extractor.extrair_texto, conteudo)
    except pdf_extractor.PDFProtegidoPorSenha:
        # Bug real: antes caia no "corrompido" generico abaixo -- mensagem
        # enganosa, o arquivo esta perfeito, so precisa da senha.
        raise HTTPException(
            status_code=400,
            detail="Este PDF está protegido por senha. Remova a senha e envie de novo.",
        )
    except Exception:
        raise HTTPException(
            status_code=400,
            detail="Não foi possível ler este arquivo como PDF. Ele pode estar corrompido.",
        )


@app.get("/health")
def health():
    return {"status": "ok"}


def _logar_extracao(
    resultado_texto: pdf_extractor.TextoExtraido,
    resultado: ExtractionResult,
    duracao_ms: float,
    arquivo: str | None,
    extra: dict | None = None,
) -> None:
    """Uma linha de log por extracao, em JSON (grep + json.loads ja basta
    pra analisar -- sem infraestrutura de log pesada, ver
    app/logging_config.py). Cobre exatamente o que o diagnostico de
    maturidade do projeto apontava como faltando: quanto tempo levou, qual
    extrator foi escolhido (so faz sentido no modo basico -- no modo IA
    quem decide e a LLM), e quais campos mais ficam vazios ou com confianca
    baixa/media."""
    dados = {
        "arquivo": arquivo,
        "duracao_ms": round(duracao_ms, 1),
        "numero_paginas": resultado_texto.numero_paginas,
        "origem_texto": resultado.origem_texto,
        "modo_extracao": resultado.modo_extracao,
    }
    if resultado.confiancas:
        dados["campos_media_ou_baixa"] = sorted(
            chave for chave, nivel in resultado.confiancas.items() if nivel in ("media", "baixa")
        )
    obrigatorios = CAMPOS_OBRIGATORIOS.get(resultado.documento.tipo_documento, [])
    campos_vazios = [c for c in obrigatorios if vazio(getattr(resultado.documento, c))]
    if campos_vazios:
        dados["campos_obrigatorios_vazios"] = campos_vazios
    if extra:
        dados.update(extra)
    logger.info(json.dumps(dados, ensure_ascii=False))


@app.post("/extract-document", response_model=ExtractionResult)
async def extract_document(file: UploadFile):
    inicio = time.perf_counter()
    resultado_texto = await _ler_pdf(file)

    if resultado_texto.numero_paginas == 0:
        # Bug real: sem este `if`, um PDF de 0 paginas caia no
        # `parece_escaneado` abaixo (texto vazio e vazio, mesma condicao) e
        # dizia "parece ser uma imagem escaneada" -- diagnostico errado,
        # nao ha imagem nenhuma, nao ha pagina nenhuma pra ter.
        aviso = "Este PDF não tem nenhuma página."
        resultado = ExtractionResult(
            modo_extracao="basico",
            origem_texto=resultado_texto.origem,
            avisos=[aviso],
            aviso=aviso,
            documento=basic_extractor.extrair(""),
        )
        _logar_extracao(resultado_texto, resultado, (time.perf_counter() - inicio) * 1000, file.filename,
                         {"sem_paginas": True})
        return resultado

    if resultado_texto.parece_escaneado:
        if resultado_texto.ocr_disponivel:
            aviso = (
                "Este PDF parece ser uma imagem escaneada, e a leitura de imagem (OCR) "
                "não conseguiu extrair texto legível dele."
            )
        else:
            # Quem ve isto na demo publica nao tem como instalar nada: a mensagem
            # diz o que fazer com o documento. A instalacao do Tesseract esta no README.
            aviso = (
                "Este PDF parece ser uma imagem escaneada, e a leitura de imagem (OCR) "
                "não está disponível neste servidor. Envie um PDF com texto digital "
                "(gerado pelo sistema emissor, não escaneado)."
            )
        resultado = ExtractionResult(
            modo_extracao="basico",
            origem_texto=resultado_texto.origem,
            avisos=[aviso],
            aviso=aviso,
            documento=basic_extractor.extrair(""),
        )
        _logar_extracao(resultado_texto, resultado, (time.perf_counter() - inicio) * 1000, file.filename,
                         {"escaneado": True})
        return resultado

    if config.ia_disponivel():
        try:
            # llm_extractor.extrair faz uma chamada de rede sincrona (SDK da
            # Anthropic) -- mesma razao do run_in_threadpool acima: sem isso,
            # o event loop fica bloqueado esperando a rede, nao so a CPU.
            documento = await run_in_threadpool(llm_extractor.extrair, resultado_texto.texto)
            resultado = ExtractionResult(
                modo_extracao="ia", origem_texto=resultado_texto.origem, documento=documento
            )
            _logar_extracao(resultado_texto, resultado, (time.perf_counter() - inicio) * 1000, file.filename)
            return resultado
        except Exception as exc:
            # Captura ampla e intencional: qualquer falha da IA (rede, auth,
            # rate limit, resposta fora do schema) deve cair para o modo
            # basico em vez de virar erro para quem esta usando o sistema.
            resultado_extracao = await run_in_threadpool(
                basic_extractor.extrair_com_metadados, resultado_texto.texto, resultado_texto.paginas_palavras
            )
            resultado = _resultado_basico(
                resultado_texto,
                resultado_extracao,
                aviso_extra=(
                    f"A extração por IA não está disponível agora "
                    f"({type(exc).__name__}); o documento foi lido no modo básico."
                ),
            )
            nome_extrator = await run_in_threadpool(
                basic_extractor.nome_extrator, resultado_texto.texto, resultado_texto.paginas_palavras
            )
            _logar_extracao(resultado_texto, resultado, (time.perf_counter() - inicio) * 1000, file.filename,
                             {"extrator": nome_extrator, "fallback_ia": type(exc).__name__})
            return resultado

    resultado_extracao = await run_in_threadpool(
        basic_extractor.extrair_com_metadados, resultado_texto.texto, resultado_texto.paginas_palavras
    )
    resultado = _resultado_basico(resultado_texto, resultado_extracao)
    nome_extrator = await run_in_threadpool(
        basic_extractor.nome_extrator, resultado_texto.texto, resultado_texto.paginas_palavras
    )
    _logar_extracao(resultado_texto, resultado, (time.perf_counter() - inicio) * 1000, file.filename,
                     {"extrator": nome_extrator})
    return resultado


@app.post("/debug/extract-text")
async def debug_extract_text(file: UploadFile):
    """Endpoint de debug: devolve o texto bruto que o pdfplumber extraiu do
    PDF, sem nenhuma extracao de campos em cima. Usado para inspecionar
    como os rotulos aparecem de verdade num documento real antes de
    ajustar as regex/heuristicas do modo basico."""
    resultado_texto = await _ler_pdf(file)

    return {
        "numero_paginas": resultado_texto.numero_paginas,
        "parece_escaneado": resultado_texto.parece_escaneado,
        "origem": resultado_texto.origem,
        "ocr_disponivel": resultado_texto.ocr_disponivel,
        "texto": resultado_texto.texto,
        "linhas": resultado_texto.texto.splitlines(),
    }


@app.post("/debug/extract-words")
async def debug_extract_words(file: UploadFile):
    """Endpoint de debug: devolve as palavras de cada pagina com sua
    posicao (x0/x1/top/bottom), sem nenhuma extracao de campos em cima.
    Usado para diagnosticar layouts em grade (ex: tabela de itens de uma
    DANFE) antes de escrever a logica de reconstrucao por coordenadas --
    texto corrido nao basta pra isso, precisa da posicao real de cada
    palavra tal como o pdfplumber extraiu.

    So funciona para PDFs com texto digital (nao PDF escaneado sem OCR
    bem-sucedido) -- posicao de palavra so existe quando ha uma camada de
    texto real no PDF."""
    resultado_texto = await _ler_pdf(file)

    if resultado_texto.paginas_palavras is None:
        raise HTTPException(
            status_code=400,
            detail=(
                "Este PDF não tem posição de palavra disponível (provavelmente "
                "é uma imagem escaneada, sem camada de texto digital)."
            ),
        )

    return {
        "numero_paginas": resultado_texto.numero_paginas,
        "paginas": [
            [
                {"texto": p.texto, "x0": p.x0, "x1": p.x1, "top": p.top, "bottom": p.bottom}
                for p in pagina
            ]
            for pagina in resultado_texto.paginas_palavras
        ],
    }


def _numerar(linhas: list[str]) -> list[str]:
    return [f"{i:02d}: {linha}" for i, linha in enumerate(linhas)]


def montar_debug_entrada_do_extrator(resultado_texto: pdf_extractor.TextoExtraido) -> dict:
    """Tudo que o extrator recebe e decide para um PDF, no mesmo caminho
    de codigo de /extract-document (basic_extractor.montar_contexto +
    selecionar_extrator + extrair). Separado do endpoint pra poder ser
    testado sem subir o servidor."""
    contexto = basic_extractor.montar_contexto(
        resultado_texto.texto, resultado_texto.paginas_palavras
    )
    extrator = extractors.selecionar_extrator(contexto)
    resultado = extrator.extrair(contexto)

    if resultado_texto.paginas_texto:
        paginas = [
            {"pagina": numero, "linhas": _numerar(texto_pagina.splitlines())}
            for numero, texto_pagina in enumerate(resultado_texto.paginas_texto, start=1)
        ]
        # A juncao das paginas so difere do texto do extrator por um
        # .strip() das pontas; se um dia divergir de verdade, isto denuncia.
        exibicao_fiel = "\n".join(resultado_texto.paginas_texto).strip() == contexto.texto
    else:
        paginas = [{"pagina": 1, "linhas": _numerar(contexto.linhas)}]
        exibicao_fiel = True

    tabela_itens = None
    if isinstance(extrator, DanfeExtractor) and contexto.paginas_palavras:
        tabela_itens = []
        montar_tabela_itens(contexto.paginas_palavras, diagnostico=tabela_itens)

    return {
        "extrator_escolhido": type(extrator).__name__,
        "pontuacoes": extractors.pontuacoes(contexto),
        "origem_texto": resultado_texto.origem,
        "numero_paginas": resultado_texto.numero_paginas,
        "caracteres_girados_descartados": resultado_texto.caracteres_girados_descartados,
        "exibicao_fiel_ao_texto_do_extrator": exibicao_fiel,
        "paginas": paginas,
        "resultado": {
            "documento": resultado.documento.model_dump(),
            "confiancas": resultado.confiancas,
            "avisos": resultado.avisos,
        },
        "tabela_itens": tabela_itens,
    }


@app.post("/debug/extractor-input")
async def debug_extractor_input(file: UploadFile):
    """Endpoint de debug: mostra o texto EXATO que o extrator recebe (linha
    por linha, separado por pagina, com o numero da linha), qual extrator
    foi escolhido e por que, o que ele extraiu, e -- pra DANFE -- cada
    decisao da reconstrucao da tabela de itens (cabecalho achado? onde a
    tabela parou e por que?). Usa o mesmo caminho de /extract-document, so
    sem chamar a IA. Ver CLAUDE.md ("Diagnosticando um PDF real")."""
    resultado_texto = await _ler_pdf(file)
    return await run_in_threadpool(montar_debug_entrada_do_extrator, resultado_texto)


@app.post("/export-excel")
async def export_excel(requisicao: ExportarExcelRequest):
    """Uma lista de documentos (a tela envia 1; a estrutura ja comporta lote).
    run_in_threadpool: montar o .xlsx e sincrono e, com muitos itens, pode
    demorar dezenas de segundos -- mesma razao do _ler_pdf acima (ver
    tests/test_concorrencia.py)."""
    buffer = await run_in_threadpool(excel_exporter.gerar_excel, requisicao.documentos)
    nome = "documento_extraido.xlsx" if len(requisicao.documentos) == 1 else "documentos_extraidos.xlsx"
    return StreamingResponse(
        buffer,
        media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        headers={"Content-Disposition": f'attachment; filename="{nome}"'},
    )


# Rotas da API (extract-document, export-excel, ...) devem ser adicionadas
# ACIMA deste mount. O StaticFiles em "/" e um catch-all: qualquer rota
# definida depois dele nunca seria alcancada.
app.mount("/", StaticFiles(directory=FRONTEND_DIR, html=True), name="frontend")
