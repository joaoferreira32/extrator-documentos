from pathlib import Path

from fastapi import FastAPI, HTTPException, UploadFile
from fastapi.responses import StreamingResponse
from fastapi.staticfiles import StaticFiles

from app import basic_extractor, config, excel_exporter, extractors, llm_extractor, pdf_extractor
from app.extractors.danfe import DanfeExtractor
from app.extractors.danfe_tabela import montar_tabela_itens
from app.schemas import ExtractionResult

BASE_DIR = Path(__file__).resolve().parent.parent.parent
FRONTEND_DIR = BASE_DIR / "frontend"

TAMANHO_MAXIMO_BYTES = 20 * 1024 * 1024  # 20 MB

app = FastAPI(title="Extrator Inteligente de Documentos")


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
        aviso=" ".join(avisos) if avisos else None,
        documento=resultado_extracao.documento,
    )


async def _ler_pdf(file: UploadFile, limitar_tamanho: bool = False) -> pdf_extractor.TextoExtraido:
    """Valida o upload e le o PDF. TODOS os endpoints que leem PDF passam
    por aqui (extract-document e os de debug), entao nao existe "outro
    jeito de ler" -- o que os endpoints de debug mostram e exatamente o
    que o extrator recebe."""
    if file.content_type != "application/pdf":
        raise HTTPException(status_code=400, detail="Envie um arquivo PDF.")

    conteudo = await file.read()
    if not conteudo:
        raise HTTPException(status_code=400, detail="Arquivo vazio.")
    if limitar_tamanho and len(conteudo) > TAMANHO_MAXIMO_BYTES:
        raise HTTPException(status_code=400, detail="Arquivo maior que 20 MB.")

    try:
        return pdf_extractor.extrair_texto(conteudo)
    except Exception:
        raise HTTPException(
            status_code=400,
            detail="Nao foi possivel ler este arquivo como PDF. Ele pode estar corrompido.",
        )


@app.get("/health")
def health():
    return {"status": "ok"}


@app.post("/extract-document", response_model=ExtractionResult)
async def extract_document(file: UploadFile):
    resultado_texto = await _ler_pdf(file, limitar_tamanho=True)

    if resultado_texto.parece_escaneado:
        if resultado_texto.ocr_disponivel:
            aviso = (
                "Este PDF parece ser uma imagem escaneada e o OCR nao "
                "conseguiu extrair texto legivel dele."
            )
        else:
            aviso = (
                "Este PDF parece ser uma imagem escaneada. OCR nao esta "
                "disponivel neste momento (Tesseract nao encontrado no "
                "sistema) -- veja o README para instalar."
            )
        return ExtractionResult(
            modo_extracao="basico",
            origem_texto=resultado_texto.origem,
            aviso=aviso,
            documento=basic_extractor.extrair(""),
        )

    if config.ia_disponivel():
        try:
            documento = llm_extractor.extrair(resultado_texto.texto)
            return ExtractionResult(
                modo_extracao="ia", origem_texto=resultado_texto.origem, documento=documento
            )
        except Exception as exc:
            # Captura ampla e intencional: qualquer falha da IA (rede, auth,
            # rate limit, resposta fora do schema) deve cair para o modo
            # basico em vez de virar erro para quem esta usando o sistema.
            resultado_extracao = basic_extractor.extrair_com_metadados(
                resultado_texto.texto, resultado_texto.paginas_palavras
            )
            return _resultado_basico(
                resultado_texto,
                resultado_extracao,
                aviso_extra=(
                    f"Extracao por IA indisponivel no momento "
                    f"({type(exc).__name__}), usando modo basico."
                ),
            )

    resultado_extracao = basic_extractor.extrair_com_metadados(
        resultado_texto.texto, resultado_texto.paginas_palavras
    )
    return _resultado_basico(resultado_texto, resultado_extracao)


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
                "Este PDF nao tem posicao de palavra disponivel (provavelmente "
                "e uma imagem escaneada, sem camada de texto digital)."
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
    return montar_debug_entrada_do_extrator(resultado_texto)


@app.post("/export-excel")
async def export_excel(resultado: ExtractionResult):
    buffer = excel_exporter.gerar_excel(resultado)
    return StreamingResponse(
        buffer,
        media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        headers={"Content-Disposition": 'attachment; filename="documento_extraido.xlsx"'},
    )


# Rotas da API (extract-document, export-excel, ...) devem ser adicionadas
# ACIMA deste mount. O StaticFiles em "/" e um catch-all: qualquer rota
# definida depois dele nunca seria alcancada.
app.mount("/", StaticFiles(directory=FRONTEND_DIR, html=True), name="frontend")
