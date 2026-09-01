"""Extracao estruturada via API da Anthropic (modo "ia", opcional).

So deve ser chamado quando app.config.ia_disponivel() for True. Erros (rede,
autenticacao, rate limit, resposta que nao valida contra o schema) devem
propagar -- e main.py que decide o fallback para o modo basico, para que a
falta ou falha da IA nunca vire um erro para quem esta usando o sistema.
"""
import anthropic

from app.config import ANTHROPIC_API_KEY
from app.schemas import DocumentoExtraido

MODEL = "claude-sonnet-5"

SYSTEM_PROMPT = (
    "Voce extrai dados estruturados de documentos comerciais em portugues "
    "(notas fiscais, pedidos de compra, relatorios). Preencha os campos que "
    "conseguir identificar no texto do documento e deixe nulos os que nao "
    "aparecerem. Nao invente valores que nao estao no texto."
)


def extrair(texto: str) -> DocumentoExtraido:
    client = anthropic.Anthropic(api_key=ANTHROPIC_API_KEY)
    response = client.messages.parse(
        model=MODEL,
        max_tokens=4096,
        system=SYSTEM_PROMPT,
        messages=[
            {
                "role": "user",
                "content": f"Extraia os dados estruturados deste documento:\n\n{texto}",
            }
        ],
        output_format=DocumentoExtraido,
    )
    return response.parsed_output
