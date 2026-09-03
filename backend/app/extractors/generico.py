"""GenericExtractor: fallback para documentos que nao sao boleto nem DANFE
(pedido de compra, relatorio, ou qualquer texto nao reconhecido).

Reaproveita os mesmos helpers de `comum.py` que os outros extratores, mas
com um universo de rotulos mais enxuto -- sem termos bancarios como
"Cedente"/"Sacado" (exclusivos de boleto), que nao fazem sentido aqui.
Sempre devolve uma pontuacao de deteccao baixa e fixa: so vence quando
nenhum extrator mais especifico reconhece o documento.
"""
from app.extractors import comum
from app.extractors.base import ContextoExtracao, ExtratorDocumento, ResultadoExtracao
from app.schemas import DocumentoExtraido

PONTUACAO_FALLBACK = 0.1

ROTULOS_EMISSOR = ["Emitente", "Fornecedor"]
ROTULOS_DESTINATARIO = ["Destinatário", "Destinatario", "Cliente"]
ROTULOS_NUMERO_DOCUMENTO = [
    "Nr do documento",
    "Nr. do documento",
    "Nº do documento",
    "Número do documento",
    "Numero do documento",
    "Pedido de Compra Nº",
    "Pedido de Compra N.",
]
ROTULOS_DATA_EMISSAO = ["Data de Emissão", "Data de Emissao", "Emissão", "Emissao"]
ROTULOS_VENCIMENTO = ["Vencimento"]

TODOS_ROTULOS = (
    ROTULOS_EMISSOR + ROTULOS_DESTINATARIO + ROTULOS_NUMERO_DOCUMENTO + ROTULOS_DATA_EMISSAO + ROTULOS_VENCIMENTO
)

PALAVRAS_CHAVE_TIPO = [
    (["nota fiscal"], "nota_fiscal"),
    (["pedido de compra", "purchase order"], "pedido_compra"),
    (["pedido"], "pedido_compra"),
    (["relat"], "relatorio"),  # cobre "relatorio"/"relatório"
]


class GenericExtractor(ExtratorDocumento):
    tipo = "desconhecido"

    def pontuacao_deteccao(self, contexto: ContextoExtracao) -> float:
        return PONTUACAO_FALLBACK

    def extrair(self, contexto: ContextoExtracao) -> ResultadoExtracao:
        texto = contexto.texto
        linhas = contexto.linhas
        confiancas: dict[str, str] = {}

        numero_documento = self._extrair_numero_documento(linhas, texto, confiancas)

        data_emissao = comum.extrair_data(linhas, ROTULOS_DATA_EMISSAO, TODOS_ROTULOS)
        if data_emissao:
            confiancas["data_emissao"] = "alta"

        data_vencimento = comum.extrair_data(linhas, ROTULOS_VENCIMENTO, TODOS_ROTULOS)
        if data_vencimento:
            confiancas["data_vencimento"] = "alta"

        emissor = comum.extrair_entidade(linhas, ROTULOS_EMISSOR, TODOS_ROTULOS)
        if emissor:
            confiancas["emissor"] = "alta"

        destinatario = comum.extrair_entidade(linhas, ROTULOS_DESTINATARIO, TODOS_ROTULOS)
        if destinatario:
            confiancas["destinatario"] = "alta"

        valor_total_bruto = comum.valor_por_total_ou_ultimo(texto)
        valor_total = comum.para_numero(valor_total_bruto) if valor_total_bruto else None
        if valor_total is not None:
            confiancas["valor_total"] = "baixa"  # sempre fallback generico, nunca rotulo dedicado

        documento = DocumentoExtraido(
            tipo_documento=self._detectar_tipo(texto.lower()),
            numero_documento=numero_documento,
            data_emissao=data_emissao,
            data_vencimento=data_vencimento,
            emissor=emissor,
            destinatario=destinatario,
            valor_total=valor_total,
            itens=[],
            campos_adicionais=[],
        )
        return ResultadoExtracao(documento=documento, confiancas=confiancas)

    def _detectar_tipo(self, texto_lower: str) -> str:
        for palavras, tipo in PALAVRAS_CHAVE_TIPO:
            if any(palavra in texto_lower for palavra in palavras):
                return tipo
        return "desconhecido"

    def _extrair_numero_documento(
        self, linhas: list[str], texto: str, confiancas: dict[str, str]
    ) -> str | None:
        encontrado = comum.localizar_rotulo(
            linhas, ROTULOS_NUMERO_DOCUMENTO, TODOS_ROTULOS, validador=comum.parece_numero
        )
        if encontrado:
            _, bruto = encontrado
            partes = bruto.split()
            valor = partes[0].strip(".:-") if partes else None
            if valor:
                confiancas["numero_documento"] = "alta"
                return valor

        m = comum.NUM_GENERICO_RE.search(texto)
        if m:
            confiancas["numero_documento"] = "baixa"
            return m.group(1)
        return None
