"""BoletoExtractor: extracao de campos de boleto bancario.

Migrado de app/basic_extractor.py -- comportamento identico ao que existia
antes da refatoracao pra padrao Strategy. Usa os helpers genericos de
`comum.py`, fornecendo o proprio universo de rotulos (Cedente, Sacado,
Nosso Número, ...).
"""
import re

from app.extractors import comum
from app.extractors.base import ContextoExtracao, ExtratorDocumento, ResultadoExtracao
from app.schemas import CampoAdicional, DocumentoExtraido

# "Nosso Número" as vezes vem como "02 / 10200000001-9" -- os digitos
# antes da barra sao so um prefixo/carteira, o numero real e depois dela.
NUMERO_COM_BARRA_RE = re.compile(r"(\d[\d.\-]*)\s*/\s*(\d[\d.\-]*)")

# Formato padrao da linha digitavel de boleto bancario: grupos de
# 5.5.5.6.5.6.1.14 digitos (47 no total), separados por ponto ou espaco.
LINHA_DIGITAVEL_RE = re.compile(
    r"\d{5}[.\s]\d{5}[.\s]\d{5}[.\s]\d{6}[.\s]\d{5}[.\s]\d{6}[.\s]\d[.\s]\d{14}"
)

ROTULOS_EMISSOR = ["Beneficiário", "Beneficiario", "Cedente", "Emitente", "Fornecedor"]
ROTULOS_DESTINATARIO = ["Sacado", "Pagador", "Destinatário", "Destinatario", "Cliente"]
ROTULOS_NUMERO_DOCUMENTO = [
    "Nr do documento",
    "Nr. do documento",
    "Nº do documento",
    "Número do documento",
    "Numero do documento",
    "Pedido de Compra Nº",
    "Pedido de Compra N.",
]
# Propositalmente SEM "Número"/"Numero" soltos: sao ambiguos demais (batem
# em "Número do Banco", "Nosso Número" e qualquer outro campo que mencione
# a palavra) e ja causaram contaminacao cruzada com Nosso Número.
# "Nosso Número" e um identificador BANCARIO (atribuido pelo banco pra
# controle interno), diferente do numero do documento em si -- por isso
# tem lista propria e vai pra campos_adicionais, nao para numero_documento.
ROTULOS_NOSSO_NUMERO = ["Nosso Número", "Nosso Numero"]
ROTULOS_DATA_EMISSAO = ["Data de Emissão", "Data de Emissao", "Emissão", "Emissao"]
ROTULOS_VENCIMENTO = ["Vencimento"]
ROTULOS_VALOR_DOCUMENTO = ["Valor Documento", "Valor do Documento"]
ROTULOS_DESCONTO = ["Desconto"]
ROTULOS_VALOR_A_PAGAR = ["Valor a Pagar", "Valor Pagar"]
ROTULOS_PARCELA = ["Parcela"]
ROTULOS_AGENCIA_CODIGO = [
    "Agência/Código Beneficiário",
    "Agencia/Codigo Beneficiario",
    "Agência/Cód. Beneficiário",
    "Agencia/Cod. Beneficiario",
    "Ag/Código Beneficiário",
    "Ag/Codigo Beneficiario",
]

TODOS_ROTULOS = (
    ROTULOS_EMISSOR
    + ROTULOS_DESTINATARIO
    + ROTULOS_NUMERO_DOCUMENTO
    + ROTULOS_NOSSO_NUMERO
    + ROTULOS_DATA_EMISSAO
    + ROTULOS_VENCIMENTO
    + ROTULOS_VALOR_DOCUMENTO
    + ROTULOS_DESCONTO
    + ROTULOS_VALOR_A_PAGAR
    + ROTULOS_PARCELA
    + ROTULOS_AGENCIA_CODIGO
)

PALAVRAS_CHAVE_DETECCAO = [
    "linha digitavel",
    "linha digitável",
    "cedente",
    "sacado",
    "beneficiario",
    "beneficiário",
    "pagador",
    "nosso numero",
    "nosso número",
]


def _linha_digitavel(texto: str) -> str | None:
    m = LINHA_DIGITAVEL_RE.search(texto)
    if m:
        return m.group(0)

    # Fallback: layouts de boleto variam (convenio tem 4 blocos de ~12
    # digitos, por exemplo). Aceita qualquer linha composta so por digitos/
    # pontos/espacos/tracos que tenha entre 44 e 48 digitos no total.
    for linha in texto.splitlines():
        candidata = linha.strip()
        if not candidata or not re.fullmatch(r"[\d.\s-]+", candidata):
            continue
        apenas_digitos = re.sub(r"[^\d]", "", candidata)
        if 44 <= len(apenas_digitos) <= 48:
            return candidata
    return None


def _extrair_valor_numerico_por_rotulo(linhas: list[str], rotulos: list[str]) -> str | None:
    """Busca um rotulo (ex: "Nosso Número") e devolve o numero associado.
    Confiamos no valor mesmo com poucos digitos -- o rotulo ja e a
    garantia de que nao e um numero solto. Trata o formato "XX / YYYYY"
    (prefixo de carteira / numero real, comum em "Nosso Número"), usando a
    parte depois da barra."""
    encontrado = comum.localizar_rotulo(
        linhas, rotulos, TODOS_ROTULOS, validador=comum.parece_numero
    )
    if not encontrado:
        return None
    _, bruto = encontrado
    m_barra = NUMERO_COM_BARRA_RE.match(bruto)
    if m_barra:
        return m_barra.group(2)
    partes = bruto.split()
    return partes[0].strip(".:-") if partes else None


class BoletoExtractor(ExtratorDocumento):
    tipo = "boleto"

    def pontuacao_deteccao(self, contexto: ContextoExtracao) -> float:
        texto_lower = contexto.texto.lower()
        return 0.9 if any(p in texto_lower for p in PALAVRAS_CHAVE_DETECCAO) else 0.0

    def extrair(self, contexto: ContextoExtracao) -> ResultadoExtracao:
        texto = contexto.texto
        linhas = contexto.linhas
        confiancas: dict[str, str] = {}
        campos_adicionais: list[CampoAdicional] = []

        linha_dig = _linha_digitavel(texto)
        if linha_dig:
            campos_adicionais.append(CampoAdicional(campo="Linha digitável", valor=linha_dig))
            confiancas["Linha digitável"] = "alta"

        nosso_numero = _extrair_valor_numerico_por_rotulo(linhas, ROTULOS_NOSSO_NUMERO)
        if nosso_numero:
            campos_adicionais.append(CampoAdicional(campo="Nosso Número", valor=nosso_numero))
            confiancas["Nosso Número"] = "alta"

        valor_documento = comum.extrair_valor_rotulo(texto, ROTULOS_VALOR_DOCUMENTO)
        if valor_documento:
            campos_adicionais.append(
                CampoAdicional(campo="Valor do Documento", valor=valor_documento)
            )
            confiancas["Valor do Documento"] = "alta"

        desconto = comum.extrair_valor_rotulo(texto, ROTULOS_DESCONTO)
        if desconto:
            campos_adicionais.append(CampoAdicional(campo="Desconto", valor=desconto))
            confiancas["Desconto"] = "alta"

        valor_a_pagar = comum.extrair_valor_rotulo(texto, ROTULOS_VALOR_A_PAGAR)
        if valor_a_pagar:
            campos_adicionais.append(CampoAdicional(campo="Valor a Pagar", valor=valor_a_pagar))
            confiancas["Valor a Pagar"] = "alta"

        parcela = comum.localizar_rotulo(linhas, ROTULOS_PARCELA, TODOS_ROTULOS)
        if parcela:
            campos_adicionais.append(CampoAdicional(campo="Parcela", valor=parcela[1]))
            confiancas["Parcela"] = "alta"

        agencia_codigo = comum.localizar_rotulo(linhas, ROTULOS_AGENCIA_CODIGO, TODOS_ROTULOS)
        if agencia_codigo:
            campos_adicionais.append(
                CampoAdicional(campo="Agência/Código Beneficiário", valor=agencia_codigo[1])
            )
            confiancas["Agência/Código Beneficiário"] = "alta"

        numero_documento, confianca_numero = self._extrair_numero_documento(linhas, texto)
        if numero_documento:
            confiancas["numero_documento"] = confianca_numero

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

        valor_total_bruto, confianca_valor = self._valor_total(texto)
        valor_total = comum.para_numero(valor_total_bruto) if valor_total_bruto else None
        if valor_total is not None:
            confiancas["valor_total"] = confianca_valor

        documento = DocumentoExtraido(
            tipo_documento="boleto",
            numero_documento=numero_documento,
            data_emissao=data_emissao,
            data_vencimento=data_vencimento,
            emissor=emissor,
            destinatario=destinatario,
            valor_total=valor_total,
            itens=[],
            campos_adicionais=campos_adicionais,
        )
        return ResultadoExtracao(documento=documento, confiancas=confiancas)

    def _extrair_numero_documento(self, linhas: list[str], texto: str) -> tuple[str | None, str]:
        # Sem rotulo conhecido, exige pelo menos 4 digitos no fallback por
        # regex solto -- e o que evita pegar pedacos curtos soltos (tipo
        # "02") em qualquer lugar do texto. Confianca alta quando vem de
        # rotulo explicito, baixa quando vem do fallback.
        encontrado = comum.localizar_rotulo(
            linhas, ROTULOS_NUMERO_DOCUMENTO, TODOS_ROTULOS, validador=comum.parece_numero
        )
        if encontrado:
            _, bruto = encontrado
            partes = bruto.split()
            valor = partes[0].strip(".:-") if partes else None
            if valor:
                return valor, "alta"

        m = comum.NUM_GENERICO_RE.search(texto)
        return (m.group(1), "baixa") if m else (None, "")

    def _valor_total(self, texto: str) -> tuple[str | None, str]:
        """Em boleto, "Valor a Pagar" e o valor final confiavel (confianca
        alta). Sem esse rotulo, cai no fallback generico (confianca
        baixa)."""
        valor_a_pagar = comum.extrair_valor_rotulo(texto, ROTULOS_VALOR_A_PAGAR)
        if valor_a_pagar:
            return valor_a_pagar, "alta"

        valor_fallback = comum.valor_por_total_ou_ultimo(texto)
        return (valor_fallback, "baixa") if valor_fallback else (None, "")
