"""Extracao heuristica por regex, sem depender de LLM.

Modo "basico": sempre disponivel, nao precisa de chave de API. Bom para
capturar datas, valores monetarios, numero do documento e emissor/
destinatario quando o texto do PDF traz rotulos razoavelmente previsiveis
(nota fiscal, pedido de compra, boleto bancario). Nao tenta reconstruir a
tabela de itens -- texto corrido de PDF sem estrutura nao da pra confiar
nisso via regex. `itens` fica vazio; o modo IA existe justamente para
cobrir esse caso melhor.

Estrategia: em vez de so procurar padroes soltos no texto inteiro, primeiro
tentamos achar um ROTULO conhecido ("Beneficiário:", "Sacado:", "Nosso
Número:", ...) e usar o texto logo depois dele -- e mais confiavel do que
regex solto, que tende a pegar pedacos de outros numeros (linha digitavel,
CNPJ, etc). Regex solto continua existindo como fallback quando nenhum
rotulo e encontrado.
"""
import re
from typing import Optional

from app.schemas import CampoAdicional, DocumentoExtraido

DATA_RE = re.compile(r"\b(\d{2}[/-]\d{2}[/-]\d{4})\b")
VALOR_RE = re.compile(r"R\$\s*([\d.]+,\d{2})")
CNPJ_RE = re.compile(r"\b\d{2}\.\d{3}\.\d{3}/\d{4}-\d{2}\b")
CPF_RE = re.compile(r"\b\d{3}\.\d{3}\.\d{3}-\d{2}\b")

# Formato padrao da linha digitavel de boleto bancario: grupos de
# 5.5.5.6.5.6.1.14 digitos (47 no total), separados por ponto ou espaco.
LINHA_DIGITAVEL_RE = re.compile(
    r"\d{5}[.\s]\d{5}[.\s]\d{5}[.\s]\d{6}[.\s]\d{5}[.\s]\d{6}[.\s]\d[.\s]\d{14}"
)

# Fallback quando nenhum rotulo de numero de documento e encontrado.
# Exige pelo menos 4 digitos para nao pegar numeros soltos curtos (ex: "02").
NUM_DOCUMENTO_RE = re.compile(
    r"(?:N[º°o.]{1,3}|N[uú]mero)\s*[:\-]?\s*(\d{4,}[\w\-./]*)", re.IGNORECASE
)

ROTULOS_EMISSOR = ["Beneficiário", "Beneficiario", "Cedente", "Emitente", "Fornecedor"]
ROTULOS_DESTINATARIO = ["Sacado", "Pagador", "Destinatário", "Destinatario", "Cliente"]
ROTULOS_NUMERO_DOCUMENTO = [
    "Nosso Número",
    "Nosso Numero",
    "Número do documento",
    "Numero do documento",
    "Pedido de Compra Nº",
    "Pedido de Compra N.",
    "Número",
    "Numero",
]
ROTULOS_DATA_EMISSAO = ["Data de Emissão", "Data de Emissao", "Emissão", "Emissao"]
ROTULOS_VENCIMENTO = ["Vencimento"]

TODOS_ROTULOS = (
    ROTULOS_EMISSOR
    + ROTULOS_DESTINATARIO
    + ROTULOS_NUMERO_DOCUMENTO
    + ROTULOS_DATA_EMISSAO
    + ROTULOS_VENCIMENTO
)

# Palavras-chave para classificar o tipo de documento. Ordem importa: do
# mais especifico para o mais generico, porque o primeiro grupo que bater
# ganha.
PALAVRAS_CHAVE_TIPO = [
    (["danfe", "nf-e", "nfe", "chave de acesso", "cfop"], "nota_fiscal"),
    (["nota fiscal"], "nota_fiscal"),
    (
        [
            "linha digitavel",
            "linha digitável",
            "cedente",
            "sacado",
            "beneficiario",
            "beneficiário",
            "pagador",
            "nosso numero",
            "nosso número",
        ],
        "boleto",
    ),
    (["pedido de compra", "purchase order"], "pedido_compra"),
    (["pedido"], "pedido_compra"),
    (["relat"], "relatorio"),  # cobre "relatorio"/"relatório"
]


def para_numero(bruto: str) -> float | str:
    """Converte um valor no formato BR ("1.234,56") para float; devolve a
    string original quando a conversao falha."""
    normalizado = bruto.strip().replace(".", "").replace(",", ".")
    try:
        return float(normalizado)
    except ValueError:
        return bruto.strip()


def _contar_digitos(s: str) -> int:
    return sum(ch.isdigit() for ch in s)


def _detectar_tipo_documento(texto_lower: str) -> str:
    for palavras, tipo in PALAVRAS_CHAVE_TIPO:
        if any(palavra in texto_lower for palavra in palavras):
            return tipo
    return "desconhecido"


def _valor_total(texto: str) -> Optional[str]:
    """Prioriza uma linha que mencione "total"; senao usa a ultima
    ocorrencia de valor monetario no texto."""
    for linha in texto.splitlines():
        if "total" in linha.lower():
            m = VALOR_RE.search(linha)
            if m:
                return m.group(1)
    ocorrencias = VALOR_RE.findall(texto)
    return ocorrencias[-1] if ocorrencias else None


def _linha_digitavel(texto: str) -> Optional[str]:
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


def _remainder_apos_rotulo(linha: str, rotulo: str) -> Optional[str]:
    m = re.search(re.escape(rotulo) + r"\s*:?\s*(.*)", linha, re.IGNORECASE)
    return m.group(1).strip() if m else None


def _paradas_excluindo(rotulos_atual: list[str]) -> list[str]:
    """Palavras de corte para truncar um valor capturado: todos os rotulos
    conhecidos + CNPJ/CPF, EXCETO os da propria categoria que esta sendo
    buscada -- senao um nome que comeca com a palavra do proprio rotulo
    (ex: destinatario "Cliente Exemplo S.A.") seria cortado para vazio."""
    atuais_lower = {r.lower() for r in rotulos_atual}
    return ["CNPJ", "CPF"] + [r for r in TODOS_ROTULOS if r.lower() not in atuais_lower]


def _truncar_em_proximo_rotulo(valor: str, paradas: list[str]) -> str:
    """Corta o valor capturado assim que outro rotulo conhecido (ou CNPJ/
    CPF) aparece na mesma linha -- evita grudar dois campos juntos quando
    estao lado a lado (ex: "Cedente: Empresa XYZ  CNPJ: 12.345...")."""
    corte = len(valor)
    valor_lower = valor.lower()
    for parada in paradas:
        idx = valor_lower.find(parada.lower())
        if idx != -1 and idx < corte:
            corte = idx
    return valor[:corte].strip(" -:\t")


def _parece_rotulo(linha: str) -> bool:
    normalizado = linha.strip().rstrip(":").lower()
    return any(normalizado == rotulo.lower() for rotulo in TODOS_ROTULOS)


def _parece_nome(valor: str) -> bool:
    """Rejeita candidatos que sao so digitos/pontuacao/espaco -- carimbos de
    data/hora, numeros soltos, etc. Um nome de pessoa ou empresa sempre tem
    pelo menos uma letra."""
    return any(ch.isalpha() for ch in valor)


def _sempre_valido(_valor: str) -> bool:
    return True


def _localizar_rotulo(
    linhas: list[str],
    rotulos: list[str],
    validador=_sempre_valido,
) -> Optional[tuple[int, str]]:
    """Procura por qualquer um dos rotulos no documento inteiro.

    A ordem de busca e por PRIORIDADE do rotulo (a ordem da lista), nao por
    posicao no documento: primeiro tenta achar o rotulo mais especifico
    (ex: "Nosso Número") em qualquer linha; so tenta o proximo rotulo da
    lista (mais generico, ex: "Número") se o mais especifico nao aparecer
    em lugar nenhum. Isso evita que um rotulo generico que aparece mais
    cedo na pagina (ex: "Número do Banco") vença um rotulo especifico e
    confiavel que aparece mais tarde.

    Se o resto da linha apos o rotulo estiver vazio (ou nao passar no
    `validador`), tenta a linha seguinte -- layout comum em boletos
    extraidos de PDF, onde rotulo e valor saem em linhas separadas.
    """
    paradas = _paradas_excluindo(rotulos)
    for rotulo in rotulos:
        for i, linha in enumerate(linhas):
            if rotulo.lower() not in linha.lower():
                continue

            resto = _remainder_apos_rotulo(linha, rotulo)
            valor = _truncar_em_proximo_rotulo(resto, paradas) if resto else ""
            if valor and validador(valor):
                return i, valor

            if i + 1 < len(linhas):
                proxima = linhas[i + 1].strip()
                if proxima and not _parece_rotulo(proxima):
                    valor = _truncar_em_proximo_rotulo(proxima, paradas)
                    if valor and validador(valor):
                        return i, valor
    return None


def _documento_fiscal_proximo(linhas: list[str], indice: int, janela: int = 3) -> Optional[str]:
    """Procura um CNPJ ou CPF nas linhas ao redor de onde um rotulo (ex:
    "Cedente") foi encontrado, para associar ao nome extraido."""
    trecho = " ".join(linhas[indice : indice + janela])
    m = CNPJ_RE.search(trecho)
    if m:
        return f"CNPJ {m.group(0)}"
    m = CPF_RE.search(trecho)
    if m:
        return f"CPF {m.group(0)}"
    return None


def _extrair_entidade(linhas: list[str], rotulos: list[str]) -> Optional[str]:
    encontrado = _localizar_rotulo(linhas, rotulos, validador=_parece_nome)
    if not encontrado:
        return None
    indice, nome = encontrado

    documento_fiscal = _documento_fiscal_proximo(linhas, indice)
    if documento_fiscal and documento_fiscal.split()[-1] not in nome:
        return f"{nome} ({documento_fiscal})"
    return nome


def _extrair_data(linhas: list[str], rotulos: list[str]) -> Optional[str]:
    encontrado = _localizar_rotulo(linhas, rotulos)
    if not encontrado:
        return None
    _, bruto = encontrado
    m = DATA_RE.search(bruto)
    return m.group(1) if m else None


def _extrair_numero_documento(linhas: list[str], texto: str) -> Optional[str]:
    # Vindo de um rotulo conhecido (ex: "Nosso Número:"), confiamos no
    # valor mesmo com poucos digitos -- o rotulo ja e a garantia de que
    # nao e um numero solto. Sem rotulo, exigimos pelo menos 4 digitos
    # (e o que evita pegar pedacos curtos soltos, tipo "02", em qualquer
    # lugar do texto).
    encontrado = _localizar_rotulo(linhas, ROTULOS_NUMERO_DOCUMENTO)
    if encontrado:
        _, bruto = encontrado
        partes = bruto.split()
        token = partes[0].strip(".:-") if partes else ""
        if _contar_digitos(token) >= 1:
            return token

    m = NUM_DOCUMENTO_RE.search(texto)
    return m.group(1) if m else None


def extrair(texto: str) -> DocumentoExtraido:
    linhas = texto.splitlines()
    texto_lower = texto.lower()

    valor_total_bruto = _valor_total(texto)

    campos_adicionais: list[CampoAdicional] = []

    linha_dig = _linha_digitavel(texto)
    if linha_dig:
        campos_adicionais.append(CampoAdicional(campo="Linha digitável", valor=linha_dig))

    vencimento = _extrair_data(linhas, ROTULOS_VENCIMENTO)
    if vencimento:
        campos_adicionais.append(CampoAdicional(campo="Vencimento", valor=vencimento))

    return DocumentoExtraido(
        tipo_documento=_detectar_tipo_documento(texto_lower),
        numero_documento=_extrair_numero_documento(linhas, texto),
        data_emissao=_extrair_data(linhas, ROTULOS_DATA_EMISSAO),
        emissor=_extrair_entidade(linhas, ROTULOS_EMISSOR),
        destinatario=_extrair_entidade(linhas, ROTULOS_DESTINATARIO),
        valor_total=para_numero(valor_total_bruto) if valor_total_bruto else None,
        itens=[],
        campos_adicionais=campos_adicionais,
    )
