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
import logging
import re
from typing import Optional

from app.schemas import CampoAdicional, DocumentoExtraido

logger = logging.getLogger(__name__)

# OCR confunde com frequencia certas letras com digitos parecidos
# (O/o<->0, I/i<->1, S/s<->5). Campos que devem ser puramente numericos
# (CNPJ, CPF, valores) usam essa classe no lugar de \d para nao perder o
# campo inteiro so porque um caractere saiu errado do Tesseract; o valor
# capturado passa por `_corrigir_confusao_ocr` antes de ser usado.
_DIG = "0-9OoIiSs"

DATA_RE = re.compile(r"\b(\d{2}[/-]\d{2}[/-]\d{4})\b")
VALOR_RE = re.compile(rf"R\$\s*([{_DIG}]{{1,3}}(?:\.[{_DIG}]{{3}})*,[{_DIG}]{{2}})")
# Valor monetario BR sem exigir prefixo "R$" -- boletos costumam mostrar
# "Valor Documento 1.000,00" sem o "R$" na frente.
VALOR_NUM_RE = re.compile(rf"([{_DIG}]{{1,3}}(?:\.[{_DIG}]{{3}})*,[{_DIG}]{{2}})")
CNPJ_RE = re.compile(rf"\b[{_DIG}]{{2}}\.[{_DIG}]{{3}}\.[{_DIG}]{{3}}/[{_DIG}]{{4}}-[{_DIG}]{{2}}\b")
CPF_RE = re.compile(rf"\b[{_DIG}]{{3}}\.[{_DIG}]{{3}}\.[{_DIG}]{{3}}-[{_DIG}]{{2}}\b")
# "Nosso Número" as vezes vem como "02 / 10200000001-9" -- os digitos
# antes da barra sao so um prefixo/carteira, o numero real e depois dela.
NUMERO_COM_BARRA_RE = re.compile(r"(\d[\d.\-]*)\s*/\s*(\d[\d.\-]*)")

# Formato padrao da linha digitavel de boleto bancario: grupos de
# 5.5.5.6.5.6.1.14 digitos (47 no total), separados por ponto ou espaco.
LINHA_DIGITAVEL_RE = re.compile(
    r"\d{5}[.\s]\d{5}[.\s]\d{5}[.\s]\d{6}[.\s]\d{5}[.\s]\d{6}[.\s]\d[.\s]\d{14}"
)

# Fallback quando nenhum rotulo de numero de documento e encontrado.
# Exige pelo menos 4 digitos para nao pegar numeros soltos curtos (ex: "02").
# Propositalmente so as formas abreviadas (N./Nº/N°) -- a palavra por
# extenso "Número" tambem aparece dentro de "Nosso Número", e usa-la aqui
# recontaminava numero_documento com o valor do Nosso Número.
NUM_DOCUMENTO_RE = re.compile(r"N[º°o.]{1,3}\s*[:\-]?\s*(\d{4,}[\w\-./]*)", re.IGNORECASE)

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
# Propositalmente SEM "Número"/"Numero" soltos aqui: sao ambiguos demais
# (batem em "Número do Banco", "Nosso Número" e qualquer outro campo que
# mencione a palavra) e ja causaram contaminacao cruzada com Nosso
# Número. O fallback por regex (NUM_DOCUMENTO_RE) cobre o resto do texto
# quando nenhum rotulo especifico aparece.
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


_TABELA_CONFUSAO_OCR = str.maketrans({"O": "0", "o": "0", "I": "1", "i": "1", "S": "5", "s": "5"})


def _corrigir_confusao_ocr(texto: str) -> str:
    """Corrige as confusoes mais comuns do OCR em campos que ja sabemos
    que devem ser puramente numericos (CNPJ, CPF, valores): O/o -> 0,
    I/i -> 1, S/s -> 5. So chamar sobre um trecho ja identificado como
    numerico por regex -- aplicar isso no texto inteiro destruiria
    palavras normais."""
    return texto.translate(_TABELA_CONFUSAO_OCR)


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


def _extrair_valor_rotulo(texto: str, rotulos: list[str]) -> Optional[str]:
    """Procura um rotulo de valor monetario (ex: "Desconto", "Valor a
    Pagar") e o numero BR logo depois dele na mesma linha -- com ou sem
    "R$"/"=" no meio (boletos variam: "Valor Documento 1.000,00" vs
    "Desconto = R$ 100,00")."""
    for linha in texto.splitlines():
        linha_lower = linha.lower()
        for rotulo in rotulos:
            idx = linha_lower.find(rotulo.lower())
            if idx == -1:
                continue
            resto = linha[idx + len(rotulo) :]
            m = VALOR_NUM_RE.search(resto)
            if m:
                return _corrigir_confusao_ocr(m.group(1))
    return None


def _valor_total(texto: str) -> Optional[str]:
    """Em boleto, "Valor a Pagar" e o valor final confiavel. Sem esse
    rotulo, prioriza uma linha que mencione "total"; senao usa a ultima
    ocorrencia de valor monetario no texto."""
    valor_a_pagar = _extrair_valor_rotulo(texto, ROTULOS_VALOR_A_PAGAR)
    if valor_a_pagar:
        return valor_a_pagar

    for linha in texto.splitlines():
        if "total" in linha.lower():
            m = VALOR_RE.search(linha)
            if m:
                return _corrigir_confusao_ocr(m.group(1))
    ocorrencias = VALOR_RE.findall(texto)
    return _corrigir_confusao_ocr(ocorrencias[-1]) if ocorrencias else None


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


# Trechos de cabecalho/rodape de boleto que tem letras (entao passariam
# num teste "so tem letra") mas claramente nao sao um nome de pessoa/
# empresa. "Recibo do Sacado" e um TITULO DE SECAO que contem a palavra
# "Sacado" -- o rotulo bate nele antes de bater no campo de verdade, e o
# rodape "Pág: 1 de 1" que costuma vir logo depois tem letra o suficiente
# pra passar despercebido sem esse bloqueio.
MARCADORES_CABECALHO_RODAPE = [
    "pág",
    "pag:",
    "página",
    "pagina",
    "recibo do sacado",
    "ficha de compensação",
    "ficha de compensacao",
]


def _parece_nome(valor: str) -> bool:
    """Rejeita candidatos que sao so digitos/pontuacao/espaco (carimbos de
    data/hora, numeros soltos) ou que batem em marcadores conhecidos de
    cabecalho/rodape de boleto. Um nome de pessoa ou empresa sempre tem
    pelo menos uma letra E nao e um desses marcadores."""
    if not any(ch.isalpha() for ch in valor):
        return False
    valor_lower = valor.lower()
    return not any(marcador in valor_lower for marcador in MARCADORES_CABECALHO_RODAPE)


def _parece_data(valor: str) -> bool:
    """Rejeita candidatos onde nao ha nenhuma data reconhecivel -- evita
    que uma mencao solta ao rotulo (ex: um aviso de "vencimento" no rodape,
    sem data por perto) trave a busca antes de achar o campo de verdade."""
    return bool(DATA_RE.search(valor))


def _sempre_valido(_valor: str) -> bool:
    return True


def _limpar_prefixo_rotulos(valor: str, rotulos: list[str]) -> str:
    """Boletos costumam mostrar um rotulo combinado tipo "Sacado/Pagador"
    grudado no nome sem separador: depois de remover "Sacado", sobra
    "/PagadorFulano de Tal". So mexe quando ha uma BARRA logo no
    inicio (sinal inequivoco de rotulo combinado) -- sem essa barra,
    qualquer palavra do valor que por acaso bata com outro rotulo da
    mesma categoria fica intacta (ex: emissor "Fornecedor Exemplo Ltda"
    nao pode perder o "Fornecedor" so por ele tambem ser um rotulo)."""
    valor = valor.lstrip(" -:\t")
    if not valor.startswith("/"):
        return valor
    resto = valor[1:].lstrip(" -:\t")
    for rotulo in rotulos:
        if resto.lower().startswith(rotulo.lower()):
            return resto[len(rotulo) :].lstrip(" /-:\t")
    return resto


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
            if resto:
                resto = _limpar_prefixo_rotulos(resto, rotulos)
            valor = _truncar_em_proximo_rotulo(resto, paradas) if resto else ""
            if valor and validador(valor):
                logger.debug(
                    "ACEITO (mesma linha) rotulo=%r linha=%d texto_linha=%r -> valor=%r",
                    rotulo, i, linha, valor,
                )
                return i, valor
            if valor:
                logger.debug(
                    "rejeitado (mesma linha) rotulo=%r linha=%d texto_linha=%r -> valor=%r",
                    rotulo, i, linha, valor,
                )

            if i + 1 < len(linhas):
                proxima_bruta = linhas[i + 1].strip()
                if proxima_bruta and not _parece_rotulo(proxima_bruta):
                    proxima = _limpar_prefixo_rotulos(proxima_bruta, rotulos)
                    valor = _truncar_em_proximo_rotulo(proxima, paradas) if proxima else ""
                    if valor and validador(valor):
                        logger.debug(
                            "ACEITO (linha seguinte) rotulo=%r linha=%d texto_linha=%r "
                            "proxima_linha=%r -> valor=%r",
                            rotulo, i, linha, proxima_bruta, valor,
                        )
                        return i, valor
                    if valor:
                        logger.debug(
                            "rejeitado (linha seguinte) rotulo=%r linha=%d "
                            "proxima_linha=%r -> valor=%r",
                            rotulo, i, proxima_bruta, valor,
                        )
    logger.debug("NENHUM candidato encontrado para rotulos=%r", rotulos)
    return None


def _documento_fiscal_proximo(linhas: list[str], indice: int, janela: int = 3) -> Optional[str]:
    """Procura um CNPJ ou CPF nas linhas ao redor de onde um rotulo (ex:
    "Cedente") foi encontrado, para associar ao nome extraido."""
    trecho = " ".join(linhas[indice : indice + janela])
    m = CNPJ_RE.search(trecho)
    if m:
        return f"CNPJ {_corrigir_confusao_ocr(m.group(0))}"
    m = CPF_RE.search(trecho)
    if m:
        return f"CPF {_corrigir_confusao_ocr(m.group(0))}"
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
    encontrado = _localizar_rotulo(linhas, rotulos, validador=_parece_data)
    if not encontrado:
        return None
    _, bruto = encontrado
    m = DATA_RE.search(bruto)
    return m.group(1) if m else None


def _parece_numero(valor: str) -> bool:
    primeiro_token = valor.split()[0] if valor.split() else ""
    return _contar_digitos(primeiro_token) >= 1


def _extrair_valor_numerico_por_rotulo(linhas: list[str], rotulos: list[str]) -> Optional[str]:
    """Busca um rotulo e devolve o numero associado. Confiamos no valor
    mesmo com poucos digitos -- o rotulo ja e a garantia de que nao e um
    numero solto (o validador so confirma que ha ALGUM digito; se um
    rotulo bater em algo sem digito nenhum, a busca continua para o
    proximo rotulo da lista em vez de desistir). Trata o formato
    "XX / YYYYY" (prefixo de carteira / numero real, comum em "Nosso
    Número"), usando a parte depois da barra."""
    encontrado = _localizar_rotulo(linhas, rotulos, validador=_parece_numero)
    if not encontrado:
        return None
    _, bruto = encontrado
    m_barra = NUMERO_COM_BARRA_RE.match(bruto)
    if m_barra:
        return m_barra.group(2)
    partes = bruto.split()
    return partes[0].strip(".:-") if partes else None


def _extrair_numero_documento(linhas: list[str], texto: str) -> Optional[str]:
    # Sem rotulo conhecido, exige pelo menos 4 digitos no fallback por
    # regex solto -- e o que evita pegar pedacos curtos soltos (tipo "02")
    # em qualquer lugar do texto.
    valor = _extrair_valor_numerico_por_rotulo(linhas, ROTULOS_NUMERO_DOCUMENTO)
    if valor:
        return valor

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

    nosso_numero = _extrair_valor_numerico_por_rotulo(linhas, ROTULOS_NOSSO_NUMERO)
    if nosso_numero:
        campos_adicionais.append(CampoAdicional(campo="Nosso Número", valor=nosso_numero))

    valor_documento = _extrair_valor_rotulo(texto, ROTULOS_VALOR_DOCUMENTO)
    if valor_documento:
        campos_adicionais.append(CampoAdicional(campo="Valor do Documento", valor=valor_documento))

    desconto = _extrair_valor_rotulo(texto, ROTULOS_DESCONTO)
    if desconto:
        campos_adicionais.append(CampoAdicional(campo="Desconto", valor=desconto))

    valor_a_pagar = _extrair_valor_rotulo(texto, ROTULOS_VALOR_A_PAGAR)
    if valor_a_pagar:
        campos_adicionais.append(CampoAdicional(campo="Valor a Pagar", valor=valor_a_pagar))

    parcela = _localizar_rotulo(linhas, ROTULOS_PARCELA)
    if parcela:
        campos_adicionais.append(CampoAdicional(campo="Parcela", valor=parcela[1]))

    agencia_codigo = _localizar_rotulo(linhas, ROTULOS_AGENCIA_CODIGO)
    if agencia_codigo:
        campos_adicionais.append(
            CampoAdicional(campo="Agência/Código Beneficiário", valor=agencia_codigo[1])
        )

    return DocumentoExtraido(
        tipo_documento=_detectar_tipo_documento(texto_lower),
        numero_documento=_extrair_numero_documento(linhas, texto),
        data_emissao=_extrair_data(linhas, ROTULOS_DATA_EMISSAO),
        data_vencimento=_extrair_data(linhas, ROTULOS_VENCIMENTO),
        emissor=_extrair_entidade(linhas, ROTULOS_EMISSOR),
        destinatario=_extrair_entidade(linhas, ROTULOS_DESTINATARIO),
        valor_total=para_numero(valor_total_bruto) if valor_total_bruto else None,
        itens=[],
        campos_adicionais=campos_adicionais,
    )
