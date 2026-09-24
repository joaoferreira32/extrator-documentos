"""Helpers de extracao compartilhados por qualquer ExtratorDocumento:
regex de CNPJ/CPF/valor/data, busca por rotulo com prioridade, associacao
de CNPJ/CPF a nome mais proximo, etc.

Migrado de app/basic_extractor.py (comportamento identico) como parte da
refatoracao pra padrao Strategy -- a diferenca e que funcoes que antes liam
uma lista global de rotulos do modulo (`TODOS_ROTULOS`) agora recebem essa
lista como parametro, porque cada extrator (boleto, DANFE, ...) tem seu
proprio universo de rotulos.
"""
import logging
import re
from typing import Optional

logger = logging.getLogger(__name__)

# OCR confunde com frequencia certas letras com digitos parecidos
# (O/o<->0, I/i<->1, S/s<->5). Campos que devem ser puramente numericos
# (CNPJ, CPF, valores) usam essa classe no lugar de \d para nao perder o
# campo inteiro so porque um caractere saiu errado do Tesseract; o valor
# capturado passa por `corrigir_confusao_ocr` antes de ser usado.
_DIG = "0-9OoIiSs"

# Dia/mes com 1 ou 2 digitos -- confirmado num dump real de DANFE
# ("15/4/2026", mes sem zero a esquerda). Exigir \d{2} rejeitava datas
# validas com um digito so.
DATA_RE = re.compile(r"\b(\d{1,2}[/-]\d{1,2}[/-]\d{4})\b")
VALOR_RE = re.compile(rf"R\$\s*([{_DIG}]{{1,3}}(?:\.[{_DIG}]{{3}})*,[{_DIG}]{{2}})")
# Valor monetario BR sem exigir prefixo "R$" -- boletos/DANFE costumam
# mostrar o valor sem o "R$" na frente em varios campos.
VALOR_NUM_RE = re.compile(rf"([{_DIG}]{{1,3}}(?:\.[{_DIG}]{{3}})*,[{_DIG}]{{2}})")
CNPJ_RE = re.compile(rf"\b[{_DIG}]{{2}}\.[{_DIG}]{{3}}\.[{_DIG}]{{3}}/[{_DIG}]{{4}}-[{_DIG}]{{2}}\b")
CPF_RE = re.compile(rf"\b[{_DIG}]{{3}}\.[{_DIG}]{{3}}\.[{_DIG}]{{3}}-[{_DIG}]{{2}}\b")

# Fallback quando nenhum rotulo especifico de numero e encontrado. Exige
# pelo menos 4 digitos para nao pegar numeros soltos curtos (ex: "02").
# Propositalmente so as formas abreviadas (N./Nº/N°) -- a palavra por
# extenso "Número" tambem aparece dentro de "Nosso Número" e de outros
# campos que mencionem a palavra, e usa-la aqui contaminava um campo com o
# valor de outro.
NUM_GENERICO_RE = re.compile(r"N[º°o.]{1,3}\s*[:\-]?\s*(\d{4,}[\w\-./]*)", re.IGNORECASE)

# Trechos de cabecalho/rodape que tem letra (passariam num teste "so tem
# letra") mas claramente nao sao um nome de pessoa/empresa. Ex real: no
# boleto, "Recibo do Sacado" e um TITULO DE SECAO que contem a palavra
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

_TABELA_CONFUSAO_OCR = str.maketrans({"O": "0", "o": "0", "I": "1", "i": "1", "S": "5", "s": "5"})


def corrigir_confusao_ocr(texto: str) -> str:
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


def formatar_valor_br(valor: float) -> str:
    """Inverso de para_numero: 1234.5 -> "1.234,50". Usado quando um valor
    ja foi parseado como float (ex: de uma grade de totais) mas precisa
    virar texto de novo pra entrar em campos_adicionais (que exige string)."""
    return f"{valor:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")


def contar_digitos(s: str) -> int:
    return sum(ch.isdigit() for ch in s)


def extrair_valor_rotulo(
    texto: str, rotulos: list[str], aceitar_linha_seguinte: bool = False
) -> Optional[str]:
    """Procura um rotulo de valor monetario (ex: "Desconto", "Valor a
    Pagar", "Valor do ICMS") e o numero BR logo depois dele na mesma linha
    -- com ou sem "R$"/"=" no meio.

    `aceitar_linha_seguinte` (opt-in, default False): se o rotulo aparece
    mas NAO ha valor depois dele na mesma linha, aceita o primeiro valor
    monetario da linha SEGUINTE. E o layout de cabecalho de tabela: rotulo
    numa linha, valor na proxima -- ex. real do boleto:
        "Valor Documento (-) desconto (-) outras deduções ..."
        "1.000,00"
    So o boleto liga isso (pro "Valor do Documento"). Fica desligado por
    padrao porque numa DANFE toda linha vizinha de um rotulo tem valores de
    OUTROS campos -- ali a leitura e por posicao na grade de totais
    (`danfe.extrair_totais_grade`)."""
    linhas = texto.splitlines()
    for i, linha in enumerate(linhas):
        linha_lower = linha.lower()
        for rotulo in rotulos:
            idx = linha_lower.find(rotulo.lower())
            if idx == -1:
                continue
            resto = linha[idx + len(rotulo) :]
            m = VALOR_NUM_RE.search(resto)
            if m:
                return corrigir_confusao_ocr(m.group(1))
            if aceitar_linha_seguinte and i + 1 < len(linhas):
                m = VALOR_NUM_RE.search(linhas[i + 1])
                if m:
                    return corrigir_confusao_ocr(m.group(1))
    return None


def valor_por_total_ou_ultimo(texto: str) -> Optional[str]:
    """Fallback generico quando nenhum rotulo especifico de valor total e
    encontrado: prioriza uma linha que mencione "total"; senao usa a
    ultima ocorrencia de valor monetario no texto."""
    for linha in texto.splitlines():
        if "total" in linha.lower():
            m = VALOR_RE.search(linha)
            if m:
                return corrigir_confusao_ocr(m.group(1))
    ocorrencias = VALOR_RE.findall(texto)
    return corrigir_confusao_ocr(ocorrencias[-1]) if ocorrencias else None


def remainder_apos_rotulo(linha: str, rotulo: str) -> Optional[str]:
    m = re.search(re.escape(rotulo) + r"\s*:?\s*(.*)", linha, re.IGNORECASE)
    return m.group(1).strip() if m else None


def paradas_excluindo(todos_rotulos: list[str], rotulos_atual: list[str]) -> list[str]:
    """Palavras de corte para truncar um valor capturado: todos os rotulos
    conhecidos (do extrator chamador) + CNPJ/CPF, EXCETO os da propria
    categoria que esta sendo buscada -- senao um nome que comeca com a
    palavra do proprio rotulo (ex: destinatario "Cliente Exemplo S.A.")
    seria cortado para vazio."""
    atuais_lower = {r.lower() for r in rotulos_atual}
    return ["CNPJ", "CPF"] + [r for r in todos_rotulos if r.lower() not in atuais_lower]


def truncar_em_proximo_rotulo(valor: str, paradas: list[str]) -> str:
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


def parece_rotulo(linha: str, todos_rotulos: list[str]) -> bool:
    normalizado = linha.strip().rstrip(":").lower()
    return any(normalizado == rotulo.lower() for rotulo in todos_rotulos)


def parece_nome(valor: str) -> bool:
    """Rejeita candidatos que sao so digitos/pontuacao/espaco (carimbos de
    data/hora, numeros soltos) ou que batem em marcadores conhecidos de
    cabecalho/rodape. Um nome de pessoa ou empresa sempre tem pelo menos
    uma letra E nao e um desses marcadores."""
    if not any(ch.isalpha() for ch in valor):
        return False
    valor_lower = valor.lower()
    return not any(marcador in valor_lower for marcador in MARCADORES_CABECALHO_RODAPE)


def parece_data(valor: str) -> bool:
    """Rejeita candidatos onde nao ha nenhuma data reconhecivel -- evita
    que uma mencao solta ao rotulo (ex: um aviso no rodape, sem data por
    perto) trave a busca antes de achar o campo de verdade."""
    return bool(DATA_RE.search(valor))


def parece_numero(valor: str) -> bool:
    primeiro_token = valor.split()[0] if valor.split() else ""
    return contar_digitos(primeiro_token) >= 1


def sempre_valido(_valor: str) -> bool:
    return True


def limpar_prefixo_rotulos(valor: str, rotulos: list[str]) -> str:
    """Documentos costumam mostrar um rotulo combinado tipo
    "Sacado/Pagador" grudado no valor sem separador: depois de remover
    "Sacado", sobra "/PagadorFulano de Tal". So mexe quando ha uma
    BARRA logo no inicio (sinal inequivoco de rotulo combinado) -- sem
    essa barra, qualquer palavra do valor que por acaso bata com outro
    rotulo da mesma categoria fica intacta (ex: emissor "Fornecedor
    Exemplo Ltda" nao pode perder o "Fornecedor" so por ele tambem ser um
    rotulo)."""
    valor = valor.lstrip(" -:\t")
    if not valor.startswith("/"):
        return valor
    resto = valor[1:].lstrip(" -:\t")
    for rotulo in rotulos:
        if resto.lower().startswith(rotulo.lower()):
            return resto[len(rotulo) :].lstrip(" /-:\t")
    return resto


def localizar_rotulo(
    linhas: list[str],
    rotulos: list[str],
    todos_rotulos: list[str],
    validador=sempre_valido,
) -> Optional[tuple[int, str]]:
    """Procura por qualquer um dos rotulos no documento inteiro.

    A ordem de busca e por PRIORIDADE do rotulo (a ordem da lista), nao por
    posicao no documento: primeiro tenta achar o rotulo mais especifico no
    documento inteiro; so tenta o proximo rotulo da lista (mais generico)
    se o especifico nao aparecer em lugar nenhum. Isso evita que um rotulo
    generico que aparece mais cedo na pagina vença um rotulo especifico e
    confiavel que aparece mais tarde.

    Se o resto da linha apos o rotulo estiver vazio (ou nao passar no
    `validador`), tenta a linha seguinte -- layout comum em documentos
    extraidos de PDF, onde rotulo e valor saem em linhas separadas.

    `todos_rotulos` e o universo completo de rotulos do extrator chamador
    (usado so para calcular as palavras de corte e para reconhecer quando
    a linha seguinte e, ela mesma, outro rotulo em vez de um valor).
    """
    paradas = paradas_excluindo(todos_rotulos, rotulos)
    for rotulo in rotulos:
        for i, linha in enumerate(linhas):
            if rotulo.lower() not in linha.lower():
                continue

            resto = remainder_apos_rotulo(linha, rotulo)
            if resto:
                resto = limpar_prefixo_rotulos(resto, rotulos)
            valor = truncar_em_proximo_rotulo(resto, paradas) if resto else ""
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
                if proxima_bruta and not parece_rotulo(proxima_bruta, todos_rotulos):
                    proxima = limpar_prefixo_rotulos(proxima_bruta, rotulos)
                    valor = truncar_em_proximo_rotulo(proxima, paradas) if proxima else ""
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


def documento_fiscal_proximo(linhas: list[str], indice: int, janela: int = 3) -> Optional[str]:
    """Procura um CNPJ ou CPF nas linhas ao redor de onde um rotulo (ex:
    "Cedente", "Emitente") foi encontrado, para associar ao nome extraido."""
    trecho = " ".join(linhas[indice : indice + janela])
    m = CNPJ_RE.search(trecho)
    if m:
        return f"CNPJ {corrigir_confusao_ocr(m.group(0))}"
    m = CPF_RE.search(trecho)
    if m:
        return f"CPF {corrigir_confusao_ocr(m.group(0))}"
    return None


def extrair_entidade(
    linhas: list[str], rotulos: list[str], todos_rotulos: list[str]
) -> Optional[str]:
    """Localiza um nome (emissor/destinatario/etc.) por rotulo e associa o
    CNPJ/CPF mais proximo, se houver."""
    encontrado = localizar_rotulo(linhas, rotulos, todos_rotulos, validador=parece_nome)
    if not encontrado:
        return None
    indice, nome = encontrado

    documento_fiscal = documento_fiscal_proximo(linhas, indice)
    if documento_fiscal and documento_fiscal.split()[-1] not in nome:
        return f"{nome} ({documento_fiscal})"
    return nome


def extrair_data(linhas: list[str], rotulos: list[str], todos_rotulos: list[str]) -> Optional[str]:
    encontrado = localizar_rotulo(linhas, rotulos, todos_rotulos, validador=parece_data)
    if not encontrado:
        return None
    _, bruto = encontrado
    m = DATA_RE.search(bruto)
    return normalizar_data(m.group(1)) if m else None


def normalizar_data(data: str) -> str:
    """"15/4/2026" -> "15/04/2026" (dia e mes sempre com 2 digitos, mantendo
    o separador original). DANFE real imprime o mes sem zero a esquerda;
    normalizar aqui deixa o campo no mesmo formato em qualquer extrator."""
    m = re.fullmatch(r"(\d{1,2})([/-])(\d{1,2})\2(\d{4})", data)
    if not m:
        return data
    dia, separador, mes, ano = m.groups()
    return f"{int(dia):02d}{separador}{int(mes):02d}{separador}{ano}"
