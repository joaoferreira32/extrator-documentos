"""DanfeExtractor: extracao de campos de DANFE (Documento Auxiliar da Nota
Fiscal Eletronica).

- Deteccao por marcador (palavra-chave no texto).
- Validacao da chave de acesso (44 digitos + digito verificador modulo
  11): algoritmo publico e padronizado pela SEFAZ, nao depende de layout.
- Tabela de itens por coordenadas (`danfe_tabela.py`), quando o PDF tem
  palavras posicionadas disponiveis (`contexto.paginas_palavras` -- None
  quando a origem e OCR, que nao tem posicao confiavel de palavra). A
  estrutura de colunas foi confirmada com um dump real via
  `/debug/extract-words` (nota fiscal da Dell), nao suposta.
- CFOP predominante: moda dos CFOPs encontrados nas linhas da tabela.
- Validacao cruzada: soma dos `valor_total` dos itens comparada com
  "Valor Total dos Produtos" (rotulo padronizado nacionalmente pelo
  Manual de Orientacao do Contribuinte -- nao e suposicao de layout de um
  emissor especifico, e um campo obrigatorio em toda DANFE). Se nao
  bater, gera aviso em vez de devolver a tabela calada.

Campos fiscais que ainda dependem de confirmar o rotulo exato num dump
real (serie, natureza da operacao, data de saida, IE do emitente, base de
calculo e valor do ICMS/frete como campos de documento) ficam para uma
proxima iteracao -- ver TODO no fim do arquivo.

Reaproveita `GenericExtractor` para os campos universais (emissor,
destinatario, numero_documento, data_emissao, valor_total) -- mesma
qualidade que uma DANFE ja tinha antes desta refatoracao, sem regressao.
"""
import re
import unicodedata

from app.extractors import comum
from app.extractors.base import ContextoExtracao, ExtratorDocumento, ResultadoExtracao
from app.extractors.danfe_tabela import montar_tabela_itens
from app.extractors.generico import GenericExtractor
from app.schemas import CampoAdicional

MARCADORES_DETECCAO = [
    "danfe",
    "documento auxiliar da nota fiscal eletr",  # cobre "eletronica"/"eletrônica"
    "chave de acesso",
    "protocolo de autorizacao de uso",
    "protocolo de autorização de uso",
    "nf-e",
    "nfe",
    "cfop",
]

ROTULOS_VALOR_TOTAL_PRODUTOS = ["Valor Total dos Produtos", "Valor total dos produtos"]
ROTULOS_VALOR_TOTAL_NOTA = ["Valor Total da Nota", "Valor total da nota", "Valor Total da NF-e"]

# GRADE DE TOTAIS: cada "linha" da grade sai do pdfplumber como UMA linha de
# N rotulos seguida de UMA linha de N valores (confirmado com a saida de
# /debug/extractor-input de uma DANFE real):
#   "BASE DE CÁLCULO DO ICMS VALOR DO ICMS BASE DE CÁLCULO ICMS ST VALOR DO ICMS SUBSTITUIÇÃO VALOR TOTAL DOS PRODUTOS"
#   "229,00 41,22 0,00 0,00 215,03"
#   "VALOR DO FRETE VALOR DO SEGURO DESCONTO OUTRAS DESPESAS ACESSÓRIAS VALOR TOTAL DO I.P.I. VALOR TOTAL DA NOTA"
#   "0,00 0,00 0,00 0,00 13,97 229,00"
# O valor de um rotulo e o de MESMA POSICAO na linha de valores (5º rotulo ->
# 5º valor), nao "o numero mais proximo": cada linha vizinha tem valores de
# OUTROS campos (o total da nota so parecia certo porque o primeiro valor da
# linha de cima, a base do ICMS, coincidia com ele: nota = produtos + IPI).
# Chaves sem acento/maiusculas -- a linha e normalizada antes de comparar.
_ROTULOS_GRADE_TOTAIS = {
    "BASE DE CALCULO DO ICMS": "base_icms",
    "VALOR DO ICMS": "valor_icms",
    "BASE DE CALCULO ICMS ST": "base_icms_st",
    "VALOR DO ICMS SUBSTITUICAO": "valor_icms_st",
    "VALOR TOTAL DOS PRODUTOS": "valor_produtos",
    "VALOR DO FRETE": "frete",
    "VALOR DO SEGURO": "seguro",
    "DESCONTO": "desconto",
    "OUTRAS DESPESAS ACESSORIAS": "outras_despesas",
    "VALOR TOTAL DO I.P.I.": "ipi",
    "VALOR TOTAL DO IPI": "ipi",
    "VALOR TOTAL DA NOTA": "valor_nota",
}
_RE_ROTULO_GRADE = re.compile(
    r"(?<!\w)(?:"
    + "|".join(re.escape(r) for r in sorted(_ROTULOS_GRADE_TOTAIS, key=len, reverse=True))
    + r")(?!\w)"
)

# O CNPJ do emitente fica no bloco do emitente, varias linhas depois do nome
# (endereco, IE...): no PDF real, 10 linhas depois -- fora da janela de 10 que
# `documento_fiscal_proximo` olhava. O bloco vai ate o rotulo do destinatario
# (onde comeca outra entidade, cujo CNPJ/CPF NAO pode ser tomado pelo do
# emitente), com um teto de linhas de seguranca.
_LIMITE_LINHAS_BLOCO_EMITENTE = 40

# --- Rotulos reais de uma DANFE (confirmados com a saida de /debug/
# extractor-input de uma DANFE real; ver CLAUDE.md, "Layout real de uma
# DANFE") -----------------------------------------------------------------
#
# EMITENTE: a legenda "Identificação do emitente" e a linha SEGUINTE traz o
# nome, mas na mesma linha do nome vem texto da caixa vizinha (o titulo):
#   "Identificação do emitente DANFE"
#   "DELL COMPUTADORES DO BRASIL LTDA Documento Auxiliar da"
# Por isso o nome e cortado no titulo da caixa vizinha.
_ROTULO_EMITENTE_RE = re.compile(r"identifica[çc][ãa]o\s+do\s+emitente", re.IGNORECASE)
_TITULO_CAIXA_VIZINHA_RE = re.compile(r"documento\s+auxiliar|danfe", re.IGNORECASE)

# DESTINATARIO: pdfplumber junta a linha de rotulos e a linha de valores num
# unico texto quando estao proximas na vertical:
#   "NOME/RAZÃO SOCIAL CNPJ/CPF DATA DA EMISSÃO FULANO DE TAL 000.000.000-00
#    15/4/2026 ENDEREÇO BAIRRO/DISTRITO CEP ... RUA EXEMPLO, nº 1 CENTRO ..."
# O nome fica entre o ultimo rotulo da sequencia inicial e o CPF/CNPJ.
_ROTULO_NOME_DESTINATARIO_RE = re.compile(r"nome\s*/\s*raz[ãa]o\s+social", re.IGNORECASE)
ROTULOS_BLOCO_DESTINATARIO = [
    "CNPJ/CPF",
    "DATA DA EMISSÃO",
    "DATA DA EMISSAO",
    "ENDEREÇO",
    "ENDERECO",
    "BAIRRO/DISTRITO",
    "CEP",
    "DATA DA ENTRADA/SAÍDA",
    "DATA DA ENTRADA/SAIDA",
    "MUNICÍPIO",
    "MUNICIPIO",
    "FONE/FAX",
    "UF",
    "INSCRIÇÃO ESTADUAL",
    "INSCRICAO ESTADUAL",
    "HORA DA ENTRADA/SAÍDA",
    "HORA DA ENTRADA/SAIDA",
]
# Limites de palavra explicitos: "UF" nao pode casar dentro de "UFRJ".
_RE_ROTULO_BLOCO = re.compile(
    r"(?<!\w)(?:"
    + "|".join(re.escape(r) for r in sorted(ROTULOS_BLOCO_DESTINATARIO, key=len, reverse=True))
    + r")(?!\w)",
    re.IGNORECASE,
)

# Chave de acesso: 44 digitos, geralmente impressos em grupos separados por
# espaco/ponto NA MESMA LINHA. Aceita ate 90 caracteres de largura pra
# cobrir espacamento generoso (confirmado num dump real: 11 blocos de 4
# digitos separados por espaco), e confirma o total de 44 digitos so
# depois de remover os separadores. Usa " " (espaco), nao \s -- bug real:
# \s tambem casa quebra de linha, entao o candidato "vazava" pra tras
# pegando os ultimos digitos da linha anterior (ex: "...0010-01" seguido
# de quebra de linha e a chave) quando essa linha anterior tambem termina
# em digito, inflando a contagem pra mais de 44 e descartando a chave de
# verdade.
_CANDIDATO_CHAVE_RE = re.compile(r"\d[\d .]{40,90}\d")

# Tolerancia de arredondamento na validacao da soma dos itens vs valor
# total dos produtos.
_TOLERANCIA_SOMA = 0.02


def _digito_verificador(chave_43_digitos: str) -> str:
    """Digito verificador da chave de acesso da NFe: modulo 11 com pesos
    ciclicos de 2 a 9, da direita pra esquerda. Algoritmo publico definido
    pela SEFAZ (Manual de Orientacao do Contribuinte), nao depende do
    layout de nenhum documento especifico."""
    pesos = [2, 3, 4, 5, 6, 7, 8, 9]
    soma = sum(
        int(digito) * pesos[posicao % 8] for posicao, digito in enumerate(reversed(chave_43_digitos))
    )
    resto = soma % 11
    return "0" if resto < 2 else str(11 - resto)


def encontrar_chave_acesso(texto: str) -> tuple[str | None, str]:
    """Procura uma sequencia de 44 digitos (a chave de acesso) e valida o
    digito verificador. Devolve (chave, confianca): confianca "alta"
    quando o DV bate, "baixa" quando acha 44 digitos mas o DV nao confere
    (mais provavel de ser OCR ruim ou coincidencia do que a chave de
    verdade -- por isso nao descarta, so avisa)."""
    for candidato in _CANDIDATO_CHAVE_RE.finditer(texto):
        apenas_digitos = re.sub(r"\D", "", candidato.group(0))
        if len(apenas_digitos) != 44:
            continue
        dv_calculado = _digito_verificador(apenas_digitos[:43])
        confianca = "alta" if dv_calculado == apenas_digitos[43] else "baixa"
        return apenas_digitos, confianca
    return None, ""


def partes_da_chave(chave: str) -> dict[str, str]:
    """Campos embutidos na chave de acesso (layout padronizado pela SEFAZ:
    cUF(2) AAMM(4) CNPJ(14) modelo(2) serie(3) nNF(9) tpEmis(1) cNF(8)
    DV(1)). Serve pra CONFERIR o que foi lido do resto do documento
    (numero da nota, CNPJ do emitente) -- evidencia independente de layout."""
    return {
        "uf": chave[0:2],
        "aamm": chave[2:6],
        "cnpj": chave[6:20],
        "modelo": chave[20:22],
        "serie": chave[22:25],
        "numero": chave[25:34],
    }


def _cortar_em_titulo(texto: str) -> str:
    """Corta o texto no titulo da caixa vizinha ("Documento Auxiliar da",
    "DANFE") que o pdfplumber cola na mesma linha do nome do emitente."""
    m = _TITULO_CAIXA_VIZINHA_RE.search(texto)
    return texto[: m.start()] if m else texto


def _pular_rotulos_iniciais(texto: str) -> str:
    """Remove a sequencia de rotulos do bloco do destinatario colados no
    inicio do texto ("CNPJ/CPF DATA DA EMISSÃO ..."), deixando o que vem
    depois do ultimo deles -- o nome."""
    texto = texto.lstrip(" :-")
    while True:
        m = _RE_ROTULO_BLOCO.match(texto)
        if not m:
            return texto
        texto = texto[m.end() :].lstrip(" :-")


def _cortar_nome_no_documento(texto: str) -> tuple[str, str | None]:
    """Corta o nome no primeiro CPF/CNPJ, data ou rotulo do bloco.
    Devolve (nome, documento_fiscal): `documento_fiscal` ("CPF 000..."/"CNPJ
    ...") so vem preenchido quando o corte foi NO documento -- e a evidencia
    de que o nome esta encaixado entre os rotulos e o CPF/CNPJ dele."""
    cortes = []
    m_doc, tipo_doc = None, None
    for tipo, regex in (("CNPJ", comum.CNPJ_RE), ("CPF", comum.CPF_RE)):
        m = regex.search(texto)
        if m and (m_doc is None or m.start() < m_doc.start()):
            m_doc, tipo_doc = m, tipo
    if m_doc:
        cortes.append(m_doc.start())
    for regex in (comum.DATA_RE, _RE_ROTULO_BLOCO):
        m = regex.search(texto)
        if m:
            cortes.append(m.start())
    fim = min(cortes) if cortes else len(texto)
    documento = None
    if m_doc and m_doc.start() == fim:
        documento = f"{tipo_doc} {comum.corrigir_confusao_ocr(m_doc.group(0))}"
    return texto[:fim].strip(" :-"), documento


def extrair_destinatario(linhas: list[str]) -> tuple[str | None, str | None, str]:
    """Destinatario a partir de "NOME/RAZÃO SOCIAL". Tenta o resto da MESMA
    linha (layout real: rotulos e valores colados numa linha so) e, se ali
    so sobrarem rotulos, a linha seguinte (rotulo e valor em linhas
    separadas). Devolve (nome, documento_fiscal, confianca): "alta" quando o
    nome fica encaixado antes do CPF/CNPJ (estrutura confirmada), "media"
    senao (documento_fiscal None)."""
    for i, linha in enumerate(linhas):
        m = _ROTULO_NOME_DESTINATARIO_RE.search(linha)
        if not m:
            continue
        candidatos = [linha[m.end() :]]
        if i + 1 < len(linhas):
            candidatos.append(linhas[i + 1])
        for texto in candidatos:
            nome, documento = _cortar_nome_no_documento(_pular_rotulos_iniciais(texto))
            if nome and comum.parece_nome(nome):
                return nome, documento, "alta" if documento else "media"
    return None, None, ""


def _cnpjs_do_bloco_do_emitente(linhas: list[str], indice_nome: int) -> list[str]:
    """CNPJs (so os numeros formatados) do bloco do emitente: da linha do
    nome ate o rotulo do destinatario (ou o teto de linhas)."""
    fim = min(indice_nome + _LIMITE_LINHAS_BLOCO_EMITENTE, len(linhas))
    for j in range(indice_nome + 1, fim):
        if _ROTULO_NOME_DESTINATARIO_RE.search(linhas[j]):
            fim = j
            break
    trecho = "\n".join(linhas[indice_nome:fim])
    return [comum.corrigir_confusao_ocr(m.group(0)) for m in comum.CNPJ_RE.finditer(trecho)]


def extrair_emissor(linhas: list[str], cnpj_chave: str | None = None) -> tuple[str | None, str | None]:
    """Nome do emitente e seu CNPJ ("CNPJ 72.381...", se houver) a partir da
    legenda "Identificação do emitente": o nome esta na mesma linha (se
    sobrar algo depois do titulo) ou na linha seguinte.

    O CNPJ e procurado no bloco do emitente inteiro (ate o rotulo do
    destinatario). Havendo mais de um, prefere o que bate com `cnpj_chave`
    (os 14 digitos embutidos na chave de acesso)."""
    for i, linha in enumerate(linhas):
        m = _ROTULO_EMITENTE_RE.search(linha)
        if not m:
            continue
        candidatos = [(i, linha[m.end() :])]
        if i + 1 < len(linhas):
            candidatos.append((i + 1, linhas[i + 1]))
        for indice, texto in candidatos:
            nome = _cortar_em_titulo(texto).strip(" :-")
            if nome and comum.parece_nome(nome):
                cnpjs = _cnpjs_do_bloco_do_emitente(linhas, indice)
                escolhido = None
                if cnpj_chave:
                    escolhido = next((c for c in cnpjs if re.sub(r"\D", "", c) == cnpj_chave), None)
                if escolhido is None and cnpjs:
                    escolhido = cnpjs[0]
                return nome, f"CNPJ {escolhido}" if escolhido else None
    return None, None


def _sem_acentos_maiusculo(texto: str) -> str:
    sem_acento = "".join(c for c in unicodedata.normalize("NFKD", texto) if not unicodedata.combining(c))
    return " ".join(sem_acento.upper().split())


def _so_valores(texto: str) -> bool:
    """True quando o texto tem apenas valores monetarios e espacos."""
    return not comum.VALOR_NUM_RE.sub(" ", texto).strip()


def extrair_totais_grade(linhas: list[str]) -> dict[str, float]:
    """Le a grade de totais da DANFE casando cada rotulo com o valor de
    MESMA POSICAO (5º rotulo -> 5º valor), nunca "o numero mais proximo".

    Uma linha so conta como linha de rotulos da grade se, tirando os rotulos
    conhecidos, sobram apenas valores/espacos; e so e aceita se a quantidade
    de valores (na mesma linha ou na linha seguinte) for IGUAL a de rotulos
    -- senao nao adivinha (rotulo desconhecido na linha, valor faltando...)
    e simplesmente nao devolve aqueles campos.

    Chaves devolvidas: base_icms, valor_icms, base_icms_st, valor_icms_st,
    valor_produtos, frete, seguro, desconto, outras_despesas, ipi, valor_nota."""
    totais: dict[str, float] = {}
    for i, linha in enumerate(linhas):
        norm = _sem_acentos_maiusculo(linha)
        rotulos = [_ROTULOS_GRADE_TOTAIS[m.group(0)] for m in _RE_ROTULO_GRADE.finditer(norm)]
        if not rotulos:
            continue
        resto = _RE_ROTULO_GRADE.sub(" ", norm)
        if not _so_valores(resto):
            continue  # a linha tem texto que nao e rotulo da grade
        valores = comum.VALOR_NUM_RE.findall(resto)
        # Valores na linha SEGUINTE so valem pra uma linha de grade (2+ rotulos, onde
        # a contagem N = N confirma o casamento). Com 1 rotulo so, o numero da
        # proxima linha e ambiguo (pode ser o deste campo ou do proximo) -- nao chuta.
        if not valores and len(rotulos) > 1 and i + 1 < len(linhas) and _so_valores(linhas[i + 1]):
            valores = comum.VALOR_NUM_RE.findall(linhas[i + 1])
        if len(valores) != len(rotulos):
            continue
        for chave, bruto in zip(rotulos, valores):
            valor = comum.para_numero(comum.corrigir_confusao_ocr(bruto))
            if isinstance(valor, float):
                totais.setdefault(chave, valor)
    return totais


def totais_fecham(totais: dict[str, float]) -> bool | None:
    """Confere a formula do total da nota (padrao da SEFAZ): produtos -
    desconto + ICMS ST + frete + seguro + outras despesas + IPI. None quando
    faltam produtos ou total da nota pra conferir."""
    if "valor_produtos" not in totais or "valor_nota" not in totais:
        return None
    esperado = (
        totais["valor_produtos"]
        - totais.get("desconto", 0.0)
        + totais.get("valor_icms_st", 0.0)
        + totais.get("frete", 0.0)
        + totais.get("seguro", 0.0)
        + totais.get("outras_despesas", 0.0)
        + totais.get("ipi", 0.0)
    )
    return abs(esperado - totais["valor_nota"]) <= _TOLERANCIA_SOMA


class DanfeExtractor(ExtratorDocumento):
    tipo = "nota_fiscal"

    def pontuacao_deteccao(self, contexto: ContextoExtracao) -> float:
        texto_lower = contexto.texto.lower()
        return 0.9 if any(marcador in texto_lower for marcador in MARCADORES_DETECCAO) else 0.0

    def extrair(self, contexto: ContextoExtracao) -> ResultadoExtracao:
        # Base: mesma extracao generica que uma DANFE ja recebia antes
        # desta refatoracao (sem regressao nos campos universais).
        resultado = GenericExtractor().extrair(contexto)
        resultado.documento.tipo_documento = "nota_fiscal"

        chave, confianca_chave = encontrar_chave_acesso(contexto.texto)
        if chave:
            resultado.documento.campos_adicionais.append(
                CampoAdicional(campo="Chave de Acesso", valor=chave)
            )
            resultado.confiancas["Chave de Acesso"] = confianca_chave
            if confianca_chave == "baixa":
                resultado.avisos.append(
                    "A chave de acesso encontrada tem 44 dígitos, mas o dígito "
                    "verificador não confere — confira a chave manualmente."
                )
        # So uma chave com DV valido serve de evidencia pra conferir os
        # outros campos (uma chave com DV errado pode ser OCR ruim).
        partes_chave = partes_da_chave(chave) if chave and confianca_chave == "alta" else None

        self._extrair_emissor(contexto, resultado, partes_chave)
        self._extrair_destinatario(contexto, resultado)
        self._conferir_numero_documento(resultado, partes_chave)

        totais = extrair_totais_grade(contexto.linhas)
        self._extrair_valor_total(contexto, resultado, totais)

        if contexto.paginas_palavras:
            self._extrair_tabela_itens(contexto, resultado, totais)

        # TODO: numero de serie, natureza da operacao, data de saida, IE
        # do emitente, e os campos de ICMS/frete como valores de
        # documento (nao agregados da tabela) -- precisam do rotulo exato
        # confirmado num dump real antes de implementar, pra nao repetir
        # o erro de supor layout sem material real.

        return resultado

    def _extrair_emissor(
        self, contexto: ContextoExtracao, resultado: ResultadoExtracao, partes_chave: dict | None
    ) -> None:
        """Emissor a partir de "Identificação do emitente" (ver
        `extrair_emissor`). O emissor do GenericExtractor e SEMPRE
        descartado aqui: o rotulo generico "Emitente" bate como substring
        dentro dessa legenda e capturava o titulo "DANFE" da caixa vizinha.

        Confianca por EVIDENCIA, nao por ter achado um rotulo: "alta" so
        quando o CNPJ lido junto do nome bate com o CNPJ embutido na chave
        de acesso (que tem DV valido); "media" sem essa conferencia;
        "baixa" (e aviso) quando o CNPJ contradiz a chave."""
        gerado_pelo_generico = resultado.documento.emissor
        resultado.documento.emissor = None
        resultado.confiancas.pop("emissor", None)

        cnpj_chave = partes_chave["cnpj"] if partes_chave else None
        nome, documento_fiscal = extrair_emissor(contexto.linhas, cnpj_chave)
        if nome:
            resultado.documento.emissor = f"{nome} ({documento_fiscal})" if documento_fiscal else nome
            confianca = "media"
            if partes_chave and documento_fiscal and documento_fiscal.startswith("CNPJ"):
                if re.sub(r"\D", "", documento_fiscal) == partes_chave["cnpj"]:
                    confianca = "alta"
                else:
                    confianca = "baixa"
                    resultado.avisos.append(
                        "O CNPJ do emitente lido no documento difere do CNPJ embutido na "
                        "chave de acesso — confira o emissor manualmente."
                    )
            resultado.confiancas["emissor"] = confianca
        elif gerado_pelo_generico and not _TITULO_CAIXA_VIZINHA_RE.search(gerado_pelo_generico):
            # Sem a legenda "Identificação do emitente" (outro layout): mantem
            # o do generico, mas sem fingir certeza.
            resultado.documento.emissor = gerado_pelo_generico
            resultado.confiancas["emissor"] = "baixa"

    def _extrair_destinatario(self, contexto: ContextoExtracao, resultado: ResultadoExtracao) -> None:
        """Destinatario no MESMO formato do emissor e do boleto, "NOME
        (CPF 000...)"/"NOME (CNPJ ...)", a partir de "NOME/RAZÃO SOCIAL"
        (ver `extrair_destinatario`). O do GenericExtractor e descartado:
        buscava "Destinatário", que bate no titulo da secao, e devolvia uma
        linha inteira de rotulos."""
        gerado_pelo_generico = resultado.documento.destinatario
        resultado.documento.destinatario = None
        resultado.confiancas.pop("destinatario", None)

        nome, documento_fiscal, confianca = extrair_destinatario(contexto.linhas)
        if nome:
            resultado.documento.destinatario = f"{nome} ({documento_fiscal})" if documento_fiscal else nome
            resultado.confiancas["destinatario"] = confianca
        elif gerado_pelo_generico and not _RE_ROTULO_BLOCO.search(gerado_pelo_generico):
            resultado.documento.destinatario = gerado_pelo_generico
            resultado.confiancas["destinatario"] = "baixa"

    def _extrair_valor_total(
        self, contexto: ContextoExtracao, resultado: ResultadoExtracao, totais: dict[str, float]
    ) -> None:
        """Total da nota pela POSICAO na grade de totais (ver
        `extrair_totais_grade`). Confianca "alta" quando a formula do total
        fecha (produtos - desconto + ST + frete + seguro + outras + IPI =
        total da nota) ou nao ha como conferir; "media" + aviso quando a
        grade foi lida mas nao fecha.

        Sem a grade (outro layout), tenta "Valor Total da Nota <valor>" na
        MESMA linha; e sem isso fica o fallback generico (que exige "R$")."""
        nota = totais.get("valor_nota")
        if nota is not None:
            resultado.documento.valor_total = nota
            if totais_fecham(totais) is False:
                resultado.confiancas["valor_total"] = "media"
                resultado.avisos.append(
                    "Os valores da grade de totais não fecham (produtos − desconto + ICMS ST + "
                    "frete + seguro + outras despesas + IPI difere do total da nota) — "
                    "confira o valor total manualmente."
                )
            else:
                resultado.confiancas["valor_total"] = "alta"
            return

        valor_total_bruto = comum.extrair_valor_rotulo(contexto.texto, ROTULOS_VALOR_TOTAL_NOTA)
        if valor_total_bruto:
            resultado.documento.valor_total = comum.para_numero(valor_total_bruto)
            resultado.confiancas["valor_total"] = "alta"

    def _conferir_numero_documento(
        self, resultado: ResultadoExtracao, partes_chave: dict | None
    ) -> None:
        """Confere o numero da nota lido do documento com o embutido na
        chave de acesso. Sem essa conferencia o numero achado pelo
        fallback "Nº..." saia "baixa" mesmo estando certo; contradicao
        vira "baixa" + aviso."""
        numero = resultado.documento.numero_documento
        if not partes_chave or not numero:
            return
        digitos = re.sub(r"\D", "", str(numero))
        if digitos and int(digitos) == int(partes_chave["numero"]):
            resultado.confiancas["numero_documento"] = "alta"
        else:
            resultado.confiancas["numero_documento"] = "baixa"
            resultado.avisos.append(
                "O número da nota lido no documento difere do número embutido na "
                "chave de acesso — confira manualmente."
            )

    def _extrair_tabela_itens(
        self, contexto: ContextoExtracao, resultado: ResultadoExtracao, totais: dict[str, float]
    ) -> None:
        tabela = montar_tabela_itens(contexto.paginas_palavras)
        if not tabela.itens:
            return

        resultado.documento.itens = tabela.itens
        # Heuristica posicional -> "media". Sobe pra "alta" abaixo SO quando a
        # soma dos itens fecha com o Valor Total dos Produtos (evidencia
        # independente: o total vem de outra parte do documento).
        resultado.confiancas["itens"] = "media"

        if tabela.cfop_predominante:
            resultado.documento.campos_adicionais.append(
                CampoAdicional(campo="CFOP", valor=tabela.cfop_predominante)
            )
            resultado.confiancas["CFOP"] = "media"

        # Valor Total dos Produtos: pela posicao na grade de totais; sem a
        # grade, "rotulo <valor>" na mesma linha.
        valor_total_produtos = totais.get("valor_produtos")
        if valor_total_produtos is None:
            bruto = comum.extrair_valor_rotulo(contexto.texto, ROTULOS_VALOR_TOTAL_PRODUTOS)
            valor_total_produtos = comum.para_numero(bruto) if bruto else None

        if isinstance(valor_total_produtos, float):
            # Exposto como campo_adicional pra quem exporta pra Excel poder
            # conferir a soma dos itens numero-contra-numero (aba Campos
            # adicionais), em vez de so confiar no aviso de texto abaixo --
            # antes desse campo, esse valor nunca aparecia em lugar nenhum
            # do arquivo exportado, so era usado internamente pra gerar aviso.
            resultado.documento.campos_adicionais.append(
                CampoAdicional(campo="Valor Total dos Produtos", valor=comum.formatar_valor_br(valor_total_produtos))
            )
            resultado.confiancas["Valor Total dos Produtos"] = "alta"

        if not isinstance(valor_total_produtos, float) or tabela.soma_valor_total is None:
            return

        if abs(valor_total_produtos - tabela.soma_valor_total) > _TOLERANCIA_SOMA:
            resultado.avisos.append(
                f"A soma dos itens (R$ {comum.formatar_valor_br(tabela.soma_valor_total)}) não bate com "
                f"o Valor Total dos Produtos informado (R$ {comum.formatar_valor_br(valor_total_produtos)}) "
                "— confira a tabela manualmente."
            )
        elif all(isinstance(item.valor_total, float) for item in tabela.itens):
            # Todos os itens com total numerico E a soma bate: sem essa
            # segunda condicao, um item sem total (contando como 0) poderia
            # "fechar" a soma por acaso.
            resultado.confiancas["itens"] = "alta"
