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


def _cortar_nome_no_documento(texto: str) -> tuple[str, bool]:
    """Corta o nome no primeiro CPF/CNPJ, data ou rotulo do bloco.
    Devolve (nome, cortou_num_documento) -- o segundo valor e a evidencia de
    que o nome esta encaixado entre os rotulos e o CPF/CNPJ dele."""
    cortes = []
    m_doc = None
    for regex in (comum.CNPJ_RE, comum.CPF_RE):
        m = regex.search(texto)
        if m and (m_doc is None or m.start() < m_doc.start()):
            m_doc = m
    if m_doc:
        cortes.append(m_doc.start())
    for regex in (comum.DATA_RE, _RE_ROTULO_BLOCO):
        m = regex.search(texto)
        if m:
            cortes.append(m.start())
    fim = min(cortes) if cortes else len(texto)
    return texto[:fim].strip(" :-"), bool(m_doc) and m_doc.start() == fim


def extrair_destinatario(linhas: list[str]) -> tuple[str | None, str]:
    """Nome do destinatario a partir de "NOME/RAZÃO SOCIAL". Tenta o resto
    da MESMA linha (layout real: rotulos e valores colados numa linha so) e,
    se ali so sobrarem rotulos, a linha seguinte (rotulo e valor em linhas
    separadas). Devolve (nome, confianca): "alta" quando o nome fica
    encaixado antes do CPF/CNPJ (estrutura confirmada), "media" senao."""
    for i, linha in enumerate(linhas):
        m = _ROTULO_NOME_DESTINATARIO_RE.search(linha)
        if not m:
            continue
        candidatos = [linha[m.end() :]]
        if i + 1 < len(linhas):
            candidatos.append(linhas[i + 1])
        for texto in candidatos:
            nome, com_documento = _cortar_nome_no_documento(_pular_rotulos_iniciais(texto))
            if nome and comum.parece_nome(nome):
                return nome, "alta" if com_documento else "media"
    return None, ""


def extrair_emissor(linhas: list[str]) -> tuple[str | None, str | None]:
    """Nome do emitente (e o CPF/CNPJ que vier depois, se houver) a partir
    da legenda "Identificação do emitente": o nome esta na mesma linha (se
    sobrar algo depois do titulo) ou na linha seguinte."""
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
                # janela maior que o default (3): o CNPJ vem depois do
                # bloco de endereco/IE (varias linhas depois do nome).
                return nome, comum.documento_fiscal_proximo(linhas, indice, janela=10)
    return None, None


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
                    "A chave de acesso encontrada tem 44 digitos mas o digito "
                    "verificador nao confere -- confira manualmente."
                )
        # So uma chave com DV valido serve de evidencia pra conferir os
        # outros campos (uma chave com DV errado pode ser OCR ruim).
        partes_chave = partes_da_chave(chave) if chave and confianca_chave == "alta" else None

        self._extrair_emissor(contexto, resultado, partes_chave)
        self._extrair_destinatario(contexto, resultado)
        self._conferir_numero_documento(resultado, partes_chave)

        # "Valor Total da Nota" e o rotulo padronizado nacionalmente pro
        # valor total do documento -- mais confiavel que o fallback
        # generico (que exige "R$" na frente do valor, mas uma DANFE
        # normalmente imprime so o numero, sem o simbolo). preferir_linha_
        # anterior=True porque a grade de totais real imprime o valor
        # ACIMA da legenda, nao abaixo (ver docstring de extrair_valor_rotulo).
        valor_total_bruto = comum.extrair_valor_rotulo(
            contexto.texto, ROTULOS_VALOR_TOTAL_NOTA, preferir_linha_anterior=True
        )
        if valor_total_bruto:
            resultado.documento.valor_total = comum.para_numero(valor_total_bruto)
            resultado.confiancas["valor_total"] = "alta"

        if contexto.paginas_palavras:
            self._extrair_tabela_itens(contexto, resultado)

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

        nome, documento_fiscal = extrair_emissor(contexto.linhas)
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
                        "chave de acesso -- confira o emissor manualmente."
                    )
            resultado.confiancas["emissor"] = confianca
        elif gerado_pelo_generico and not _TITULO_CAIXA_VIZINHA_RE.search(gerado_pelo_generico):
            # Sem a legenda "Identificação do emitente" (outro layout): mantem
            # o do generico, mas sem fingir certeza.
            resultado.documento.emissor = gerado_pelo_generico
            resultado.confiancas["emissor"] = "baixa"

    def _extrair_destinatario(self, contexto: ContextoExtracao, resultado: ResultadoExtracao) -> None:
        """Destinatario (so o nome/razao social, sem CPF/CNPJ) a partir de
        "NOME/RAZÃO SOCIAL" (ver `extrair_destinatario`). O do
        GenericExtractor e descartado: buscava "Destinatário", que bate no
        titulo da secao, e devolvia uma linha inteira de rotulos."""
        gerado_pelo_generico = resultado.documento.destinatario
        resultado.documento.destinatario = None
        resultado.confiancas.pop("destinatario", None)

        nome, confianca = extrair_destinatario(contexto.linhas)
        if nome:
            resultado.documento.destinatario = nome
            resultado.confiancas["destinatario"] = confianca
        elif gerado_pelo_generico and not _RE_ROTULO_BLOCO.search(gerado_pelo_generico):
            resultado.documento.destinatario = gerado_pelo_generico
            resultado.confiancas["destinatario"] = "baixa"

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
                "O numero da nota lido no documento difere do numero embutido na "
                "chave de acesso -- confira manualmente."
            )

    def _extrair_tabela_itens(self, contexto: ContextoExtracao, resultado: ResultadoExtracao) -> None:
        tabela = montar_tabela_itens(contexto.paginas_palavras)
        if not tabela.itens:
            return

        resultado.documento.itens = tabela.itens
        resultado.confiancas["itens"] = "media"  # heuristica posicional

        if tabela.cfop_predominante:
            resultado.documento.campos_adicionais.append(
                CampoAdicional(campo="CFOP", valor=tabela.cfop_predominante)
            )
            resultado.confiancas["CFOP"] = "media"

        valor_total_produtos_bruto = comum.extrair_valor_rotulo(
            contexto.texto, ROTULOS_VALOR_TOTAL_PRODUTOS, preferir_linha_anterior=True
        )
        if valor_total_produtos_bruto is None or tabela.soma_valor_total is None:
            return

        valor_total_produtos = comum.para_numero(valor_total_produtos_bruto)
        if not isinstance(valor_total_produtos, float):
            return

        if abs(valor_total_produtos - tabela.soma_valor_total) > _TOLERANCIA_SOMA:
            resultado.avisos.append(
                f"A soma dos itens (R$ {tabela.soma_valor_total:.2f}) nao bate com "
                f"o Valor Total dos Produtos informado (R$ {valor_total_produtos:.2f}) "
                "-- confira a tabela manualmente."
            )
