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
from app.extractors.generico import TODOS_ROTULOS as _TODOS_ROTULOS_GENERICO
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

# Rotulo real do campo de nome do destinatario numa DANFE (confirmado num
# dump real via /debug/extract-text) -- diferente do generico
# "Destinatário"/"Cliente" que o GenericExtractor usa, que nao aparece
# nesse layout: a caixa do destinatario tem "NOME/RAZÃO SOCIAL" como
# legenda, com o rotulo ANTES do valor (padrao normal, ja coberto pelo
# fallback de proxima linha existente em localizar_rotulo).
ROTULOS_DESTINATARIO_DANFE = ["Nome/Razão Social", "Nome/Razao Social"]
TODOS_ROTULOS_DANFE = _TODOS_ROTULOS_GENERICO + ROTULOS_DESTINATARIO_DANFE

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

        emissor = self._extrair_emissor(contexto)
        if emissor:
            resultado.documento.emissor = emissor
            resultado.confiancas["emissor"] = "alta"

        destinatario = comum.extrair_entidade(
            contexto.linhas, ROTULOS_DESTINATARIO_DANFE, TODOS_ROTULOS_DANFE
        )
        if destinatario:
            resultado.documento.destinatario = destinatario
            resultado.confiancas["destinatario"] = "alta"

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

    def _extrair_emissor(self, contexto: ContextoExtracao) -> str | None:
        """O nome do emitente numa DANFE aparece nas primeiras linhas do
        documento, ANTES de qualquer rotulo -- confirmado num dump real
        via /debug/extract-text. A legenda "Identificação do Emitente"
        funciona como legenda por baixo da caixa (mesmo padrao de "VALOR
        TOTAL DA NOTA"), entao buscar por rotulo generico
        ("Emitente"/"Fornecedor", como o GenericExtractor faz) e o que
        causava o bug real: a legenda batia com o rotulo, mas a linha
        seguinte pertencia a uma caixa vizinha (o titulo "DANFE"), nao ao
        emitente de verdade. Pular linhas que sao o proprio marcador de
        titulo da DANFE evita reincidir nesse bug caso o nome nao seja
        mesmo a primeira linha em algum layout."""
        for i, linha in enumerate(contexto.linhas):
            candidato = linha.strip()
            if not candidato:
                continue
            if any(marcador in candidato.lower() for marcador in MARCADORES_DETECCAO):
                continue
            if not comum.parece_nome(candidato):
                continue
            # janela maior que o default (3): confirmado num dump real que
            # o CNPJ do emitente vem depois do bloco de endereco (nome,
            # rua, bairro/CEP, municipio/UF -- ate 7 linhas antes do CNPJ).
            documento_fiscal = comum.documento_fiscal_proximo(contexto.linhas, i, janela=10)
            if documento_fiscal and documento_fiscal.split()[-1] not in candidato:
                return f"{candidato} ({documento_fiscal})"
            return candidato
        return None

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
