"""DanfeExtractor: extracao de campos de DANFE (Documento Auxiliar da Nota
Fiscal Eletronica).

ESTADO ATUAL (etapa 1-2 do plano, antes do dump real de uma DANFE via
`/debug/extract-words`): so implementa o que NAO depende de suposicao de
layout --

- Deteccao por marcador (palavra-chave no texto): mecanismo identico ao
  que ja existia no basic_extractor.py monolitico, so que agora escopado
  nesta classe.
- Validacao da chave de acesso (44 digitos + digito verificador modulo
  11): e um algoritmo publico e padronizado pela SEFAZ, nao uma suposicao
  de como o pdfplumber extrai o texto de uma DANFE especifica -- da pra
  implementar com confianca sem ver um documento real.

Os campos especificos de DANFE pedidos (numero da NF por rotulo dedicado,
serie, natureza da operacao, CFOP predominante, datas, ICMS, frete, e
principalmente a tabela de itens por coordenadas) dependem de como o
pdfplumber realmente linhariza o layout em grade de uma DANFE real -- já
foi bug real nesse projeto (boleto) confiar em suposicao de layout em vez
de material real, entao esses campos ficam para a proxima etapa, apos o
dump de `/debug/extract-words`.

Enquanto isso, reaproveita `GenericExtractor` para os campos universais
(emissor, destinatario, numero_documento, data_emissao, valor_total) --
mesma qualidade que uma DANFE ja tinha antes desta refatoracao (fallback
generico), sem regressao. So adiciona chave de acesso por cima.
"""
import re

from app.extractors.base import ContextoExtracao, ExtratorDocumento, ResultadoExtracao
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

# Chave de acesso: 44 digitos, geralmente impressos em grupos separados por
# espaco/ponto. Aceita ate 80 caracteres de largura pra cobrir espacamento
# generoso, e confirma o total de 44 digitos so depois de remover os
# separadores.
_CANDIDATO_CHAVE_RE = re.compile(r"\d[\d\s.]{40,90}\d")


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

        # TODO (etapa 3, apos dump real via /debug/extract-words): numero
        # da NF/serie/natureza da operacao por rotulo dedicado, CFOP
        # predominante, data de saida, IE do emitente, base de calculo e
        # valor do ICMS, valor dos produtos, valor do frete, e a tabela de
        # itens por coordenadas (danfe_tabela.py, ainda nao existe).

        return resultado
