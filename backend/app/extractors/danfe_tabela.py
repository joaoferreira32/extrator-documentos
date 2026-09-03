"""Reconstrucao da tabela de itens da DANFE por coordenadas (x0/x1/top/
bottom de cada palavra), nao por regex sobre texto corrido -- texto
corrido perde a estrutura de colunas de uma tabela.

Estrutura de colunas e as duas armadilhas de tokenizacao abaixo foram
confirmadas com um dump real (`/debug/extract-words`) de uma DANFE real
(nota da Dell), nao inventadas por suposicao -- ver CLAUDE.md.
"""
import re
import statistics
from dataclasses import dataclass
from typing import Optional

from app.extractors.base import Palavra
from app.extractors import comum
from app.schemas import ItemDocumento

TOLERANCIA_LINHA = 3.0
"""Palavras cujo `top` difere em ate essa distancia (pontos PDF) sao
consideradas da mesma linha."""

VALOR_MONETARIO_RE = re.compile(r"\d{1,3}(?:\.\d{3})*,\d{2}")

# Colunas da tabela de itens de uma DANFE, na ordem padrao nacional
# (Manual de Orientacao do Contribuinte / layout da SEFAZ) -- confirmada
# pelo dump real. Cada coluna tem uma palavra-chave curta o suficiente pra
# sobreviver a diferencas de tokenizacao (pdfplumber so quebra em
# espaco em branco; "B.CALC.ICMS" e "ICMS/IPI", por exemplo, saem como um
# unico token porque nao tem espaco interno).
COLUNAS = [
    ("codigo", ["código", "codigo"]),
    ("descricao", ["descri"]),
    ("ncm", ["ncm"]),
    ("cst", ["cst"]),
    ("cfop", ["cfop"]),
    ("unidade", ["unid"]),
    ("quantidade", ["quant"]),
    ("valor_unitario", ["unit"]),
    ("valor_total", ["total"]),
    ("base_icms", ["calc"]),
    ("valor_icms", ["icms"]),
    ("valor_ipi", ["i.p.i", "ipi"]),  # "I.P.I." tem pontos entre as letras -- "ipi" sozinho nao bate
    ("aliquota", ["alíq", "aliq"]),
]

# Regra confirmada no dump real: rotulos de secao impressos girados
# (ex: "SOTUDORP" = "PRODUTOS" ao contrario) ficam numa faixa estreita de
# x0 na margem esquerda e tem caixa delimitadora alta e estreita (o
# oposto de uma palavra horizontal normal, que e larga e baixa).
_X0_MIN_TEXTO_VERTICAL = 15
_X0_MAX_TEXTO_VERTICAL = 85
_RAZAO_ALTURA_LARGURA_MINIMA = 2.0


def _eh_texto_vertical(p: Palavra) -> bool:
    largura = p.x1 - p.x0
    altura = p.bottom - p.top
    if not (_X0_MIN_TEXTO_VERTICAL <= p.x0 <= _X0_MAX_TEXTO_VERTICAL):
        return False
    if largura <= 0:
        return False
    return altura > _RAZAO_ALTURA_LARGURA_MINIMA * largura


def _agrupar_linhas(palavras: list[Palavra]) -> list[list[Palavra]]:
    """Agrupa palavras por proximidade vertical de `top` -- cada grupo e
    uma linha visual da pagina, na ordem em que aparecem de cima pra
    baixo."""
    candidatas = [p for p in palavras if not _eh_texto_vertical(p)]
    candidatas.sort(key=lambda p: p.top)

    linhas: list[list[Palavra]] = []
    for p in candidatas:
        if linhas and abs(p.top - linhas[-1][0].top) <= TOLERANCIA_LINHA:
            linhas[-1].append(p)
        else:
            linhas.append([p])
    for linha in linhas:
        linha.sort(key=lambda p: p.x0)
    return linhas


def _identificar_colunas(linha_cabecalho: list[Palavra]) -> list[tuple[str, float]]:
    """Acha, na linha de cabecalho, a palavra que melhor identifica cada
    coluna esperada (na ordem esquerda->direita da lista COLUNAS) e usa o
    x0 dessa palavra como ancora. A busca avanca sequencialmente (nunca
    volta pra tras) -- e o que evita que um rotulo ambiguo tipo "ICMS"
    (que aparece dentro de "B.CALC.ICMS", "VALOR ICMS" e "ICMS/IPI") seja
    associado a coluna errada: cada rotulo so pode casar com a primeira
    palavra ainda nao usada que bater, da esquerda pra direita."""
    indice_min = 0
    ancoras: list[tuple[str, float]] = []
    for nome, variantes in COLUNAS:
        for i in range(indice_min, len(linha_cabecalho)):
            texto_lower = linha_cabecalho[i].texto.lower()
            if any(v in texto_lower for v in variantes):
                ancoras.append((nome, linha_cabecalho[i].x0))
                indice_min = i + 1
                break
    return ancoras


def _achar_linha_cabecalho(linhas: list[list[Palavra]]) -> Optional[list[Palavra]]:
    """A linha de cabecalho da tabela de itens e a que tem mais palavras
    batendo com os rotulos de coluna esperados."""
    melhor_linha = None
    melhor_pontuacao = 0
    for linha in linhas:
        pontuacao = len(_identificar_colunas(linha))
        if pontuacao > melhor_pontuacao:
            melhor_pontuacao = pontuacao
            melhor_linha = linha
    # Exige pelo menos metade das colunas conhecidas pra confiar que e
    # mesmo o cabecalho da tabela (e nao uma linha qualquer que por acaso
    # contem uma ou duas palavras parecidas).
    if melhor_pontuacao < len(COLUNAS) // 2:
        return None
    return melhor_linha


def _limites_colunas(ancoras: list[tuple[str, float]]) -> list[tuple[str, float, float]]:
    """Transforma [(nome, x0_ancora), ...] em [(nome, inicio, fim), ...],
    usando o ponto medio entre ancoras consecutivas como fronteira. A
    primeira coluna nao tem limite inferior e a ultima nao tem limite
    superior (dado real pode comecar antes do x0 do rotulo do cabecalho --
    ex: token "460-BCZS" em x0=85, antes do rotulo "CODIGO" em x0=98.7)."""
    ancoras_ordenadas = sorted(ancoras, key=lambda a: a[1])
    limites = []
    for i, (nome, x0) in enumerate(ancoras_ordenadas):
        inicio = float("-inf") if i == 0 else (ancoras_ordenadas[i - 1][1] + x0) / 2
        fim = (
            float("inf")
            if i == len(ancoras_ordenadas) - 1
            else (x0 + ancoras_ordenadas[i + 1][1]) / 2
        )
        limites.append((nome, inicio, fim))
    return limites


def _coluna_da_palavra(p: Palavra, limites: list[tuple[str, float, float]]) -> Optional[str]:
    """Usa o x0 da palavra (a borda esquerda) pra decidir a coluna --
    coerente com como as ancoras do cabecalho tambem sao o x0 de onde o
    rotulo comeca."""
    for nome, inicio, fim in limites:
        if inicio <= p.x0 < fim:
            return nome
    return None


def _tipo_linha(valores_coluna: dict[str, list[Palavra]]) -> str:
    """Classifica uma linha abaixo do cabecalho da tabela:

    - "item": tem conteudo na coluna CODIGO -- todo item de verdade tem
      codigo de produto, entao esse e o sinal mais confiavel de que a
      linha e um item novo (mais confiavel do que "tem algum numero em
      alguma coluna", que rodape/secao seguinte tambem pode ter por
      coincidencia de posicao -- foi bug real: "Valor Total dos
      Produtos..." caindo na faixa de x da coluna NCM virava um item
      fantasma vazio).
    - "continuacao": so tem conteudo em DESCRICAO -- quebra de linha
      dentro da celula de descricao de um item que ja comecou.
    - "fim": qualquer outra coisa (rodape, secao seguinte, etc.) -- sinal
      de que a tabela acabou.
    """
    if valores_coluna.get("codigo"):
        return "item"
    chaves = set(valores_coluna.keys())
    if chaves and chaves <= {"descricao"}:
        return "continuacao"
    return "fim"


def _separar_valores_colados(texto: str) -> list[str]:
    """Uma coluna estreita pode fazer o pdfplumber juntar dois valores
    monetarios num so token sem espaco entre eles (ex: "13,9718,00" =
    "13,97" + "18,00"). Separa por PADRAO de valor monetario, nao por
    posicao x -- mais robusto do que tentar adivinhar uma fronteira de
    pixel pra uma sub-coluna que o cabecalho nem rotula separadamente."""
    valores = VALOR_MONETARIO_RE.findall(texto)
    return valores if valores else [texto]


def _montar_item(valores_coluna: dict[str, list[Palavra]]) -> ItemDocumento:
    descricao_palavras = sorted(valores_coluna.get("descricao", []), key=lambda p: p.x0)
    descricao = " ".join(p.texto for p in descricao_palavras).strip()

    def _valor_numerico(nome_coluna: str):
        palavras = valores_coluna.get(nome_coluna, [])
        if not palavras:
            return None
        bruto = "".join(p.texto for p in sorted(palavras, key=lambda p: p.x0))
        partes = _separar_valores_colados(bruto)
        return comum.para_numero(partes[0]) if partes else None

    return ItemDocumento(
        descricao=descricao or "(descrição não identificada)",
        quantidade=_valor_numerico("quantidade"),
        valor_unitario=_valor_numerico("valor_unitario"),
        valor_total=_valor_numerico("valor_total"),
    )


@dataclass
class TabelaItensResultado:
    itens: list[ItemDocumento]
    cfop_predominante: Optional[str]
    soma_valor_total: Optional[float]


def montar_tabela_itens(paginas_palavras: list[list[Palavra]]) -> TabelaItensResultado:
    """Reconstroi a tabela de itens de uma DANFE a partir das palavras
    posicionadas de cada pagina. Nunca levanta excecao -- se nao achar uma
    linha de cabecalho reconhecivel, devolve uma tabela vazia (quem chama
    decide se isso vira aviso)."""
    itens: list[ItemDocumento] = []
    cfops: list[str] = []

    for palavras_pagina in paginas_palavras:
        linhas = _agrupar_linhas(palavras_pagina)
        linha_cabecalho = _achar_linha_cabecalho(linhas)
        if linha_cabecalho is None:
            continue

        ancoras = _identificar_colunas(linha_cabecalho)
        limites = _limites_colunas(ancoras)
        top_cabecalho = linha_cabecalho[0].top

        item_atual: dict[str, list[Palavra]] | None = None
        for linha in linhas:
            if linha[0].top <= top_cabecalho:
                continue  # cabecalho ou algo acima dele

            valores_coluna: dict[str, list[Palavra]] = {}
            for p in linha:
                nome_coluna = _coluna_da_palavra(p, limites)
                if nome_coluna:
                    valores_coluna.setdefault(nome_coluna, []).append(p)

            tipo = _tipo_linha(valores_coluna)

            if tipo == "continuacao":
                if item_atual:
                    item_atual.setdefault("descricao", []).extend(valores_coluna["descricao"])
                continue

            if tipo == "fim":
                break

            # tipo == "item": comeca um item novo
            if item_atual:
                itens.append(_montar_item(item_atual))
            item_atual = valores_coluna

            cfop_palavras = valores_coluna.get("cfop")
            if cfop_palavras:
                cfops.append("".join(p.texto for p in cfop_palavras))

        if item_atual:
            itens.append(_montar_item(item_atual))

    cfop_predominante = statistics.mode(cfops) if cfops else None
    valores_totais = [item.valor_total for item in itens if isinstance(item.valor_total, float)]
    soma = round(sum(valores_totais), 2) if valores_totais else None

    return TabelaItensResultado(itens=itens, cfop_predominante=cfop_predominante, soma_valor_total=soma)
