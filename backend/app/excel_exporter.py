"""Exportacao para Excel (.xlsx) com openpyxl.

Recebe uma LISTA de documentos (hoje a tela manda 1; a estrutura ja comporta
lote) e gera 4 abas, sempre presentes -- estrutura estavel pra quem usa Power
Query/formulas, mesmo que uma aba fique so com o cabecalho:

- Resumo:            uma linha por documento
- Itens:             todos os itens de todos os documentos
- Campos adicionais: documento, campo, valor, confianca
- Avisos:            um aviso por linha

Toda aba tem a coluna ID (sequencial, 1..N) como CHAVE DE LIGACAO -- tipo +
numero pode colidir entre emissores diferentes -- e a coluna Documento
(descricao legivel: "Nota fiscal 000012345").

Decisoes que evitam bugs reais de planilha:
- FORMATO MONETARIO: o codigo guardado no arquivo usa a notacao en-US
  (`"R$" #,##0.00`); o Excel/LibreOffice LOCALIZAM na exibicao (pt-BR mostra
  `R$ 1.234,56`). Escrever `#.##0,00` literalmente quebraria o formato.
- SO valores monetarios viram numero: valores/totais do documento e dos itens
  (quando ja sao float) e 3 campos adicionais conhecidos. Chave de acesso (44
  digitos; Excel so guarda 15 de precisao), Nosso Numero, CFOP, linha
  digitavel... ficam TEXTO. Nunca "parece numero -> converte".
- TEXTO E SEMPRE TEXTO: o openpyxl trata string que comeca com "=" como
  FORMULA, e o texto vem de PDF (nao confiavel): um emissor "=HYPERLINK(...)"
  viraria formula executavel na maquina de quem abre. Forcamos data_type "s".
  Caracteres de controle (que o openpyxl rejeita) sao removidos.
- Datas viram data de verdade (dd/mm/aaaa); o que nao der pra converter fica
  como texto (o modo IA pode devolver formatos inesperados).

Identidade visual (etapa 7): cabecalho, zebra, bordas, alinhamento por tipo,
destaque do Valor total, linha de total dos Itens, congelar coluna ID e cor
da guia -- ver docstrings de `_escrever_aba` e `_escrever_linha_total_itens`.
NADA disso muda estrutura de dados ou tipo de celula, so apresentacao.

Acabamento "senior" (etapa 8), tambem so apresentacao:
- Cada aba de dados (Resumo/Itens/Campos adicionais/Avisos) vira uma TABELA
  NOMEADA do Excel (`openpyxl.worksheet.table.Table`), no lugar do
  auto_filter solto de antes -- aparece no Gerenciador de Nomes e serve de
  origem pronta pra formula (`=SOMA(Itens[Valor total])`) ou Tabela
  Dinamica. O estilo visual PROPRIO da Tabela vem todo desligado
  (`TableStyleInfo(showRowStripes=False, ...)`) de proposito: sem isso, o
  listrado embutido do Excel brigaria com a zebra e os destaques de
  confianca que ja existem. Confirmado empiricamente (script isolado, nao
  faz parte da suite) que Tabela + freeze_panes + auto_filter antigo
  convivem sem erro no round-trip do openpyxl, mas manter os dois seria
  redundante -- fica so a Tabela, do mesmo jeito que o Excel faz sozinho
  quando voce usa Inserir > Tabela em cima de um auto_filter existente.
- Aba "Legenda" (5a aba, sempre por ultimo, sem cor de guia -- as outras 4
  tem a mesma cor do cabecalho, essa fica neutra de proposito pra sinalizar
  "isto e texto de apoio, nao dado"): explica as 3 cores que aparecem nas
  outras abas (media/baixa/corrigido), com a cor de verdade ao lado do
  texto, nao so descrita em palavras.
- Metadados do arquivo (`wb.properties`): titulo, autor e data de geracao.
  `creator` e o NOME DO APLICATIVO, nunca uma pessoa -- o arquivo sai de uma
  demo publica e pode conter documento de qualquer um, carimbar um nome
  pessoal no metadado nao faz sentido nesse contexto.
- Impressao: area de impressao (so a tabela real, sem a nota de geracao),
  cabecalho repetido em toda pagina, paisagem e ajuste de escala na largura
  nas 5 abas.
- Coerencia Resumo x Itens: uma comparacao literal "soma dos itens == Valor
  total do Resumo" foi COGITADA E DESCARTADA -- esta errada pra DANFE (o
  Valor total da nota inclui ICMS ST, frete e IPI, nunca deveria bater com
  a soma pura dos produtos; um alarme assim dispararia em documento
  correto). Em vez disso: todo documento com 1+ avisos ganha um comentario
  na celula "Documento" do Resumo apontando pra aba Avisos (sinal que ja
  existia, so nao aparecia fora da aba Avisos); e o extrator da DANFE
  (`danfe.py`) passou a expor "Valor Total dos Produtos" como mais um
  campos_adicionais (mesmo padrao de CFOP/Chave de Acesso -- nao muda a
  estrutura do schema), pra dar um numero de verdade pra comparar com a
  soma dos itens em vez de so confiar no texto do aviso.

Decisao importante: NAO existe linha de titulo mesclada acima do cabecalho.
Foi cogitada e testada (empiricamente, com pandas.read_excel simulando um
leitor automatico tipo Power Query) e DESCARTADA: auto_filter e freeze_panes
sobrevivem tecnicamente, mas qualquer ferramenta que assuma "linha 1 =
cabecalho" (Power Query "usar primeira linha como cabecalhos", pandas com
header=0 default) le a linha de titulo como se fosse o cabecalho de
verdade -- os nomes de coluna reais (ID, Documento...) viram uma linha de
DADO. Isso corrompe a leitura automatica pra exatamente o publico que este
arquivo promete atender (ver topo do docstring: "Power Query/formulas").
Em vez disso, a data de geracao vai numa nota discreta na aba Resumo, na
propria linha do cabecalho, 2 colunas a direita da tabela -- fora do
auto_filter e das colunas reais, entao um leitor automatico ve no maximo
2 colunas extras no fim (nao corrompe as colunas verdadeiras).
"""
import re
from dataclasses import dataclass
from datetime import date, datetime
from io import BytesIO
from typing import Any, Optional

from openpyxl import Workbook
from openpyxl.comments import Comment
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.table import Table, TableStyleInfo

from app.confianca import NIVEIS, resumo_confianca, vazio
from app.schemas import DocumentoParaExportar

NOME_APLICATIVO = "Extrator Inteligente de Documentos"

FORMATO_MOEDA = '"R$" #,##0.00'
FORMATO_MOEDA_UNITARIO = '"R$" #,##0.00##'  # preco unitario pode ter ate 4 casas
FORMATO_DATA = "dd/mm/yyyy"
FORMATO_TEXTO = "@"

# Mesmos tons da interface (style.css): warning-bg, danger-bg e accent-soft.
FUNDOS = {"media": "FDF3E0", "baixa": "FBECEB", "corrigido": "E7F0EF"}

# Identidade visual: mesmos tokens de cor do frontend (style.css), duplicados
# aqui de proposito -- nao ha arquivo de tokens compartilhado entre Python e
# CSS neste projeto (mesma duplicacao intencional que ja existe pra regra de
# confianca geral, ver app/confianca.py).
COR_CABECALHO_FUNDO = "0F4C4C"  # --color-accent
COR_CABECALHO_TEXTO = "FFFFFF"  # --color-accent-contrast
COR_ZEBRA = "F5F6F7"  # --color-bg
COR_BORDA = "E3E6EA"  # --color-border
COR_DESTAQUE_VALOR_TOTAL = "0A3838"  # --color-accent-dark
COR_NOTA_GERACAO = "667080"  # --color-text-muted
ALTURA_LINHA_CABECALHO = 24
TAMANHO_FONTE_VALOR_TOTAL = 13

ROTULOS_TIPO = {
    "boleto": "Boleto",
    "nota_fiscal": "Nota fiscal",
    "pedido_compra": "Pedido de compra",
    "relatorio": "Relatório",
    "desconhecido": "Desconhecido",
}
ROTULOS_NIVEL = {"alta": "Alta", "media": "Média", "baixa": "Baixa", "corrigido": "Corrigido"}

# Campos adicionais que sao VALOR MONETARIO (viram numero). O resto e codigo/texto.
CAMPOS_MONETARIOS = {"Valor do Documento", "Desconto", "Valor a Pagar", "Valor Total dos Produtos"}

CABECALHOS_RESUMO = [
    "ID", "Arquivo", "Documento", "Tipo", "Número", "Data de emissão", "Data de vencimento",
    "Emissor", "Emissor CNPJ/CPF", "Destinatário", "Destinatário CNPJ/CPF", "Valor total", "Confiança geral",
]
CABECALHOS_ITENS = ["ID", "Documento", "Descrição", "Quantidade", "Valor unitário", "Valor total"]
CABECALHOS_CAMPOS = ["ID", "Documento", "Campo", "Valor", "Confiança"]
CABECALHOS_AVISOS = ["ID", "Documento", "Aviso"]

# nome da aba -> nome da Tabela nomeada (Excel nao aceita espaço em nome de tabela)
NOMES_TABELA = {"Resumo": "Resumo", "Itens": "Itens", "Campos adicionais": "CamposAdicionais", "Avisos": "Avisos"}

# aba de apoio (Legenda): cor de verdade ao lado do texto, nao so descrita em palavras
LEGENDA_LINHAS = [
    ("media", 'Confiança média: valor inferido pela posição no documento, sem rótulo explícito. Confira.'),
    ("baixa", "Confiança baixa: sem rótulo que confirme o valor. Confira e corrija antes de usar."),
    ("corrigido", "Corrigido manualmente na tela antes de exportar (também em itálico). "
                  "O valor original extraído fica no comentário da própria célula."),
]

LARGURA_MINIMA, LARGURA_MAXIMA = 8, 60
_CONTROLE_RE = re.compile(r"[\x00-\x08\x0b\x0c\x0e-\x1f]")
_LIMITE_TEXTO_EXCEL = 32767
_DOC_FISCAL_RE = re.compile(r"^(?P<nome>.*?)\s*\((?P<tipo>CNPJ|CPF)\s+(?P<doc>[^()]+?)\)\s*$", re.DOTALL)
_DATA_BR_RE = re.compile(r"^(\d{1,2})[/-](\d{1,2})[/-](\d{4})$")
_DATA_ISO_RE = re.compile(r"^(\d{4})-(\d{2})-(\d{2})$")


@dataclass
class Celula:
    valor: Any = None
    formato: Optional[str] = None
    estado: Optional[str] = None  # "media" | "baixa" | "corrigido": fundo (+ italico se corrigido)
    comentario: Optional[str] = None


# ---------- conversoes ----------


def _texto_seguro(texto: str) -> str:
    return _CONTROLE_RE.sub("", texto)[:_LIMITE_TEXTO_EXCEL]


def _numero_br(texto: str) -> Optional[float]:
    """"1.234,56" | "1234,56" | "1234.56" | "R$ 1.234,56" -> float; senao None
    (mesmas regras do parseNumeroBR da tela)."""
    t = re.sub(r"\s+", "", re.sub(r"^R\$\s*", "", texto.strip(), flags=re.IGNORECASE))
    if not t:
        return None
    try:
        if re.fullmatch(r"-?\d{1,3}(\.\d{3})+(,\d+)?", t):
            return float(t.replace(".", "").replace(",", "."))
        if re.fullmatch(r"-?\d+,\d+", t):
            return float(t.replace(",", "."))
        if re.fullmatch(r"-?\d+(\.\d+)?", t):
            return float(t)
    except ValueError:
        pass
    return None


def _para_data(texto: str) -> Optional[date]:
    t = texto.strip()
    try:
        m = _DATA_BR_RE.match(t)
        if m:
            return date(int(m.group(3)), int(m.group(2)), int(m.group(1)))
        m = _DATA_ISO_RE.match(t)
        if m:
            return date(int(m.group(1)), int(m.group(2)), int(m.group(3)))
    except ValueError:  # dia/mes inexistente (31/02...)
        pass
    return None


def separar_documento_fiscal(valor: Optional[str]) -> tuple[Optional[str], Optional[str]]:
    """"NOME (CNPJ 72.381...)" -> ("NOME", "72.381..."). Sem o parentese
    final, o valor inteiro e o nome. Pega o ULTIMO parentese, entao nomes com
    parenteses proprios ("EMPRESA (FILIAL) LTDA (CNPJ ...)") funcionam."""
    if vazio(valor):
        return None, None
    m = _DOC_FISCAL_RE.match(valor.strip())
    if not m or not m.group("nome").strip():
        return valor.strip(), None
    return m.group("nome").strip(), m.group("doc").strip()


def _texto_original(valor) -> str:
    if vazio(valor):
        return "(vazio)"
    if isinstance(valor, float):
        return f"{valor:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")
    return str(valor)


def rotulo_documento(tipo: str, numero: Optional[str], arquivo: Optional[str]) -> str:
    """Descricao legivel: "Nota fiscal 000012345". Sem numero cai no nome do
    arquivo (ou "(sem numero)")."""
    nome_tipo = ROTULOS_TIPO.get(tipo, tipo)
    if not vazio(numero):
        return f"{nome_tipo} {str(numero).strip()}"
    if not vazio(arquivo):
        return f"{nome_tipo} — {arquivo}"
    return f"{nome_tipo} (sem número)"


# ---------- montagem das celulas ----------


def _comentario(estado: Optional[str], original=None) -> Optional[str]:
    if estado == "corrigido":
        return f"Corrigido pelo usuário. Valor extraído: {_texto_original(original)}"
    if estado == "media":
        return "Confiança média: inferido pela posição no documento. Confira."
    if estado == "baixa":
        return "Confiança baixa: sem rótulo que confirme. Confira e corrija."
    return None


def _estado_visual(chave: str, valor, doc: DocumentoParaExportar) -> Optional[str]:
    """Estado que pinta a celula: corrigido (mesmo se o usuario esvaziou o
    campo) > baixa > media. "alta" e sem confianca nao pintam."""
    if chave in doc.corrigidos:
        return "corrigido"
    if vazio(valor):
        return None
    nivel = doc.resultado.confiancas.get(chave)
    return nivel if nivel in ("media", "baixa") else None


def _celula(valor, chave: str, doc: DocumentoParaExportar, formato: Optional[str] = None) -> Celula:
    estado = _estado_visual(chave, valor, doc)
    return Celula(valor, formato, estado, _comentario(estado, doc.corrigidos.get(chave)))


def _celula_data(texto: Optional[str], chave: str, doc: DocumentoParaExportar) -> Celula:
    if vazio(texto):
        return _celula(None, chave, doc)
    data = _para_data(texto)
    return _celula(data if data else texto, chave, doc, FORMATO_DATA if data else None)


def _celula_numero(valor, chave: str, doc: DocumentoParaExportar) -> Celula:
    """float -> numero moeda; string (conversao ja falhou antes) fica texto."""
    if isinstance(valor, float) or isinstance(valor, int):
        return _celula(valor, chave, doc, FORMATO_MOEDA)
    return _celula(valor, chave, doc)


def _linha_resumo(id_doc: int, doc: DocumentoParaExportar, rotulo: str) -> list[Celula]:
    documento = doc.resultado.documento
    nome_emissor, doc_emissor = separar_documento_fiscal(documento.emissor)
    nome_dest, doc_dest = separar_documento_fiscal(documento.destinatario)
    resumo = resumo_confianca(doc.resultado, doc.corrigidos.keys())

    def par(chave, valor, nome, doc_fiscal):
        """Nome + CNPJ/CPF: as duas celulas levam o mesmo destaque; o comentario
        (com o valor original completo) fica so na do nome."""
        c_nome = _celula(nome, chave, doc)
        c_doc = Celula(doc_fiscal, None, c_nome.estado, None)
        return c_nome, c_doc

    emissor, emissor_doc = par("emissor", documento.emissor, nome_emissor, doc_emissor)
    destinatario, destinatario_doc = par("destinatario", documento.destinatario, nome_dest, doc_dest)

    # coerencia Resumo x Itens/Avisos: quem so olha o Resumo nao teria como
    # saber que este documento tem aviso (ex: soma dos itens divergente) sem
    # esse comentario -- ver docstring do modulo sobre a comparacao literal
    # de valores, que foi descartada por estar errada pra DANFE
    avisos_doc = doc.resultado.avisos or ([doc.resultado.aviso] if doc.resultado.aviso else [])
    comentario_avisos = None
    if avisos_doc:
        plural = "s" if len(avisos_doc) > 1 else ""
        comentario_avisos = f"Este documento tem {len(avisos_doc)} aviso{plural} -- ver a aba Avisos."

    return [
        Celula(id_doc, "0"),
        Celula(doc.arquivo),
        Celula(valor=rotulo, comentario=comentario_avisos),
        _celula(documento.tipo_documento, "tipo_documento", doc),  # valor interno (nota_fiscal), bom pra filtrar
        _celula(documento.numero_documento, "numero_documento", doc),
        _celula_data(documento.data_emissao, "data_emissao", doc),
        _celula_data(documento.data_vencimento, "data_vencimento", doc),
        emissor,
        emissor_doc,
        destinatario,
        destinatario_doc,
        _celula_numero(documento.valor_total, "valor_total", doc),
        Celula(resumo.texto if resumo else "—"),
    ]


def _linhas_itens(id_doc: int, doc: DocumentoParaExportar, rotulo: str) -> list[list[Celula]]:
    nivel = doc.resultado.confiancas.get("itens")
    estado = nivel if nivel in ("media", "baixa") else None  # tabela de itens lida por posicao
    linhas = []
    for i, item in enumerate(doc.resultado.documento.itens):
        def numero(valor, formato):
            # so numero de verdade recebe formato; texto (conversao falhou) fica texto
            return Celula(valor, formato if isinstance(valor, (int, float)) else None, estado)

        linhas.append([
            Celula(id_doc, "0"),
            Celula(rotulo),
            # um comentario so, na 1a linha de cada documento (evita poluir a aba)
            Celula(item.descricao, None, estado, _comentario(estado) if i == 0 else None),
            numero(item.quantidade, None),
            numero(item.valor_unitario, FORMATO_MOEDA_UNITARIO),
            numero(item.valor_total, FORMATO_MOEDA),
        ])
    return linhas


def _linhas_campos(id_doc: int, doc: DocumentoParaExportar, rotulo: str) -> list[list[Celula]]:
    linhas = []
    for extra in doc.resultado.documento.campos_adicionais:
        chave, valor = extra.campo, extra.valor
        estado = _estado_visual(chave, valor, doc)
        if chave in CAMPOS_MONETARIOS and _numero_br(valor) is not None:
            celula_valor = Celula(_numero_br(valor), FORMATO_MOEDA, estado, _comentario(estado, doc.corrigidos.get(chave)))
        else:
            celula_valor = Celula(valor, None, estado, _comentario(estado, doc.corrigidos.get(chave)))

        if chave in doc.corrigidos:
            nivel = "corrigido"
        else:
            nivel = doc.resultado.confiancas.get(chave)
        texto_nivel = ROTULOS_NIVEL.get(nivel, "—") if nivel in NIVEIS or nivel == "corrigido" else "—"
        linhas.append([
            Celula(id_doc, "0"),
            Celula(rotulo),
            Celula(chave),
            celula_valor,
            Celula(texto_nivel, None, estado),
        ])
    return linhas


def _linhas_avisos(id_doc: int, doc: DocumentoParaExportar, rotulo: str) -> list[list[Celula]]:
    resultado = doc.resultado
    avisos = resultado.avisos or ([resultado.aviso] if resultado.aviso else [])
    return [[Celula(id_doc, "0"), Celula(rotulo), Celula(aviso)] for aviso in avisos]


# ---------- escrita e formatacao ----------


def _largura_visual(celula: Celula) -> int:
    valor = celula.valor
    if valor is None:
        return 0
    if isinstance(valor, date):
        return 10
    if isinstance(valor, (int, float)) and not isinstance(valor, bool):
        return len(f"{valor:,.2f}") + 4 if celula.formato and "R$" in celula.formato else len(str(valor))
    return max((len(linha) for linha in str(valor).splitlines()), default=0)


def _fonte(negrito: bool = False, italico: bool = False, cor: Optional[str] = None, tamanho: Optional[int] = None) -> Font:
    """Monta a fonte de uma celula com todos os atributos de uma vez --
    existe pra combinar negrito/italico/cor/tamanho sem que uma atribuicao
    de Font() apague a anterior (ex: destaque do Valor total + corrigido)."""
    args: dict[str, Any] = {"bold": negrito, "italic": italico}
    if cor:
        args["color"] = cor
    if tamanho:
        args["size"] = tamanho
    return Font(**args)


def _borda_padrao() -> Border:
    lado = Side(style="thin", color=COR_BORDA)
    return Border(left=lado, right=lado, top=lado, bottom=lado)


def _tabela_sem_estilo_proprio(nome: str, ref: str) -> Table:
    """Tabela nomeada do Excel com o estilo visual PROPRIO desligado --
    sem isso, o listrado embutido do Excel brigaria com a zebra e os
    destaques de confianca que ja existem (ver docstring do modulo)."""
    tabela = Table(displayName=nome, ref=ref)
    tabela.tableStyleInfo = TableStyleInfo(
        name=None, showRowStripes=False, showColumnStripes=False, showFirstColumn=False, showLastColumn=False
    )
    return tabela


def _configurar_impressao(ws, ultima_coluna: str, ultima_linha: int) -> None:
    """Area de impressao, cabecalho repetido em toda pagina, paisagem e
    ajuste de escala na largura -- uma planilha de nota fiscal vai ser
    impressa. `pageSetUpPr.fitToPage` precisa ser ligado a parte: sem ele o
    Excel ignora fitToWidth/fitToHeight (pegadinha conhecida do formato)."""
    ws.print_area = f"A1:{ultima_coluna}{ultima_linha}"
    ws.print_title_rows = "1:1"
    ws.page_setup.orientation = "landscape"
    ws.page_setup.fitToWidth = 1
    ws.page_setup.fitToHeight = 0
    ws.sheet_properties.pageSetUpPr.fitToPage = True
    ws.print_options.horizontalCentered = True


def _escrever_aba(ws, cabecalhos: list[str], linhas: list[list[Celula]], nome_tabela: str) -> None:
    """Cabecalho com fundo escuro/texto branco, zebra nas linhas de dado
    (substituida -- nunca somada -- pelo destaque de confianca quando a
    celula tem um), bordas finas em toda celula, alinhamento por tipo
    (texto a esquerda, numero/data a direita, com recuo), congela cabecalho
    E coluna ID (`freeze_panes = "B2"`), vira uma Tabela nomeada (no lugar
    do auto_filter solto) e pinta a guia da aba."""
    borda = _borda_padrao()
    for coluna, titulo in enumerate(cabecalhos, start=1):
        celula = ws.cell(row=1, column=coluna, value=titulo)
        celula.font = _fonte(negrito=True, cor=COR_CABECALHO_TEXTO)
        celula.fill = PatternFill("solid", start_color=COR_CABECALHO_FUNDO, end_color=COR_CABECALHO_FUNDO)
        celula.border = borda
        celula.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
    ws.row_dimensions[1].height = ALTURA_LINHA_CABECALHO

    for numero_linha, linha in enumerate(linhas, start=2):
        zebra = numero_linha % 2 == 1  # 1a linha de dado (par) sem zebra, alterna dai
        for coluna, dados in enumerate(linha, start=1):
            celula = ws.cell(row=numero_linha, column=coluna)
            valor = dados.valor
            if isinstance(valor, str):
                valor = _texto_seguro(valor)
                if valor:
                    celula.value = valor
                    celula.data_type = "s"  # NUNCA formula (ver docstring do modulo)
                    celula.number_format = FORMATO_TEXTO
                celula.alignment = Alignment(horizontal="left", vertical="center", wrap_text=True, indent=1)
            elif valor is not None:
                celula.value = valor
                if dados.formato:
                    celula.number_format = dados.formato
                celula.alignment = Alignment(horizontal="right", vertical="center", indent=1)
            else:
                celula.alignment = Alignment(horizontal="left", vertical="center")
            celula.border = borda

            if dados.estado in FUNDOS:
                cor = FUNDOS[dados.estado]
                celula.fill = PatternFill("solid", start_color=cor, end_color=cor)
                if dados.estado == "corrigido":
                    celula.font = _fonte(italico=True)
            elif zebra:
                celula.fill = PatternFill("solid", start_color=COR_ZEBRA, end_color=COR_ZEBRA)
            if dados.comentario:
                comentario = Comment(dados.comentario, "Extrator")
                comentario.width, comentario.height = 280, 90
                celula.comment = comentario

    for coluna, titulo in enumerate(cabecalhos, start=1):
        maior = max([len(titulo)] + [_largura_visual(linha[coluna - 1]) for linha in linhas])
        ws.column_dimensions[get_column_letter(coluna)].width = min(max(maior + 2, LARGURA_MINIMA), LARGURA_MAXIMA)

    ws.freeze_panes = "B2"  # cabecalho E coluna ID congelados
    # So cria a Tabela quando ha pelo menos 1 linha de dados. Bug real: uma
    # Tabela com ref so no cabecalho (aba vazia -- ex: Avisos de quase todo
    # documento, Itens de boleto) e valida pelo schema, passa no openpyxl e no
    # pandas, e o Excel de verdade recusa o arquivo e oferece "reparar",
    # descartando a formatacao. O proprio Excel nunca grava tabela sem linha de
    # dados. Aba sem dados fica sem Tabela (nao ha o que filtrar). Ver
    # tests/ooxml.py.
    if linhas:
        ws.add_table(_tabela_sem_estilo_proprio(nome_tabela, f"A1:{get_column_letter(len(cabecalhos))}{len(linhas) + 1}"))
    ws.sheet_properties.tabColor = COR_CABECALHO_FUNDO


def _destacar_valor_total(ws, linhas: list[list[Celula]]) -> None:
    """Valor total (Resumo) em negrito, fonte maior e cor de destaque --
    e o campo mais consultado do relatorio. Roda DEPOIS de _escrever_aba:
    so troca a fonte (mantem o fundo, seja zebra ou destaque de confianca)
    e preserva o italico de "corrigido"."""
    coluna = CABECALHOS_RESUMO.index("Valor total") + 1
    for numero_linha, linha in enumerate(linhas, start=2):
        dados = linha[coluna - 1]
        celula = ws.cell(row=numero_linha, column=coluna)
        celula.font = _fonte(
            negrito=True, italico=(dados.estado == "corrigido"), cor=COR_DESTAQUE_VALOR_TOTAL, tamanho=TAMANHO_FONTE_VALOR_TOTAL
        )


def _nota_geracao(ws, num_colunas: int, data_geracao: date) -> None:
    """"Gerado em dd/mm/aaaa" discreto, na propria linha do cabecalho, 2
    colunas a direita da tabela (so na aba Resumo). Ficou de fora do
    auto_filter e das colunas reais de proposito -- ver docstring do modulo
    sobre por que NAO existe uma linha de titulo mesclada acima do
    cabecalho: leitores automaticos (Power Query, pandas) que assumem
    "linha 1 = cabecalho" continuam lendo as colunas verdadeiras direito;
    no maximo enxergam 2 colunas extras no fim."""
    celula = ws.cell(row=1, column=num_colunas + 2, value=f"Gerado em {data_geracao.strftime('%d/%m/%Y')}")
    celula.font = _fonte(italico=True, cor=COR_NOTA_GERACAO, tamanho=9)
    celula.alignment = Alignment(horizontal="left", vertical="center")


def _escrever_linha_total_itens(ws, n_itens: int) -> None:
    """Uma linha de total ao FINAL da aba Itens (geral, nao por documento --
    subtotal por documento fica pra quando existir interface de lote).
    Soma Quantidade e Valor total com formula nativa do Excel (=SOMA, que
    ignora celulas de texto automaticamente -- um item com quantidade "1 un"
    em vez de numero nao quebra a soma). A formula e gravada em ingles
    (`SUM`) porque o formato OOXML guarda formulas em sintaxe canonica
    independente de idioma -- o Excel em pt-BR EXIBE "=SOMA(...)" sozinho,
    mesmo padrao ja usado pro formato de moeda (grava en-US, Excel
    localiza). Valor unitario fica de fora: nao faz sentido somar preco
    unitario de itens diferentes."""
    col_qtd = CABECALHOS_ITENS.index("Quantidade") + 1
    col_total = CABECALHOS_ITENS.index("Valor total") + 1
    primeira, ultima = 2, n_itens + 1
    linha_total = n_itens + 2

    ws.merge_cells(start_row=linha_total, start_column=1, end_row=linha_total, end_column=col_qtd - 1)
    borda_topo = Border(top=Side(style="thin", color=COR_BORDA))
    for coluna in range(1, len(CABECALHOS_ITENS) + 1):
        ws.cell(row=linha_total, column=coluna).border = borda_topo

    rotulo = ws.cell(row=linha_total, column=1, value="Total")
    rotulo.font = _fonte(negrito=True)
    rotulo.alignment = Alignment(horizontal="right", vertical="center", indent=1)

    for coluna, formato in ((col_qtd, None), (col_total, FORMATO_MOEDA)):
        letra = get_column_letter(coluna)
        celula = ws.cell(row=linha_total, column=coluna, value=f"=SUM({letra}{primeira}:{letra}{ultima})")
        if formato:
            celula.number_format = formato
        celula.font = _fonte(negrito=True)
        celula.alignment = Alignment(horizontal="right", vertical="center", indent=1)


def _escrever_legenda(ws) -> None:
    """Aba de apoio, estatica, sem Tabela/filtro/zebra/cor de guia (sinaliza
    "isto e texto de apoio, nao dado"). Uma celula com a cor DE VERDADE ao
    lado do texto -- ver docstring do modulo sobre a decisao de usar uma
    aba em vez de repetir a explicacao em rodape nas 4 abas de dados."""
    for coluna, titulo in enumerate(("Cor", "Significado"), start=1):
        celula = ws.cell(row=1, column=coluna, value=titulo)
        celula.font = _fonte(negrito=True, cor=COR_CABECALHO_TEXTO)
        celula.fill = PatternFill("solid", start_color=COR_CABECALHO_FUNDO, end_color=COR_CABECALHO_FUNDO)
        celula.border = _borda_padrao()
    ws.row_dimensions[1].height = ALTURA_LINHA_CABECALHO

    for numero_linha, (estado, texto) in enumerate(LEGENDA_LINHAS, start=2):
        cor = FUNDOS[estado]
        # sem value: value="" vira uma celula tipada como texto SEM texto
        # (<c t="inlineStr"/> sem <is>), um estado que o Excel nao produz
        cel_cor = ws.cell(row=numero_linha, column=1)
        cel_cor.fill = PatternFill("solid", start_color=cor, end_color=cor)
        cel_cor.border = _borda_padrao()
        cel_texto = ws.cell(row=numero_linha, column=2, value=texto)
        cel_texto.alignment = Alignment(horizontal="left", vertical="center", wrap_text=True, indent=1)
        cel_texto.border = _borda_padrao()
        if estado == "corrigido":
            cel_texto.font = _fonte(italico=True)

    ws.column_dimensions["A"].width = 10
    ws.column_dimensions["B"].width = 80
    ws.freeze_panes = "A2"
    # sem Tabela/filtro (nao e dado filtravel) e sem tabColor (guia neutra
    # de proposito, ao contrario das 4 abas de dados)


def gerar_excel(documentos: list[DocumentoParaExportar], data_geracao: Optional[date] = None) -> BytesIO:
    resumo, itens, campos, avisos = [], [], [], []
    for id_doc, doc in enumerate(documentos, start=1):
        documento = doc.resultado.documento
        rotulo = rotulo_documento(documento.tipo_documento, documento.numero_documento, doc.arquivo)
        resumo.append(_linha_resumo(id_doc, doc, rotulo))
        itens += _linhas_itens(id_doc, doc, rotulo)
        campos += _linhas_campos(id_doc, doc, rotulo)
        avisos += _linhas_avisos(id_doc, doc, rotulo)

    data_geracao = data_geracao or date.today()

    wb = Workbook()
    # metadados do arquivo (propriedades do documento): autor e o NOME DO
    # APLICATIVO, nunca uma pessoa -- ver docstring do modulo
    wb.properties.creator = NOME_APLICATIVO
    wb.properties.lastModifiedBy = NOME_APLICATIVO
    wb.properties.title = f"Relatório de extração — {NOME_APLICATIVO}"
    wb.properties.subject = "Extração de documentos fiscais"
    wb.properties.created = datetime.combine(data_geracao, datetime.min.time())
    # `modified` NAO e setado aqui de proposito: o openpyxl (writer/excel.py,
    # save_workbook()) sobrescreve com datetime.now() UTC no instante exato
    # do wb.save(), incondicionalmente -- nao ha como fixar via propriedade
    # antes de salvar (verificado lendo o codigo-fonte da biblioteca).
    # Decisao: deixar como esta. E o horario real de geracao do arquivo, nao
    # um dado do documento nem da pessoa que fez upload -- e exatamente o
    # que "modified" significa no padrao OOXML (todo .xlsx feito por
    # qualquer programa tem isso). Diferente de "created" (que reflete a
    # data logica do relatorio e por isso e controlavel/testavel), nao ha
    # motivo de privacidade pra mascarar quando o processo rodou de verdade.

    ws_resumo = wb.active
    ws_resumo.title = "Resumo"
    _escrever_aba(ws_resumo, CABECALHOS_RESUMO, resumo, NOMES_TABELA["Resumo"])
    _destacar_valor_total(ws_resumo, resumo)
    _nota_geracao(ws_resumo, len(CABECALHOS_RESUMO), data_geracao)
    _configurar_impressao(ws_resumo, get_column_letter(len(CABECALHOS_RESUMO)), max(len(resumo) + 1, 1))

    ws_itens = wb.create_sheet("Itens")
    _escrever_aba(ws_itens, CABECALHOS_ITENS, itens, NOMES_TABELA["Itens"])
    if itens:
        _escrever_linha_total_itens(ws_itens, len(itens))
    # area de impressao inclui a linha de total (ela e conteudo de leitura,
    # mesmo ficando fora da Tabela/filtro, que e sobre dado filtravel)
    _configurar_impressao(ws_itens, get_column_letter(len(CABECALHOS_ITENS)), ws_itens.max_row)

    ws_campos = wb.create_sheet("Campos adicionais")
    _escrever_aba(ws_campos, CABECALHOS_CAMPOS, campos, NOMES_TABELA["Campos adicionais"])
    _configurar_impressao(ws_campos, get_column_letter(len(CABECALHOS_CAMPOS)), max(len(campos) + 1, 1))

    ws_avisos = wb.create_sheet("Avisos")
    _escrever_aba(ws_avisos, CABECALHOS_AVISOS, avisos, NOMES_TABELA["Avisos"])
    _configurar_impressao(ws_avisos, get_column_letter(len(CABECALHOS_AVISOS)), max(len(avisos) + 1, 1))

    ws_legenda = wb.create_sheet("Legenda")
    _escrever_legenda(ws_legenda)
    _configurar_impressao(ws_legenda, "B", len(LEGENDA_LINHAS) + 1)

    buffer = BytesIO()
    wb.save(buffer)
    buffer.seek(0)
    return buffer
