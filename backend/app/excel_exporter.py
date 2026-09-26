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

Recursos nativos do Excel (etapa 10) -- detalhes e motivos nas docstrings de
cada funcao:
- Total do lote (`_escrever_linha_total_documentos`) e grafico de barras de
  Valor total por documento (`_grafico_valor_por_documento`): so com 2+
  documentos (2+ com valor numerico, no caso do grafico).
- Listas suspensas (`_lista_suspensa`) em Documentos!Tipo e Campos
  adicionais!Confianca -- as unicas colunas com conjunto fechado de valores.
- Formatacao condicional nativa SO na coluna Confianca de Campos adicionais
  (`_colorir_coluna_confianca`), onde o texto da celula e o proprio valor. Nas
  demais celulas destacadas a cor documenta a PROVENIENCIA da extracao e
  continua fixa: reescrever o campo nao pode apagar o rastro de "veio com
  confianca baixa". Cogitado e DESCARTADO trocar todos os destaques por
  regras: a maioria das celulas destacadas (Emissor, Valor total, Itens) nao
  tem, na propria linha, o dado de confianca que uma regra poderia ler.
- Protecao de planilha sem senha (`_proteger_planilha`): cabecalho e totais
  travados, dados livres, filtro/ordenacao liberados.

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
from openpyxl.chart import BarChart
from openpyxl.chart.data_source import AxDataSource, NumDataSource, NumRef, StrRef
from openpyxl.chart.label import DataLabelList
from openpyxl.chart.series import Series, SeriesLabel
from openpyxl.chart.shapes import GraphicalProperties
from openpyxl.comments import Comment
from openpyxl.drawing.line import LineProperties
from openpyxl.formatting.rule import CellIsRule
from openpyxl.styles import Alignment, Border, Font, PatternFill, Protection, Side
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.datavalidation import DataValidation
from openpyxl.worksheet.pagebreak import Break
from openpyxl.worksheet.table import Table, TableStyleInfo

from app.confianca import CAMPOS_OBRIGATORIOS, NIVEIS, e_obrigatorio, resumo_confianca, vazio
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
SEM_CONFIANCA = "—"  # campo sem informacao de confianca (ex: modo IA)

# Campos adicionais que sao VALOR MONETARIO (viram numero). O resto e codigo/texto.
CAMPOS_MONETARIOS = {"Valor do Documento", "Desconto", "Valor a Pagar", "Valor Total dos Produtos"}

CABECALHOS_RESUMO = [
    "ID", "Arquivo", "Documento", "Tipo", "Número", "Data de emissão", "Data de vencimento",
    "Emissor", "Emissor CNPJ/CPF", "Destinatário", "Destinatário CNPJ/CPF", "Valor total", "Confiança geral",
]
CABECALHOS_ITENS = ["ID", "Documento", "Descrição", "Quantidade", "Valor unitário", "Valor total"]
CABECALHOS_CAMPOS = ["ID", "Documento", "Campo", "Valor", "Confiança"]
CABECALHOS_AVISOS = ["ID", "Documento", "Aviso"]

# nome da aba -> nome da Tabela nomeada (Excel nao aceita espaço em nome de
# tabela). Chave = titulo real da aba (por isso "Documentos", nao "Resumo" --
# a aba foi renomeada na etapa 9). Identificadores internos que NAO sao nome
# de aba (CABECALHOS_RESUMO, _linha_resumo, ws_resumo...) continuam com o
# nome antigo de proposito, pra nao inflar o diff com renomeacoes que nao
# mudam comportamento nenhum.
NOME_ABA_DOCUMENTOS = "Documentos"
NOMES_TABELA = {
    NOME_ABA_DOCUMENTOS: "Documentos", "Itens": "Itens", "Campos adicionais": "CamposAdicionais", "Avisos": "Avisos"
}

# legenda de cor (media/baixa/corrigido): cor de verdade ao lado do texto, nao
# so descrita em palavras. Etapa 9: nao e mais uma aba a parte (Legenda foi
# removida) -- vira o rodape da aba Relatorio, e SO com as cores que aparecem
# de verdade no lote exportado (ver _rodape_legenda).
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
_CHAVE_ACESSO_RE = re.compile(r"^\d{44}$")


@dataclass
class Celula:
    valor: Any = None
    formato: Optional[str] = None
    estado: Optional[str] = None  # "media" | "baixa" | "corrigido": fundo (+ italico se corrigido)
    comentario: Optional[str] = None


# ---------- conversoes ----------


def _texto_seguro(texto: str) -> str:
    return _CONTROLE_RE.sub("", texto)[:_LIMITE_TEXTO_EXCEL]


def _formatar_chave_acesso(valor: str) -> str:
    """Mesma exibicao em blocos de 4 digitos que a tela usa pra Chave de
    Acesso -- so exibicao (na aba Campos adicionais E no Relatorio); o valor
    exportado no JSON/API continua sem espaco."""
    digitos = re.sub(r"\D", "", valor)
    if not _CHAVE_ACESSO_RE.match(digitos):
        return valor
    return " ".join(digitos[i:i + 4] for i in range(0, 44, 4))


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
        comentario_avisos = f"Este documento tem {len(avisos_doc)} aviso{plural} — veja a aba Avisos."

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
            # Chave de Acesso em blocos de 4 digitos -- mesma exibicao da
            # tela e do Relatorio (so exibicao; o valor exportado no
            # JSON/API continua sem espaco, ver DocumentoExtraido).
            valor_exibido = _formatar_chave_acesso(valor) if chave == "Chave de Acesso" else valor
            celula_valor = Celula(valor_exibido, None, estado, _comentario(estado, doc.corrigidos.get(chave)))

        if chave in doc.corrigidos:
            nivel = "corrigido"
        else:
            nivel = doc.resultado.confiancas.get(chave)
        texto_nivel = ROTULOS_NIVEL.get(nivel, SEM_CONFIANCA) if nivel in NIVEIS or nivel == "corrigido" else SEM_CONFIANCA
        linhas.append([
            Celula(id_doc, "0"),
            Celula(rotulo),
            Celula(chave),
            celula_valor,
            # SEM fundo fixo: a cor desta coluna vem de regras de formatacao
            # condicional sobre o proprio texto (ver _colorir_coluna_confianca)
            Celula(texto_nivel),
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
            celula.protection = Protection(locked=False)  # dado editavel; ver _proteger_planilha

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


def _escrever_linha_total_documentos(ws, n_documentos: int) -> None:
    """Total do LOTE ao final da aba Documentos: soma de "Valor total" com
    `=SUM` (ignora texto, igual aos Itens: um valor que nao virou numero nao
    quebra a soma). So e chamada com 2+ documentos -- com 1, o total seria o
    proprio valor da unica linha, ruido.

    Mesmo padrao dos Itens (linha logo abaixo da Tabela, FORA da ref dela),
    mas SEM celulas mescladas: o rotulo fica na coluna imediatamente antes do
    Valor total, alinhado a direita. Mesclar A:K atravessaria a divisa do
    congelamento (`freeze_panes = "B2"`) numa faixa larga, e nao ha ganho
    visual que compense."""
    col_total = CABECALHOS_RESUMO.index("Valor total") + 1
    linha_total = n_documentos + 2

    borda_topo = Border(top=Side(style="thin", color=COR_BORDA))
    for coluna in range(1, len(CABECALHOS_RESUMO) + 1):
        ws.cell(row=linha_total, column=coluna).border = borda_topo

    rotulo = ws.cell(row=linha_total, column=col_total - 1, value="Total do lote")
    rotulo.font = _fonte(negrito=True)
    rotulo.alignment = Alignment(horizontal="right", vertical="center", indent=1)

    letra = get_column_letter(col_total)
    total = ws.cell(row=linha_total, column=col_total, value=f"=SUM({letra}2:{letra}{n_documentos + 1})")
    total.number_format = FORMATO_MOEDA
    total.font = _fonte(negrito=True, cor=COR_DESTAQUE_VALOR_TOTAL, tamanho=TAMANHO_FONTE_VALOR_TOTAL)
    total.alignment = Alignment(horizontal="right", vertical="center", indent=1)


# ---------- protecao de planilha (etapa 10) ----------


def _proteger_planilha(ws) -> None:
    """Trava contra engano, SEM senha (Revisao > Desproteger planilha, 1
    clique): cabecalho, linhas de total (formulas) e tudo fora da tabela ficam
    bloqueados; so as celulas de dado (desbloqueadas em `_escrever_aba`)
    aceitam edicao. Protecao nao e criptografia nem seguranca -- e so o aviso
    do Excel antes de alguem apagar um `=SUM` sem querer.

    A semantica dos atributos no formato e INVERTIDA: `autoFilter=False`
    significa filtro LIBERADO (`autoFilter="0"` no XML). Prioridade do projeto:
    filtro e ordenacao da Tabela funcionando valem mais do que a trava --
    por isso os dois sao liberados, e o verificador (tests/ooxml.py) acusa uma
    aba protegida com Tabela e filtro bloqueado. Colunas e linhas tambem podem
    ser redimensionadas. Graficos ficam editaveis (`objects` desligado): quem
    monta uma apresentacao precisa selecionar e copiar. Consequencia
    conhecida: com a aba protegida a Tabela nao cresce (nao da pra digitar uma
    linha nova embaixo dela) -- e preciso desproteger antes."""
    ws.protection.sheet = True
    ws.protection.autoFilter = False
    ws.protection.sort = False
    ws.protection.formatColumns = False
    ws.protection.formatRows = False


# ---------- grafico do lote (etapa 10) ----------
TITULO_GRAFICO = "Valor total por documento"
MINIMO_DOCUMENTOS_NO_GRAFICO = 2
LARGURA_GRAFICO_CM, ALTURA_GRAFICO_CM = 24, 9.5
LINHAS_ENTRE_TOTAL_E_GRAFICO = 2


def _referencia_de_linhas(nome_aba: str, coluna: int, linhas: list[int]) -> str:
    """Referencia de celulas de UMA coluna nas `linhas` dadas (ordenadas):
    `'Documentos'!$L$2:$L$6` quando sao contiguas -- o caso normal -- ou, com
    lacunas, a uniao entre parenteses `('Documentos'!$L$2:$L$3,'Documentos'!$L$5)`
    (o Excel aceita e grava exatamente assim). Existe pra deixar de fora do
    grafico o documento cujo Valor total e texto: o Excel plota texto como
    ZERO, o que mostraria uma barra "R$ 0,00" que nao existe."""
    letra = get_column_letter(coluna)
    aba = "'" + nome_aba.replace("'", "''") + "'"
    faixas: list[list[int]] = []
    for linha in linhas:
        if faixas and linha == faixas[-1][-1] + 1:
            faixas[-1].append(linha)
        else:
            faixas.append([linha])
    partes = [
        f"{aba}!${letra}${f[0]}" if len(f) == 1 else f"{aba}!${letra}${f[0]}:${letra}${f[-1]}" for f in faixas
    ]
    return partes[0] if len(partes) == 1 else "(" + ",".join(partes) + ")"


def _grafico_valor_por_documento(ws, resumo: list[list[Celula]]) -> bool:
    """Grafico de barras "Valor total por documento" ABAIXO da tabela (e da
    linha de total), so quando ha 2+ documentos com Valor total numerico --
    com 1 barra so, o grafico seria enfeite. Devolve se foi criado.

    - Le direto das celulas da aba (nada de tabela auxiliar oculta): editar um
      valor na planilha atualiza o grafico, e filtrar a Tabela (linhas
      ocultas) tambem (`plotVisOnly`, o padrao).
    - Ancorado na coluna B, nao na A: a A (ID) esta congelada, e um grafico
      atravessando a divisa do congelamento ficaria cortado ao rolar.
    - FORA da area de impressao de proposito (que e "so a tabela real"): um
      grafico pode cair sobre a quebra de pagina e sair cortado.
    - Eixos com `delete=False` EXPLICITO: o openpyxl 3.1 nao escreve o
      elemento, e o Excel 365 entao esconde os eixos. Categorias como
      `strRef` (sao nomes, nao numeros). Rotulo de valor em cada barra, com
      as 6 flags de `dLbls` explicitas (as omitidas o Excel pode ligar)."""
    col_valor = CABECALHOS_RESUMO.index("Valor total")
    col_nome = CABECALHOS_RESUMO.index("Documento")
    linhas = [
        numero_linha
        for numero_linha, linha in enumerate(resumo, start=2)
        if isinstance(linha[col_valor].valor, (int, float)) and not isinstance(linha[col_valor].valor, bool)
    ]
    if len(linhas) < MINIMO_DOCUMENTOS_NO_GRAFICO:
        return False

    aba = ws.title
    serie = Series(
        tx=SeriesLabel(strRef=StrRef(f=_referencia_de_linhas(aba, col_valor + 1, [1]))),
        cat=AxDataSource(strRef=StrRef(f=_referencia_de_linhas(aba, col_nome + 1, linhas))),
        val=NumDataSource(numRef=NumRef(f=_referencia_de_linhas(aba, col_valor + 1, linhas))),
    )
    serie.graphicalProperties = GraphicalProperties(solidFill=COR_CABECALHO_FUNDO)
    serie.graphicalProperties.line = LineProperties(solidFill=COR_CABECALHO_FUNDO)
    serie.dLbls = DataLabelList(
        showVal=True, showSerName=False, showCatName=False, showLegendKey=False, showPercent=False, showBubbleSize=False,
        numFmt=FORMATO_MOEDA,
    )

    grafico = BarChart()
    grafico.type = "col"
    grafico.title = TITULO_GRAFICO
    grafico.title.overlay = False  # sem isso o titulo pode ser desenhado por cima da area do grafico
    grafico.roundedCorners = False  # ausente, o Excel assume cantos arredondados
    grafico.legend = None  # uma serie so: a legenda repetiria o titulo
    grafico.gapWidth = 70
    grafico.width, grafico.height = LARGURA_GRAFICO_CM, ALTURA_GRAFICO_CM
    grafico.series.append(serie)
    grafico.x_axis.delete = False
    grafico.y_axis.delete = False
    grafico.x_axis.axPos, grafico.y_axis.axPos = "b", "l"
    grafico.y_axis.number_format = '"R$" #,##0'
    grafico.y_axis.majorGridlines.spPr = GraphicalProperties(ln=LineProperties(solidFill=COR_BORDA))

    linha_da_ancora = len(resumo) + 2 + LINHAS_ENTRE_TOTAL_E_GRAFICO + 1  # apos a linha de total
    ws.add_chart(grafico, f"B{linha_da_ancora}")
    return True


# ---------- validacao de dados e formatacao condicional (etapa 10) ----------
#
# Listas suspensas so onde a coluna tem um conjunto FECHADO de valores, e a
# lista e exatamente o dominio que o exportador escreve nela (senao o proprio
# arquivo violaria a regra que carrega):
# - Campos adicionais!Confianca: Alta/Media/Baixa/Corrigido e "—" (sem info).
# - Documentos!Tipo: os valores INTERNOS (nota_fiscal...), nao os rotulos
#   amigaveis -- e o que a coluna guarda (bom pra filtrar).
# "Confianca geral" (Documentos) NAO entra: e uma proporcao em texto
# ("6 de 7 alta"), nao um conjunto fechado.
# Limites do Excel que, violados, o fazem pedir "reparar" (checados em
# tests/ooxml.py): lista literal <= 255 caracteres, titulo do erro <= 32,
# mensagem <= 255. `showDropDown=True` no openpyxl ESCONDE a seta (semantica
# invertida do formato) -- por isso nunca e ligado.
VALORES_CONFIANCA = [*ROTULOS_NIVEL.values(), SEM_CONFIANCA]
VALORES_TIPO = list(ROTULOS_TIPO)
TITULO_ERRO_LISTA = "Valor inválido"
LIMITE_LISTA_LITERAL = 255


def _lista_suspensa(ws, faixa: str, valores: list[str]) -> None:
    if any("," in v for v in valores):
        raise ValueError("valor com virgula nao cabe numa lista literal do Excel (a virgula separa itens)")
    literal = ",".join(valores)
    if len(literal) > LIMITE_LISTA_LITERAL:
        raise ValueError(f"lista literal de {len(literal)} caracteres passa do limite do Excel ({LIMITE_LISTA_LITERAL})")
    validacao = DataValidation(
        type="list",
        formula1=f'"{literal}"',
        allow_blank=True,
        showErrorMessage=True,
        errorStyle="stop",
        errorTitle=TITULO_ERRO_LISTA,
        error="Escolha um valor da lista: " + ", ".join(valores),
    )
    validacao.add(faixa)
    ws.add_data_validation(validacao)


def _colorir_coluna_confianca(ws, faixa: str) -> None:
    """Regras de formatacao condicional (nativas) sobre o TEXTO da coluna
    Confianca de Campos adicionais: e a unica coluna em que o texto da celula
    e o proprio valor, entao a cor pode acompanha-lo (editou "Baixa" pra
    "Alta", a cor some). Nas demais celulas destacadas (Emissor, Valor total,
    linhas de Itens...) a cor documenta a PROVENIENCIA da extracao -- o rastro
    de "isto veio com confianca baixa" -- e NAO pode depender do valor atual
    da celula (reescrever o campo apagaria o rastro); la continua o fundo
    fixo + comentario. Mesmos tons de FUNDOS; "Corrigido" tambem em italico."""
    for texto, estado in (("Média", "media"), ("Baixa", "baixa"), ("Corrigido", "corrigido")):
        cor = FUNDOS[estado]
        ws.conditional_formatting.add(
            faixa,
            CellIsRule(
                operator="equal",
                formula=[f'"{texto}"'],
                fill=PatternFill(fill_type="solid", start_color=cor, end_color=cor),
                font=Font(italic=True) if estado == "corrigido" else None,
            ),
        )


# ---------- Relatorio (etapa 9): aba de LEITURA, nao de dado tabular ----------
#
# As outras abas (Documentos/Itens/Campos adicionais/Avisos) continuam
# existindo -- servem pra filtro, Tabela Dinamica, Power Query. O Relatorio e
# a MESMA informacao, reorganizada em blocos (emitente, destinatario,
# valores...) pra quem so quer LER o documento, sem interpretar uma tabela.
# Por isso e a 1a aba (a que abre) e a unica com cabecalho de pagina em cima
# -- aqui NAO ha o risco de confundir leitor automatico que ja descartou a
# linha de titulo mesclada nas abas de dados (ver docstring do modulo): esta
# aba nao e uma tabela, ninguem vai ler ela por Power Query.
COR_RUBRICA = "667080"  # --color-text-muted: rotulo em cinza, discreto
COR_AUSENTE = "98A2B3"  # --color-text-faint: "nao encontrado" (campo obrigatorio vazio)
TAMANHO_FONTE_APLICATIVO = 18
TAMANHO_FONTE_TITULO_DOCUMENTO = 13
TAMANHO_FONTE_VALOR_PRINCIPAL = 20
ALTURA_LINHA_VALOR_PRINCIPAL = 28  # generoso o bastante pra fonte 20 nao cortar (ver _linha_campo)
LARGURAS_RELATORIO = {"A": 3, "B": 22, "C": 15, "D": 15, "E": 15, "F": 15}
ROTULOS_MODO = {"basico": "Modo básico", "ia": "Modo IA"}
ROTULOS_ORIGEM = {"digital": "Texto digital", "ocr": "OCR"}


@dataclass
class CampoRelatorio:
    """Uma linha rotulo+valor do Relatorio. `celula` None + `ausente` True =
    campo OBRIGATORIO vazio ("não encontrado"); o proprio CampoRelatorio
    None (nao um valor deste tipo) = campo OPCIONAL vazio, a linha some."""
    rotulo: str
    celula: Optional[Celula]
    ausente: bool = False


def _campo_relatorio(
    rotulo: str, chave: str, valor, doc: DocumentoParaExportar, tipo: str, formato: Optional[str] = None
) -> Optional[CampoRelatorio]:
    """None (linha omitida) quando opcional e vazio; "não encontrado" quando
    obrigatorio e vazio -- mesma lista de obrigatorios por tipo que a tela
    usa (`app.confianca.CAMPOS_OBRIGATORIOS`, testada e2e contra a tela).
    `formato` so se aplica quando o valor realmente vira numero (ex: Valor
    total -> FORMATO_MOEDA); string (conversao ja falhou antes) ignora."""
    if vazio(valor):
        return CampoRelatorio(rotulo, None, ausente=True) if e_obrigatorio(tipo, chave) else None
    return CampoRelatorio(rotulo, _celula(valor, chave, doc, formato))


def _campo_relatorio_data(rotulo: str, chave: str, texto: Optional[str], doc: DocumentoParaExportar, tipo: str) -> Optional[CampoRelatorio]:
    if vazio(texto):
        return CampoRelatorio(rotulo, None, ausente=True) if e_obrigatorio(tipo, chave) else None
    data = _para_data(texto)
    return CampoRelatorio(rotulo, _celula(data if data else texto, chave, doc, FORMATO_DATA if data else None))


def _campos_relatorio_fiscal(
    rotulo_nome: str, chave: str, valor_completo: Optional[str], doc: DocumentoParaExportar, tipo: str
) -> tuple[Optional[CampoRelatorio], Optional[CampoRelatorio]]:
    """Como `_campo_relatorio`, mas pra emissor/destinatario: separa "NOME
    (CNPJ x)" em nome + documento fiscal (mesmo formato das outras abas).
    Os dois levam o MESMO destaque de confianca; o comentario (com o valor
    original) fica so no nome -- mesmo padrao da aba Documentos."""
    if vazio(valor_completo):
        campo = CampoRelatorio(rotulo_nome, None, ausente=True) if e_obrigatorio(tipo, chave) else None
        return campo, None
    nome, doc_fiscal = separar_documento_fiscal(valor_completo)
    c_nome = _celula(nome, chave, doc)
    campo_nome = CampoRelatorio(rotulo_nome, c_nome)
    if vazio(doc_fiscal):
        return campo_nome, None
    campo_doc = CampoRelatorio("CNPJ/CPF", Celula(doc_fiscal, None, c_nome.estado, None))
    return campo_nome, campo_doc


def _titulo_secao(ws, linha: int, titulo: str) -> int:
    """Titulo de bloco (DOCUMENTO/EMISSOR/...): maiusculo, negrito, linha
    fina na cor de destaque embaixo -- hierarquia sem precisar de mais cor."""
    celula = ws.cell(row=linha, column=2, value=titulo.upper())
    celula.font = _fonte(negrito=True, cor=COR_CABECALHO_FUNDO, tamanho=10)
    borda = Border(bottom=Side(style="thin", color=COR_CABECALHO_FUNDO))
    for coluna in range(2, 7):
        ws.cell(row=linha, column=coluna).border = borda
    return linha + 1


def _espaco(ws, linha: int, altura: int = 6) -> int:
    ws.row_dimensions[linha].height = altura
    return linha + 1


def _linha_campo(
    ws, linha: int, campo: Optional[CampoRelatorio], *,
    tamanho: int = 11, negrito: bool = False, cor_valor: Optional[str] = None, altura: Optional[int] = None,
) -> int:
    """Rotulo em cinza (coluna B) + valor (C:F mescladas) com o MESMO
    destaque de confianca (fundo + comentario) das abas de dados. `campo`
    None nao escreve nada (campo opcional vazio -- a linha inteira some).

    `altura`: o Excel NAO recalcula a altura da linha sozinho quando a
    celula com a fonte maior esta MESCLADA -- limitacao conhecida do
    proprio Excel, nao do openpyxl (o resto do arquivo confia no auto-ajuste
    porque nunca combina fonte grande com celula mesclada, exceto aqui).
    Bug real: sem isso, "R$ 900,00" em fonte 20 aparecia cortado pela
    metade. So o Valor total passa isso; o resto usa None (auto)."""
    if campo is None:
        return linha
    vertical = "top"
    if altura:
        ws.row_dimensions[linha].height = altura
        vertical = "center"  # linha alta de proposito (fonte grande): centraliza em vez de "flutuar" no topo

    c_rotulo = ws.cell(row=linha, column=2, value=campo.rotulo)
    c_rotulo.font = _fonte(cor=COR_RUBRICA, tamanho=9)
    c_rotulo.alignment = Alignment(horizontal="left", vertical=vertical, wrap_text=True, indent=1)

    ws.merge_cells(start_row=linha, start_column=3, end_row=linha, end_column=6)
    c_valor = ws.cell(row=linha, column=3)
    estado = campo.celula.estado if campo.celula else None

    if campo.ausente:
        c_valor.value = "não encontrado"
        c_valor.data_type = "s"
        c_valor.font = _fonte(italico=True, cor=COR_AUSENTE, tamanho=tamanho)
    else:
        valor = campo.celula.valor
        if isinstance(valor, str):
            texto = _texto_seguro(valor)
            if texto:
                c_valor.value = texto
                c_valor.data_type = "s"  # NUNCA formula (ver docstring do modulo)
                c_valor.number_format = FORMATO_TEXTO
            c_valor.font = _fonte(negrito=negrito, cor=cor_valor, tamanho=tamanho)
        elif valor is not None:
            c_valor.value = valor
            if campo.celula.formato:
                c_valor.number_format = campo.celula.formato
            c_valor.font = _fonte(negrito=negrito, cor=cor_valor, tamanho=tamanho)
    c_valor.alignment = Alignment(horizontal="left", vertical=vertical, wrap_text=True, indent=1)

    if estado in FUNDOS:
        cor = FUNDOS[estado]
        for coluna in (2, 3):
            ws.cell(row=linha, column=coluna).fill = PatternFill("solid", start_color=cor, end_color=cor)
        if estado == "corrigido":
            c_valor.font = _fonte(italico=True, negrito=negrito, cor=cor_valor, tamanho=tamanho)
    if campo.celula and campo.celula.comentario:
        comentario = Comment(campo.celula.comentario, "Extrator")
        comentario.width, comentario.height = 280, 90
        c_valor.comment = comentario

    return linha + 1


def _bloco_avisos(ws, linha: int, avisos: list[str]) -> int:
    if not avisos:
        return linha
    linha = _titulo_secao(ws, linha, "Pontos de atenção")
    for aviso in avisos:
        ws.merge_cells(start_row=linha, start_column=2, end_row=linha, end_column=6)
        celula = ws.cell(row=linha, column=2)
        texto = _texto_seguro(f"⚠ {aviso}")
        if texto:
            celula.value = texto
            celula.data_type = "s"
            celula.number_format = FORMATO_TEXTO
        celula.font = _fonte(cor="8A5A00", tamanho=10)  # --color-warning
        celula.fill = PatternFill("solid", start_color="FDF3E0", end_color="FDF3E0")  # --color-warning-bg
        celula.alignment = Alignment(horizontal="left", vertical="center", wrap_text=True, indent=1)
        linha += 1
    return _espaco(ws, linha, altura=8)


def _cabecalho_documento(ws, linha: int, doc: DocumentoParaExportar, rotulo: str) -> int:
    celula = ws.cell(row=linha, column=2)
    texto_rotulo = _texto_seguro(rotulo)
    if texto_rotulo:
        celula.value = texto_rotulo
        celula.data_type = "s"  # NUNCA formula (rotulo embute numero_documento, vindo do PDF)
        celula.number_format = FORMATO_TEXTO
    celula.font = _fonte(negrito=True, cor=COR_DESTAQUE_VALOR_TOTAL, tamanho=TAMANHO_FONTE_TITULO_DOCUMENTO)
    linha += 1

    partes = [
        doc.arquivo,
        ROTULOS_MODO.get(doc.resultado.modo_extracao),
        ROTULOS_ORIGEM.get(doc.resultado.origem_texto),
    ]
    subtitulo = " · ".join(p for p in partes if p)
    if subtitulo:
        celula = ws.cell(row=linha, column=2, value=_texto_seguro(subtitulo))
        celula.data_type = "s"
        celula.font = _fonte(italico=True, cor=COR_RUBRICA, tamanho=8)
        linha += 1
    return _espaco(ws, linha, altura=8)


def _bloco_itens(ws, linha: int, doc: DocumentoParaExportar) -> int:
    itens = doc.resultado.documento.itens
    if not itens:
        return linha
    linha = _titulo_secao(ws, linha, "Itens")

    ws.merge_cells(start_row=linha, start_column=2, end_row=linha, end_column=3)
    for coluna, titulo in ((2, "Descrição"), (4, "Qtd."), (5, "Vl. unit."), (6, "Vl. total")):
        celula = ws.cell(row=linha, column=coluna, value=titulo)
        celula.font = _fonte(negrito=True, cor=COR_RUBRICA, tamanho=8)
        celula.alignment = Alignment(horizontal="left" if coluna == 2 else "right", vertical="center", indent=1)
    linha += 1

    nivel = doc.resultado.confiancas.get("itens")
    estado = nivel if nivel in ("media", "baixa") else None
    for item in itens:
        ws.merge_cells(start_row=linha, start_column=2, end_row=linha, end_column=3)
        c_desc = ws.cell(row=linha, column=2)
        texto = _texto_seguro(item.descricao) if item.descricao else ""
        if texto:
            c_desc.value = texto
            c_desc.data_type = "s"
            c_desc.number_format = FORMATO_TEXTO
        c_desc.font = _fonte(tamanho=10)
        c_desc.alignment = Alignment(horizontal="left", vertical="top", wrap_text=True, indent=1)

        for coluna, valor, formato in (
            (4, item.quantidade, None),
            (5, item.valor_unitario, FORMATO_MOEDA_UNITARIO),
            (6, item.valor_total, FORMATO_MOEDA),
        ):
            celula = ws.cell(row=linha, column=coluna)
            if isinstance(valor, (int, float)) and not isinstance(valor, bool):
                celula.value = valor
                if formato:
                    celula.number_format = formato
            elif isinstance(valor, str):
                texto = _texto_seguro(valor)
                if texto:
                    celula.value = texto
                    celula.data_type = "s"
                    celula.number_format = FORMATO_TEXTO
            celula.font = _fonte(tamanho=10)
            celula.alignment = Alignment(horizontal="right", vertical="top", indent=1)

        if estado in FUNDOS:
            cor = FUNDOS[estado]
            for coluna in range(2, 7):
                ws.cell(row=linha, column=coluna).fill = PatternFill("solid", start_color=cor, end_color=cor)
        linha += 1
    return _espaco(ws, linha, altura=8)


def _escrever_documento_relatorio(ws, linha: int, doc: DocumentoParaExportar, rotulo: str) -> tuple[int, set[str]]:
    """Escreve o bloco de UM documento a partir de `linha`. Devolve
    (proxima_linha, estados_de_confianca_usados) -- alimenta o rodape da
    legenda, que so mostra a cor que realmente apareceu no lote."""
    documento = doc.resultado.documento
    tipo = documento.tipo_documento
    estados: set[str] = set()

    def com(*campos: Optional[CampoRelatorio]) -> None:
        for campo in campos:
            if campo and campo.celula and campo.celula.estado:
                estados.add(campo.celula.estado)

    linha = _cabecalho_documento(ws, linha, doc, rotulo)

    avisos_doc = doc.resultado.avisos or ([doc.resultado.aviso] if doc.resultado.aviso else [])
    linha = _bloco_avisos(ws, linha, avisos_doc)

    # DOCUMENTO
    linha = _titulo_secao(ws, linha, "Documento")
    c_tipo = ws.cell(row=linha, column=2, value="Tipo")
    c_tipo.font = _fonte(cor=COR_RUBRICA, tamanho=9)
    c_tipo.alignment = Alignment(horizontal="left", vertical="top", indent=1)
    ws.merge_cells(start_row=linha, start_column=3, end_row=linha, end_column=6)
    c_valor_tipo = ws.cell(row=linha, column=3)
    # tipo normalmente e um dos valores conhecidos, mas o modo IA pode
    # inventar qualquer string -- sanitiza igual a qualquer texto do PDF
    texto_tipo = _texto_seguro(ROTULOS_TIPO.get(tipo, tipo))
    if texto_tipo:
        c_valor_tipo.value = texto_tipo
        c_valor_tipo.data_type = "s"
        c_valor_tipo.number_format = FORMATO_TEXTO
    c_valor_tipo.alignment = Alignment(horizontal="left", vertical="top", indent=1)
    linha += 1
    campo_numero = _campo_relatorio("Número", "numero_documento", documento.numero_documento, doc, tipo)
    campo_emissao = _campo_relatorio_data("Data de emissão", "data_emissao", documento.data_emissao, doc, tipo)
    campo_vencimento = _campo_relatorio_data("Data de vencimento", "data_vencimento", documento.data_vencimento, doc, tipo)
    com(campo_numero, campo_emissao, campo_vencimento)
    linha = _linha_campo(ws, linha, campo_numero)
    linha = _linha_campo(ws, linha, campo_emissao)
    linha = _linha_campo(ws, linha, campo_vencimento)
    linha = _espaco(ws, linha, altura=8)

    # EMISSOR / DESTINATARIO
    for titulo, chave, valor in (("Emissor", "emissor", documento.emissor), ("Destinatário", "destinatario", documento.destinatario)):
        campo_nome, campo_doc_fiscal = _campos_relatorio_fiscal("Nome", chave, valor, doc, tipo)
        if campo_nome is None and campo_doc_fiscal is None:
            continue  # opcional e vazio (nunca acontece hoje -- os dois sao obrigatorios em quase todo tipo -- mas nao inventa secao vazia se um dia acontecer)
        linha = _titulo_secao(ws, linha, titulo)
        com(campo_nome, campo_doc_fiscal)
        linha = _linha_campo(ws, linha, campo_nome)
        linha = _linha_campo(ws, linha, campo_doc_fiscal)
        linha = _espaco(ws, linha, altura=8)

    # VALORES
    linha = _titulo_secao(ws, linha, "Valores")
    campo_total = _campo_relatorio("Valor total", "valor_total", documento.valor_total, doc, tipo, formato=FORMATO_MOEDA)
    com(campo_total)
    linha = _linha_campo(
        ws, linha, campo_total, tamanho=TAMANHO_FONTE_VALOR_PRINCIPAL, negrito=True,
        cor_valor=COR_DESTAQUE_VALOR_TOTAL, altura=ALTURA_LINHA_VALOR_PRINCIPAL,
    )
    for extra in documento.campos_adicionais:
        if extra.campo not in CAMPOS_MONETARIOS:
            continue
        estado = _estado_visual(extra.campo, extra.valor, doc)
        numero = _numero_br(extra.valor)
        celula = Celula(
            numero if numero is not None else extra.valor,
            FORMATO_MOEDA if numero is not None else None,
            estado,
            _comentario(estado, doc.corrigidos.get(extra.campo)),
        )
        campo = CampoRelatorio(extra.campo, celula)
        com(campo)
        linha = _linha_campo(ws, linha, campo)
    linha = _espaco(ws, linha, altura=8)

    # DADOS ADICIONAIS (o resto dos campos_adicionais, nao monetarios)
    outros = [e for e in documento.campos_adicionais if e.campo not in CAMPOS_MONETARIOS]
    if outros:
        linha = _titulo_secao(ws, linha, "Dados adicionais")
        for extra in outros:
            estado = _estado_visual(extra.campo, extra.valor, doc)
            valor_exibido = _formatar_chave_acesso(extra.valor) if extra.campo == "Chave de Acesso" else extra.valor
            celula = Celula(valor_exibido, None, estado, _comentario(estado, doc.corrigidos.get(extra.campo)))
            campo = CampoRelatorio(extra.campo, celula)
            com(campo)
            linha = _linha_campo(ws, linha, campo)
        linha = _espaco(ws, linha, altura=8)

    linha = _bloco_itens(ws, linha, doc)
    if doc.resultado.confiancas.get("itens") in ("media", "baixa"):
        estados.add(doc.resultado.confiancas["itens"])

    return _espaco(ws, linha, altura=16), estados


def _rodape_legenda(ws, linha: int, estados_presentes: set[str]) -> None:
    """So as cores que aparecem DE VERDADE neste lote -- nunca as 3 juntas
    "por via das dúvidas". Se nenhuma apareceu (tudo "alta", sem correcao),
    nao ha rodape nenhum."""
    relevantes = [(estado, texto) for estado, texto in LEGENDA_LINHAS if estado in estados_presentes]
    if not relevantes:
        return
    linha = _espaco(ws, linha, altura=12)  # respiro antes do rodape (senao cola no ultimo bloco)
    for estado, texto in relevantes:
        cor = FUNDOS[estado]
        ws.cell(row=linha, column=2).fill = PatternFill("solid", start_color=cor, end_color=cor)
        ws.merge_cells(start_row=linha, start_column=3, end_row=linha, end_column=6)
        celula = ws.cell(row=linha, column=3, value=texto)
        celula.data_type = "s"
        celula.font = _fonte(cor=COR_RUBRICA, tamanho=8, italico=(estado == "corrigido"))
        celula.alignment = Alignment(horizontal="left", vertical="center", wrap_text=True, indent=1)
        linha += 1


def _escrever_relatorio(ws, documentos: list[DocumentoParaExportar], data_geracao: date) -> None:
    celula = ws.cell(row=1, column=2, value=NOME_APLICATIVO)
    celula.font = _fonte(negrito=True, cor=COR_DESTAQUE_VALOR_TOTAL, tamanho=TAMANHO_FONTE_APLICATIVO)
    celula = ws.cell(row=2, column=2, value=f"Relatório de extração · Gerado em {data_geracao.strftime('%d/%m/%Y')}")
    celula.font = _fonte(italico=True, cor=COR_RUBRICA, tamanho=9)
    linha = _espaco(ws, 3, altura=12)

    estados_todos: set[str] = set()
    for indice, doc in enumerate(documentos):
        documento = doc.resultado.documento
        rotulo = rotulo_documento(documento.tipo_documento, documento.numero_documento, doc.arquivo)
        linha, estados = _escrever_documento_relatorio(ws, linha, doc, rotulo)
        estados_todos |= estados
        if indice < len(documentos) - 1:
            ws.row_breaks.append(Break(id=linha - 1))  # quebra de pagina entre documentos, so no lote

    _rodape_legenda(ws, linha, estados_todos)

    ws.sheet_view.showGridLines = False  # e o que mais tira a "cara de planilha"
    for coluna, largura in LARGURAS_RELATORIO.items():
        ws.column_dimensions[coluna].width = largura
    ws.page_setup.orientation = "portrait"  # layout estreito, ao contrario das abas de dados (paisagem)
    ws.page_setup.fitToWidth = 1
    ws.page_setup.fitToHeight = 0
    ws.sheet_properties.pageSetUpPr.fitToPage = True
    ws.print_options.horizontalCentered = True
    ws.sheet_properties.tabColor = COR_CABECALHO_FUNDO  # unica aba colorida -- "comece por aqui"


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

    # Relatorio: 1a aba (a que abre), a unica de LEITURA -- ver secao acima.
    ws_relatorio = wb.active
    ws_relatorio.title = "Relatório"
    _escrever_relatorio(ws_relatorio, documentos, data_geracao)

    ws_resumo = wb.create_sheet(NOME_ABA_DOCUMENTOS)
    _escrever_aba(ws_resumo, CABECALHOS_RESUMO, resumo, NOMES_TABELA[NOME_ABA_DOCUMENTOS])
    _destacar_valor_total(ws_resumo, resumo)
    _nota_geracao(ws_resumo, len(CABECALHOS_RESUMO), data_geracao)
    if resumo:
        col_tipo = get_column_letter(CABECALHOS_RESUMO.index("Tipo") + 1)
        _lista_suspensa(ws_resumo, f"{col_tipo}2:{col_tipo}{len(resumo) + 1}", VALORES_TIPO)
    if len(resumo) > 1:
        _escrever_linha_total_documentos(ws_resumo, len(resumo))
        _grafico_valor_por_documento(ws_resumo, resumo)
    # como nos Itens, a area de impressao inclui a linha de total
    _configurar_impressao(ws_resumo, get_column_letter(len(CABECALHOS_RESUMO)), ws_resumo.max_row)

    ws_itens = wb.create_sheet("Itens")
    _escrever_aba(ws_itens, CABECALHOS_ITENS, itens, NOMES_TABELA["Itens"])
    if itens:
        _escrever_linha_total_itens(ws_itens, len(itens))
    # area de impressao inclui a linha de total (ela e conteudo de leitura,
    # mesmo ficando fora da Tabela/filtro, que e sobre dado filtravel)
    _configurar_impressao(ws_itens, get_column_letter(len(CABECALHOS_ITENS)), ws_itens.max_row)

    ws_campos = wb.create_sheet("Campos adicionais")
    _escrever_aba(ws_campos, CABECALHOS_CAMPOS, campos, NOMES_TABELA["Campos adicionais"])
    if campos:
        col_conf = get_column_letter(CABECALHOS_CAMPOS.index("Confiança") + 1)
        faixa_conf = f"{col_conf}2:{col_conf}{len(campos) + 1}"
        _lista_suspensa(ws_campos, faixa_conf, VALORES_CONFIANCA)
        _colorir_coluna_confianca(ws_campos, faixa_conf)
    _configurar_impressao(ws_campos, get_column_letter(len(CABECALHOS_CAMPOS)), max(len(campos) + 1, 1))

    ws_avisos = wb.create_sheet("Avisos")
    _escrever_aba(ws_avisos, CABECALHOS_AVISOS, avisos, NOMES_TABELA["Avisos"])
    _configurar_impressao(ws_avisos, get_column_letter(len(CABECALHOS_AVISOS)), max(len(avisos) + 1, 1))

    # Abas SEM NENHUM DADO ficam ocultas (nao removidas): a estrutura do
    # arquivo continua estavel pra quem usa Power Query/formulas com o nome
    # da aba fixo (ver docstring do modulo), mas quem abre no Excel nao ve
    # uma aba vazia -- ex: Itens e Avisos de um boleto sem soma divergente.
    # A aba Legenda foi removida (etapa 9): virou o rodape do Relatorio, so
    # com as cores que aparecem de verdade (ver _rodape_legenda).
    for aba, linhas in ((ws_itens, itens), (ws_campos, campos), (ws_avisos, avisos)):
        if not linhas:
            aba.sheet_state = "hidden"

    # Relatorio (leitura, sem formula nem dado a editar) fica fora da protecao
    for aba in (ws_resumo, ws_itens, ws_campos, ws_avisos):
        _proteger_planilha(aba)

    buffer = BytesIO()
    wb.save(buffer)
    buffer.seek(0)
    return buffer
