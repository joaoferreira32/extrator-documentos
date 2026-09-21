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
"""
import re
from dataclasses import dataclass
from datetime import date
from io import BytesIO
from typing import Any, Optional

from openpyxl import Workbook
from openpyxl.comments import Comment
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter

from app.confianca import NIVEIS, resumo_confianca, vazio
from app.schemas import DocumentoParaExportar

FORMATO_MOEDA = '"R$" #,##0.00'
FORMATO_MOEDA_UNITARIO = '"R$" #,##0.00##'  # preco unitario pode ter ate 4 casas
FORMATO_DATA = "dd/mm/yyyy"
FORMATO_TEXTO = "@"

# Mesmos tons da interface (style.css): warning-bg, danger-bg e accent-soft.
FUNDOS = {"media": "FDF3E0", "baixa": "FBECEB", "corrigido": "E7F0EF"}
FUNDO_CABECALHO = "EEF0F2"

ROTULOS_TIPO = {
    "boleto": "Boleto",
    "nota_fiscal": "Nota fiscal",
    "pedido_compra": "Pedido de compra",
    "relatorio": "Relatório",
    "desconhecido": "Desconhecido",
}
ROTULOS_NIVEL = {"alta": "Alta", "media": "Média", "baixa": "Baixa", "corrigido": "Corrigido"}

# Campos adicionais que sao VALOR MONETARIO (viram numero). O resto e codigo/texto.
CAMPOS_MONETARIOS = {"Valor do Documento", "Desconto", "Valor a Pagar"}

CABECALHOS_RESUMO = [
    "ID", "Arquivo", "Documento", "Tipo", "Número", "Data de emissão", "Data de vencimento",
    "Emissor", "Emissor CNPJ/CPF", "Destinatário", "Destinatário CNPJ/CPF", "Valor total", "Confiança geral",
]
CABECALHOS_ITENS = ["ID", "Documento", "Descrição", "Quantidade", "Valor unitário", "Valor total"]
CABECALHOS_CAMPOS = ["ID", "Documento", "Campo", "Valor", "Confiança"]
CABECALHOS_AVISOS = ["ID", "Documento", "Aviso"]

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

    return [
        Celula(id_doc, "0"),
        Celula(doc.arquivo),
        Celula(rotulo),
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


def _escrever_aba(ws, cabecalhos: list[str], linhas: list[list[Celula]]) -> None:
    lado = Side(style="thin", color="C4C9D0")
    for coluna, titulo in enumerate(cabecalhos, start=1):
        celula = ws.cell(row=1, column=coluna, value=titulo)
        celula.font = Font(bold=True)
        celula.fill = PatternFill("solid", start_color=FUNDO_CABECALHO, end_color=FUNDO_CABECALHO)
        celula.border = Border(bottom=lado)
        celula.alignment = Alignment(vertical="center")

    for numero_linha, linha in enumerate(linhas, start=2):
        for coluna, dados in enumerate(linha, start=1):
            celula = ws.cell(row=numero_linha, column=coluna)
            valor = dados.valor
            if isinstance(valor, str):
                valor = _texto_seguro(valor)
                if valor:
                    celula.value = valor
                    celula.data_type = "s"  # NUNCA formula (ver docstring do modulo)
                    celula.number_format = FORMATO_TEXTO
                celula.alignment = Alignment(wrap_text=True, vertical="top")
            elif valor is not None:
                celula.value = valor
                if dados.formato:
                    celula.number_format = dados.formato
                celula.alignment = Alignment(vertical="top")
            else:
                celula.alignment = Alignment(vertical="top")

            if dados.estado in FUNDOS:
                cor = FUNDOS[dados.estado]
                celula.fill = PatternFill("solid", start_color=cor, end_color=cor)
                if dados.estado == "corrigido":
                    celula.font = Font(italic=True)
            if dados.comentario:
                comentario = Comment(dados.comentario, "Extrator")
                comentario.width, comentario.height = 280, 90
                celula.comment = comentario

    for coluna, titulo in enumerate(cabecalhos, start=1):
        maior = max([len(titulo)] + [_largura_visual(linha[coluna - 1]) for linha in linhas])
        ws.column_dimensions[get_column_letter(coluna)].width = min(max(maior + 2, LARGURA_MINIMA), LARGURA_MAXIMA)

    ws.freeze_panes = "A2"
    ws.auto_filter.ref = f"A1:{get_column_letter(len(cabecalhos))}{max(len(linhas) + 1, 1)}"


def gerar_excel(documentos: list[DocumentoParaExportar]) -> BytesIO:
    resumo, itens, campos, avisos = [], [], [], []
    for id_doc, doc in enumerate(documentos, start=1):
        documento = doc.resultado.documento
        rotulo = rotulo_documento(documento.tipo_documento, documento.numero_documento, doc.arquivo)
        resumo.append(_linha_resumo(id_doc, doc, rotulo))
        itens += _linhas_itens(id_doc, doc, rotulo)
        campos += _linhas_campos(id_doc, doc, rotulo)
        avisos += _linhas_avisos(id_doc, doc, rotulo)

    wb = Workbook()
    ws_resumo = wb.active
    ws_resumo.title = "Resumo"
    _escrever_aba(ws_resumo, CABECALHOS_RESUMO, resumo)
    _escrever_aba(wb.create_sheet("Itens"), CABECALHOS_ITENS, itens)
    _escrever_aba(wb.create_sheet("Campos adicionais"), CABECALHOS_CAMPOS, campos)
    _escrever_aba(wb.create_sheet("Avisos"), CABECALHOS_AVISOS, avisos)

    buffer = BytesIO()
    wb.save(buffer)
    buffer.seek(0)
    return buffer
