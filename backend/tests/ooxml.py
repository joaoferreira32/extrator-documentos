"""Verificador de invariantes do Excel lendo o .xlsx como ZIP + XML bruto.

Existe por causa de um bug real: o .xlsx da etapa 8 passava em tudo que o
projeto conferia (recarregar com openpyxl, ler com pandas) e o Excel de verdade
recusava o arquivo e oferecia "reparar". Nenhuma das duas bibliotecas enxerga o
problema, porque as duas sao TOLERANTES -- o Excel nao e.

Por que nao so validar contra o XSD do OOXML: foi testado com o validador
oficial da Microsoft (Open XML SDK 2.20, construido a partir do schema) e ele
(a) NAO acusou o problema real -- uma Tabela sem nenhuma linha de dados e valida
pelo schema e rejeitada pelo Excel, e (b) acusou 3 falsos positivos na ORDEM dos
filhos de <font>, que o schema publicado define como `choice` (qualquer ordem;
conferido em datypic.com/sc/ooxml/t-ssml_CT_Font.html) e que aparecem inclusive
na fonte padrao que o openpyxl escreve em todo arquivo. Ou seja: o schema sozinho
nao protege contra essa classe de erro. O que protege sao as regras que o Excel
aplica ALEM do schema -- as que estao abaixo, cada uma com o motivo.

Le o XML com xml.etree (nao com openpyxl) de proposito: conferir com a mesma
biblioteca que escreveu o arquivo nao pega o que ela escreve errado.
"""
import re
import zipfile
from io import BytesIO
from xml.etree import ElementTree as ET

NS = {
    "m": "http://schemas.openxmlformats.org/spreadsheetml/2006/main",
    "r": "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
    "rel": "http://schemas.openxmlformats.org/package/2006/relationships",
    "ct": "http://schemas.openxmlformats.org/package/2006/content-types",
    "xdr": "http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing",
    "c": "http://schemas.openxmlformats.org/drawingml/2006/chart",
}
_R_ID = f"{{{NS['r']}}}id"
_NOME_TABELA_RE = re.compile(r"^[A-Za-z_\\][A-Za-z0-9_.]*$")
_PARECE_CELULA_RE = re.compile(r"^(?:[A-Za-z]{1,3}\d+|[Rr]\d*[Cc]\d*|[RrCc])$")


def _col_num(letras: str) -> int:
    n = 0
    for ch in letras.upper():
        n = n * 26 + (ord(ch) - 64)
    return n


def _faixa(ref: str) -> tuple[int, int, int, int]:
    """'A1:C4' -> (col_ini, lin_ini, col_fim, lin_fim); 'A1' -> faixa de 1 celula."""
    partes = ref.replace("$", "").split(":")
    ini, fim = partes[0], partes[-1]
    m_ini = re.match(r"([A-Za-z]+)(\d+)", ini)
    m_fim = re.match(r"([A-Za-z]+)(\d+)", fim)
    return _col_num(m_ini.group(1)), int(m_ini.group(2)), _col_num(m_fim.group(1)), int(m_fim.group(2))


def _intersecta(a, b) -> bool:
    return not (a[2] < b[0] or b[2] < a[0] or a[3] < b[1] or b[3] < a[1])


def _alvo(base: str, target: str) -> str:
    """Resolve o Target de um relacionamento (absoluto '/xl/..' ou relativo)."""
    if target.startswith("/"):
        return target[1:]
    partes = base.rsplit("/", 1)[0].split("/") if "/" in base else []
    for pedaco in target.split("/"):
        if pedaco == "..":
            partes = partes[:-1]
        elif pedaco != ".":
            partes.append(pedaco)
    return "/".join(partes)


def _rels(zf, parte: str) -> dict[str, tuple[str, str]]:
    """{rId: (tipo, caminho_da_parte_alvo)} dos relacionamentos de uma parte."""
    pasta, nome = parte.rsplit("/", 1) if "/" in parte else ("", parte)
    caminho_rels = f"{pasta}/_rels/{nome}.rels" if pasta else f"_rels/{nome}.rels"
    if caminho_rels not in zf.namelist():
        return {}
    raiz = ET.fromstring(zf.read(caminho_rels))
    return {
        rel.get("Id"): (rel.get("Type"), _alvo(parte, rel.get("Target")))
        for rel in raiz.findall("rel:Relationship", NS)
        if rel.get("TargetMode") != "External"
    }


def _texto_celula(celula) -> str | None:
    t = celula.get("t")
    if t == "inlineStr":
        return "".join(celula.itertext())
    v = celula.find("m:v", NS)
    return v.text if v is not None else None


def problemas(conteudo: bytes | BytesIO) -> list[str]:
    """Lista de violacoes (vazia = ok). Cada mensagem diz a parte e a regra."""
    dados = conteudo.getvalue() if isinstance(conteudo, BytesIO) else conteudo
    achados: list[str] = []
    with zipfile.ZipFile(BytesIO(dados)) as zf:
        nomes = set(zf.namelist())

        # 1. Toda parte precisa de content type (Default por extensao ou Override).
        tipos = ET.fromstring(zf.read("[Content_Types].xml"))
        defaults = {d.get("Extension").lower() for d in tipos.findall("ct:Default", NS)}
        overrides = {o.get("PartName").lstrip("/") for o in tipos.findall("ct:Override", NS)}
        for nome in nomes:
            if nome == "[Content_Types].xml" or nome.endswith("/"):
                continue
            ext = nome.rsplit(".", 1)[-1].lower()
            if nome not in overrides and ext not in defaults:
                achados.append(f"{nome}: sem content type em [Content_Types].xml")
        for parte in overrides:
            if parte not in nomes:
                achados.append(f"[Content_Types].xml: Override para parte inexistente {parte}")

        # 2. Workbook -> planilhas.
        rels_wb = _rels(zf, "xl/workbook.xml")
        for rid, (_tipo, alvo) in rels_wb.items():
            if alvo not in nomes:
                achados.append(f"xl/workbook.xml: {rid} aponta pra parte inexistente {alvo}")
        wb = ET.fromstring(zf.read("xl/workbook.xml"))
        planilhas = []  # [(nome_aba, parte)]
        for aba in wb.findall("m:sheets/m:sheet", NS):
            planilhas.append((aba.get("name"), rels_wb.get(aba.get(_R_ID), (None, None))[1]))
        nomes_aba = [n for n, _ in planilhas]

        # 3. Nomes definidos (area de impressao, titulos): aba e indice validos.
        for dn in wb.findall("m:definedNames/m:definedName", NS):
            local = dn.get("localSheetId")
            if local is not None and not (0 <= int(local) < len(planilhas)):
                achados.append(f"definedName {dn.get('name')}: localSheetId {local} fora das abas")
            for aba_ref in re.findall(r"'((?:[^']|'')+)'!|(?<![A-Za-z'])([A-Za-z_][\w.]*)!", dn.text or ""):
                aba = (aba_ref[0] or aba_ref[1]).replace("''", "'")
                if aba not in nomes_aba:
                    achados.append(f"definedName {dn.get('name')}: referencia aba inexistente {aba!r}")

        # estilos diferenciais (usados pela formatacao condicional): a contagem
        # declarada tem que bater e cada cfRule.dxfId tem que apontar pra um.
        n_dxfs = 0
        if "xl/styles.xml" in nomes:
            dxfs = ET.fromstring(zf.read("xl/styles.xml")).find("m:dxfs", NS)
            if dxfs is not None:
                n_dxfs = len(dxfs.findall("m:dxf", NS))
                if dxfs.get("count") is not None and int(dxfs.get("count")) != n_dxfs:
                    achados.append(f"xl/styles.xml: dxfs count={dxfs.get('count')} mas ha {n_dxfs} dxf")

        # indices de estilo (cellXfs) cuja celula e DESBLOQUEADA (<protection locked="0">)
        desbloqueados: set[int] = set()
        if "xl/styles.xml" in nomes:
            xfs = ET.fromstring(zf.read("xl/styles.xml")).find("m:cellXfs", NS)
            for i, xf in enumerate(xfs.findall("m:xf", NS) if xfs is not None else []):
                prot = xf.find("m:protection", NS)
                if prot is not None and prot.get("locked") in ("0", "false"):
                    desbloqueados.add(i)

        ids_tabela: dict[str, str] = {}
        nomes_tabela: dict[str, str] = {}
        for nome_aba, parte in planilhas:
            if parte is None or parte not in nomes:
                achados.append(f"aba {nome_aba!r}: parte da planilha nao encontrada")
                continue
            ws = ET.fromstring(zf.read(parte))
            rels_ws = _rels(zf, parte)
            for rid, (_tipo, alvo) in rels_ws.items():
                if alvo not in nomes:
                    achados.append(f"{parte}: {rid} aponta pra parte inexistente {alvo}")

            celulas = {}
            for c in ws.iter(f"{{{NS['m']}}}c"):
                celulas[c.get("r")] = c
                t = c.get("t")
                # 4. Celula tipada sem conteudo. O schema aceita (<is>/<v> sao
                #    opcionais), mas uma string sem string e um estado que o
                #    Excel nao produz -- nao arriscar.
                if t == "inlineStr" and c.find("m:is", NS) is None:
                    achados.append(f"{parte}: {c.get('r')} e inlineStr sem <is> (celula vazia tipada como texto)")
                if t == "s" and c.find("m:v", NS) is None:
                    achados.append(f"{parte}: {c.get('r')} e string compartilhada sem <v>")

            mescladas = [_faixa(m.get("ref")) for m in ws.findall("m:mergeCells/m:mergeCell", NS)]

            # 5. Tabelas da aba.
            faixas_tabela = []
            for tp in ws.findall("m:tableParts/m:tablePart", NS):
                rid = tp.get(_R_ID)
                if rid not in rels_ws:
                    achados.append(f"{parte}: tablePart {rid} sem relacionamento")
                    continue
                parte_tab = rels_ws[rid][1]
                tab = ET.fromstring(zf.read(parte_tab))
                onde = f"{parte_tab} (aba {nome_aba!r})"

                tid, nome = tab.get("id"), tab.get("displayName")
                if tid in ids_tabela:
                    achados.append(f"{onde}: id {tid} repetido (ja usado em {ids_tabela[tid]}) -- tem que ser unico no arquivo")
                ids_tabela[tid] = parte_tab
                if not nome or not _NOME_TABELA_RE.match(nome) or _PARECE_CELULA_RE.match(nome):
                    achados.append(f"{onde}: displayName {nome!r} invalido")
                elif nome.lower() in nomes_tabela:
                    achados.append(f"{onde}: displayName {nome!r} repetido")
                else:
                    nomes_tabela[nome.lower()] = parte_tab

                faixa = _faixa(tab.get("ref"))
                n_cab = int(tab.get("headerRowCount", "1"))
                n_total = int(tab.get("totalsRowCount", "0"))
                n_dados = faixa[3] - faixa[1] + 1 - n_cab - n_total
                # O bug real: Tabela com ref so no cabecalho. Valida pelo schema,
                # o Excel oferece reparar. O proprio Excel, ao criar uma tabela
                # vazia, sempre inclui 1 linha de dados (a "linha de insercao").
                if n_dados < 1:
                    achados.append(f"{onde}: ref {tab.get('ref')} sem nenhuma linha de dados (so cabecalho)")

                colunas = tab.findall("m:tableColumns/m:tableColumn", NS)
                largura = faixa[2] - faixa[0] + 1
                declarado = tab.find("m:tableColumns", NS)
                if declarado is None:
                    achados.append(f"{onde}: sem <tableColumns> (obrigatorio)")
                    continue
                if int(declarado.get("count", len(colunas))) != len(colunas) or len(colunas) != largura:
                    achados.append(f"{onde}: {len(colunas)} tableColumn(s) para uma ref de {largura} coluna(s)")
                nomes_col = [c.get("name") for c in colunas]
                if len({n.lower() for n in nomes_col if n}) != len(nomes_col):
                    achados.append(f"{onde}: nomes de coluna vazios ou repetidos {nomes_col}")
                ids_col = [c.get("id") for c in colunas]
                if len(set(ids_col)) != len(ids_col):
                    achados.append(f"{onde}: ids de tableColumn repetidos")

                # Cabecalho na planilha tem que ser texto e bater com tableColumn.
                if n_cab:
                    for i, nome_col in enumerate(nomes_col):
                        ref_cab = f"{_letra(faixa[0] + i)}{faixa[1]}"
                        celula = celulas.get(ref_cab)
                        texto = _texto_celula(celula) if celula is not None else None
                        if texto != nome_col:
                            achados.append(f"{onde}: cabecalho {ref_cab}={texto!r} difere de tableColumn {nome_col!r}")

                # Tabela nao pode cruzar celula mesclada nem outra tabela.
                for m in mescladas:
                    if _intersecta(faixa, m):
                        achados.append(f"{onde}: ref {tab.get('ref')} cruza celulas mescladas")
                for outra in faixas_tabela:
                    if _intersecta(faixa, outra):
                        achados.append(f"{onde}: ref {tab.get('ref')} sobrepoe outra tabela da mesma aba")
                faixas_tabela.append(faixa)

                auto = tab.find("m:autoFilter", NS)
                if auto is not None and auto.get("ref") != tab.get("ref"):
                    achados.append(f"{onde}: autoFilter {auto.get('ref')} difere da ref da tabela {tab.get('ref')}")

            # 6. Nao pode haver autoFilter solto na aba cruzando uma tabela.
            solto = ws.find("m:autoFilter", NS)
            if solto is not None and any(_intersecta(_faixa(solto.get("ref")), f) for f in faixas_tabela):
                achados.append(f"{parte}: autoFilter da aba cruza uma tabela (a tabela ja tem o proprio)")

            achados += _protecao(ws, parte, celulas, desbloqueados, tem_tabela=bool(faixas_tabela))
            achados += _validacoes_de_dados(ws, parte)
            achados += _formatacao_condicional(ws, parte, n_dxfs)
            achados += _graficos(zf, ws, parte, rels_ws, nomes_aba, faixas_tabela=faixas_tabela)
    return achados


# Referencia de celulas de grafico: 'Aba'!$L$2:$L$6, 'Aba'!$L$2 ou a uniao
# entre parenteses ('Aba'!$L$2:$L$3,'Aba'!$L$5).
_AREA_RE = re.compile(
    r"(?:'((?:[^']|'')+)'|([A-Za-z_][\w.]*))!\$?([A-Za-z]{1,3})\$?(\d+)(?::\$?([A-Za-z]{1,3})\$?(\d+))?"
)


def _areas(formula: str) -> tuple[list[tuple[str, tuple[int, int, int, int]]], str]:
    """([(aba, faixa)], resto_nao_reconhecido) de uma referencia de grafico."""
    areas = []
    for m in _AREA_RE.finditer(formula):
        aba = (m.group(1) or m.group(2)).replace("''", "'")
        col1, lin1 = _col_num(m.group(3)), int(m.group(4))
        col2, lin2 = (_col_num(m.group(5)), int(m.group(6))) if m.group(5) else (col1, lin1)
        areas.append((aba, (col1, lin1, col2, lin2)))
    return areas, _AREA_RE.sub("", formula).strip("(), ")


def _n_celulas(areas) -> int:
    return sum((f[2] - f[0] + 1) * (f[3] - f[1] + 1) for _, f in areas)


_FLAGS_DLBLS = ("showLegendKey", "showVal", "showCatName", "showSerName", "showPercent", "showBubbleSize")


def _graficos(zf, ws, parte: str, rels_ws, nomes_aba: list[str], faixas_tabela) -> list[str]:
    """Graficos da aba: regras que o Excel aplica alem do schema (ver o
    docstring do modulo). Cada uma existe por um problema real do
    openpyxl 3.1 ou do formato -- o motivo esta na mensagem."""
    achados: list[str] = []
    desenho = ws.find("m:drawing", NS)
    if desenho is None:
        return achados
    if desenho.get(_R_ID) not in rels_ws:
        return [f"{parte}: <drawing> {desenho.get(_R_ID)} sem relacionamento"]
    parte_desenho = rels_ws[desenho.get(_R_ID)][1]
    if parte_desenho not in zf.namelist():
        return achados  # a regra generica de relacionamentos ja acusou
    desenho_xml = ET.fromstring(zf.read(parte_desenho))
    rels_desenho = _rels(zf, parte_desenho)

    # o grafico nao pode ser ancorado em cima de uma tabela (esconderia os dados)
    for ancora in list(desenho_xml):
        de = ancora.find("xdr:from", NS)
        if de is not None:
            col, lin = int(de.find("xdr:col", NS).text) + 1, int(de.find("xdr:row", NS).text) + 1
            if any(f[0] <= col <= f[2] and f[1] <= lin <= f[3] for f in faixas_tabela):
                achados.append(f"{parte_desenho}: grafico ancorado em cima de uma tabela (celula {_letra(col)}{lin})")

    for quadro in desenho_xml.iter(f"{{{NS['c']}}}chart"):
        rid = quadro.get(_R_ID)
        if rid not in rels_desenho:
            achados.append(f"{parte_desenho}: grafico {rid} sem relacionamento")
            continue
        parte_grafico = rels_desenho[rid][1]
        if parte_grafico not in zf.namelist():
            continue
        grafico = ET.fromstring(zf.read(parte_grafico))
        onde = parte_grafico

        for serie in grafico.iter(f"{{{NS['c']}}}ser"):
            contagens = {}
            for papel in ("tx", "cat", "val"):
                for f in serie.findall(f"c:{papel}//c:f", NS):
                    areas, resto = _areas(f.text or "")
                    if not areas or resto:
                        achados.append(f"{onde}: referencia {f.text!r} de {papel} nao reconhecida")
                        continue
                    for aba, faixa in areas:
                        if aba not in nomes_aba:
                            achados.append(f"{onde}: {papel} referencia a aba inexistente {aba!r}")
                        if faixa[2] < faixa[0] or faixa[3] < faixa[1]:
                            achados.append(f"{onde}: intervalo invertido em {f.text!r}")
                    contagens[papel] = _n_celulas(areas)
            if "cat" in contagens and "val" in contagens and contagens["cat"] != contagens["val"]:
                achados.append(f"{onde}: serie com {contagens['cat']} categoria(s) e {contagens['val']} valor(es)")
            rotulos = serie.find("c:dLbls", NS)
            if rotulos is not None and rotulos.find("c:delete", NS) is None:
                faltando = [f for f in _FLAGS_DLBLS if rotulos.find(f"c:{f}", NS) is None]
                if faltando:
                    achados.append(f"{onde}: dLbls sem as flags {faltando} (o Excel pode ligar as omitidas)")

        eixos = {}
        for tipo in ("catAx", "valAx", "dateAx", "serAx"):
            for eixo in grafico.iter(f"{{{NS['c']}}}{tipo}"):
                eixos[eixo.find("c:axId", NS).get("val")] = eixo
        area_do_grafico = grafico.find("c:chart/c:plotArea", NS)
        for plot in (list(area_do_grafico) if area_do_grafico is not None else []):
            for ax in plot.findall("c:axId", NS):
                if ax.get("val") not in eixos:
                    achados.append(f"{onde}: axId {ax.get('val')} usado pelo grafico mas nao definido")
        for ident, eixo in eixos.items():
            cruza = eixo.find("c:crossAx", NS)
            if cruza is None or cruza.get("val") not in eixos:
                achados.append(f"{onde}: eixo {ident} cruza um eixo inexistente")
            apagado = eixo.find("c:delete", NS)
            # openpyxl 3.1 nao escreve <delete>; o Excel 365 entao esconde o eixo
            if apagado is None or apagado.get("val") not in ("0", "false"):
                achados.append(f"{onde}: eixo {ident} sem <delete val=0> explicito (o eixo some no Excel 365)")

        titulo = grafico.find("c:chart/c:title", NS)
        if titulo is not None and titulo.find("c:overlay", NS) is None:
            achados.append(f"{onde}: titulo sem <overlay val=0> (pode ser desenhado por cima do grafico)")
    return achados


def _protecao(ws, parte: str, celulas: dict, desbloqueados: set[int], tem_tabela: bool) -> list[str]:
    """So vale em aba PROTEGIDA (<sheetProtection sheet="1">). O formato usa
    semantica invertida: autoFilter="1" (o padrao) significa filtro BLOQUEADO."""
    achados: list[str] = []
    protecao = ws.find("m:sheetProtection", NS)
    if protecao is None or protecao.get("sheet") not in ("1", "true"):
        return achados
    for ref, c in celulas.items():
        if c.find("m:f", NS) is not None and int(c.get("s", 0)) in desbloqueados:
            achados.append(f"{parte}: formula em {ref} desbloqueada numa aba protegida (a trava nao protege o total)")
    if tem_tabela:
        for atributo, o_que in (("autoFilter", "as setas de filtro da Tabela"), ("sort", "a ordenacao da Tabela")):
            if protecao.get(atributo, "1") in ("1", "true"):  # ausente = padrao do formato = bloqueado
                achados.append(f"{parte}: aba protegida com Tabela e {atributo} bloqueado ({o_que} ficaria inutilizavel)")
    return achados


# Limites do Excel para validacao de dados: passar deles nao e erro de schema,
# mas o Excel remove a validacao inteira e oferece "reparar".
_LIMITE_LISTA_LITERAL = 255
_LIMITE_TITULO = 32
_LIMITE_MENSAGEM = 255


def _validacoes_de_dados(ws, parte: str) -> list[str]:
    achados: list[str] = []
    bloco = ws.find("m:dataValidations", NS)
    if bloco is None:
        return achados
    itens = bloco.findall("m:dataValidation", NS)
    if bloco.get("count") is not None and int(bloco.get("count")) != len(itens):
        achados.append(f"{parte}: dataValidations count={bloco.get('count')} mas ha {len(itens)} dataValidation")
    faixas_vistas: list[tuple[int, int, int, int]] = []
    for dv in itens:
        sqref = (dv.get("sqref") or "").split()
        if not sqref:
            achados.append(f"{parte}: dataValidation sem sqref")
        for pedaco in sqref:
            faixa = _faixa(pedaco)
            if any(_intersecta(faixa, outra) for outra in faixas_vistas):
                achados.append(f"{parte}: dataValidation {pedaco} sobrepoe outra validacao da aba (Excel pede reparo)")
            faixas_vistas.append(faixa)
        if dv.get("type") == "list":
            if dv.get("showDropDown") in ("1", "true"):
                achados.append(f"{parte}: dataValidation {dv.get('sqref')} com showDropDown=1 (no formato isso ESCONDE a seta da lista)")
            f1 = dv.find("m:formula1", NS)
            texto = (f1.text or "").strip() if f1 is not None else ""
            if not texto:
                achados.append(f"{parte}: dataValidation {dv.get('sqref')} do tipo list sem formula1")
            elif texto.startswith('"') and texto.endswith('"') and len(texto) - 2 > _LIMITE_LISTA_LITERAL:
                achados.append(f"{parte}: lista literal de {len(texto) - 2} caracteres (limite do Excel: {_LIMITE_LISTA_LITERAL})")
        for atributo, limite in (
            ("errorTitle", _LIMITE_TITULO), ("promptTitle", _LIMITE_TITULO), ("error", _LIMITE_MENSAGEM), ("prompt", _LIMITE_MENSAGEM),
        ):
            valor = dv.get(atributo)
            if valor is not None and len(valor) > limite:
                achados.append(f"{parte}: dataValidation {atributo} com {len(valor)} caracteres (limite do Excel: {limite})")
    return achados


def _formatacao_condicional(ws, parte: str, n_dxfs: int) -> list[str]:
    achados: list[str] = []
    prioridades: set[str] = set()
    for cf in ws.findall("m:conditionalFormatting", NS):
        if not (cf.get("sqref") or "").strip():
            achados.append(f"{parte}: conditionalFormatting sem sqref")
        for regra in cf.findall("m:cfRule", NS):
            tipo, dxf, prioridade = regra.get("type"), regra.get("dxfId"), regra.get("priority")
            onde = f"{parte}: cfRule {tipo} em {cf.get('sqref')}"
            if dxf is None:
                if tipo in ("cellIs", "expression"):
                    achados.append(f"{onde} sem dxfId (a regra nao pintaria nada)")
            elif not (0 <= int(dxf) < n_dxfs):
                achados.append(f"{onde} com dxfId={dxf}, mas styles.xml tem {n_dxfs} dxf")
            if prioridade is None:
                achados.append(f"{onde} sem priority")
            elif prioridade in prioridades:
                achados.append(f"{onde} com priority={prioridade} repetida na aba")
            else:
                prioridades.add(prioridade)
            formulas = regra.findall("m:formula", NS)
            if tipo == "cellIs":
                operador = regra.get("operator")
                if not operador:
                    achados.append(f"{onde} sem operator")
                esperado = 2 if operador in ("between", "notBetween") else 1
                if len(formulas) != esperado:
                    achados.append(f"{onde}: {len(formulas)} formula(s), esperado {esperado} pro operador {operador!r}")
            for f in formulas:
                if not (f.text or "").strip():
                    achados.append(f"{onde} com <formula> vazia")
    return achados


def _letra(n: int) -> str:
    s = ""
    while n:
        n, resto = divmod(n - 1, 26)
        s = chr(65 + resto) + s
    return s
