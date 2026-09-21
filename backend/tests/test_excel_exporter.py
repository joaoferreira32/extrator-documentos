"""Excel multi-aba (etapa 6): estrutura, formatos, destaques e seguranca.

Le o .xlsx de volta com openpyxl -- o que o usuario abriria no Excel.
"""
import asyncio
import zipfile
from datetime import date, datetime
from io import BytesIO

import openpyxl
import pytest
from pydantic import ValidationError

from app import main
from app.excel_exporter import (
    FORMATO_DATA,
    FORMATO_MOEDA,
    FORMATO_MOEDA_UNITARIO,
    gerar_excel,
    rotulo_documento,
    separar_documento_fiscal,
)
from app.schemas import (
    DocumentoExtraido,
    DocumentoParaExportar,
    ExportarExcelRequest,
    ExtractionResult,
)

CHAVE_44 = "35260472381189001001550010000123451123456786"


def _doc(arquivo="a.pdf", confiancas=None, avisos=None, aviso=None, corrigidos=None, **documento):
    documento.setdefault("tipo_documento", "nota_fiscal")
    return DocumentoParaExportar(
        arquivo=arquivo,
        resultado=ExtractionResult(
            modo_extracao="basico",
            confiancas=confiancas or {},
            avisos=avisos or [],
            aviso=aviso,
            documento=DocumentoExtraido(**documento),
        ),
        corrigidos=corrigidos or {},
    )


def _abrir(*docs):
    buffer = gerar_excel(list(docs))
    return openpyxl.load_workbook(buffer), buffer


def _linhas(ws):
    cabecalho = [c.value for c in ws[1]]
    return [dict(zip(cabecalho, linha)) for linha in ws.iter_rows(min_row=2, values_only=True)]


def _celula(ws, coluna, linha=2):
    cabecalho = [c.value for c in ws[1]]
    return ws.cell(row=linha, column=cabecalho.index(coluna) + 1)


def _fundo(celula):
    return celula.fill.fgColor.rgb[-6:] if celula.fill.fill_type == "solid" else None


# ---------- estrutura ----------


def test_quatro_abas_sempre_presentes_na_ordem_mesmo_sem_itens_campos_e_avisos():
    wb, _ = _abrir(_doc(numero_documento="1"))
    assert wb.sheetnames == ["Resumo", "Itens", "Campos adicionais", "Avisos"]
    for nome in ("Itens", "Campos adicionais", "Avisos"):
        assert wb[nome].max_row == 1, f"{nome} sem dados deve ter so o cabecalho"


def test_cabecalho_negrito_congelado_e_com_filtro_em_todas_as_abas():
    wb, _ = _abrir(_doc(numero_documento="1", itens=[{"descricao": "x", "valor_total": 1.0}],
                        campos_adicionais=[{"campo": "CFOP", "valor": "5102"}], avisos=["a"]))
    for ws in wb:
        assert all(c.font.bold for c in ws[1]), f"{ws.title}: cabecalho em negrito"
        assert ws.freeze_panes == "A2", f"{ws.title}: cabecalho congelado"
        ultima_coluna = openpyxl.utils.get_column_letter(ws.max_column)
        assert ws.auto_filter.ref == f"A1:{ultima_coluna}{ws.max_row}", f"{ws.title}: filtro cobre a tabela toda"


def test_toda_aba_tem_id_sequencial_e_documento_legivel():
    wb, _ = _abrir(_doc(numero_documento="000012345", itens=[{"descricao": "x"}],
                        campos_adicionais=[{"campo": "CFOP", "valor": "5102"}], avisos=["a"]))
    for ws in wb:
        cabecalho = [c.value for c in ws[1]]
        assert cabecalho[:2] == ["ID", "Documento"] or (ws.title == "Resumo" and cabecalho[:3] == ["ID", "Arquivo", "Documento"])
        linha = _linhas(ws)[0]
        assert linha["ID"] == 1 and linha["Documento"] == "Nota fiscal 000012345"


def test_lote_de_dois_documentos_ids_documentos_e_linhas_ligados_pelo_id():
    d1 = _doc(arquivo="um.pdf", numero_documento="1", itens=[{"descricao": "a"}, {"descricao": "b"}],
              campos_adicionais=[{"campo": "CFOP", "valor": "5102"}], avisos=["aviso do 1"])
    d2 = _doc(arquivo="dois.pdf", tipo_documento="boleto", numero_documento="1",  # MESMO numero de proposito
              itens=[{"descricao": "c"}], campos_adicionais=[{"campo": "Parcela", "valor": "1/2"}])
    wb, _ = _abrir(d1, d2)

    resumo = _linhas(wb["Resumo"])
    assert [(r["ID"], r["Arquivo"], r["Documento"]) for r in resumo] == [
        (1, "um.pdf", "Nota fiscal 1"), (2, "dois.pdf", "Boleto 1"),
    ]
    assert [(r["ID"], r["Descrição"]) for r in _linhas(wb["Itens"])] == [(1, "a"), (1, "b"), (2, "c")]
    assert [(r["ID"], r["Campo"]) for r in _linhas(wb["Campos adicionais"])] == [(1, "CFOP"), (2, "Parcela")]
    assert [(r["ID"], r["Aviso"]) for r in _linhas(wb["Avisos"])] == [(1, "aviso do 1")]


def test_rotulo_documento_com_e_sem_numero():
    assert rotulo_documento("boleto", "123", "a.pdf") == "Boleto 123"
    assert rotulo_documento("nota_fiscal", None, "a.pdf") == "Nota fiscal — a.pdf"
    assert rotulo_documento("nota_fiscal", "  ", None) == "Nota fiscal (sem número)"
    assert rotulo_documento("tipo_inventado", "9", None) == "tipo_inventado 9"


# ---------- Resumo ----------


def test_emissor_e_destinatario_sem_o_documento_entre_parenteses_e_em_colunas_proprias():
    wb, _ = _abrir(_doc(emissor="DELL LTDA (CNPJ 72.381.189/0010-01)", destinatario="FULANO DE TAL (CPF 000.000.000-00)"))
    r = _linhas(wb["Resumo"])[0]
    assert r["Emissor"] == "DELL LTDA" and r["Emissor CNPJ/CPF"] == "72.381.189/0010-01"
    assert r["Destinatário"] == "FULANO DE TAL" and r["Destinatário CNPJ/CPF"] == "000.000.000-00"


@pytest.mark.parametrize(
    "entrada, esperado",
    [
        ("EMPRESA SEM DOC LTDA", ("EMPRESA SEM DOC LTDA", None)),
        ("EMPRESA (FILIAL) LTDA (CNPJ 11.222.333/0001-81)", ("EMPRESA (FILIAL) LTDA", "11.222.333/0001-81")),
        ("(CNPJ 11.222.333/0001-81)", ("(CNPJ 11.222.333/0001-81)", None)),  # sem nome: nao inventa
        (None, (None, None)),
        ("  ", (None, None)),
    ],
)
def test_separar_documento_fiscal(entrada, esperado):
    assert separar_documento_fiscal(entrada) == esperado


def test_numero_do_documento_fica_texto_e_preserva_zeros_a_esquerda():
    wb, _ = _abrir(_doc(numero_documento="000012345"))
    celula = _celula(wb["Resumo"], "Número")
    assert celula.value == "000012345" and celula.data_type == "s"


def test_datas_viram_data_de_verdade_e_o_que_nao_converte_fica_texto():
    wb, _ = _abrir(
        _doc(data_emissao="15/04/2026", data_vencimento="2026-05-01"),
        _doc(data_emissao="31/02/2026", data_vencimento="quando der"),
    )
    ws = wb["Resumo"]
    emissao, vencimento = _celula(ws, "Data de emissão", 2), _celula(ws, "Data de vencimento", 2)
    assert isinstance(emissao.value, datetime) and emissao.value.date() == date(2026, 4, 15)
    assert emissao.number_format == FORMATO_DATA and emissao.is_date
    assert vencimento.value.date() == date(2026, 5, 1)  # ISO (o modo IA pode devolver)
    # dia inexistente e texto livre: permanecem texto, sem quebrar
    assert _celula(ws, "Data de emissão", 3).value == "31/02/2026"
    assert _celula(ws, "Data de vencimento", 3).value == "quando der"


def test_valor_total_float_vira_numero_em_moeda_e_string_fica_texto():
    wb, _ = _abrir(_doc(valor_total=1234.5), _doc(valor_total="a combinar"))
    ws = wb["Resumo"]
    numero = _celula(ws, "Valor total", 2)
    assert numero.value == 1234.5 and numero.number_format == FORMATO_MOEDA and numero.data_type == "n"
    assert _celula(ws, "Valor total", 3).value == "a combinar"


def test_formato_monetario_usa_a_notacao_interna_que_o_excel_localiza():
    """O codigo guardado no arquivo e en-US; em pt-BR o Excel mostra R$ 1.234,56.
    Escrever "#.##0,00" literalmente quebraria o formato."""
    assert FORMATO_MOEDA == '"R$" #,##0.00'
    assert "#.##0,00" not in FORMATO_MOEDA and "#.##0,00" not in FORMATO_MOEDA_UNITARIO


def test_tipo_no_resumo_e_o_valor_interno_e_a_confianca_geral_segue_a_regra_da_tela():
    d = _doc(numero_documento="1", emissor="E (CNPJ 1)", valor_total=10.0,
             campos_adicionais=[{"campo": "CFOP", "valor": "5102"}],
             confiancas={"numero_documento": "alta", "emissor": "alta", "valor_total": "media", "CFOP": "media"})
    r = _linhas(_abrir(d)[0]["Resumo"])[0]
    assert r["Tipo"] == "nota_fiscal"  # interno (bom pra filtrar), nao "Nota fiscal"
    assert r["Confiança geral"] == "2 de 4 alta"


def test_confianca_geral_no_modo_ia_e_traco():
    r = _linhas(_abrir(_doc(numero_documento="1"))[0]["Resumo"])[0]
    assert r["Confiança geral"] == "—"


def test_confianca_geral_deixa_os_corrigidos_fora_do_x_de_y():
    d = _doc(numero_documento="1", emissor="E", valor_total=10.0,
             confiancas={"numero_documento": "alta", "emissor": "alta", "valor_total": "alta"},
             corrigidos={"emissor": "outro"})
    assert _linhas(_abrir(d)[0]["Resumo"])[0]["Confiança geral"] == "2 de 2 alta"


# ---------- Campos adicionais ----------


def test_campos_monetarios_viram_numero_e_codigos_ficam_texto():
    d = _doc(campos_adicionais=[
        {"campo": "Valor do Documento", "valor": "1.000,00"},
        {"campo": "Desconto", "valor": "100,00"},
        {"campo": "Valor a Pagar", "valor": "sem valor"},  # nao parseavel: texto
        {"campo": "Chave de Acesso", "valor": CHAVE_44},
        {"campo": "Nosso Número", "valor": "10200000001-9"},
        {"campo": "CFOP", "valor": "5102"},
    ])
    ws = _abrir(d)[0]["Campos adicionais"]
    por_campo = {ws.cell(row=r, column=3).value: ws.cell(row=r, column=4) for r in range(2, ws.max_row + 1)}
    assert por_campo["Valor do Documento"].value == 1000.0 and por_campo["Valor do Documento"].number_format == FORMATO_MOEDA
    assert por_campo["Desconto"].value == 100.0
    assert por_campo["Valor a Pagar"].value == "sem valor"
    # 44 digitos: como numero o Excel guardaria so 15 de precisao -- tem que ser texto INTEIRO
    assert por_campo["Chave de Acesso"].value == CHAVE_44 and por_campo["Chave de Acesso"].data_type == "s"
    assert por_campo["Nosso Número"].value == "10200000001-9"
    assert por_campo["CFOP"].value == "5102" and por_campo["CFOP"].data_type == "s"


def test_coluna_confianca_dos_campos_adicionais_em_texto():
    d = _doc(campos_adicionais=[{"campo": "A", "valor": "1"}, {"campo": "B", "valor": "2"},
                                {"campo": "C", "valor": "3"}, {"campo": "D", "valor": "4"}],
             confiancas={"A": "alta", "B": "media", "C": "baixa"}, corrigidos={"D": "x"})
    linhas = _linhas(_abrir(d)[0]["Campos adicionais"])
    assert [(l["Campo"], l["Confiança"]) for l in linhas] == [
        ("A", "Alta"), ("B", "Média"), ("C", "Baixa"), ("D", "Corrigido"),
    ]


# ---------- Itens ----------


def test_itens_numericos_com_formato_e_texto_quando_nao_converteu():
    d = _doc(itens=[{"descricao": "Mochila", "quantidade": 2.5, "valor_unitario": 215.0312, "valor_total": 537.58},
                    {"descricao": "Servico", "quantidade": "1 un", "valor_unitario": None, "valor_total": "cortesia"}])
    ws = _abrir(d)[0]["Itens"]
    assert _celula(ws, "Quantidade", 2).value == 2.5
    unit = _celula(ws, "Valor unitário", 2)
    assert unit.value == 215.0312 and unit.number_format == FORMATO_MOEDA_UNITARIO  # ate 4 casas
    assert _celula(ws, "Valor total", 2).number_format == FORMATO_MOEDA
    assert _celula(ws, "Quantidade", 3).value == "1 un" and _celula(ws, "Valor total", 3).value == "cortesia"


# ---------- Avisos ----------


def test_avisos_uma_linha_por_aviso_e_cai_na_string_aviso_quando_nao_ha_lista():
    wb, _ = _abrir(_doc(numero_documento="1", avisos=["primeiro", "segundo"]),
                   _doc(numero_documento="2", aviso="so a string antiga"))
    assert [(l["ID"], l["Aviso"]) for l in _linhas(wb["Avisos"])] == [
        (1, "primeiro"), (1, "segundo"), (2, "so a string antiga"),
    ]


# ---------- destaques ----------


def test_fundo_e_comentario_para_media_baixa_e_corrigido():
    d = _doc(numero_documento="1", data_emissao="15/04/2026", emissor="E (CNPJ 11.222.333/0001-81)",
             valor_total=10.0,
             confiancas={"numero_documento": "alta", "data_emissao": "media", "emissor": "baixa", "valor_total": "alta"},
             corrigidos={"valor_total": 9.0})
    ws = _abrir(d)[0]["Resumo"]

    assert _fundo(_celula(ws, "Número")) is None and _celula(ws, "Número").comment is None  # alta: sem destaque
    data = _celula(ws, "Data de emissão")
    assert _fundo(data) == "FDF3E0" and "média" in data.comment.text
    emissor, emissor_doc = _celula(ws, "Emissor"), _celula(ws, "Emissor CNPJ/CPF")
    assert _fundo(emissor) == "FBECEB" and "baixa" in emissor.comment.text
    assert _fundo(emissor_doc) == "FBECEB" and emissor_doc.comment is None  # CNPJ leva a cor, o comentario fica no nome
    total = _celula(ws, "Valor total")
    assert _fundo(total) == "E7F0EF" and total.font.italic  # corrigido: outra cor + italico
    assert total.comment.text == "Corrigido pelo usuário. Valor extraído: 9,00"


def test_corrigido_para_vazio_ainda_e_marcado_com_o_valor_original():
    d = _doc(numero_documento=None, corrigidos={"numero_documento": "000012345"})
    celula = _celula(_abrir(d)[0]["Resumo"], "Número")
    assert celula.value is None and _fundo(celula) == "E7F0EF"
    assert "000012345" in celula.comment.text


def test_confianca_media_da_tabela_de_itens_pinta_as_linhas_com_um_comentario_so():
    d = _doc(itens=[{"descricao": "a", "valor_total": 1.0}, {"descricao": "b", "valor_total": 2.0}],
             confiancas={"itens": "media"})
    ws = _abrir(d)[0]["Itens"]
    assert _fundo(_celula(ws, "Descrição", 2)) == "FDF3E0" and _fundo(_celula(ws, "Valor total", 3)) == "FDF3E0"
    assert _celula(ws, "Descrição", 2).comment is not None
    assert _celula(ws, "Descrição", 3).comment is None


def test_modo_ia_nao_pinta_nada():
    ws = _abrir(_doc(numero_documento="1", valor_total=1.0))[0]["Resumo"]
    assert all(_fundo(c) is None for c in ws[2])


# ---------- seguranca e robustez ----------


def test_texto_que_parece_formula_fica_texto_e_nao_vira_formula():
    """O openpyxl trata string com "=" no inicio como formula; o texto vem de
    PDF (nao confiavel), entao um emissor "=HYPERLINK(...)" nao pode executar."""
    perigoso = '=HYPERLINK("http://x","clique")'
    d = _doc(emissor=perigoso, destinatario="=1+1", numero_documento="=2+2",
             itens=[{"descricao": "=SUM(A1:A2)"}], campos_adicionais=[{"campo": "Nota", "valor": "=A1"}],
             avisos=["=cmd|' /C calc'!A0"])
    wb, buffer = _abrir(d)

    assert _celula(wb["Resumo"], "Emissor").value == perigoso
    assert _celula(wb["Resumo"], "Destinatário").value == "=1+1"
    assert _celula(wb["Itens"], "Descrição").value == "=SUM(A1:A2)"
    for ws in wb:
        for linha in ws.iter_rows():
            assert all(c.data_type != "f" for c in linha), f"{ws.title}: celula virou formula"

    buffer.seek(0)
    with zipfile.ZipFile(buffer) as zf:  # e no XML nao existe nenhum <f>
        for nome in zf.namelist():
            if nome.startswith("xl/worksheets/sheet"):
                xml = zf.read(nome).decode("utf-8")
                assert "<f>" not in xml and "<f " not in xml, nome


def test_caracteres_de_controle_sao_removidos_em_vez_de_quebrar_o_arquivo():
    d = _doc(emissor="EMPRESA\x00\x01\x0b LTDA", numero_documento="12\x1f3")
    ws = _abrir(d)[0]["Resumo"]
    assert _celula(ws, "Emissor").value == "EMPRESA LTDA"
    assert _celula(ws, "Número").value == "123"


def test_texto_gigante_e_truncado_no_limite_do_excel():
    ws = _abrir(_doc(emissor="X" * 40000))[0]["Resumo"]
    assert len(_celula(ws, "Emissor").value) == 32767


def test_larguras_ajustadas_ao_conteudo_com_teto():
    d = _doc(numero_documento="1", emissor="E" * 200, destinatario="Curto",
             campos_adicionais=[{"campo": "Chave de Acesso", "valor": CHAVE_44}])
    wb, _ = _abrir(d)
    resumo = wb["Resumo"]
    largura = lambda ws, letra: ws.column_dimensions[letra].width
    assert largura(resumo, "H") == 60  # Emissor: 200 chars -> teto
    assert largura(resumo, "C") >= len("Nota fiscal 1")  # Documento cabe
    assert largura(resumo, "A") <= 10  # ID: estreita
    assert largura(wb["Campos adicionais"], "D") >= len(CHAVE_44)  # a chave inteira cabe (44 < 60)


# ---------- endpoint ----------


def _corpo(resposta):
    async def ler():
        return b"".join([pedaco async for pedaco in resposta.body_iterator])

    return asyncio.run(ler())


def test_endpoint_devolve_xlsx_com_nome_de_um_documento_ou_de_lote():
    um = asyncio.run(main.export_excel(ExportarExcelRequest(documentos=[_doc(numero_documento="1")])))
    assert um.headers["content-disposition"] == 'attachment; filename="documento_extraido.xlsx"'
    assert openpyxl.load_workbook(BytesIO(_corpo(um))).sheetnames == ["Resumo", "Itens", "Campos adicionais", "Avisos"]

    varios = asyncio.run(main.export_excel(ExportarExcelRequest(documentos=[_doc(), _doc()])))
    assert varios.headers["content-disposition"] == 'attachment; filename="documentos_extraidos.xlsx"'


def test_requisicao_sem_documentos_e_rejeitada():
    with pytest.raises(ValidationError):
        ExportarExcelRequest(documentos=[])
