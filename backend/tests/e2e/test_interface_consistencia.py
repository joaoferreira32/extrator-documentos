"""Testes e2e: o MESMO valor aparece formatado igual na tela e no Excel.

Divergencias achadas na revisao final (antes deste arquivo): a tela mostrava
dinheiro sem "R$" (o Excel com), data como veio (o modo IA pode devolver ISO;
o usuario pode digitar "1/5/2026"; o Excel converte pra dd/mm/aaaa) e valores
monetarios de campos adicionais crus ("1000,00" x "R$ 1.000,00" no Excel).
A formatacao da tela e so EXIBICAO: o valor no JSON (e o que o Excel recebe)
continua o de sempre -- os testes conferem os dois lados.

Diferenca que continua DE PROPOSITO: a coluna Tipo da aba Documentos guarda o
valor interno (nota_fiscal), bom pra filtrar; a aba Relatorio mostra o rotulo.
"""
from datetime import date

import pytest

pytest.importorskip("playwright.sync_api", reason="instale requirements-dev.txt")

pytestmark = pytest.mark.e2e

RESPOSTA = {
    "modo_extracao": "ia", "origem_texto": "digital", "confiancas": {}, "avisos": [], "aviso": None,
    "documento": {
        "tipo_documento": "boleto", "numero_documento": "0042", "data_emissao": "2026-04-15",
        "data_vencimento": "31/2/2026", "emissor": "ESCOLA EXEMPLO (CNPJ 11.222.333/0001-81)",
        "destinatario": "FULANO EXEMPLO (CPF 000.000.000-00)", "valor_total": 1234.5,
        "itens": [{"descricao": "Mensalidade", "quantidade": 1.0, "valor_unitario": 1234.5, "valor_total": 1234.5}],
        "campos_adicionais": [
            {"campo": "Valor do Documento", "valor": "1234,50"},
            {"campo": "Nosso Número", "valor": "10200000001-9"},
        ],
    },
}


def _extrair_resposta_simulada(tela):
    tela.page.goto(tela.url)
    tela.page.route("**/extract-document", lambda rota: rota.fulfill(json=RESPOSTA))
    tela.page.set_input_files("#file-input", str(tela.pdfs["boleto"]))
    tela.page.click("#btn-extrair")
    tela.page.wait_for_selector("#resultado:not([hidden])")


def _texto(tela, id_campo):
    return tela.campo(id_campo).locator(".valor-texto").inner_text()


def _id_do_extra(tela, rotulo):
    return next(i for i, c in enumerate(RESPOSTA["documento"]["campos_adicionais"]) if c["campo"] == rotulo)


def test_dinheiro_datas_e_campos_monetarios_iguais_na_tela_e_no_excel(tela):
    _extrair_resposta_simulada(tela)

    # tela
    assert _texto(tela, "valor_total") == "R$ 1.234,50"
    assert _texto(tela, "data_emissao") == "15/04/2026"  # ISO do modo IA -> dd/mm/aaaa
    assert _texto(tela, "data_vencimento") == "31/2/2026"  # dia inexistente: fica como veio, sem completar zero (Excel idem)
    assert _texto(tela, f"extra-{_id_do_extra(tela, 'Valor do Documento')}") == "R$ 1.234,50"
    assert _texto(tela, f"extra-{_id_do_extra(tela, 'Nosso Número')}") == "10200000001-9"  # codigo nao e dinheiro
    celulas = tela.page.locator("#tabela-itens tbody tr:first-child td").all_inner_texts()
    assert celulas == ["Mensalidade", "1", "R$ 1.234,50", "R$ 1.234,50"]

    # o valor interno nao mudou (so a exibicao)
    doc = tela.json_bruto()["documento"]
    assert doc["data_emissao"] == "2026-04-15" and doc["valor_total"] == 1234.5
    assert doc["campos_adicionais"][0]["valor"] == "1234,50"
    assert tela.page.locator(".chip-corrigido").count() == 0  # exibir formatado nao e "corrigir"

    # Excel: os mesmos valores, com a mesma aparencia (formato de celula)
    planilha = tela.baixar_excel()
    r = planilha.resumo
    assert r["Data de emissão"].date() == date(2026, 4, 15)
    assert planilha.celula("Documentos", "Data de emissão").number_format == "dd/mm/yyyy"
    assert r["Data de vencimento"] == "31/2/2026"  # texto, como na tela
    assert r["Valor total"] == 1234.5 and planilha.celula("Documentos", "Valor total").number_format == '"R$" #,##0.00'
    assert planilha.campos["Valor do Documento"]["Valor"] == 1234.5
    assert planilha.itens[0]["Valor total"] == 1234.5


def test_data_digitada_pelo_usuario_aparece_como_o_excel_vai_mostrar(tela):
    _extrair_resposta_simulada(tela)
    tela.page.click('[data-campo="data_emissao"] .valor-btn')
    tela.campo("data_emissao").locator("textarea").fill("1/5/2026")
    tela.page.keyboard.press("Enter")

    assert _texto(tela, "data_emissao") == "01/05/2026"
    assert tela.json_bruto()["documento"]["data_emissao"] == "1/5/2026"  # guardado como digitado
    assert tela.campo("data_emissao").locator(".chip-corrigido").count() == 1
    assert tela.baixar_excel().resumo["Data de emissão"].date() == date(2026, 5, 1)


def test_abrir_e_confirmar_sem_mudar_nao_vira_correcao_nos_campos_formatados(tela):
    """Regressao: exibir formatado nao pode, sozinho, virar correcao nem mudar o
    valor guardado (o original "2026-04-15" continua no JSON)."""
    _extrair_resposta_simulada(tela)
    for id_campo in ("data_emissao", "valor_total", f"extra-{_id_do_extra(tela, 'Valor do Documento')}"):
        tela.page.click(f'[data-campo="{id_campo}"] .valor-btn')
        tela.page.keyboard.press("Enter")
    assert tela.page.locator(".chip-corrigido").count() == 0
    doc = tela.json_bruto()["documento"]
    assert doc["data_emissao"] == "2026-04-15" and doc["campos_adicionais"][0]["valor"] == "1234,50"
