"""Testes e2e (Chromium real) da interface: confianca por campo, resumo, banner
de avisos, edicao, e a exibicao formatada (pt-BR, tipo amigavel, chave em
blocos). PDFs FICTICIOS gerados por PyMuPDF (ver helpers.py).

Rode com: `pytest tests/e2e --e2e -v` (ver requirements-dev.txt).

Cada teste abre uma pagina nova e falha se houver erro de console/JS (fixture
`tela`): um erro no meio de um handler ja deixou a tela "quase certa".
"""
import pytest

pytest.importorskip("playwright.sync_api", reason="instale requirements-dev.txt")

import helpers  # noqa: E402

pytestmark = pytest.mark.e2e

BASE_OBRIGATORIOS = ["emissor", "destinatario", "numero_documento", "data_emissao", "valor_total"]
OBRIGATORIOS = {
    "nota_fiscal": BASE_OBRIGATORIOS,
    "boleto": BASE_OBRIGATORIOS + ["data_vencimento"],
    "pedido_compra": BASE_OBRIGATORIOS,
}


def esperado(resposta):
    """(alta, total, revisar, vazios) calculado A PARTIR DO JSON, em Python,
    de forma independente do frontend."""
    doc, conf = resposta["documento"], resposta["confiancas"]
    obrigatorios = OBRIGATORIOS.get(doc["tipo_documento"], [])
    alta = revisar = vazios = 0
    for chave in ["numero_documento", "data_emissao", "data_vencimento", "emissor", "destinatario", "valor_total"]:
        valor = doc[chave]
        if valor in (None, ""):
            vazios += chave in obrigatorios
        elif conf.get(chave) == "alta":
            alta += 1
        elif conf.get(chave) in ("media", "baixa"):
            revisar += 1
    for extra in doc["campos_adicionais"]:
        c = conf.get(extra["campo"])
        alta += c == "alta"
        revisar += c in ("media", "baixa")
    if doc["itens"] and conf.get("itens"):
        alta += conf["itens"] == "alta"
        revisar += conf["itens"] in ("media", "baixa")
    return alta, alta + revisar, revisar, vazios


# ---------- indicador de confianca e resumo ----------


def test_danfe_ok_chips_com_texto_e_icone_resumo_e_itens_alta(tela):
    resposta = tela.extrair("danfe_ok")
    alta, total, revisar, _ = esperado(resposta)
    resumo = tela.resumo()

    assert f"{alta} de {total} campos com alta confiança" in resumo
    assert (f"{revisar} para revisar" in resumo) == (revisar > 0)
    assert "vazio" not in resumo
    assert tela.page.locator("#banner-avisos").is_hidden()

    emissor = tela.campo("emissor")
    assert emissor.locator(".chip-alta").inner_text().strip() == "Alta"  # texto, nao so cor
    assert emissor.locator(".chip .icone svg").count() == 1  # e icone
    assert tela.page.locator("#chip-itens .chip-alta").count() == 1  # soma dos itens fecha
    assert tela.campo("extra-1").locator(".chip-media").count() == 1  # CFOP
    assert tela.page.locator("#legenda-confianca").is_visible()
    assert tela.page.locator("#nota-confianca").is_hidden()


def test_opcional_vazio_aparece_neutro_fora_do_resumo(tela):
    tela.extrair("danfe_ok")
    vencimento = tela.campo("data_vencimento")  # opcional na nota fiscal
    assert vencimento.locator(".valor-texto").inner_text() == "—"
    assert vencimento.locator(".chip").count() == 0
    assert "is-opcional-vazio" in vencimento.get_attribute("class")


def test_danfe_com_problemas_banner_com_avisos_e_obrigatorio_vazio_aberto(tela):
    resposta = tela.extrair("danfe_problemas")
    alta, total, revisar, vazios = esperado(resposta)
    resumo = tela.resumo()

    assert f"{alta} de {total} campos com alta confiança" in resumo
    assert f"{revisar} para revisar" in resumo
    assert vazios == 1 and "1 vazio" in resumo

    assert tela.page.locator("#banner-avisos").is_visible()
    assert tela.page.locator("#banner-avisos .banner-titulo").inner_text() == "Atenção"
    assert tela.page.locator("#banner-avisos-corpo li").count() == len(resposta["avisos"]) == 2

    numero = tela.campo("numero_documento")  # obrigatorio e vazio
    assert "is-vazio" in numero.get_attribute("class")
    assert numero.locator("textarea").count() == 1  # ja aberto
    assert numero.locator(".chip-vazio").count() == 1
    assert numero.locator(".dica").inner_text() == "Preencha este campo."
    assert tela.campo("valor_total").locator(".chip-media").count() == 1  # totais nao fecham
    assert tela.page.locator("#chip-itens .chip-media").count() == 1  # soma dos itens nao fecha

    numero.locator("textarea").fill("000012345")
    tela.page.keyboard.press("Tab")  # blur confirma
    assert tela.campo("numero_documento").locator(".chip-corrigido").count() == 1
    resumo = tela.resumo()
    assert "vazio" not in resumo and "1 corrigido" in resumo


def test_boleto_tudo_alta_sem_campo_aberto(tela):
    resposta = tela.extrair("boleto")
    alta, total, revisar, vazios = esperado(resposta)
    resumo = tela.resumo()
    assert f"{alta} de {total} campos com alta confiança" in resumo
    assert "para revisar" not in resumo and "vazio" not in resumo
    assert tela.page.locator("textarea").count() == 0  # so botoes de valor


def test_generico_baixa_e_obrigatorios_vazios(tela):
    resposta = tela.extrair("generico")
    alta, total, revisar, vazios = esperado(resposta)
    resumo = tela.resumo()
    assert f"{alta} de {total} campos com alta confiança" in resumo
    assert f"{revisar} para revisar" in resumo and f"{vazios} vazios" in resumo

    numero = tela.campo("numero_documento")  # confianca "baixa"
    assert "is-baixa" in numero.get_attribute("class")
    assert numero.locator("textarea").count() == 1
    assert numero.locator(".chip-baixa").inner_text().strip() == "Baixa"
    assert numero.locator(".dica").inner_text() == "Confira e corrija."
    assert numero.evaluate("e => getComputedStyle(e).borderLeftWidth") == "3px"  # destaque alem do chip


# ---------- edicao ----------


def test_editar_por_clique_enter_atualiza_resumo_json_e_excel(tela):
    resposta = tela.extrair("danfe_ok")
    alta, total, _, _ = esperado(resposta)

    tela.page.click('[data-campo="emissor"] .valor-btn')
    editor = tela.campo("emissor").locator("textarea")
    assert editor.evaluate("e => document.activeElement === e"), "clique deve abrir o editor com foco"
    editor.fill("EMISSOR CORRIGIDO LTDA")
    tela.page.keyboard.press("Enter")

    assert tela.campo("emissor").locator(".chip-corrigido").inner_text().strip() == "Corrigido"
    resumo = tela.resumo()
    # corrigido sai do "X de Y" e entra a parte
    assert f"{alta - 1} de {total - 1} campos com alta confiança" in resumo and "1 corrigido" in resumo
    assert tela.json_bruto()["documento"]["emissor"] == "EMISSOR CORRIGIDO LTDA"
    # bug real ja visto: foco tem que voltar ao botao depois do Enter
    assert tela.campo("emissor").locator(".valor-btn").evaluate("e => document.activeElement === e")
    assert tela.baixar_excel()["Emissor"] == "EMISSOR CORRIGIDO LTDA"


def test_escape_descarta_a_edicao(tela):
    tela.extrair("danfe_ok")
    tela.page.click('[data-campo="destinatario"] .valor-btn')
    tela.campo("destinatario").locator("textarea").fill("TEXTO DESCARTADO")
    tela.page.keyboard.press("Escape")

    assert "DESCARTADO" not in tela.page.locator("#json-bruto").text_content()
    assert tela.campo("destinatario").locator(".chip-alta").count() == 1


def test_opcional_vazio_edita_ao_clicar_e_vira_corrigido(tela):
    tela.extrair("danfe_ok")
    tela.page.click('[data-campo="data_vencimento"] .valor-btn')
    tela.campo("data_vencimento").locator("textarea").fill("20/04/2026")
    tela.page.keyboard.press("Enter")
    assert tela.campo("data_vencimento").locator(".chip-corrigido").count() == 1
    assert tela.json_bruto()["documento"]["data_vencimento"] == "20/04/2026"


def test_teclado_tab_enter_e_escape(tela):
    tela.extrair("danfe_ok")
    tela.page.focus("#btn-excel")
    tela.page.keyboard.press("Tab")
    assert tela.page.evaluate("document.activeElement.className") == "valor-btn"
    tela.page.keyboard.press("Enter")
    assert tela.page.evaluate("document.activeElement.tagName") == "SELECT"  # 1o campo: tipo
    tela.page.keyboard.press("Escape")
    assert tela.page.evaluate("document.activeElement.className") == "valor-btn"


def test_valor_total_em_formato_br_vira_numero_no_json_e_no_excel(tela):
    tela.extrair("generico")
    tela.campo("valor_total").locator("textarea").fill("1.500,50")
    tela.page.keyboard.press("Enter")

    valor = tela.json_bruto()["documento"]["valor_total"]
    assert valor == 1500.5 and isinstance(valor, float)
    assert tela.baixar_excel()["Valor total"] == 1500.5  # numero de verdade no Excel
    assert tela.campo("valor_total").locator(".valor-texto").inner_text() == "1.500,50"

    tela.page.click('[data-campo="valor_total"] .valor-btn')  # nao parseavel: fica texto
    tela.campo("valor_total").locator("textarea").fill("a combinar")
    tela.page.keyboard.press("Enter")
    assert tela.json_bruto()["documento"]["valor_total"] == "a combinar"

    tela.page.click('[data-campo="valor_total"] .valor-btn')  # apagar obrigatorio -> null + Vazio
    tela.campo("valor_total").locator("textarea").fill("")
    tela.page.keyboard.press("Enter")
    assert tela.json_bruto()["documento"]["valor_total"] is None
    assert tela.page.locator('[data-campo="valor_total"].is-vazio').count() == 1


def test_mudar_o_tipo_muda_os_campos_obrigatorios(tela):
    tela.extrair("generico")
    assert tela.page.locator('[data-campo="data_vencimento"].is-vazio').count() == 0  # opcional na nota
    tela.page.click('[data-campo="tipo_documento"] .valor-btn')
    tela.page.select_option('[data-campo="tipo_documento"] select', "boleto")
    assert tela.page.locator('[data-campo="data_vencimento"].is-vazio').count() == 1  # obrigatorio no boleto
    assert tela.campo("tipo_documento").locator(".chip-corrigido").count() == 1


# ---------- exibicao formatada (o valor interno nao muda) ----------


def test_itens_em_pt_br_e_valores_numericos_intactos_no_json_e_no_excel(tela):
    import openpyxl

    tela.extrair("danfe_ok")
    celulas = tela.page.locator("#tabela-itens tbody tr:first-child td").all_inner_texts()
    assert celulas == ["Mochila", "1", "215,03", "215,03"]  # nunca "215.03"

    item = tela.json_bruto()["documento"]["itens"][0]
    assert item["valor_total"] == 215.03 and item["valor_unitario"] == 215.03  # numeros no JSON
    tela.baixar_excel()
    itens = openpyxl.load_workbook(tela.tmp_path / "exportado.xlsx")["Itens"]
    linha = [c.value for c in next(itens.iter_rows(min_row=2))]
    assert linha[3] == 215.03  # numero somavel no Excel, nao texto formatado


def test_tipo_com_rotulo_amigavel_e_valor_interno_no_json_e_no_excel(tela):
    tela.extrair("danfe_ok")
    assert tela.campo("tipo_documento").locator(".valor-texto").inner_text() == "Nota fiscal"
    assert tela.json_bruto()["documento"]["tipo_documento"] == "nota_fiscal"
    assert tela.baixar_excel()["Tipo de documento"] == "nota_fiscal"

    tela.page.click('[data-campo="tipo_documento"] .valor-btn')
    opcoes = tela.page.locator('[data-campo="tipo_documento"] option')
    assert opcoes.all_inner_texts() == ["Boleto", "Nota fiscal", "Pedido de compra", "Relatório", "Desconhecido"]
    assert opcoes.evaluate_all("els => els.map(e => e.value)") == [
        "boleto", "nota_fiscal", "pedido_compra", "relatorio", "desconhecido",
    ]

    tela.page.select_option('[data-campo="tipo_documento"] select', "pedido_compra")
    assert tela.campo("tipo_documento").locator(".valor-texto").inner_text() == "Pedido de compra"
    assert tela.json_bruto()["documento"]["tipo_documento"] == "pedido_compra"
    assert tela.baixar_excel()["Tipo de documento"] == "pedido_compra"


def test_boleto_mostra_tipo_amigavel(tela):
    tela.extrair("boleto")
    assert tela.campo("tipo_documento").locator(".valor-texto").inner_text() == "Boleto"


def test_chave_de_acesso_em_blocos_de_4_com_valor_intacto(tela):
    tela.extrair("danfe_ok")
    chave = tela.campo("extra-0")
    assert chave.locator(".valor-texto").inner_text() == helpers.CHAVE  # 11 blocos de 4
    assert tela.json_bruto()["documento"]["campos_adicionais"][0]["valor"] == helpers.CHAVE_SEM_ESPACOS
    assert tela.baixar_excel()["Chave de Acesso"] == helpers.CHAVE_SEM_ESPACOS

    # abrir e confirmar sem mudar NAO e correcao (os espacos sao so exibicao)
    tela.page.click('[data-campo="extra-0"] .valor-btn')
    editor = tela.campo("extra-0").locator("textarea")
    assert editor.input_value() == helpers.CHAVE
    tela.page.keyboard.press("Enter")
    assert tela.campo("extra-0").locator(".chip-alta").count() == 1
    assert "corrigido" not in tela.resumo()

    # editar guarda so digitos
    tela.page.click('[data-campo="extra-0"] .valor-btn')
    tela.campo("extra-0").locator("textarea").fill(helpers.CHAVE[:-1] + "0")
    tela.page.keyboard.press("Enter")
    assert tela.campo("extra-0").locator(".chip-corrigido").count() == 1
    valor = tela.json_bruto()["documento"]["campos_adicionais"][0]["valor"]
    assert valor == helpers.CHAVE_SEM_ESPACOS[:-1] + "0" and " " not in valor


# ---------- modo IA, acessibilidade, contraste, celular ----------


def test_modo_ia_sem_chips_de_confianca_mas_vazios_e_edicao_funcionam(tela):
    resposta_ia = {
        "modo_extracao": "ia", "origem_texto": "digital", "confiancas": {}, "avisos": [], "aviso": None,
        "documento": {
            "tipo_documento": "boleto", "numero_documento": "1234567890", "data_emissao": "08/07/2026",
            "data_vencimento": None, "emissor": "INSTITUICAO EXEMPLO", "destinatario": None, "valor_total": 900.0,
            "itens": [], "campos_adicionais": [{"campo": "Parcela", "valor": "2/5"}],
        },
    }
    tela.page.goto(tela.url)
    tela.page.route("**/extract-document", lambda rota: rota.fulfill(json=resposta_ia))
    tela.page.set_input_files("#file-input", str(tela.pdfs["boleto"]))
    tela.page.click("#btn-extrair")
    tela.page.wait_for_selector("#resultado:not([hidden])")

    assert tela.page.locator(".chip-alta, .chip-media, .chip-baixa").count() == 0
    assert tela.page.locator(".chip-vazio").count() == 2  # destinatario e data_vencimento (boleto)
    assert tela.page.locator("#nota-confianca").is_visible()
    assert tela.page.locator("#legenda-confianca").is_hidden()
    resumo = tela.resumo()
    assert "alta confiança" not in resumo and "2 vazios" in resumo

    tela.campo("destinatario").locator("textarea").fill("Fulano Exemplo")
    tela.page.keyboard.press("Enter")
    assert "1 corrigido" in tela.resumo()


def test_acessibilidade_basica(tela):
    tela.extrair("generico")
    ids = tela.page.evaluate("[...document.querySelectorAll('[id]')].map(e => e.id)")
    assert len(ids) == len(set(ids)), "ids duplicados"
    sem_nome = tela.page.evaluate("""() => [...document.querySelectorAll('#resultado textarea, #resultado select, #resultado .valor-btn')]
        .filter(e => {
          const ref = e.getAttribute('aria-labelledby');
          const nome = ref ? (document.getElementById(ref)?.textContent || '') : (e.getAttribute('aria-label') || '');
          return !nome.trim();
        }).length""")
    assert sem_nome == 0, "todo textarea/select/botao de valor precisa de nome acessivel"
    assert tela.page.locator("#resumo").get_attribute("aria-live") == "polite"
    assert tela.page.locator("#banner-avisos").get_attribute("role") == "status"
    assert tela.page.evaluate("[...document.querySelectorAll('.chip')].filter(c => !c.textContent.trim()).length") == 0


def test_contraste_de_chips_dicas_resumo_e_banner_wcag_aa(tela):
    def medir():
        return tela.page.evaluate("""() => [...document.querySelectorAll('.chip, .dica, .resumo-item, .banner')].map(e => {
            const cs = getComputedStyle(e); let bg = cs.backgroundColor, el = e;
            while ((bg === 'rgba(0, 0, 0, 0)' || bg === 'transparent') && el.parentElement) {
              el = el.parentElement; bg = getComputedStyle(el).backgroundColor;
            }
            return { classe: e.className, cor: cs.color, fundo: bg };
        })""")

    medidas = []
    tela.extrair("danfe_problemas")  # alta, media, vazio, banner
    tela.campo("numero_documento").locator("textarea").fill("000012345")
    tela.page.keyboard.press("Tab")  # + corrigido
    medidas += medir()
    tela.extrair("generico")  # baixa
    medidas += medir()

    assert {m["classe"].split()[-1] for m in medidas} >= {
        "chip-alta", "chip-media", "chip-baixa", "chip-vazio", "chip-corrigido", "banner", "dica",
    }
    for m in medidas:
        razao = helpers.contraste(helpers.rgb(m["cor"]), helpers.rgb(m["fundo"]))
        assert razao >= 4.5, f"{m['classe']}: contraste {razao:.2f}:1 < 4,5:1 ({m['cor']} sobre {m['fundo']})"


def test_celular_sem_rolagem_horizontal(tela):
    contexto = tela.page.context.browser.new_context(viewport={"width": 390, "height": 844})
    try:
        pagina = contexto.new_page()
        pagina.goto(tela.url)
        pagina.set_input_files("#file-input", str(tela.pdfs["danfe_problemas"]))
        pagina.click("#btn-extrair")
        pagina.wait_for_selector("#resultado:not([hidden])")
        assert pagina.evaluate("document.documentElement.scrollWidth") <= 390
    finally:
        contexto.close()
