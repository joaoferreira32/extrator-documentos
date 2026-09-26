"""Testes do DanfeExtractor e da reconstrucao de tabela por coordenadas.

SOBRE A FIXTURE DE TEXTO (danfe_real_anonimizado.txt) -- leia antes de
confiar nela: e um COMPOSTO de trechos reais, NAO as 84 linhas completas
que o extrator recebe.
- Vem da saida de /debug/extractor-input de uma DANFE real (anonimizada
  pelo usuario): linhas 00-02, 14-15, 38 (texto girado do canhoto), 39-41
  (emitente), 49-50 (IE + CNPJ do emitente), 51-52 (destinatario, com
  rotulos e valores COLADOS na mesma linha), 53-56 (grade de totais: linha
  de N rotulos, linha de N valores), 75-76 (cabecalho da tabela e produto).
- Endereco e chave de acesso vem de um trecho colado ANTES. A adjacencia
  entre os blocos e ASSUMIDA -- em particular, no PDF real o CNPJ fica 10
  linhas depois do nome do emitente; aqui fica mais perto (o teste
  test_cnpj_do_emissor_a_10_linhas_do_nome cobre a distancia real com
  preenchimento SINTETICO).
- Nao ha linha de numero da nota ("Nº ...") nem o resto das linhas.
Quando o usuario fornecer as 84 linhas completas, esta fixture deve ser
substituida por elas (ver CLAUDE.md, "Diagnosticando um PDF real").

Historico: uma fixture anterior tratava um trecho como se fosse o texto
inteiro e os testes passavam enquanto o PDF real falhava (emissor
"FOLHA 1/", destinatario com linha de rotulos, itens vazio).

Dados do destinatario sao ficticios (pessoa fisica de terceiro).
Emitente/CNPJ/chave de acesso/tabela de itens sao reais e publicos.
Trechos marcados "SINTETICO" abaixo sao montados a mao de proposito e nao
fingem ser texto real.
"""
import json
from pathlib import Path

from app.extractors.base import ContextoExtracao, Palavra
from app.extractors.danfe import (
    DanfeExtractor,
    extrair_destinatario,
    extrair_emissor,
    extrair_totais_grade,
    partes_da_chave,
    totais_fecham,
)
from app.extractors.danfe_tabela import _eh_texto_vertical, montar_tabela_itens

FIXTURE_TEXTO = Path(__file__).parent / "fixtures" / "danfe_real_anonimizado.txt"
FIXTURE_PALAVRAS = Path(__file__).parent / "fixtures" / "danfe_palavras_anonimizado.json"


def _carregar_paginas_palavras() -> list[list[Palavra]]:
    dados = json.loads(FIXTURE_PALAVRAS.read_text(encoding="utf-8"))
    return [
        [Palavra(texto=p["texto"], x0=p["x0"], x1=p["x1"], top=p["top"], bottom=p["bottom"]) for p in pagina]
        for pagina in dados["paginas"]
    ]


def _contexto() -> ContextoExtracao:
    texto = FIXTURE_TEXTO.read_text(encoding="utf-8")
    return ContextoExtracao(
        texto=texto, linhas=texto.splitlines(), paginas_palavras=_carregar_paginas_palavras()
    )


def test_deteccao_danfe():
    contexto = _contexto()
    assert DanfeExtractor().pontuacao_deteccao(contexto) > 0.5


def test_emissor_vem_da_linha_apos_identificacao_do_emitente():
    """Layout real: a legenda tem o titulo da caixa vizinha colado
    ("...emitente DANFE") e o nome, na linha seguinte, tem o titulo da
    outra caixa no fim ("... Documento Auxiliar da"). O rotulo generico
    "Emitente" capturava "DANFE" da primeira; e a "FOLHA 1/" vinha do texto
    girado que precede tudo."""
    linhas = [
        "FOLHA 1/",
        "e-FN",
        "Identificação do emitente DANFE",
        "DELL COMPUTADORES DO BRASIL LTDA Documento Auxiliar da",
        "Nota Fiscal Eletrônica",
        "72.381.189/0010-01",
    ]
    nome, documento = extrair_emissor(linhas)
    assert nome == "DELL COMPUTADORES DO BRASIL LTDA"
    assert documento == "CNPJ 72.381.189/0010-01"


def test_emissor_sem_a_legenda_nao_inventa():
    """SINTETICO: sem "Identificação do emitente" nao ha ancora -- devolve
    None em vez de chutar a primeira linha (que era "FOLHA 1/")."""
    assert extrair_emissor(["FOLHA 1/", "DANFE", "EMPRESA EXEMPLO LTDA"]) == (None, None)


def test_destinatario_com_rotulos_e_valores_colados_na_mesma_linha():
    """Layout real (linha 51): os rotulos vem primeiro, todos colados, e os
    valores depois -- o nome fica entre o ultimo rotulo e o CPF."""
    linhas = [
        "NOME/RAZÃO SOCIAL CNPJ/CPF DATA DA EMISSÃO FULANO DE TAL SILVA 000.000.000-00 "
        "15/4/2026 ENDEREÇO BAIRRO/DISTRITO CEP DATA DA ENTRADA/SAÍDA RUA EXEMPLO, nº 1 CENTRO 00000-000",
    ]
    assert extrair_destinatario(linhas) == ("FULANO DE TAL SILVA", "CPF 000.000.000-00", "alta")


def test_destinatario_com_rotulo_e_valor_em_linhas_separadas():
    """SINTETICO: outro layout (rotulos numa linha, valores na seguinte)."""
    linhas = ["NOME/RAZÃO SOCIAL CNPJ/CPF DATA DA EMISSÃO", "FULANO DE TAL 000.000.000-00 15/4/2026"]
    assert extrair_destinatario(linhas) == ("FULANO DE TAL", "CPF 000.000.000-00", "alta")


def test_destinatario_pessoa_juridica_devolve_cnpj():
    """SINTETICO."""
    linhas = ["NOME/RAZÃO SOCIAL CNPJ/CPF DATA DA EMISSÃO EMPRESA EXEMPLO LTDA 11.222.333/0001-81 15/4/2026"]
    assert extrair_destinatario(linhas) == ("EMPRESA EXEMPLO LTDA", "CNPJ 11.222.333/0001-81", "alta")


def test_destinatario_sem_cpf_e_media_e_nome_comecando_com_rotulo_curto_nao_e_cortado():
    """SINTETICO: sem CPF/CNPJ apos o nome nao ha a evidencia estrutural
    ("media"). "UFRJ ..." comeca com "UF" (um rotulo do bloco) mas nao
    pode ser tratado como rotulo -- limite de palavra."""
    linhas = ["NOME/RAZÃO SOCIAL CNPJ/CPF UFRJ COMERCIO LTDA"]
    assert extrair_destinatario(linhas) == ("UFRJ COMERCIO LTDA", None, "media")


def test_destinatario_so_com_rotulos_devolve_none():
    assert extrair_destinatario(["NOME/RAZÃO SOCIAL CNPJ/CPF DATA DA EMISSÃO"]) == (None, None, "")


def test_partes_da_chave():
    partes = partes_da_chave("35260472381189001001550010000123451123456786")
    assert partes["cnpj"] == "72381189001001"
    assert partes["numero"] == "000012345"
    assert partes["serie"] == "001"
    assert partes["aamm"] == "2604"


def _resultado_com_linhas(texto: str):
    contexto = ContextoExtracao(texto=texto, linhas=texto.splitlines(), paginas_palavras=None)
    return DanfeExtractor().extrair(contexto)


def test_numero_documento_conferido_com_a_chave_sobe_para_alta():
    """SINTETICO (a fixture composta nao tem linha de numero): o fallback
    "Nº..." acha o numero, mas so com "baixa"; a chave de acesso (DV
    valido) traz o mesmo numero -- evidencia independente."""
    texto = FIXTURE_TEXTO.read_text(encoding="utf-8") + "Nº000012345\n"
    resultado = _resultado_com_linhas(texto)
    assert resultado.documento.numero_documento == "000012345"
    assert resultado.confiancas["numero_documento"] == "alta"


def test_numero_documento_que_contradiz_a_chave_vira_baixa_com_aviso():
    texto = FIXTURE_TEXTO.read_text(encoding="utf-8") + "Nº999999999\n"
    resultado = _resultado_com_linhas(texto)
    assert resultado.confiancas["numero_documento"] == "baixa"
    assert any("número da nota" in aviso for aviso in resultado.avisos)


def test_emissor_com_cnpj_diferente_da_chave_vira_baixa_com_aviso():
    texto = FIXTURE_TEXTO.read_text(encoding="utf-8").replace("72.381.189/0010-01", "11.222.333/0001-81")
    resultado = _resultado_com_linhas(texto)
    assert resultado.confiancas["emissor"] == "baixa"
    assert any("CNPJ do emitente" in aviso for aviso in resultado.avisos)


def test_emissor_do_generico_com_titulo_e_descartado_e_sem_titulo_vira_baixa():
    """SINTETICO: sem a legenda real, o generico ("Emitente: X") so
    sobrevive como "baixa"; se ele capturou o titulo "DANFE" (o bug real),
    e descartado."""
    ok = _resultado_com_linhas("DANFE\nEmitente: EMPRESA EXEMPLO LTDA\n")
    assert ok.documento.emissor == "EMPRESA EXEMPLO LTDA"
    assert ok.confiancas["emissor"] == "baixa"

    ruim = _resultado_com_linhas("Emitente DANFE\n")
    assert ruim.documento.emissor is None
    assert "emissor" not in ruim.confiancas


def test_danfe_completa():
    resultado = DanfeExtractor().extrair(_contexto())
    doc = resultado.documento

    print(doc.model_dump_json(indent=2))
    print("confiancas:", resultado.confiancas)
    print("avisos:", resultado.avisos)

    assert doc.tipo_documento == "nota_fiscal"

    campos = {c.campo: c.valor for c in doc.campos_adicionais}
    # Chave de acesso real (44 digitos, DV valido -- conferido a mao com o
    # mesmo algoritmo de test_chave_acesso.py).
    assert campos.get("Chave de Acesso") == "35260472381189001001550010000123451123456786"
    assert resultado.confiancas.get("Chave de Acesso") == "alta"

    # emissor: a fixture comeca com texto girado do canhoto ("FOLHA 1/",
    # "e-FN", ...); o nome vem da linha apos "Identificação do emitente",
    # sem o "Documento Auxiliar da" da caixa vizinha. Confianca "alta"
    # porque o CNPJ lido bate com o CNPJ embutido na chave de acesso.
    assert doc.emissor == "DELL COMPUTADORES DO BRASIL LTDA (CNPJ 72.381.189/0010-01)"
    assert resultado.confiancas.get("emissor") == "alta"

    # destinatario: mesmo formato do emissor ("NOME (CPF ...)"), da mesma
    # linha dos rotulos, entre o ultimo rotulo e o CPF.
    assert doc.destinatario == "FULANO DE TAL SILVA (CPF 000.000.000-00)"
    assert resultado.confiancas.get("destinatario") == "alta"

    # data_emissao normalizada pra dd/mm/aaaa (o documento imprime "15/4/2026").
    assert doc.data_emissao == "15/04/2026"
    assert resultado.confiancas.get("data_emissao") == "alta"

    # valor_total: pela POSICAO na grade de totais (6º rotulo -> 6º valor da
    # linha 56), e "alta" porque a formula do total fecha.
    assert doc.valor_total == 229.0
    assert resultado.confiancas.get("valor_total") == "alta"

    # --- Tabela de itens por coordenadas ---
    assert len(doc.itens) == 1
    item = doc.itens[0]
    assert item.descricao == "Mochila Dell Gaming Backpack 17, GM1720PM"
    assert item.quantidade == 1.0
    assert item.valor_unitario == 215.03
    assert item.valor_total == 215.03
    # A soma dos itens (215,03) fecha com o Valor Total dos Produtos da grade.
    assert resultado.confiancas.get("itens") == "alta"

    assert campos.get("CFOP") == "5102"

    # Valor Total dos Produtos exposto como campo_adicional (pra quem
    # exporta pra Excel poder conferir a soma dos itens numero-contra-numero
    # na aba Campos adicionais, sem depender so do texto do aviso).
    assert campos.get("Valor Total dos Produtos") == "215,03"
    assert resultado.confiancas.get("Valor Total dos Produtos") == "alta"

    # Soma dos itens (215,03) bate com "Valor Total dos Produtos" (215,03)
    # na fixture -- nao deve gerar aviso de divergencia.
    assert resultado.avisos == []


def test_soma_divergente_gera_aviso():
    # "Valor Total dos Produtos" e o 5º valor da linha de valores da grade
    # (229,00 41,22 0,00 0,00 215,03) -- troca so ele.
    texto = FIXTURE_TEXTO.read_text(encoding="utf-8").replace(
        "229,00 41,22 0,00 0,00 215,03", "229,00 41,22 0,00 0,00 999,99"
    )
    contexto = ContextoExtracao(
        texto=texto, linhas=texto.splitlines(), paginas_palavras=_carregar_paginas_palavras()
    )
    resultado = DanfeExtractor().extrair(contexto)
    (aviso,) = [a for a in resultado.avisos if "não bate" in a]
    # valores em pt-BR, igual a tela e ao Excel (antes saia "R$ 215.03")
    assert "(R$ 215,03)" in aviso and "(R$ 999,99)" in aviso
    # o campo continua exposto mesmo quando diverge -- o valor "errado" (do
    # jeito que veio no documento) e o que interessa mostrar, nao um valor
    # corrigido por nos
    campos = {c.campo: c.valor for c in resultado.documento.campos_adicionais}
    assert campos.get("Valor Total dos Produtos") == "999,99"


def test_montar_tabela_itens_descarta_texto_vertical():
    paginas = _carregar_paginas_palavras()
    tabela = montar_tabela_itens(paginas)

    assert len(tabela.itens) == 1
    # Nenhum dos rotulos verticais (SOTUDORP, LISARB, ...) deve aparecer
    # em nenhum campo do item.
    item_texto = tabela.itens[0].descricao
    for ruido in ["SOTUDORP", "LISARB", "ADARITER", "SERODATUPMOC", "OIRÁTANITSED"]:
        assert ruido not in item_texto


def test_montar_tabela_itens_para_em_informacoes_complementares():
    """Bug real: com o boleto de teste em navegador, o bloco 'INFORMACOES
    COMPLEMENTARES' que vem depois da tabela de itens (secao padronizada
    nacionalmente, nao especifica de um emissor) caia por coincidencia de
    x na coluna CODIGO e virava uma sequencia de itens fantasmas. As
    coordenadas exatas do bloco na DANFE real do usuario nao foram
    fornecidas (so o teste em navegador reportou top aproximado
    491-517) -- aqui adicionamos linhas representativas dessa secao por
    cima da fixture real (mesma pagina, coordenadas plausiveis, top maior
    que o item real) so pra validar que o corte por conteudo funciona,
    independente da posicao vertical exata."""
    paginas = _carregar_paginas_palavras()
    pagina = paginas[0]

    def w(texto, x0, top):
        return Palavra(texto=texto, x0=x0, x1=x0 + len(texto) * 6, top=top, bottom=top + 7)

    pagina.extend(
        [
            w("INFORMAÇÕES", 40, 491),
            w("COMPLEMENTARES", 110, 491),
            w("Reservado", 40, 505),
            w("ao", 100, 505),
            w("Fisco", 115, 505),
        ]
    )

    tabela = montar_tabela_itens(paginas)
    assert len(tabela.itens) == 1
    assert "Reservado" not in tabela.itens[0].descricao


def test_montar_tabela_itens_separa_valores_colados():
    """O token real "13,9718,00" (VALOR I.P.I. + ALIQUOTA ICMS colados)
    nao pode vazar pro valor_total do item nem corromper o valor_unitario
    -- eles vem de colunas antes do token colado, entao devem sair
    intactos."""
    tabela = montar_tabela_itens(_carregar_paginas_palavras())
    item = tabela.itens[0]
    assert item.valor_unitario == 215.03 or item.valor_unitario == 215.0300  # mesma coisa
    assert item.valor_total == 215.03


def test_eh_texto_vertical():
    vertical = Palavra(texto="SOTUDORP", x0=20, x1=28, top=150, bottom=210)
    horizontal = Palavra(texto="460-BCZS", x0=85, x1=100, top=422.1, bottom=429.4)
    assert _eh_texto_vertical(vertical) is True
    assert _eh_texto_vertical(horizontal) is False


def test_eh_texto_vertical_fora_da_faixa_x_nao_e_descartado():
    """A regra so vale na faixa de x0 15-85 (margem lateral) -- uma
    palavra alta e estreita fora dessa faixa nao deveria ser confundida
    com um rotulo rotacionado."""
    fora_da_faixa = Palavra(texto="ALGO", x0=300, x1=308, top=150, bottom=210)
    assert _eh_texto_vertical(fora_da_faixa) is False


def _palavra(texto, x0, top, largura=8):
    return Palavra(texto=texto, x0=x0, x1=x0 + largura, top=top, bottom=top + 7.3)


def test_fragmentos_do_cabecalho_abaixo_dele_nao_encerram_a_tabela():
    """Reproduz o FORMATO do bug real (trace de /debug/extractor-input):
    cabecalho em 413.6 e, logo abaixo (416.7), uma linha so de fragmentos
    do cabecalho ("IC M S IP I") -- sem CODIGO e nao e so descricao, entao
    era "fim" e a tabela saia vazia sem nunca chegar no produto.

    Aqui usa 416.8: o cabecalho da fixture esta em 413.7 e a diferenca
    precisa passar de 3.0 (tolerancia de agrupamento de linha), como no
    real (3.1). As posicoes x dos fragmentos sao plausiveis (coluna de
    ICMS/IPI), nao medidas -- o que importa e o top e a ausencia de
    CODIGO."""
    paginas = _carregar_paginas_palavras()
    paginas[0].extend(
        [
            _palavra("IC", 705.0, 416.8),
            _palavra("M", 714.0, 416.8),
            _palavra("S", 721.0, 416.8),
            _palavra("IP", 749.7, 416.8),
            _palavra("I", 760.0, 416.8),
        ]
    )
    diagnostico: list[dict] = []
    tabela = montar_tabela_itens(paginas, diagnostico=diagnostico)

    assert len(tabela.itens) == 1
    assert tabela.itens[0].descricao == "Mochila Dell Gaming Backpack 17, GM1720PM"
    decisoes = [linha["decisao"] for linha in diagnostico[0]["linhas"]]
    assert decisoes[0].startswith("pulada: fragmento do cabecalho")
    assert decisoes[1] == "item"


def test_linha_fim_depois_do_primeiro_item_continua_encerrando_a_tabela():
    """A tolerancia a fragmentos vale so ANTES do primeiro item: depois
    dele, uma linha sem CODIGO (rodape/secao seguinte) ainda encerra a
    tabela, mesmo estando perto do cabecalho."""
    paginas = _carregar_paginas_palavras()
    paginas[0].extend([_palavra("Rodape", 705.0, 430.0), _palavra("qualquer", 749.7, 430.0)])
    diagnostico: list[dict] = []
    tabela = montar_tabela_itens(paginas, diagnostico=diagnostico)

    assert len(tabela.itens) == 1
    assert diagnostico[0]["linhas"][-1]["decisao"].startswith("fim: sem CODIGO")


# --- Grade de totais: rotulo -> valor de MESMA POSICAO ----------------------------

LINHAS_GRADE_REAL = [
    "BASE DE CÁLCULO DO ICMS VALOR DO ICMS BASE DE CÁLCULO ICMS ST VALOR DO ICMS SUBSTITUIÇÃO VALOR TOTAL DOS PRODUTOS",
    "229,00 41,22 0,00 0,00 215,03",
    "VALOR DO FRETE VALOR DO SEGURO DESCONTO OUTRAS DESPESAS ACESSÓRIAS VALOR TOTAL DO I.P.I. VALOR TOTAL DA NOTA",
    "0,00 0,00 0,00 0,00 13,97 229,00",
]


def test_totais_da_grade_casados_por_posicao():
    """Linhas reais 53-56. Antes, o valor era "o numero mais proximo do
    rotulo": produtos saia 229,00 (a base do ICMS, 1º valor da linha) em vez
    de 215,03, e o total da nota so estava certo por coincidencia (nota =
    produtos + IPI = 215,03 + 13,97 = 229,00 = base do ICMS)."""
    totais = extrair_totais_grade(LINHAS_GRADE_REAL)

    assert totais["valor_produtos"] == 215.03
    assert totais["ipi"] == 13.97
    assert totais["valor_nota"] == 229.0
    assert totais["base_icms"] == 229.0
    assert totais["valor_icms"] == 41.22
    assert totais["base_icms_st"] == 0.0
    assert totais["valor_icms_st"] == 0.0
    assert totais["frete"] == 0.0
    assert totais["seguro"] == 0.0
    assert totais["desconto"] == 0.0
    assert totais["outras_despesas"] == 0.0
    assert totais_fecham(totais) is True


def test_valor_total_segue_a_posicao_e_nao_o_numero_mais_proximo():
    """Prova de que e posicional: muda so o 6º valor da 2ª linha. Se o
    extrator ainda pegasse o "vizinho" (229,00 da linha de cima), o
    total continuaria 229,00."""
    texto = FIXTURE_TEXTO.read_text(encoding="utf-8").replace(
        "0,00 0,00 0,00 0,00 13,97 229,00", "0,00 0,00 0,00 0,00 13,97 300,00"
    )
    resultado = _resultado_com_linhas(texto)

    assert resultado.documento.valor_total == 300.0
    # A formula nao fecha (215,03 + 13,97 != 300,00): confianca cai e avisa.
    assert resultado.confiancas["valor_total"] == "media"
    assert any("não fecham" in aviso for aviso in resultado.avisos)


def test_grade_com_quantidade_de_valores_diferente_nao_chuta():
    """SINTETICO: 5 rotulos e 4 valores -- nao adivinha qual faltou."""
    linhas = [
        "BASE DE CÁLCULO DO ICMS VALOR DO ICMS BASE DE CÁLCULO ICMS ST VALOR DO ICMS SUBSTITUIÇÃO VALOR TOTAL DOS PRODUTOS",
        "229,00 41,22 0,00 215,03",
    ]
    assert extrair_totais_grade(linhas) == {}


def test_grade_com_rotulo_desconhecido_na_linha_nao_e_lida():
    """SINTETICO: um rotulo que a gente nao conhece (ex: imposto de
    importacao) muda a contagem/posicao -- a linha inteira e ignorada em vez
    de deslocar os valores."""
    linhas = [
        "VALOR DO FRETE VALOR DO IMPOSTO DE IMPORTAÇÃO VALOR TOTAL DA NOTA",
        "0,00 5,00 229,00",
    ]
    assert extrair_totais_grade(linhas) == {}


def test_valores_e_rotulos_na_mesma_linha():
    """SINTETICO: quando o pdfplumber cola rotulos e valores numa linha so
    (como no bloco do destinatario), a ordem e a mesma."""
    linhas = ["VALOR DO FRETE VALOR TOTAL DA NOTA 3,50 229,00"]
    assert extrair_totais_grade(linhas) == {"frete": 3.5, "valor_nota": 229.0}


def test_valor_total_na_mesma_linha_do_rotulo_quando_nao_ha_grade():
    """SINTETICO: outro layout ("Valor Total da Nota 229,00") continua
    funcionando pelo fallback da mesma linha."""
    resultado = _resultado_com_linhas("DANFE\nValor Total da Nota 229,00\n")
    assert resultado.documento.valor_total == 229.0
    assert resultado.confiancas["valor_total"] == "alta"


# --- CNPJ do emissor ------------------------------------------------------------------

def _linhas_emitente_com_preenchimento(n_preenchimento: int, cnpj: str = "72.381.189/0010-01") -> list[str]:
    return (
        ["Identificação do emitente DANFE", "DELL COMPUTADORES DO BRASIL LTDA Documento Auxiliar da"]
        + ["Nota Fiscal Eletrônica"]
        + ["linha de endereco/IE"] * n_preenchimento
        + ["INSCRIÇÃO ESTADUAL INSCR. ESTADUAL DO SUBST. TRIBUT. CNPJ", f"748241245113 {cnpj}"]
        + ["NOME/RAZÃO SOCIAL CNPJ/CPF DATA DA EMISSÃO FULANO DE TAL SILVA 000.000.000-00 15/4/2026"]
    )


def test_cnpj_do_emissor_a_10_linhas_do_nome():
    """Bug real: no PDF real o nome esta na linha 40 e o CNPJ na 50. A janela
    fixa de 10 linhas (`documento_fiscal_proximo`) cobria 40..49 e perdia
    o CNPJ por UMA linha -- emissor saia sem CNPJ e com confianca "media".
    (Preenchimento SINTETICO: so a distancia e real.)"""
    linhas = _linhas_emitente_com_preenchimento(7)
    assert linhas.index("748241245113 72.381.189/0010-01") - 1 == 10  # nome -> CNPJ = 10 linhas
    assert extrair_emissor(linhas) == ("DELL COMPUTADORES DO BRASIL LTDA", "CNPJ 72.381.189/0010-01")


def test_cnpj_do_emissor_bem_longe_do_nome_tambem_e_achado():
    linhas = _linhas_emitente_com_preenchimento(25)
    assert extrair_emissor(linhas)[1] == "CNPJ 72.381.189/0010-01"


def test_cnpj_do_destinatario_nao_e_tomado_pelo_do_emissor():
    """SINTETICO: o bloco do emitente termina no rotulo do destinatario. Sem
    CNPJ no bloco do emitente, o CNPJ que vem depois de NOME/RAZÃO SOCIAL
    (outra entidade) nao pode virar o do emissor."""
    linhas = [
        "Identificação do emitente DANFE",
        "DELL COMPUTADORES DO BRASIL LTDA Documento Auxiliar da",
        "NOME/RAZÃO SOCIAL CNPJ/CPF EMPRESA CLIENTE LTDA 11.222.333/0001-81 15/4/2026",
    ]
    assert extrair_emissor(linhas) == ("DELL COMPUTADORES DO BRASIL LTDA", None)


def test_com_dois_cnpjs_no_bloco_prefere_o_que_bate_com_a_chave():
    """SINTETICO: ex. CNPJ de uma filial/substituto tributario no bloco."""
    linhas = [
        "Identificação do emitente DANFE",
        "DELL COMPUTADORES DO BRASIL LTDA Documento Auxiliar da",
        "CNPJ SUBSTITUTO 99.888.777/0001-66",
        "748241245113 72.381.189/0010-01",
    ]
    assert extrair_emissor(linhas)[1] == "CNPJ 99.888.777/0001-66"  # sem chave: o primeiro
    assert extrair_emissor(linhas, cnpj_chave="72381189001001")[1] == "CNPJ 72.381.189/0010-01"


def test_emissor_no_pdf_real_tem_cnpj_conferido_e_confianca_alta():
    """Fim a fim com a distancia real (nome -> CNPJ = 10 linhas) e uma chave
    valida: o emissor sai "NOME (CNPJ ...)" e "alta"."""
    chave = "3526 0472 3811 8900 1001 5500 1000 0123 4511 2345 6786"
    linhas = _linhas_emitente_com_preenchimento(6) + [chave]
    resultado = _resultado_com_linhas("\n".join(linhas))

    assert resultado.documento.emissor == "DELL COMPUTADORES DO BRASIL LTDA (CNPJ 72.381.189/0010-01)"
    assert resultado.confiancas["emissor"] == "alta"


def test_rotulo_sozinho_na_linha_nao_pega_o_numero_da_linha_seguinte():
    """SINTETICO: com UM rotulo por linha o numero da linha seguinte e
    ambiguo (valor deste campo ou do proximo). Um layout "valor, legenda,
    valor, legenda..." fazia "VALOR TOTAL DOS PRODUTOS" ler o 0,00 do campo
    seguinte -- e gerava um aviso falso de soma. Nao chuta."""
    linhas = [
        "215,03",
        "VALOR TOTAL DOS PRODUTOS",
        "0,00",
        "VALOR DO FRETE",
    ]
    assert extrair_totais_grade(linhas) == {}


def test_danfe_nao_usa_a_linha_seguinte_pra_valores_de_rotulo():
    """A DANFE nao liga `aceitar_linha_seguinte`: com a leitura por
    posicao na grade, o resultado nao depende de "o numero da linha ao
    lado". (Garante que o ajuste do boleto nao vazou pra ca.)"""
    resultado = DanfeExtractor().extrair(_contexto())
    assert resultado.documento.valor_total == 229.0
    assert resultado.avisos == []


def test_confianca_dos_itens_e_alta_so_quando_a_soma_fecha():
    """A tabela por coordenada e heuristica posicional ("media"); vira "alta"
    quando a soma dos itens fecha com o Valor Total dos Produtos."""
    texto = FIXTURE_TEXTO.read_text(encoding="utf-8")
    contexto_ok = ContextoExtracao(
        texto=texto, linhas=texto.splitlines(), paginas_palavras=_carregar_paginas_palavras()
    )
    assert DanfeExtractor().extrair(contexto_ok).confiancas["itens"] == "alta"

    # Soma NAO fecha (produtos 999,99 na grade): continua "media" e avisa.
    divergente = texto.replace("229,00 41,22 0,00 0,00 215,03", "229,00 41,22 0,00 0,00 999,99")
    contexto_div = ContextoExtracao(
        texto=divergente, linhas=divergente.splitlines(), paginas_palavras=_carregar_paginas_palavras()
    )
    resultado = DanfeExtractor().extrair(contexto_div)
    assert resultado.confiancas["itens"] == "media"
    assert any("não bate" in aviso for aviso in resultado.avisos)


def test_confianca_dos_itens_fica_media_sem_valor_total_dos_produtos_pra_conferir():
    """SINTETICO: sem "Valor Total dos Produtos" no texto nao ha com o que
    conferir a soma -- nao sobe pra "alta"."""
    texto = "DANFE" + chr(10) + "DELL COMPUTADORES" + chr(10)
    contexto = ContextoExtracao(
        texto=texto, linhas=texto.splitlines(), paginas_palavras=_carregar_paginas_palavras()
    )
    resultado = DanfeExtractor().extrair(contexto)
    assert resultado.documento.itens  # a tabela foi lida
    assert resultado.confiancas["itens"] == "media"
