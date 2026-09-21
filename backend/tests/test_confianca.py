"""Regras da confianca geral em Python -- a mesma conta da tela (script.js)."""
from app.confianca import estado_do_campo, resumo_confianca
from app.schemas import DocumentoExtraido, ExtractionResult


def _resultado(confiancas, **documento):
    return ExtractionResult(
        modo_extracao="basico", confiancas=confiancas, documento=DocumentoExtraido(**documento)
    )


def test_conta_alta_e_revisar_so_de_campos_com_valor_e_confianca():
    r = _resultado(
        {"numero_documento": "alta", "emissor": "media", "valor_total": "baixa", "CFOP": "alta", "itens": "alta"},
        numero_documento="1", emissor="E", valor_total=10.0, data_emissao=None,  # data vazia: nao conta
        campos_adicionais=[{"campo": "CFOP", "valor": "5102"}],
        itens=[{"descricao": "x"}],
    )
    resumo = resumo_confianca(r)
    assert (resumo.alta, resumo.revisar, resumo.total) == (3, 2, 5)  # numero, CFOP, itens | emissor, valor
    assert resumo.texto == "3 de 5 alta"


def test_tipo_do_documento_nunca_conta():
    r = _resultado({"tipo_documento": "alta"}, tipo_documento="boleto")
    assert resumo_confianca(r).total == 0


def test_corrigido_fica_fora_do_x_de_y_e_conta_a_parte():
    r = _resultado({"numero_documento": "alta", "emissor": "alta"}, numero_documento="1", emissor="E")
    resumo = resumo_confianca(r, {"emissor"})
    assert (resumo.alta, resumo.total, resumo.corrigidos) == (1, 1, 1)


def test_corrigido_e_depois_esvaziado_nao_conta_em_lugar_nenhum():
    r = _resultado({"emissor": "alta"}, emissor=None)
    resumo = resumo_confianca(r, {"emissor"})
    assert (resumo.alta, resumo.total, resumo.corrigidos) == (0, 0, 0)


def test_itens_sem_itens_nao_contam_mesmo_com_confianca():
    r = _resultado({"itens": "media"}, itens=[])
    assert resumo_confianca(r).total == 0


def test_modo_ia_sem_confiancas_nao_tem_resumo():
    assert resumo_confianca(_resultado({}, numero_documento="1")) is None


def test_estado_do_campo_ordem_vazio_corrigido_confianca():
    conf = {"a": "media"}
    assert estado_do_campo("", "a", conf, {"a"}) is None  # vazio vence
    assert estado_do_campo("v", "a", conf, {"a"}) == "corrigido"
    assert estado_do_campo("v", "a", conf, set()) == "media"
    assert estado_do_campo("v", "sem_conf", conf, set()) is None
