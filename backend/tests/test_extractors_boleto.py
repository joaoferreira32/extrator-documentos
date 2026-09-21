"""Teste de regressao com o texto bruto de um boleto real (anonimizado).

Este fixture existe porque cenarios sinteticos escritos a mao nao
reproduziam o layout real de um boleto bancario -- eles passavam mas o
extrator continuava errando no documento de verdade (o mesmo rotulo
aparece varias vezes em contextos diferentes: cabecalho do recibo,
tabela-resumo no rodape, campos de verdade). Ver CLAUDE.md.

Movido de test_basic_extractor.py quando a extracao de boleto virou
`BoletoExtractor` (padrao Strategy) -- mesmo comportamento, testado
diretamente na classe em vez de pela funcao de compatibilidade.
"""
import logging
from pathlib import Path

from app.extractors.base import ContextoExtracao
from app.extractors.boleto import BoletoExtractor
from app.extractors import comum

FIXTURE = Path(__file__).parent / "fixtures" / "boleto_real_anonimizado.txt"


def _extrair(texto: str):
    contexto = ContextoExtracao(texto=texto, linhas=texto.splitlines())
    return BoletoExtractor().extrair(contexto)


def test_boleto_real_anonimizado(caplog):
    texto = FIXTURE.read_text(encoding="utf-8")

    with caplog.at_level(logging.DEBUG, logger="app.extractors.comum"):
        resultado = _extrair(texto)
    doc = resultado.documento

    print("\n--- log de extracao (rotulo/linha escolhido por campo) ---")
    for registro in caplog.records:
        print(registro.getMessage())

    print("\n--- resultado ---")
    print(doc.model_dump_json(indent=2))
    print("confiancas:", resultado.confiancas)

    campos = {c.campo: c.valor for c in doc.campos_adicionais}

    assert doc.destinatario is not None and "Maria" in doc.destinatario, (
        f"destinatario deveria conter o nome do pagador, veio {doc.destinatario!r}"
    )
    assert doc.numero_documento == "1234567890", (
        f"numero_documento errado: {doc.numero_documento!r}"
    )
    assert campos.get("Nosso Número") == "10200000001-9", (
        f"nosso_numero errado: {campos.get('Nosso Número')!r}"
    )
    # Campos vindos de rotulo explicito devem ter confianca alta.
    assert resultado.confiancas.get("numero_documento") == "alta"
    assert resultado.confiancas.get("destinatario") == "alta"


def test_deteccao_boleto():
    texto = FIXTURE.read_text(encoding="utf-8")
    contexto = ContextoExtracao(texto=texto, linhas=texto.splitlines())
    assert BoletoExtractor().pontuacao_deteccao(contexto) > 0.5


def test_corrige_confusao_ocr_em_cnpj():
    """OCR troca com frequencia O<->0, I<->1, S<->5 em campos numericos."""
    texto = "Beneficiário: Empresa Teste Ltda\nCNPJ: 12.34S.678/OOO1-9I"
    doc = _extrair(texto).documento
    assert doc.emissor == "Empresa Teste Ltda (CNPJ 12.345.678/0001-91)"


def test_corrige_confusao_ocr_em_cpf():
    texto = "Pagador Fulano de Tal CPF: I11.222.333-44"
    doc = _extrair(texto).documento
    assert doc.destinatario == "Fulano de Tal (CPF 111.222.333-44)"


def test_corrige_confusao_ocr_em_valor():
    from app.extractors.boleto import ROTULOS_VALOR_A_PAGAR

    valor = comum.extrair_valor_rotulo("Valor a Pagar = R$ 9OO,OO", ROTULOS_VALOR_A_PAGAR)
    assert valor == "900,00"


def test_numero_documento_curto_sem_rotulo_e_rejeitado():
    """Regressao do bug original: numero solto curto (ex: '02') nao pode
    virar numero_documento sem vir de um rotulo confiavel."""
    texto = "BANCO EXEMPLO S.A.\nAceite N  02  Especie Doc DM\nCedente: Empresa ABC Ltda"
    doc = _extrair(texto).documento
    assert doc.numero_documento != "02"


def _extrair_fixture():
    texto = FIXTURE.read_text(encoding="utf-8")
    contexto = ContextoExtracao(texto=texto, linhas=texto.splitlines(), paginas_palavras=None)
    return BoletoExtractor().extrair(contexto)


def test_boleto_traz_os_7_campos_adicionais():
    """Trava os 7 campos_adicionais do boleto real, com valores, ordem e
    confianca. Regressao real: "Valor do Documento" sumiu quando
    `comum.extrair_valor_rotulo` foi "restaurado" pra so a mesma linha -- o
    boleto real tem o rotulo numa linha ("Valor Documento (-) desconto ...")
    e o valor na SEGUINTE ("1.000,00"), e so saia gracas a um fallback que
    eu tinha adicionado pra DANFE. Nenhum teste travava esse campo, entao a
    suite passou com o boleto quebrado (o codigo original, pre-refatoracao,
    tambem nao o extraia desse texto)."""
    resultado = _extrair_fixture()

    campos = [(c.campo, c.valor) for c in resultado.documento.campos_adicionais]
    assert campos == [
        ("Linha digitável", "11111.11111 11111.111111 11111.111111 1 11111111111111"),
        ("Nosso Número", "10200000001-9"),
        ("Valor do Documento", "1.000,00"),
        ("Desconto", "100,00"),
        ("Valor a Pagar", "900,00"),
        ("Parcela", "2/5"),
        ("Agência/Código Beneficiário", "0000-0/0000000"),
    ]
    for campo, _ in campos:
        assert resultado.confiancas[campo] == "alta", campo


def test_valor_do_documento_no_boleto_vem_da_linha_seguinte_ao_rotulo():
    """Fixa o motivo: o valor esta na linha depois do rotulo."""
    linhas = FIXTURE.read_text(encoding="utf-8").splitlines()
    i = next(i for i, linha in enumerate(linhas) if linha.startswith("Valor Documento"))
    assert "1.000,00" not in linhas[i]
    assert linhas[i + 1].strip() == "1.000,00"
