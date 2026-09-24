"""O .xlsx gerado tem que abrir no Excel sem "reparar" -- nao so no openpyxl.

Bug real: o arquivo da etapa 8 passava em todos os testes (que releem com
openpyxl) e o Excel de verdade recusava e oferecia reparar, perdendo a
formatacao. Causa: Tabelas nomeadas com ref so no cabecalho (aba sem dados --
Avisos de quase todo documento, Itens de boleto) e celulas da Legenda tipadas
como texto sem texto. Ver tests/ooxml.py para as regras e por que o schema
sozinho nao basta.

Duas camadas:
- `ooxml.problemas` (sempre roda): invariantes do Excel lidos do XML bruto.
- validador oficial da Microsoft (opt-in, `--ooxml-sdk`, so Windows): schema
  OOXML via Open XML SDK. Complementar -- ele NAO pegou o bug acima.
"""
import os
import subprocess
import sys
import urllib.request
import zipfile
from datetime import date
from pathlib import Path

import openpyxl
import pytest
from openpyxl.worksheet.table import Table

sys.path.insert(0, str(Path(__file__).resolve().parent))
import ooxml  # noqa: E402

from app.excel_exporter import gerar_excel  # noqa: E402
from app.schemas import DocumentoExtraido, DocumentoParaExportar, ExtractionResult  # noqa: E402

CHAVE_44 = "35260472381189001001550010000123451123456786"


def _doc(arquivo="a.pdf", confiancas=None, avisos=None, corrigidos=None, modo="basico", **documento):
    documento.setdefault("tipo_documento", "nota_fiscal")
    return DocumentoParaExportar(
        arquivo=arquivo,
        resultado=ExtractionResult(
            modo_extracao=modo, confiancas=confiancas or {}, avisos=avisos or [], documento=DocumentoExtraido(**documento)
        ),
        corrigidos=corrigidos or {},
    )


# Um cenario por formato de exportacao que existe de verdade (e alguns de
# borda). O que importa e cobrir abas VAZIAS e CHEIAS, porque foi aba vazia que
# quebrou o arquivo.
CENARIOS = {
    "boleto (sem itens nem avisos)": [
        _doc(tipo_documento="boleto", numero_documento="123", data_vencimento="10/05/2026", valor_total=900.0,
             emissor="ESCOLA (CNPJ 11.111.111/0001-11)", destinatario="FULANO (CPF 000.000.000-00)",
             campos_adicionais=[{"campo": "Nosso Número", "valor": "10200000001-9"},
                                {"campo": "Valor do Documento", "valor": "900,00"}],
             confiancas={"numero_documento": "alta", "valor_total": "alta"})
    ],
    "danfe (itens, sem avisos)": [
        _doc(numero_documento="000012345", valor_total=229.0, data_emissao="15/04/2026",
             itens=[{"descricao": "Mochila", "quantidade": 1.0, "valor_unitario": 215.03, "valor_total": 215.03}],
             campos_adicionais=[{"campo": "Chave de Acesso", "valor": CHAVE_44}, {"campo": "CFOP", "valor": "5102"}],
             confiancas={"itens": "alta", "CFOP": "media", "valor_total": "alta"})
    ],
    "tudo preenchido, com media/baixa/corrigido": [
        _doc(numero_documento="1", emissor="E (CNPJ 11.222.333/0001-81)", valor_total=10.0, data_emissao="15/04/2026",
             itens=[{"descricao": "a", "valor_total": 1.0}, {"descricao": "b", "quantidade": "1 un", "valor_total": "x"}],
             campos_adicionais=[{"campo": "CFOP", "valor": "5102"}], avisos=["a soma nao bate"],
             confiancas={"emissor": "baixa", "data_emissao": "media", "itens": "media"},
             corrigidos={"valor_total": 9.0})
    ],
    "vazio (tipo desconhecido, nada extraido)": [_doc(tipo_documento="desconhecido")],
    "modo IA (sem confiancas)": [_doc(modo="ia", numero_documento="1", valor_total=1.0, itens=[{"descricao": "x"}])],
    "lote de 3": [
        _doc(arquivo="1.pdf", numero_documento="1", itens=[{"descricao": "a", "valor_total": 1.0}]),
        _doc(arquivo="2.pdf", tipo_documento="boleto", numero_documento="2", avisos=["x"]),
        _doc(arquivo="3.pdf", tipo_documento="desconhecido"),
    ],
    "texto hostil (formula, controle, gigante)": [
        _doc(emissor='=HYPERLINK("http://x","y")', destinatario="A\x00B\x0bC", numero_documento="X" * 40000,
             itens=[{"descricao": "=SUM(A1:A2)"}], avisos=["=cmd|' /C calc'!A0"])
    ],
}


@pytest.mark.parametrize("nome", list(CENARIOS))
def test_arquivo_gerado_respeita_as_regras_do_excel(nome):
    achados = ooxml.problemas(gerar_excel(CENARIOS[nome], data_geracao=date(2026, 9, 23)))
    assert achados == [], f"{nome}: o Excel provavelmente pediria pra reparar:\n" + "\n".join(achados)


def test_aba_sem_dados_fica_sem_tabela_e_aba_com_dados_tem():
    wb = openpyxl.load_workbook(gerar_excel(CENARIOS["boleto (sem itens nem avisos)"]))
    assert wb["Itens"].tables == {} and wb["Avisos"].tables == {}
    assert "CamposAdicionais" in wb["Campos adicionais"].tables and "Resumo" in wb["Resumo"].tables


# ---------- o verificador pega o que tem que pegar (senao ele e decorativo) ----------


def _salvar(wb) -> bytes:
    from io import BytesIO

    buf = BytesIO()
    wb.save(buf)
    return buf.getvalue()


def test_verificador_acusa_tabela_so_com_cabecalho():
    """Reproduz o bug que chegou ao Excel: tem que ser acusado."""
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.append(["ID", "Aviso"])
    ws.add_table(Table(displayName="Avisos", ref="A1:B1"))
    assert any("sem nenhuma linha de dados" in a for a in ooxml.problemas(_salvar(wb)))


def test_verificador_acusa_celula_texto_sem_texto():
    wb = openpyxl.Workbook()
    wb.active["A1"] = ""
    wb.active["A1"].data_type = "s"
    # o openpyxl so escreve inlineStr vazio se a celula tiver estilo
    wb.active["A1"].font = openpyxl.styles.Font(bold=True)
    assert any("inlineStr sem <is>" in a for a in ooxml.problemas(_salvar(wb)))


def test_verificador_acusa_cabecalho_divergente_e_tabela_sobre_celula_mesclada():
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.append(["ID", "Nome"])
    ws.append([1, "a"])
    ws.append([2, "b"])
    tabela = Table(displayName="T", ref="A1:B3")
    ws.add_table(tabela)
    ws.merge_cells("A3:B3")
    conteudo = _salvar(wb)
    # adultera o nome da coluna no XML da tabela, sem mexer na celula
    from io import BytesIO

    entrada, saida = zipfile.ZipFile(BytesIO(conteudo)), BytesIO()
    with zipfile.ZipFile(saida, "w") as z:
        for item in entrada.infolist():
            dados = entrada.read(item.filename)
            if item.filename.startswith("xl/tables/"):
                dados = dados.replace(b'name="Nome"', b'name="Outro"')
            z.writestr(item, dados)
    achados = ooxml.problemas(saida.getvalue())
    assert any("difere de tableColumn" in a for a in achados)
    assert any("cruza celulas mescladas" in a for a in achados)


# ---------- validador oficial da Microsoft (opt-in) ----------

_SDK_URL = "https://www.nuget.org/api/v2/package/DocumentFormat.OpenXml/2.20.0"
_SDK_DLL = "lib/net46/DocumentFormat.OpenXml.dll"

# Falso positivo documentado: o SDK exige ordem fixa nos filhos de <font>, mas o
# schema publicado define CT_Font como `choice maxOccurs=unbounded` (qualquer
# ordem) -- conferido em datypic.com/sc/ooxml/t-ssml_CT_Font.html. Aparece ate
# na fonte padrao que o openpyxl escreve em TODO arquivo, que o Excel abre sem
# reclamar. Qualquer OUTRO erro do SDK falha o teste.
_FALSO_POSITIVO_FONTE = ("/x:styleSheet[1]/x:fonts[1]/x:font[", "unexpected child element")


def _dll_do_sdk() -> Path:
    cache = Path(os.environ.get("LOCALAPPDATA", Path.home())) / "extrator-docs" / "openxml-sdk-2.20.0"
    dll = cache / _SDK_DLL
    if not dll.exists():
        cache.mkdir(parents=True, exist_ok=True)
        pacote = cache / "pacote.nupkg"
        urllib.request.urlretrieve(_SDK_URL, pacote)
        with zipfile.ZipFile(pacote) as z:
            z.extract(_SDK_DLL, cache)
    return dll


@pytest.mark.ooxml_sdk
@pytest.mark.parametrize("nome", list(CENARIOS))
def test_validador_oficial_da_microsoft_so_acusa_o_falso_positivo_conhecido(nome, tmp_path, request):
    if not request.config.getoption("--ooxml-sdk"):
        pytest.skip("validador da Microsoft e opt-in: pytest tests/test_ooxml.py --ooxml-sdk")
    if sys.platform != "win32":
        pytest.skip("usa o Windows PowerShell 5.1 (.NET Framework) pra carregar o SDK")

    arquivo = tmp_path / "saida.xlsx"
    arquivo.write_bytes(gerar_excel(CENARIOS[nome]).getvalue())
    script = Path(__file__).resolve().parent / "ooxml_sdk.ps1"
    r = subprocess.run(
        ["powershell.exe", "-NoProfile", "-ExecutionPolicy", "Bypass", "-File", str(script),
         "-Dll", str(_dll_do_sdk()), "-Arquivo", str(arquivo)],
        capture_output=True, text=True, timeout=120,
    )
    assert r.returncode == 0, r.stderr
    erros = [linha.split("\t") for linha in r.stdout.splitlines() if linha.strip()]
    reais = [
        e for e in erros
        if not (e[1].startswith(_FALSO_POSITIVO_FONTE[0]) and _FALSO_POSITIVO_FONTE[1] in e[2])
    ]
    assert reais == [], "\n".join("\t".join(e) for e in reais)
