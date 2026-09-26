"""Todo texto que chega ao usuario (tela, Excel) em portugues correto: com
acento e sem o separador "--" de codigo.

Bug real, visivel no video de demonstracao: os avisos da DANFE e as mensagens
de erro saiam como "nao fecham", "Nao foi possivel", "-- confira", e o aviso da
soma dos itens mostrava dinheiro em padrao americano ("R$ 215.03") enquanto a
tela e o Excel mostram "215,03".

Le o CODIGO (AST), nao so os casos que um teste consegue provocar: qualquer
mensagem nova escrita sem acento em um desses lugares faz este teste falhar.
Os lugares sao os que viram texto na tela ou no Excel:
- `avisos.append(...)` nos extratores (banner de avisos, aba Avisos);
- `HTTPException(detail=...)` (cartao de erro da tela);
- atribuicoes a `aviso` e o argumento `aviso_extra` (avisos montados no main.py);
- `comentario_avisos` (comentario de celula no Excel).
Comentarios, docstrings, logs e os ROTULOS que o extrator procura no PDF (que
existem de proposito sem acento, pra casar com o texto extraido) ficam de fora.
"""
import ast
import re
from pathlib import Path

import pytest

APP = Path(__file__).resolve().parents[1] / "app"

SEM_ACENTO = re.compile(
    r"\b(nao|voce|pagina|paginas|possivel|disponivel|indisponivel|legivel|extracao|extracoes|exportacao|"
    r"informacao|numero|numeros|digito|digitos|codigo|ja|ate|tambem|so|ha|valido|invalido|obrigatorio|"
    r"conteudo|necessario|basico|sera|unico|proximo|automatico|especifico|posicao|esta protegido|"
    r"esta vazio|esta disponivel|demonstracao|publica|atencao|excecao|situacao|operacao|emissao)\b",
    re.I,
)


def _texto(no) -> str:
    """Texto literal de uma string, f-string ou concatenacao implicita; partes
    dinamicas das f-strings viram "{}"."""
    if isinstance(no, ast.Constant) and isinstance(no.value, str):
        return no.value
    if isinstance(no, ast.JoinedStr):
        return "".join(v.value if isinstance(v, ast.Constant) else "{}" for v in no.values)
    if isinstance(no, ast.BinOp) and isinstance(no.op, ast.Add):
        return _texto(no.left) + _texto(no.right)
    return ""


def _mensagens_para_o_usuario():
    for arquivo in sorted(APP.rglob("*.py")):
        arvore = ast.parse(arquivo.read_text(encoding="utf-8"))
        for no in ast.walk(arvore):
            alvos = []
            if isinstance(no, ast.Call):
                f = no.func
                if isinstance(f, ast.Attribute) and f.attr in ("append", "insert") and "aviso" in ast.unparse(f.value):
                    alvos += no.args
                if isinstance(f, ast.Name) and f.id == "HTTPException":
                    alvos += [k.value for k in no.keywords if k.arg == "detail"]
                alvos += [k.value for k in no.keywords if k.arg == "aviso_extra"]
            elif isinstance(no, ast.Assign):
                nomes = {t.id for t in no.targets if isinstance(t, ast.Name)}
                if nomes & {"aviso", "comentario_avisos"}:
                    alvos.append(no.value)
            for alvo in alvos:
                texto = _texto(alvo)
                if texto.strip():
                    yield f"{arquivo.relative_to(APP)}:{alvo.lineno}", texto


MENSAGENS = list(_mensagens_para_o_usuario())


def test_o_teste_encontra_as_mensagens_que_deveria():
    """Sem isto o teste passaria vazio se a forma de montar as mensagens mudasse."""
    textos = " | ".join(t for _, t in MENSAGENS)
    for trecho in ("protegido por senha", "corrompido", "nenhuma página", "imagem escaneada", "não fecham",
                   "não bate", "número da nota", "dígito", "modo básico", "veja a aba Avisos", "não é um PDF",
                   "está vazio", "o limite é"):
        assert trecho in textos, f"mensagem com {trecho!r} nao foi encontrada pelo AST"
    assert len(MENSAGENS) >= 15


@pytest.mark.parametrize("onde, texto", MENSAGENS, ids=[onde for onde, _ in MENSAGENS])
def test_mensagem_com_acento_e_sem_separador_de_codigo(onde, texto):
    assert not SEM_ACENTO.search(texto), f"{onde}: palavra sem acento em {texto!r}"
    assert " -- " not in texto and not texto.strip().startswith("--"), f"{onde}: use travessao (—), nao '--': {texto!r}"


def test_dinheiro_nos_avisos_nao_usa_formatacao_americana():
    """"R$ {x:.2f}" gera "R$ 215.03"; a tela e o Excel mostram "215,03"."""
    for onde, texto in MENSAGENS:
        assert "R$ {}" not in texto or "formatar_valor_br" in (APP / onde.split(":")[0]).read_text(encoding="utf-8"), onde
    fonte = (APP / "extractors" / "danfe.py").read_text(encoding="utf-8")
    assert not re.search(r"R\$ \{[^}]*:\.2f\}", fonte), "dinheiro formatado com :.2f (ponto decimal) num aviso"
