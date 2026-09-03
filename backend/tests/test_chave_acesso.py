"""Valida o digito verificador da chave de acesso da NFe (modulo 11,
pesos ciclicos 2-9) -- algoritmo publico definido pela SEFAZ, testado com
vetores sinteticos calculados a mao (nao depende de nenhuma DANFE real).
"""
from app.extractors.danfe import _digito_verificador, encontrar_chave_acesso

# 43 digitos "1": soma dos pesos (5 ciclos completos de 2..9 = 220, mais
# 2+3+4 dos 3 digitos finais = 9) = 229. 229 % 11 = 9 -> DV = 11 - 9 = 2.
CHAVE_43_UNS = "1" * 43
DV_ESPERADO = "2"


def test_digito_verificador_vetor_conhecido():
    assert _digito_verificador(CHAVE_43_UNS) == DV_ESPERADO


def test_encontrar_chave_acesso_valida():
    chave_valida = CHAVE_43_UNS + DV_ESPERADO
    texto = f"Chave de Acesso\n{chave_valida}\nProtocolo de Autorizacao"
    chave, confianca = encontrar_chave_acesso(texto)
    assert chave == chave_valida
    assert confianca == "alta"


def test_encontrar_chave_acesso_com_dv_invalido():
    chave_invalida = CHAVE_43_UNS + "9"  # DV errado de proposito
    texto = f"Chave de Acesso\n{chave_invalida}"
    chave, confianca = encontrar_chave_acesso(texto)
    assert chave == chave_invalida
    assert confianca == "baixa"


def test_encontrar_chave_acesso_tolera_espacos_entre_grupos():
    chave_valida = CHAVE_43_UNS + DV_ESPERADO
    # Formato comum de impressao: grupos de 4 digitos separados por espaco.
    agrupada = " ".join(chave_valida[i : i + 4] for i in range(0, 44, 4))
    chave, confianca = encontrar_chave_acesso(f"Chave de Acesso {agrupada}")
    assert chave == chave_valida
    assert confianca == "alta"


def test_sem_chave_de_acesso_no_texto():
    chave, confianca = encontrar_chave_acesso("Um texto qualquer sem nenhuma chave.")
    assert chave is None
    assert confianca == ""
