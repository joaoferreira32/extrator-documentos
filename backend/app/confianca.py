"""Regras da confianca "geral" de um documento -- a MESMA conta que a tela faz
(`script.js`: estadoDoCampo / contarEstados). Existe em Python porque o Excel
mostra "X de Y alta" e o backend nao pode depender do numero calculado pelo
navegador. Um teste e2e compara o numero da tela com o do Excel, entao as duas
implementacoes nao divergem sem falhar.

Regras (identicas as da tela):
- Contam so campos COM valor: os 6 do documento (exceto o tipo, que nao tem
  confianca) + cada campo adicional + `itens` (1, quando ha itens).
- Um campo CORRIGIDO pelo usuario fica de fora do "X de Y" (nao vira "alta
  confianca" so porque alguem mexeu nele) e conta a parte.
- Campo vazio nao conta em lugar nenhum aqui.
- Sem `confiancas` (modo IA) nao ha resumo.
"""
from dataclasses import dataclass
from typing import Iterable, Optional

from app.schemas import ExtractionResult

NIVEIS = ("alta", "media", "baixa")
CAMPOS_DO_DOCUMENTO = (
    "numero_documento",
    "data_emissao",
    "data_vencimento",
    "emissor",
    "destinatario",
    "valor_total",
)


def vazio(valor) -> bool:
    return valor is None or str(valor).strip() == ""


def estado_do_campo(valor, chave: str, confiancas: dict, corrigidos) -> Optional[str]:
    """"corrigido" | "alta" | "media" | "baixa" | None (vazio ou sem confianca).
    Mesma ordem da tela: vazio -> corrigido -> confianca do backend."""
    if vazio(valor):
        return None
    if chave in corrigidos:
        return "corrigido"
    nivel = confiancas.get(chave)
    return nivel if nivel in NIVEIS else None


@dataclass(frozen=True)
class ResumoConfianca:
    alta: int
    revisar: int  # media + baixa
    corrigidos: int

    @property
    def total(self) -> int:
        return self.alta + self.revisar

    @property
    def texto(self) -> str:
        return f"{self.alta} de {self.total} alta"


def resumo_confianca(resultado: ExtractionResult, corrigidos: Iterable[str] = ()) -> Optional[ResumoConfianca]:
    """None no modo IA (`confiancas` vazio): nao ha confianca por campo."""
    if not resultado.confiancas:
        return None

    corrigidos = set(corrigidos)
    documento = resultado.documento
    alta = revisar = corrigidos_n = 0

    pares = [(chave, getattr(documento, chave)) for chave in CAMPOS_DO_DOCUMENTO]
    pares += [(extra.campo, extra.valor) for extra in documento.campos_adicionais]
    for chave, valor in pares:
        estado = estado_do_campo(valor, chave, resultado.confiancas, corrigidos)
        if estado == "alta":
            alta += 1
        elif estado in ("media", "baixa"):
            revisar += 1
        elif estado == "corrigido":
            corrigidos_n += 1

    nivel_itens = resultado.confiancas.get("itens")
    if documento.itens and nivel_itens in NIVEIS:
        if nivel_itens == "alta":
            alta += 1
        else:
            revisar += 1

    return ResumoConfianca(alta=alta, revisar=revisar, corrigidos=corrigidos_n)
