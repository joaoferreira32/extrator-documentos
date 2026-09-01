"""Exportacao do resultado da extracao para Excel (.xlsx) via pandas + openpyxl.

Os campos numericos de DocumentoExtraido (quantidade, valor_unitario,
valor_total) chegam aqui como float quando a conversao deu certo -- o
pandas escreve esses valores como numero no Excel, nao como texto, entao
dao pra somar direto na planilha.
"""
from io import BytesIO

import pandas as pd

from app.schemas import ExtractionResult

ROTULOS_METADADOS = {
    "tipo_documento": "Tipo de documento",
    "numero_documento": "Numero do documento",
    "data_emissao": "Data de emissao",
    "emissor": "Emissor",
    "destinatario": "Destinatario",
    "valor_total": "Valor total",
}

COLUNAS_ITENS = ["Descricao", "Quantidade", "Valor unitario", "Valor total"]


def gerar_excel(resultado: ExtractionResult) -> BytesIO:
    documento = resultado.documento

    linhas_resumo = [
        {"Campo": rotulo, "Valor": getattr(documento, chave)}
        for chave, rotulo in ROTULOS_METADADOS.items()
    ]
    linhas_resumo += [
        {"Campo": extra.campo, "Valor": extra.valor}
        for extra in documento.campos_adicionais
    ]
    df_resumo = pd.DataFrame(linhas_resumo)

    if documento.itens:
        df_itens = pd.DataFrame(
            [
                {
                    "Descricao": item.descricao,
                    "Quantidade": item.quantidade,
                    "Valor unitario": item.valor_unitario,
                    "Valor total": item.valor_total,
                }
                for item in documento.itens
            ]
        )
    else:
        df_itens = pd.DataFrame(columns=COLUNAS_ITENS)

    buffer = BytesIO()
    with pd.ExcelWriter(buffer, engine="openpyxl") as writer:
        df_resumo.to_excel(writer, sheet_name="Resumo", index=False)
        df_itens.to_excel(writer, sheet_name="Itens", index=False)
    buffer.seek(0)
    return buffer
