"""Tela e Excel formatam os mesmos valores do mesmo jeito (a parte que da pra
conferir sem navegador; o resto esta em tests/e2e/test_interface_consistencia.py).

A tela (frontend/script.js) e o Excel (app/excel_exporter.py) decidem, cada um
por conta propria, quais campos adicionais sao dinheiro. Se as listas divergirem,
um campo aparece "R$ 1.000,00" num lugar e "1000,00" no outro.
"""
import re
from pathlib import Path

from app.excel_exporter import CAMPOS_MONETARIOS

SCRIPT = Path(__file__).resolve().parents[2] / "frontend" / "script.js"


def test_tela_e_excel_tratam_os_mesmos_campos_adicionais_como_dinheiro():
    fonte = SCRIPT.read_text(encoding="utf-8")
    bloco = re.search(r"const CAMPOS_ADICIONAIS_MONETARIOS = new Set\(\[(.*?)\]\);", fonte, re.S)
    assert bloco, "CAMPOS_ADICIONAIS_MONETARIOS nao encontrado em script.js"
    na_tela = set(re.findall(r'"([^"]+)"', bloco.group(1)))
    assert na_tela == CAMPOS_MONETARIOS
