"""Gera um PDF escaneado de teste (so imagem, sem camada de texto) para
validar o OCR manualmente na interface. Nao faz parte do app -- e so uma
ferramenta de teste manual.

Uso:
    python gerar_pdf_teste_ocr.py
"""
import pymupdf
from PIL import Image, ImageDraw, ImageFont

LINHAS = [
    "NOTA FISCAL N. 98765",
    "Data de Emissao: 10/08/2026",
    "Emitente: Comercio Exemplo Ltda",
    "CNPJ: 12.345.678/0001-90",
    "Destinatario: Cliente Exemplo S.A.",
    "Valor Total: R$ 2.500,00",
]

SAIDA_PNG = "pagina_teste_ocr.png"
SAIDA_PDF = "sample_escaneado_teste.pdf"


def main():
    img = Image.new("RGB", (1240, 500), "white")
    draw = ImageDraw.Draw(img)
    try:
        fonte = ImageFont.truetype("arial.ttf", 28)
    except Exception:
        fonte = ImageFont.load_default()

    y = 40
    for linha in LINHAS:
        draw.text((40, y), linha, fill="black", font=fonte)
        y += 60
    img.save(SAIDA_PNG)

    doc = pymupdf.open()
    pagina = doc.new_page(width=1240, height=500)
    pagina.insert_image(pagina.rect, filename=SAIDA_PNG)
    doc.save(SAIDA_PDF)
    doc.close()

    print(f"PDF gerado: {SAIDA_PDF}")
    print("Suba esse arquivo na interface (http://localhost:8000) e confira:")
    print('  - badge "Leitura: OCR"')
    print("  - tipo_documento: nota_fiscal, numero_documento: 98765, valor_total: 2500")


if __name__ == "__main__":
    main()
