````powershell
cd backend
.venv\Scripts\python.exe -m pytest tests -v
````

São 125 testes rodando por padrão, entre unitários e de regressão. Os de regressão usam
o texto bruto de um boleto e de trechos de uma DANFE reais, com os dados pessoais
trocados por fictícios. Eles existem porque os cenários que escrevi à mão não
reproduziam os bugs que apareciam no documento de verdade.

O teste de OCR real só roda se o Tesseract e o idioma português estiverem instalados.
Sem eles, ele é pulado em vez de falhar. Os outros testes de OCR usam mocks.

Há mais 25 testes de interface, que rodam num Chromium de verdade via Playwright com
PDFs fictícios gerados na hora. Eles são opcionais e têm dependências à parte:

````powershell
.venv\Scripts\python.exe -m pip install -r requirements-dev.txt
.venv\Scripts\python.exe -m playwright install chromium
.venv\Scripts\python.exe -m pytest tests\e2e --e2e -v
````

## OCR (opcional)

O sistema sempre tenta ler o texto digital primeiro. Se esse texto vier vazio ou com
menos de 20 caracteres, o que costuma indicar um PDF escaneado, cada página é
convertida em imagem pelo PyMuPDF e o Tesseract lê o resultado. Daí em diante o texto
segue pelo mesmo extrator de sempre, sem caminho separado. A resposta traz o campo
`origem_texto`, com o valor `digital` ou `ocr`, que aparece como badge na tela.

Alguns pontos do funcionamento:

- Se o Tesseract não estiver instalado, um PDF escaneado devolve o aviso "OCR não está
  disponível". O app não quebra.
- A renderização das páginas usa PyMuPDF em vez de Poppler ou ImageMagick porque ele
  instala só com pip. Assim o OCR depende de um único programa externo, o Tesseract.
- O OCR confunde caracteres parecidos, como O com 0, I com 1 e S com 5. A correção é
  aplicada só nos campos que já sabemos ser numéricos, como CNPJ, CPF e valores.
  Corrigir o texto inteiro destruiria palavras normais.
- Duas limitações: a tabela de itens da DANFE precisa de texto digital, então via OCR o
  campo `itens` fica vazio; e a correção de caracteres não cobre o número do documento
  nem texto livre.

### Instalando o Tesseract (Windows)

O Tesseract é um programa do sistema, não um pacote Python.

````powershell
winget install --id UB-Mannheim.TesseractOCR -e
````

Também dá para baixar o instalador em
[github.com/UB-Mannheim/tesseract](https://github.com/UB-Mannheim/tesseract/wiki).

Duas pegadinhas que encontrei testando:

1. O instalador não adiciona o Tesseract ao PATH. O app contorna isso tentando
   `C:\Program Files\Tesseract-OCR\tesseract.exe` antes de desistir. Se você instalou em
   outro lugar, adicione a pasta ao PATH.
2. A instalação via winget traz só o idioma inglês. Baixe o
   [`por.traineddata`](https://github.com/tesseract-ocr/tessdata_fast/raw/main/por.traineddata)
   e copie para `C:\Program Files\Tesseract-OCR\tessdata\`, usando um terminal como
   administrador. Sem permissão de administrador, salve o arquivo em qualquer pasta e
   aponte a variável `TESSDATA_PREFIX` para ela.

Para conferir, rode `& "C:\Program Files\Tesseract-OCR\tesseract.exe" --list-langs`. A
saída deve listar `por`.

## Decisões técnicas

**Um extrator por tipo de documento (padrão Strategy).** Cada tipo tem seu próprio
conjunto de rótulos. Para dar suporte a um documento novo basta criar um arquivo e
registrá-lo numa lista, sem tocar nos extratores que já funcionam.

**Busca por rótulo em vez de regex solto pelo texto.** O extrator procura o rótulo
conhecido e lê o que vem depois dele. Isso evita capturar pedaços de outros números e
permite associar um CNPJ ao nome que está mais próximo dele.

**Confiança medida por evidência.** Ter encontrado o rótulo não é garantia de nada: em
um dos testes o emissor saiu errado mesmo com o rótulo localizado. Na DANFE, o CNPJ e o
número da nota são conferidos contra a chave de acesso, os totais contra a fórmula da
nota, e a soma dos itens contra o valor total dos produtos.

**Texto girado descartado na leitura.** O canhoto da DANFE é impresso na vertical e o
pdfplumber o devolve como lixo antes do conteúdo real, o que fazia o emissor sair como
"FOLHA 1/". Descartar esses caracteres na leitura resolve para todos os extratores de
uma vez.

**Structured Outputs no modo IA.** Usando `client.messages.parse(output_format=...)`, a
resposta já chega validada pelo schema Pydantic, sem depender de pedir JSON no prompt.

**Campos numéricos flexíveis (`float | string`).** Viram número quando a conversão é
confiável e caem para texto quando não é, em vez de quebrar a extração inteira.

**No Excel, só valor monetário vira número.** A chave de acesso tem 44 dígitos e o Excel
guarda apenas 15 de precisão, o que a destruiria. CFOP e número do documento têm zeros à
esquerda. Os três ficam como texto. Todo texto vindo do PDF é gravado como texto, nunca
como fórmula.

**Frontend servido pelo próprio FastAPI**, em HTML, CSS e JavaScript puros. Um processo
só em `localhost:8000`, sem CORS e sem CDN externo.

## Limitações conhecidas e próximos passos

- Só a DANFE tem extração de itens no modo básico. Boleto e documentos genéricos ficam
  com `itens` vazio, e o modo IA cobre esses casos.
- As limitações do OCR estão na seção [OCR](#ocr-opcional).
- Na DANFE, série, natureza da operação, data de saída e inscrição estadual do emitente
  ainda não são extraídas.
- Não existe interface para processar vários PDFs de uma vez. A tela envia um documento
  por vez, embora a API e o Excel já aceitem uma lista.
- As heurísticas foram feitas para o formato comercial brasileiro, com datas em
  dd/mm/aaaa e valores em reais, e validadas com um boleto e uma DANFE reais. Outros
  layouts podem exigir ajustes.

## Autor

João Pedro Ferreira — [LinkedIn](https://www.linkedin.com/in/joaopedroferreira-2824d9324/)

## Status

MVP funcional: upload, extração básica ou via IA, tabela com indicador de confiança,
correção dos campos e download em Excel.
````
````

O que mudei, resumindo: tirei os travessões do meio das frases, desmontei os parênteses em frases inteiras, troquei os bullets com negrito e dois-pontos por parágrafos curtos nas decisões técnicas, e transformei as pegadinhas do Tesseract em lista numerada. Também escrevi em primeira pessoa nos pontos em que você conta o que aconteceu ("os cenários que escrevi à mão", "duas pegadinhas que encontrei testando"), que é o que mais diferencia um README escrito por alguém que fez o projeto.

Para aplicar, pede ao Claude Code para substituir essas seções pelo texto acima, ou edita direto no VS Code se preferir.
