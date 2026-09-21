# Extrator Inteligente de Documentos

Extrai automaticamente os dados de **notas fiscais (DANFE)** e **boletos bancários**
em PDF — emissor, destinatário, valores, datas, itens — e gera uma planilha Excel
pronta para conferência. Evita a digitação manual de documentos fiscais, com um
indicador de confiança em cada campo extraído para o usuário saber o que revisar.

![Tela do extrator com resultado extraído](docs/screenshot.png)

## Destaques

- **Funciona sem IA:** extratores por tipo de documento (padrão Strategy) leem
  campos por rótulo e reconstroem a tabela de itens da DANFE pela posição x/y das
  palavras. A IA (Claude) é um upgrade opcional, com fallback automático.
- **Na DANFE, a confiança é verificada, não chutada:** CNPJ do emissor e número da
  nota são conferidos contra a chave de acesso (com dígito verificador); totais são
  validados pela fórmula da nota.
- **Correção antes de exportar:** qualquer campo pode ser editado na tela; o Excel
  marca o que foi corrigido e guarda o valor original em comentário.
- **Excel multi-aba com estrutura pronta para lote:** Resumo, Itens, Campos adicionais e Avisos,
  com valores numéricos em R$, datas reais e chave de 44 dígitos preservada.
- **Segurança e privacidade:** proteção contra injeção de fórmula no Excel;
  nenhum dado pessoal real no repositório nem no histórico do Git.
- **150 testes:** unitários, regressão sobre um boleto e uma DANFE reais anonimizados
  e 25 testes de interface num navegador real (Playwright).
- **OCR opcional** (Tesseract) para PDFs escaneados.

## Stack

Python · FastAPI · pdfplumber · PyMuPDF · Tesseract · openpyxl · Claude API (opcional) · HTML/CSS/JS puro · Playwright

## Como funciona

O PDF passa por três etapas: **texto → campos → planilha**.

1. **Texto.** O `pdfplumber` lê o texto digital do PDF (e a posição x/y de cada
   palavra). Se o PDF for uma imagem escaneada, cai para [OCR](#ocr-opcional).
2. **Campos.** O extrator do tipo de documento estrutura o texto em JSON
   (`DocumentoExtraido`). Há dois modos, que preenchem o mesmo formato — por isso
   tela e Excel funcionam igual nos dois.
3. **Planilha.** A tela mostra o resultado, deixa corrigir e baixa o `.xlsx`.

### Modo básico (padrão, sem configuração)

- Escolhe o extrator pelo tipo do documento: **DANFE**, **boleto** ou **genérico**
  (pedido de compra, relatório).
- Captura os campos por **rótulo** conhecido (`Beneficiário`, `Sacado`, `Vencimento`…),
  não por regex solto no texto todo — isso evita roubar pedaços de outros números.
- Identifica CNPJ/CPF pelo formato e associa ao emissor/destinatário.
- Boleto: também linha digitável, nosso número e vencimento.
- DANFE: valida a chave de acesso e reconstrói a tabela de itens pela **posição**
  das palavras (regex sobre texto corrido perde a estrutura de colunas).
- Cada campo carrega uma confiança (`alta`, `media` ou `baixa`) no campo
  `confiancas` da resposta da API.

### Modo IA (opcional)

Com `ANTHROPIC_API_KEY` configurada, o Claude extrai os mesmos campos (inclusive a
lista de itens) via Structured Outputs. Se a chamada falhar
por qualquer motivo, o sistema volta sozinho ao modo básico e avisa na tela.
A resposta sempre traz `modo_extracao` (`"basico"` ou `"ia"`), exibido como badge.

### Confiança e correção na tela

Cada campo mostra **Alta**, **Média** ou **Baixa** (sempre com ícone e texto, nunca só
cor), há um resumo no topo ("X de Y campos com alta confiança") e um banner com os
avisos da extração. Campos de confiança baixa ou obrigatórios vazios já vêm abertos
para edição; **qualquer campo pode ser corrigido** antes de exportar. No modo IA não
há indicador de confiança por campo, mas a edição funciona igual.

### Excel

O botão **Baixar Excel** gera um `.xlsx` com 4 abas, sempre presentes:

| Aba | Conteúdo |
|---|---|
| **Resumo** | uma linha por documento (tipo, número, datas, emissor e destinatário com CNPJ/CPF em colunas próprias, valor total, confiança geral) |
| **Itens** | descrição, quantidade, valor unitário e total |
| **Campos adicionais** | chave de acesso, CFOP, nosso número, linha digitável… com a confiança de cada um |
| **Avisos** | o que a extração pediu para conferir |

Todas têm cabeçalho congelado, filtro, largura ajustada e uma coluna `ID` que liga as
abas. Confiança média/baixa e campos corrigidos ficam destacados, com comentário.
A estrutura (lista de documentos, `ID`) já comporta vários documentos no mesmo arquivo.

## Como rodar (Windows / PowerShell)

Requer Python 3.

```powershell
cd backend
python -m venv .venv
.venv\Scripts\python.exe -m pip install -r requirements.txt
.venv\Scripts\python.exe -m uvicorn app.main:app --reload
```

Acesse `http://localhost:8000` (interface) ou `http://localhost:8000/docs` (Swagger).

Essa forma funciona **sem ativar o venv**. Se preferir ativá-lo
(`.venv\Scripts\Activate.ps1`; o PowerShell pode bloquear scripts por política de
execução), depois basta `uvicorn app.main:app --reload`.

**Modo IA (opcional):** copie `backend\.env.example` para `backend\.env`, preencha
`ANTHROPIC_API_KEY=` com sua chave e reinicie o servidor. Sem a chave, tudo continua
funcionando em modo básico.

### Testes

```powershell
cd backend
.venv\Scripts\python.exe -m pytest tests -v
```

- **125 testes** rodam por padrão: unitários e regressão sobre o texto bruto de um
  boleto e de trechos de uma DANFE reais, com dados pessoais trocados por fictícios.
  Essas fixtures existem porque cenários escritos à mão não reproduziam os bugs que só
  apareciam no documento de verdade.
- O teste de OCR real só roda com o Tesseract e o idioma português instalados; sem
  eles é pulado (não falha). Os demais testes de OCR usam mocks.
- **25 testes de interface** (Chromium real via Playwright, com PDFs fictícios
  gerados na hora) são opt-in e usam dependências separadas:

```powershell
.venv\Scripts\python.exe -m pip install -r requirements-dev.txt
.venv\Scripts\python.exe -m playwright install chromium
.venv\Scripts\python.exe -m pytest tests\e2e --e2e -v
```

## OCR (opcional)

O sistema sempre tenta o **texto digital** primeiro. Só aciona OCR quando esse texto vem
vazio ou com menos de 20 caracteres (o suficiente para ignorar um carimbo solto): cada
página vira imagem (PyMuPDF) e o Tesseract lê o resultado. O texto do OCR segue pelo
**mesmo** extrator de sempre — não existe um caminho de extração separado. A resposta
traz `origem_texto` (`"digital"` ou `"ocr"`), exibido como badge.

- **Nunca quebra o app.** Sem o Tesseract, um PDF escaneado devolve um aviso claro
  ("OCR não está disponível"), não um erro.
- **PyMuPDF em vez de Poppler/ImageMagick** para renderizar as páginas: instala só com
  `pip`, então o OCR precisa de uma dependência externa (o Tesseract), não duas.
- **Confusão de caracteres** (`O`/`0`, `I`/`1`, `S`/`5`) é corrigida só nos campos que
  já são numéricos (CNPJ, CPF, valores). Aplicar a correção no texto todo destruiria
  palavras normais.
- **Limites:** a tabela de itens da DANFE precisa de texto digital (via OCR `itens`
  fica vazio), e a correção de confusão não cobre número do documento nem texto livre.

### Instalando o Tesseract (Windows)

O Tesseract é um programa do sistema, não um pacote Python.

```powershell
winget install --id UB-Mannheim.TesseractOCR -e
```

(ou baixe o instalador em [github.com/UB-Mannheim/tesseract](https://github.com/UB-Mannheim/tesseract/wiki)).
Duas pegadinhas reais:

- **Fora do PATH:** o instalador não adiciona o Tesseract ao PATH. O app já lida com
  isso — tenta `C:\Program Files\Tesseract-OCR\tesseract.exe` antes de desistir. Se você
  instalou em outro lugar, adicione a pasta ao PATH.
- **Sem o idioma português:** a instalação via `winget` só traz o inglês. Baixe
  [`por.traineddata`](https://github.com/tesseract-ocr/tessdata_fast/raw/main/por.traineddata)
  e copie para `C:\Program Files\Tesseract-OCR\tessdata\` (terminal como Administrador).
  Sem admin: coloque o arquivo em qualquer pasta e aponte `TESSDATA_PREFIX` para ela.

Confira com `& "C:\Program Files\Tesseract-OCR\tesseract.exe" --list-langs` — a saída
deve listar `por`.

## Decisões técnicas

- **Padrão Strategy por tipo de documento:** cada extrator tem seu próprio universo de
  rótulos. Adicionar um tipo novo é criar um arquivo e registrá-lo em uma lista, sem
  mexer nos existentes.
- **Extração por rótulo, não regex solto:** procura o rótulo conhecido e usa o texto
  logo depois dele, o que permite associar CNPJ/CPF ao nome da entidade mais próxima.
- **Confiança por evidência:** "achei o rótulo" não basta (já saiu emissor errado com
  rótulo achado). Na DANFE, o CNPJ e o número da nota são conferidos contra a chave de
  acesso e os totais contra a fórmula da nota; a soma dos itens é comparada com o total
  dos produtos.
- **Texto girado descartado na leitura:** o canhoto da DANFE é impresso girado e o
  `pdfplumber` o extrai como lixo antes do conteúdo real. Descartar na origem corrige
  todos os extratores de uma vez.
- **`client.messages.parse(output_format=...)`** (Structured Outputs) em vez de pedir
  JSON no prompt: a resposta já vem validada pelo schema Pydantic.
- **Campos numéricos flexíveis** (`float | string`): viram número quando a conversão é
  confiável e caem para texto quando não é, sem quebrar a extração.
- **Excel: só valor monetário vira número.** Chave de acesso (44 dígitos; o Excel só
  guarda 15 de precisão), CFOP e número do documento (zeros à esquerda) ficam texto.
  Todo texto vindo do PDF é gravado como texto, nunca como fórmula.
- **Frontend servido pelo próprio FastAPI**, em HTML/CSS/JS puro: um único processo em
  `localhost:8000`, sem CORS e sem CDN externo.

## Limitações conhecidas / próximos passos

- Só a DANFE tem tabela de itens no modo básico; boleto e documentos genéricos ficam
  com `itens` vazio (o modo IA cobre esses casos).
- OCR: ver os limites na seção [OCR](#ocr-opcional).
- DANFE: série, natureza da operação, data de saída e IE do emitente ainda não são
  extraídos como campos.
- Interface de lote (vários PDFs de uma vez) ainda não existe; hoje a tela envia um
  documento por vez, embora a API e o Excel já aceitem uma lista.
- As heurísticas foram desenhadas para o formato comercial brasileiro (datas
  `dd/mm/aaaa`, valores em `R$`) e validadas com um boleto e uma DANFE reais — outros
  layouts podem exigir ajustes.

## Autor

João Pedro Ferreira — [LinkedIn](https://www.linkedin.com/in/joaopedroferreira-2824d9324/)

## Status

MVP funcional de ponta a ponta: upload → extração (básica ou via IA) → tabela com
confiança → correção → download em Excel.
