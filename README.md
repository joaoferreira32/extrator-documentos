# Extrator Inteligente de Documentos

MVP de um extrator de documentos: você sobe um PDF (nota fiscal, pedido de
compra ou relatório), o sistema extrai o texto, estrutura os dados em JSON,
exibe numa tabela e permite baixar o resultado em Excel.

![Tela do extrator com resultado extraído](docs/screenshot.png)

## Stack

- **Backend:** Python 3, FastAPI
- **Extração de PDF:** [pdfplumber](https://github.com/jsvine/pdfplumber) (texto digital) +
  [PyMuPDF](https://pymupdf.readthedocs.io/) para renderizar páginas como
  imagem quando é preciso OCR
- **OCR (opcional):** [Tesseract](https://github.com/tesseract-ocr/tesseract)
  via `pytesseract`, pacote de idioma português — ver "Instalando o
  Tesseract" abaixo
- **Exportação para Excel:** pandas + openpyxl
- **Estruturação dos dados:** ver "Modo básico vs modo IA" abaixo
- **Frontend:** HTML/CSS/JS puro, sem framework, servido pelo próprio FastAPI

## Modo básico vs modo IA

A IA é um **upgrade opcional**, não um requisito para o sistema funcionar:

- **Modo básico** (padrão, sem nenhuma configuração): `pdfplumber` extrai o
  texto (e, quando digital, a posição x/y de cada palavra) do PDF, e um
  extrator dedicado por tipo de documento (padrão Strategy, em
  `app/extractors/`) captura os campos por **rótulo** conhecido
  (`Beneficiário:`, `Sacado:`, `Cedente:`, `Nosso Número:`, `Vencimento:`,
  etc.) em vez de regex solto pelo texto inteiro — isso é o que evita
  pegar pedaços soltos de outros números. Reconhece boleto bancário, DANFE
  (`DANFE`/`NF-e`/`CFOP`/chave de acesso) e pedido de compra por
  palavra-chave, identifica CNPJ/CPF por formato e associa ao
  emissor/destinatário quando aparecem perto do nome, e para boletos
  também captura linha digitável e vencimento. Pra DANFE, `itens` é
  reconstruído por **posição** (x0/x1/top/bottom de cada palavra, não
  regex sobre texto corrido) — a única forma confiável de recuperar uma
  tabela a partir de PDF sem estrutura de tabela nativa; boleto e
  documentos genéricos deixam `itens` vazio, já que não têm uma tabela de
  itens no mesmo sentido. Cada campo carrega uma confiança
  (`"alta"`/`"media"`/`"baixa"`) exposta no campo `confiancas` da
  resposta da API.
- **Modo IA** (com `ANTHROPIC_API_KEY` configurada): usa a API da Anthropic
  (Claude) via [Structured Outputs](https://docs.claude.com/) para extrair
  os mesmos campos com muito mais precisão, incluindo a lista de itens. Se a
  chamada à IA falhar por qualquer motivo, o sistema cai automaticamente
  para o modo básico em vez de mostrar um erro.

Os dois modos preenchem o mesmo formato de dados (`DocumentoExtraido`), que
é o que permite a tabela e a exportação Excel funcionarem igual
independente de como o documento foi processado. A resposta da API sempre
inclui um campo `modo_extracao` (`"basico"` ou `"ia"`), exibido como um
badge na interface.

## Origem do texto: digital vs OCR

Antes de qualquer extração de campos, o sistema decide de onde vem o
texto:

1. `pdfplumber` tenta extrair o texto digital do PDF (a camada de texto
   real, quando o PDF não é uma imagem escaneada).
2. Se esse texto vier vazio ou curto demais para ser confiável (menos de
   20 caracteres — o suficiente pra ignorar um carimbo solto), o sistema
   tenta **OCR** automaticamente: renderiza cada página como imagem
   (PyMuPDF) e roda o Tesseract em cima.
3. O texto que sobrar — digital ou vindo do OCR — segue para o **mesmo**
   `basic_extractor.py`/modo IA de sempre. Não existe um caminho de
   extração separado para texto de OCR.

A resposta da API inclui `origem_texto` (`"digital"` ou `"ocr"`), exibido
como um segundo badge ao lado do badge de modo, pra você saber como o
documento foi lido.

**OCR é opcional e nunca quebra o app.** Se o Tesseract não estiver
instalado no sistema, o PDF escaneado continua retornando um aviso claro
("OCR não está disponível") em vez de erro. Texto vindo de OCR tende a
confundir letras com números parecidos (`O`/`0`, `I`/`1`, `S`/`5`) — o
modo básico corrige isso nos campos que já sabe que devem ser numéricos
(CNPJ, CPF, valores) antes de usar o valor.

## Como rodar localmente

```bash
cd backend
python -m venv .venv
.venv\Scripts\activate      # Windows
pip install -r requirements.txt
uvicorn app.main:app --reload
```

Acesse `http://localhost:8000`.

### Habilitando o modo IA (opcional)

1. Copie `backend/.env.example` para `backend/.env`
2. Preencha `ANTHROPIC_API_KEY=` com sua chave da Anthropic
3. Reinicie o servidor

Sem a chave configurada, o sistema continua funcionando normalmente em modo
básico.

### Instalando o Tesseract (OCR, opcional — Windows)

O OCR precisa do **Tesseract instalado no sistema** (não é um pacote
Python — `pytesseract` só chama o binário). Sem ele, o app continua
funcionando normalmente; PDFs escaneados só não vão ter o texto extraído.

1. Instale via [winget](https://learn.microsoft.com/pt-br/windows/package-manager/winget/):

   ```powershell
   winget install --id UB-Mannheim.TesseractOCR -e
   ```

   ou baixe o instalador em
   [github.com/UB-Mannheim/tesseract](https://github.com/UB-Mannheim/tesseract/wiki).

2. **O instalador não adiciona o Tesseract ao PATH automaticamente**, e a
   instalação via `winget` (silenciosa) só traz o idioma inglês por
   padrão — o pacote de português (`por.traineddata`) não vem incluído.
   Duas pegadinhas reais encontradas testando isso:

   - **Binário fora do PATH:** o app já lida com isso sozinho — se
     `tesseract` não estiver no PATH, `pdf_extractor.py` tenta o caminho
     padrão do instalador (`C:\Program Files\Tesseract-OCR\tesseract.exe`)
     antes de desistir. Se você instalou em outro lugar, adicione a pasta
     ao PATH manualmente.
   - **Pacote de português ausente:** baixe `por.traineddata` em
     [github.com/tesseract-ocr/tessdata_fast](https://github.com/tesseract-ocr/tessdata_fast/raw/main/por.traineddata)
     e copie para `C:\Program Files\Tesseract-OCR\tessdata\` (precisa de
     um terminal **como Administrador** para copiar em `Program Files`).
     Alternativa sem precisar de admin: coloque o arquivo em qualquer
     pasta e defina a variável de ambiente `TESSDATA_PREFIX` apontando
     para essa pasta antes de rodar o servidor.

3. Confirme que funcionou:

   ```powershell
   & "C:\Program Files\Tesseract-OCR\tesseract.exe" --list-langs
   ```

   Deve listar `por` na saída.

### Rodando os testes

```bash
cd backend
pytest tests/ -v -s
```

Inclui um teste de regressão sobre o texto bruto de um boleto bancário real
(dados pessoais trocados por fictícios) — existe porque cenários sintéticos
escritos à mão não reproduziam bugs que só apareciam no documento de
verdade. Os testes de OCR usam mocks pra validar a orquestração
digital→OCR sem depender do Tesseract estar instalado; há um teste "de
verdade" com OCR real que roda quando o Tesseract e o idioma português
estão disponíveis, e é pulado (não falha) quando não estão.

## Decisões técnicas

- **`client.messages.parse(output_format=...)`** (Anthropic Structured
  Outputs) em vez de "pedir JSON no prompt e fazer parsing manual" — a
  resposta já vem validada contra o schema Pydantic, então não existe o
  risco clássico de a LLM devolver um JSON malformado ou incompleto.
- **Campos numéricos flexíveis** (`float | string`): `quantidade`,
  `valor_unitario` e `valor_total` viram `float` quando a conversão é
  confiável, e caem para `string` quando não é — isso é o que faz a
  planilha Excel sair com números realmente somáveis, sem quebrar a
  extração quando algum valor vem em formato inesperado.
- **Extração por rótulo, não regex solto pelo texto todo:** o modo básico
  procura primeiro um rótulo conhecido e usa o texto logo depois dele. É
  mais confiável do que casar um padrão em qualquer lugar do documento
  (que tende a "roubar" pedaços de outros números, como linha digitável ou
  CNPJ) e permite associar CNPJ/CPF ao nome da entidade mais próxima.
- **PDF escaneado (sem texto extraível):** detectado antes de tentar
  qualquer extração; o sistema tenta OCR (Tesseract) automaticamente e,
  se não estiver disponível ou não conseguir ler nada, retorna um aviso
  claro em vez de um resultado vazio sem explicação. OCR é uma
  dependência de sistema externa (não um pacote Python) — por isso tem
  que degradar sem quebrar o app.
- **PyMuPDF em vez de Poppler/ImageMagick para rasterizar páginas:**
  `pdf2image` (Poppler) e o modo de imagem do próprio pdfplumber
  (ImageMagick/Wand) exigiriam mais um binário de sistema além do
  Tesseract. PyMuPDF renderiza página→imagem só com `pip install`, então
  OCR precisa de exatamente uma dependência externa, não duas.
- **Correção de confusão de OCR só em campos já identificados como
  numéricos** (CNPJ, CPF, valores): a regex desses campos aceita `O/o`,
  `I/i`, `S/s` no lugar de dígitos e corrige antes de usar. Aplicar essa
  correção no texto inteiro destruiria palavras normais — por isso é
  escopada aos campos que a regex já confirmou que deveriam ser números.
- **Frontend servido pelo próprio FastAPI** (`StaticFiles`, um único
  processo em `localhost:8000`) em vez de dois servidores separados —
  evita configurar CORS e simplifica rodar o projeto localmente.
- **CSS puro com custom properties**, sem framework nem CDN externo:
  paleta e espaçamento centralizados em variáveis (`:root`), drag-and-drop
  nativo (`dragenter`/`drop`) sem biblioteca, e estados de carregamento/erro
  como elementos próprios da página em vez de `alert()`.

## Limitações conhecidas / próximos passos

- Modo básico extrai itens de tabela só pra DANFE (por posição x/y das
  palavras) — boleto e documentos genéricos ficam com `itens` vazio. O
  modo IA cobre também esses casos.
- A tabela de itens da DANFE exige texto de origem digital (posição
  confiável de palavra); em PDF escaneado (origem OCR) `itens` fica
  vazio nesse modo.
- DANFE: campos fiscais adicionais (série, natureza da operação, data de
  saída, IE do emitente, ICMS/frete como campos de documento) ainda não
  implementados — dependem de confirmar o rótulo exato num dump real
  antes de codar.
- OCR precisa do Tesseract instalado separadamente no sistema (ver
  "Instalando o Tesseract" acima) — sem ele, PDFs escaneados continuam
  retornando só o aviso, sem texto.
- A correção de confusão de caracteres do OCR cobre CNPJ, CPF e valores
  monetários; outros campos numéricos (ex: número do documento) não
  passam por essa correção ainda.
- Heurísticas de regex do modo básico foram desenhadas para o formato de
  documento comercial brasileiro comum (datas `dd/mm/aaaa`, valores em
  `R$`) — podem não pegar todos os formatos.

## Status

MVP funcional: upload → extração (básica ou via IA) → tabela → download em
Excel. Estrutura completa e decisões de arquitetura em `CLAUDE.md`.
