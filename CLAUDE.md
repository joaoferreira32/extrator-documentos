# extrator-docs

Extrator inteligente de documentos: usuário sobe um PDF (nota fiscal, pedido de
compra ou relatório), o sistema extrai o texto, estrutura os dados em JSON,
exibe numa tabela e permite baixar em Excel. Projeto de portfólio (estudante
de Engenharia da Computação).

## Stack

- Backend: Python 3 + FastAPI
- Extração de PDF: pdfplumber (texto digital) + PyMuPDF (rasterizar página
  para imagem) + Tesseract via `pytesseract` (OCR, dependência de sistema
  externa e opcional)
- Excel: pandas + openpyxl
- LLM (opcional): Anthropic Claude (`anthropic` SDK), modelo `claude-sonnet-5`,
  via `client.messages.parse(output_format=...)` (Structured Outputs)
- Frontend: HTML/CSS/JS puro, sem framework, servido como estático pelo FastAPI

## Estrutura

```
backend/app/
  main.py            # FastAPI app, rotas, serve frontend/ como estático
  config.py          # carrega ANTHROPIC_API_KEY (opcional) do .env
  schemas.py         # Pydantic: DocumentoExtraido, ExtractionResult, etc.
  pdf_extractor.py    # texto: pdfplumber (digital) -> OCR (Tesseract) se vazio/curto
  basic_extractor.py  # extração heurística por regex (modo "basico")
  llm_extractor.py    # extração via Claude (modo "ia", opcional)
  excel_exporter.py   # ExtractionResult -> .xlsx
backend/tests/
  fixtures/boleto_real_anonimizado.txt  # texto bruto de boleto real (dados trocados)
  test_basic_extractor.py               # roda o extrator sobre a fixture acima
  test_pdf_extractor.py                 # orquestracao digital->OCR (mocks) + 1 teste OCR real (skip se indisponivel)
frontend/
  index.html, style.css, script.js
```

## Decisão de arquitetura central: modo básico vs modo IA

A IA é um **upgrade opcional**, não um requisito de funcionamento:

- **Modo básico** (`modo_extracao: "basico"`): sempre disponível, não depende
  de chave de API. `pdfplumber` extrai o texto e `basic_extractor.py` busca
  por **rótulos conhecidos** (`Beneficiário:`, `Cedente:`, `Sacado:`,
  `Nosso Número:`, `Vencimento:`, etc.) e usa o texto logo depois deles —
  mais confiável do que regex solto pelo texto inteiro, que tende a capturar
  pedaços soltos de outros números (esse foi um bug real: `numero_documento`
  virando "02", pedaço de outro campo). Reconhece boleto bancário, nota
  fiscal e pedido de compra por palavra-chave; identifica CNPJ/CPF por
  formato e associa ao emissor/destinatário mais próximo; para boleto,
  também captura linha digitável, "Nosso Número", parcela, agência/código
  do beneficiário e os três valores (documento, desconto, valor a pagar)
  em `campos_adicionais`; `valor_total` prioriza "Valor a Pagar" quando
  existe. `itens` normalmente fica vazio nesse modo — reconstruir tabelas
  de itens por regex não é confiável. Essa é uma limitação conhecida do
  MVP, não um bug.
  - `data_emissao` e `data_vencimento` são campos separados — nunca se
    misturam porque cada um busca por um rótulo diferente
    (`ROTULOS_DATA_EMISSAO` vs `ROTULOS_VENCIMENTO`) e ambos exigem que o
    candidato tenha uma data reconhecível (`_parece_data`) antes de aceitar.
  - `numero_documento` (rótulos `Nr do documento`/`Número do documento`)
    e "Nosso Número" (`campos_adicionais`) são campos DIFERENTES de
    propósito — o segundo é um identificador bancário interno, não o
    número do documento. Listas de rótulos separadas
    (`ROTULOS_NUMERO_DOCUMENTO` vs `ROTULOS_NOSSO_NUMERO`); mistura-los foi
    um bug real. Nenhuma das duas listas inclui o rótulo genérico solto
    "Número"/"Numero" — ele bate como substring dentro de "**Nosso**
    Número" (e de qualquer outro campo que mencione a palavra), o que já
    contaminou os dois campos; o fallback por regex (`NUM_DOCUMENTO_RE`)
    também teve a variante "Número" removida pelo mesmo motivo, mantendo só
    as formas abreviadas N./Nº/N°.
  - Regra de negócio importante: `numero_documento`/`Nosso Número` só
    aceitam valores curtos/poucos dígitos quando vêm de um rótulo
    conhecido (confiável); sem rótulo, o fallback por regex solto exige no
    mínimo 4 dígitos — é o que evita recapturar o bug do "02".
  - `_localizar_rotulo` busca por **prioridade do rótulo** (ordem da lista
    `ROTULOS_*`), não por posição no documento: primeiro procura o rótulo
    mais específico no documento inteiro; só cai para um rótulo mais
    genérico se o específico não aparecer em lugar nenhum. Existe porque
    um rótulo genérico que aparece mais cedo na página (ex: "Número do
    Banco") vencia um rótulo específico e confiável que aparecia mais
    tarde.
  - Candidatos a `emissor`/`destinatario` passam por `_parece_nome`, que
    rejeita (a) valores que são só dígitos/pontuação — carimbos de
    data/hora, números soltos — e (b) marcadores de cabeçalho/rodapé de
    boleto (`MARCADORES_CABECALHO_RODAPE`: "Pág", "Página", "Recibo do
    Sacado", ...). Sem (a), um "Sacado" seguido de carimbo de data/hora
    virava destinatário; sem (b), a seção "Recibo do Sacado" (que contém a
    palavra "Sacado") fazia o rótulo bater ali em vez de no campo de
    verdade, capturando o rodapé "Pág: 1 de 1" como se fosse o nome.
  - "Nosso Número" às vezes vem como "02 / 10200000001-9" (prefixo de
    carteira / número real) — `_extrair_valor_numerico_por_rotulo`
    (compartilhada entre `numero_documento` e "Nosso Número") reconhece
    esse formato e usa a parte depois da barra.
  - Rótulos combinados colados sem espaço (ex: "Sacado/PagadorFulano de
    Tal") são tratados por `_limpar_prefixo_rotulos`, mas SÓ quando há
    uma barra `/` logo após o primeiro rótulo — sem essa barra como sinal,
    a limpeza fica desligada de propósito: já foi bug real remover
    "Fornecedor" do começo do nome de uma empresa só porque "Fornecedor"
    também é um rótulo válido de emissor.
- **Modo IA** (`modo_extracao: "ia"`): usado quando `ANTHROPIC_API_KEY` está
  configurada. Tenta primeiro; se a chamada falhar por qualquer motivo (rede,
  rate limit, resposta inválida), cai para o modo básico automaticamente e
  preenche `aviso` explicando o fallback.
- **Nunca é erro não ter `ANTHROPIC_API_KEY` configurada.** O endpoint
  `/extract-document` sempre responde 200 com algum resultado — a distinção
  de modo/qualidade vai no campo `modo_extracao`, não em status de erro.
- PDF escaneado (sem texto extraível) também não é erro: aciona OCR
  automaticamente e, se mesmo assim não conseguir texto, retorna aviso
  específico.

`DocumentoExtraido` é o mesmo contrato de dados nos dois modos — é o que
permite tabela e exportação Excel funcionarem igual independente de como o
documento foi extraído.

## OCR (fallback digital -> imagem)

`pdf_extractor.extrair_texto()` sempre tenta o texto digital do
pdfplumber primeiro. Só aciona OCR quando esse texto vem vazio ou tem
menos de `TAMANHO_MINIMO_TEXTO_DIGITAL` (20) caracteres — PDFs escaneados
às vezes têm um carimbo ou timestamp real misturado na imagem, e 20
caracteres é pouco pra confiar como "documento com texto de verdade".
Texto de OCR passa pelo **mesmo** `basic_extractor`/modo IA de sempre —
não existe lógica de extração de campos duplicada para OCR.

- **PyMuPDF em vez de `pdf2image`/Wand para rasterizar página→imagem:**
  ambos exigiriam mais um binário de sistema (Poppler ou ImageMagick)
  além do próprio Tesseract. PyMuPDF instala só com `pip`, então OCR
  precisa de exatamente uma dependência externa (o Tesseract), não duas.
  Import é `import pymupdf` (não `import fitz` — nome antigo, gera
  warning de depreciação).
- **OCR nunca pode quebrar o app.** `_tentar_ocr()` devolve
  `("", False)` — nunca levanta exceção — em qualquer cenário de falha:
  bibliotecas não instaladas (`ImportError` no topo do módulo),
  `pytesseract.TesseractNotFoundError` (Tesseract não instalado), ou
  qualquer outro erro (ex: pacote de idioma `por` ausente — vira
  `TesseractError` genérico, capturado pelo `except Exception` catch-all).
  `TextoExtraido.ocr_disponivel` distingue "OCR não disponível" de "OCR
  tentou e não achou nada" — `main.py` usa isso pra escolher a mensagem
  de aviso certa.
- **Instalador do Tesseract para Windows não adiciona o binário ao
  PATH.** Sem isso, `pytesseract` levanta `TesseractNotFoundError` mesmo
  com o Tesseract instalado, confundindo "não instalado" com "instalado
  mas não configurado". `pdf_extractor.py` detecta isso
  (`shutil.which("tesseract") is None` em `sys.platform == "win32"`) e
  tenta o caminho padrão do instalador
  (`C:\Program Files\Tesseract-OCR\tesseract.exe`) antes de desistir.
- **Instalação via `winget` (silenciosa) não traz o pacote de idioma
  português** (`por.traineddata`) — só o inglês vem por padrão. Testado
  de verdade: sem esse arquivo em `...\Tesseract-OCR\tessdata\`,
  `pytesseract` levanta `TesseractError` ("Failed loading language
  'por'"), corretamente capturado como `ocr_disponivel=False`. Ver
  instruções de instalação no README.
- **Confusão de caracteres do OCR** (`O`/`0`, `I`/`1`, `S`/`5`) é
  corrigida só em campos que a regex **já** identificou como numéricos
  (CNPJ, CPF, valores) — `_DIG = "0-9OoIiSs"` substitui `\d` nessas
  regex especificamente, e `_corrigir_confusao_ocr()` normaliza o trecho
  capturado antes de usar. Aplicar a correção no texto inteiro
  destruiria palavras normais que por acaso tenham essas letras.

## Convenções

- Chave de API sempre via variável de ambiente (`ANTHROPIC_API_KEY` em
  `backend/.env`), nunca hardcoded. `.env` está no `.gitignore`;
  `.env.example` documenta a variável.
- Campos numéricos (`quantidade`, `valor_unitario`, `valor_total`) são
  `float` quando a conversão é possível, com fallback para `string` caso
  contrário — isso é o que faz o Excel sair com números somáveis.

## Como rodar

```bash
cd backend
pip install -r requirements.txt
uvicorn app.main:app --reload
```

Abrir `http://localhost:8000` (frontend) ou `http://localhost:8000/docs`
(Swagger UI para testar os endpoints diretamente).

**Se editar o código e o comportamento não mudar** (uma rota nova não
aparece no `/docs`, uma correção parece "não ter feito efeito"), quase
certo que sobrou um processo antigo do uvicorn preso na porta 8000. Já
aconteceu duas vezes neste projeto. No Windows:
`netstat -ano | findstr :8000` para achar o PID, `taskkill /F /PID <PID>`
para matar, e reiniciar o servidor antes de testar de novo.

## Testes

```bash
cd backend
pytest tests/ -v -s
```

`tests/test_basic_extractor.py` roda `basic_extractor.extrair()` sobre
`tests/fixtures/boleto_real_anonimizado.txt` — texto bruto de um boleto
real (nome/CPF/CNPJ/valores trocados por fictícios, mas rótulos e
estrutura de linha exatamente como o pdfplumber extraiu). Esse fixture
existe porque cenários sintéticos escritos à mão não reproduziam bugs
reais: o mesmo rótulo (ex: "Sacado") aparece várias vezes no boleto em
contextos diferentes (título de seção, rodapé, tabela-resumo, campo de
verdade), e só um teste sobre o texto real pega isso. `_localizar_rotulo`
loga (nível DEBUG, logger `app.basic_extractor`) qual rótulo/linha foi
aceito ou rejeitado para cada campo — o teste imprime esse log com
`caplog`, então dá pra ver exatamente onde a extração está acertando ou
errando sem precisar adivinhar.

`tests/test_pdf_extractor.py` testa a orquestração digital->OCR com
mocks (`monkeypatch` em `_tentar_ocr`/`_OCR_IMPORTADO`), então passa
independente de o Tesseract estar instalado na máquina que roda os
testes. O último teste do arquivo (`test_ocr_real_extrai_texto...`) usa
OCR de verdade e só roda quando `_ocr_real_disponivel()` confirma que o
Tesseract + idioma `por` estão realmente disponíveis — pulado (não
falha) caso contrário. Os testes de confusão de OCR
(`test_corrige_confusao_ocr_*` em `test_basic_extractor.py`) são puros
(sem dependência externa).

## Estado atual / próximas etapas

MVP funcional de ponta a ponta (testado no navegador via Playwright):
upload de PDF -> extração (modo básico sempre; modo IA quando
`ANTHROPIC_API_KEY` está configurada, com fallback automático para básico
em caso de falha) -> tabela na tela -> download em Excel. Erros comuns
(arquivo não-PDF, PDF corrompido, arquivo muito grande) retornam mensagens
amigáveis em vez de 500.

`POST /debug/extract-text` devolve o texto bruto do pdfplumber (sem
nenhuma extração de campos em cima) — usado para inspecionar como os
rótulos aparecem de verdade num PDF real antes de ajustar regex/
heurísticas do modo básico. Foi assim que os bugs de boleto real (número
com barra, rótulo colado, vencimento não capturado, valores sem "R$") do
histórico do projeto foram diagnosticados e corrigidos.

Não implementado ainda / possíveis próximos passos:

- Modo básico não reconstrói tabela de itens (só o modo IA faz isso hoje).
- Correção de confusão de OCR cobre só CNPJ/CPF/valores — não
  numero_documento/Nosso Número nem texto livre (emissor/destinatario).
- Só há um fixture de teste automatizado com texto real (um boleto
  anonimizado). Vale adicionar mais fixtures reais (nota fiscal, pedido de
  compra) conforme aparecerem casos.

Plano original com todas as etapas em
`C:\Users\User\.claude\plans\jazzy-foraging-harbor.md`.
