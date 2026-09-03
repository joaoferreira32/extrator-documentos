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
  main.py              # FastAPI app, rotas, serve frontend/ como estático
  config.py            # carrega ANTHROPIC_API_KEY (opcional) do .env
  schemas.py           # Pydantic: DocumentoExtraido, ExtractionResult, etc.
  pdf_extractor.py     # texto: pdfplumber (digital) -> OCR (Tesseract) se vazio/curto;
                        #   tambem captura palavras posicionadas (x/y) quando digital
  basic_extractor.py   # camada fina: monta ContextoExtracao e delega pro extractors/
  extractors/
    base.py            # ExtratorDocumento (interface), ContextoExtracao, Palavra, ResultadoExtracao
    __init__.py         # registro dos extratores + selecionar_extrator()
    comum.py            # helpers compartilhados (regex CNPJ/CPF/valor/data, busca por rotulo)
    boleto.py            # BoletoExtractor
    danfe.py              # DanfeExtractor (deteccao + chave de acesso; campos fiscais e
                          #   tabela de itens ainda pendentes -- ver "Estado atual")
    generico.py           # GenericExtractor (fallback: pedido_compra, relatorio, desconhecido)
  llm_extractor.py     # extração via Claude (modo "ia", opcional)
  excel_exporter.py    # ExtractionResult -> .xlsx
backend/tests/
  fixtures/boleto_real_anonimizado.txt  # texto bruto de boleto real (dados trocados)
  test_basic_extractor.py               # compat da funcao publica extrair()
  test_extractors_boleto.py             # BoletoExtractor isolado, com a fixture real acima
  test_detector.py                      # selecionar_extrator() roteia pro extrator certo
  test_chave_acesso.py                  # digito verificador da chave de acesso (algoritmo publico)
  test_pdf_extractor.py                 # orquestracao digital->OCR (mocks) + 1 teste OCR real (skip se indisponivel)
frontend/
  index.html, style.css, script.js
```

## Decisão de arquitetura central: modo básico vs modo IA

A IA é um **upgrade opcional**, não um requisito de funcionamento:

- **Modo básico** (`modo_extracao: "basico"`): sempre disponível, não depende
  de chave de API. Delega pro extrator certo por tipo de documento (ver
  "Arquitetura de extratores" abaixo) em vez de uma heurística monolítica.
  `itens` só é populado pela DANFE (via tabela por coordenadas) e pelo modo
  IA — boleto e documentos genéricos deixam vazio, porque reconstruir
  tabela por regex não é confiável. Essa é uma limitação conhecida do MVP,
  não um bug.
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

## Arquitetura de extratores (padrão Strategy)

O modo básico era uma heurística monolítica (`basic_extractor.py` com um
único conjunto global de rótulos); virou um padrão Strategy em
`app/extractors/`: `ExtratorDocumento` é a interface (`base.py`),
`selecionar_extrator()` (`__init__.py`) escolhe o extrator com maior
`pontuacao_deteccao()` entre os registrados, e cada tipo de documento tem
seu próprio arquivo com seu próprio universo de rótulos —
`boleto.py`, `danfe.py`, `generico.py`. `basic_extractor.py` ficou uma
camada fina de compatibilidade (monta o `ContextoExtracao` e delega).

**Adicionar um tipo novo:** criar `extractors/tipo_novo.py` implementando
a interface + adicionar a classe em `_EXTRATORES` (`extractors/__init__.py`).
Não precisa tocar em nenhum extrator existente nem em
`selecionar_extrator`. Não é descoberta automática de plugins (seria
over-engineering pra este projeto) — é o menor acréscimo que ainda cumpre
"não mexer em código existente".

**`comum.py`** tem os helpers genéricos que qualquer extrator pode reusar
(regex de CNPJ/CPF/valor/data, `localizar_rotulo`, `extrair_entidade`,
etc.) — funções que antes liam uma lista global de rótulos do módulo agora
recebem `todos_rotulos` como parâmetro, porque cada extrator tem seu
próprio universo. O histórico de bugs reais documentado abaixo (rótulo
genérico vencendo rótulo específico, cabeçalho/rodapé virando nome, etc.)
foi corrigido nesses helpers genéricos, então vale pra **qualquer**
extrator que os use, não só boleto.

**Confiança por campo** (`ResultadoExtracao.confiancas`, exposta em
`ExtractionResult.confiancas` na API — ver "Estado atual" sobre o que já
está ligado): `"alta"` quando veio de rótulo explícito, `"media"` quando
veio de heurística posicional (ex: tabela de itens por coordenada),
`"baixa"` quando veio de fallback sem rótulo. Fica em `ExtractionResult`,
não dentro de `DocumentoExtraido` — colocar confiança dentro de cada campo
exigiria que o schema usado pelo `output_format` do modo IA
(`client.messages.parse()`) também carregasse essa noção, que a LLM não
tem naturalmente. Chaves do dict: nome do atributo em `DocumentoExtraido`
(ex: `"numero_documento"`) ou, para itens de `campos_adicionais`, o texto
exato do campo (ex: `"Nosso Número"`, `"Chave de Acesso"`).

### Boleto (`boleto.py`) — comportamento herdado, sem mudança

Toda a lógica e o histórico de bugs abaixo já existiam antes da
refatoração; só mudou de arquivo (comportamento idêntico, testado pela
fixture real em `test_extractors_boleto.py`):

- `data_emissao` e `data_vencimento` são campos separados — nunca se
  misturam porque cada um busca por um rótulo diferente
  (`ROTULOS_DATA_EMISSAO` vs `ROTULOS_VENCIMENTO`) e ambos exigem que o
  candidato tenha uma data reconhecível (`comum.parece_data`) antes de
  aceitar.
- `numero_documento` (rótulos `Nr do documento`/`Número do documento`)
  e "Nosso Número" (`campos_adicionais`) são campos DIFERENTES de
  propósito — o segundo é um identificador bancário interno, não o
  número do documento. Listas de rótulos separadas; misturá-los foi um
  bug real. Nenhuma das duas listas inclui o rótulo genérico solto
  "Número"/"Numero" — ele bate como substring dentro de "**Nosso**
  Número" (e de qualquer outro campo que mencione a palavra), o que já
  contaminou os dois campos; o fallback por regex (`comum.NUM_GENERICO_RE`)
  também teve a variante "Número" removida pelo mesmo motivo, mantendo só
  as formas abreviadas N./Nº/N°.
- Regra de negócio importante: `numero_documento`/`Nosso Número` só
  aceitam valores curtos/poucos dígitos quando vêm de um rótulo
  conhecido (confiável, confiança `"alta"`); sem rótulo, o fallback por
  regex solto exige no mínimo 4 dígitos e vira confiança `"baixa"` — é o
  que evita recapturar o bug do "02".
- `comum.localizar_rotulo` busca por **prioridade do rótulo** (ordem da
  lista passada), não por posição no documento: primeiro procura o
  rótulo mais específico no documento inteiro; só cai para um rótulo
  mais genérico se o específico não aparecer em lugar nenhum. Existe
  porque um rótulo genérico que aparece mais cedo na página vencia um
  rótulo específico e confiável que aparecia mais tarde.
- Candidatos a `emissor`/`destinatario` passam por `comum.parece_nome`,
  que rejeita (a) valores que são só dígitos/pontuação — carimbos de
  data/hora, números soltos — e (b) marcadores de cabeçalho/rodapé
  (`comum.MARCADORES_CABECALHO_RODAPE`: "Pág", "Página", "Recibo do
  Sacado", ...). Sem (a), um "Sacado" seguido de carimbo de data/hora
  virava destinatário; sem (b), a seção "Recibo do Sacado" (que contém a
  palavra "Sacado") fazia o rótulo bater ali em vez de no campo de
  verdade, capturando o rodapé "Pág: 1 de 1" como se fosse o nome.
- "Nosso Número" às vezes vem como "02 / 10200000001-9" (prefixo de
  carteira / número real) — `boleto._extrair_valor_numerico_por_rotulo`
  reconhece esse formato e usa a parte depois da barra.
- Rótulos combinados colados sem espaço (ex: "Sacado/PagadorFulano de
  Tal") são tratados por `comum.limpar_prefixo_rotulos`, mas SÓ
  quando há uma barra `/` logo após o primeiro rótulo — sem essa barra
  como sinal, a limpeza fica desligada de propósito: já foi bug real
  remover "Fornecedor" do começo do nome de uma empresa só porque
  "Fornecedor" também é um rótulo válido de emissor.

### DANFE (`danfe.py`) — estado parcial, ver "Estado atual"

Só o que **não depende de suposição de layout** está implementado:
detecção por marcador (`MARCADORES_DETECCAO`: `danfe`, `documento
auxiliar da nota fiscal eletr[ônica]`, `chave de acesso`, `protocolo de
autorização de uso`, mais `nf-e`/`cfop` que já existiam) e validação da
chave de acesso (44 dígitos + dígito verificador módulo 11, pesos
cíclicos 2–9 — algoritmo público da SEFAZ, `test_chave_acesso.py` tem
vetores sintéticos calculados a mão). Campos fiscais específicos (série,
natureza da operação, CFOP predominante, ICMS, frete) e a tabela de itens
por coordenadas **ainda não foram implementados** — dependem de um dump
real via `/debug/extract-words` pra não repetir o erro de supor layout
sem material real (foi assim que o boleto quebrou repetidamente antes da
fixture real). Enquanto isso, `DanfeExtractor` reusa `GenericExtractor`
para os campos universais (emissor/destinatário/número/data/valor), então
uma DANFE não regride em relação ao que já tinha antes desta refatoração
— só ganha a chave de acesso validada por cima.

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
  (CNPJ, CPF, valores) — `comum._DIG = "0-9OoIiSs"` substitui `\d` nessas
  regex especificamente, e `comum.corrigir_confusao_ocr()` normaliza o
  trecho capturado antes de usar. Aplicar a correção no texto inteiro
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

`tests/test_extractors_boleto.py` roda `BoletoExtractor` sobre
`tests/fixtures/boleto_real_anonimizado.txt` — texto bruto de um boleto
real (nome/CPF/CNPJ/valores trocados por fictícios, mas rótulos e
estrutura de linha exatamente como o pdfplumber extraiu). Esse fixture
existe porque cenários sintéticos escritos à mão não reproduziam bugs
reais: o mesmo rótulo (ex: "Sacado") aparece várias vezes no boleto em
contextos diferentes (título de seção, rodapé, tabela-resumo, campo de
verdade), e só um teste sobre o texto real pega isso. `comum.localizar_rotulo`
loga (nível DEBUG, logger `app.extractors.comum`) qual rótulo/linha foi
aceito ou rejeitado para cada campo — o teste imprime esse log com
`caplog`, então dá pra ver exatamente onde a extração está acertando ou
errando sem precisar adivinhar. `tests/test_basic_extractor.py` testa só
a camada fina de compatibilidade (`extrair()`/`extrair_com_metadados()`);
`tests/test_detector.py` testa o roteamento; `tests/test_chave_acesso.py`
testa o dígito verificador com vetores sintéticos (sem depender de DANFE
real nenhuma).

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

`POST /debug/extract-words` (novo) devolve as palavras de cada página com
posição x0/x1/top/bottom (`pagina.extract_words()`), sem extração em cima
— usado especificamente pra diagnosticar layouts em grade (tabela de
itens da DANFE) antes de escrever a lógica de reconstrução por
coordenadas. Só funciona com texto digital (`paginas_palavras` é `None`
quando a origem é OCR — retorna 400 nesse caso).

**Refatoração por padrão Strategy em andamento** (ver "Arquitetura de
extratores" acima) — etapas 1 e 2 do plano concluídas: extratores
separados por tipo com boleto migrado sem regressão (mesma fixture real,
mesmos resultados), captura de palavras posicionadas, e os dois endpoints
de debug. **Bloqueado na etapa 3** (campos fiscais da DANFE + tabela de
itens por coordenadas) esperando um dump anonimizado real via
`/debug/extract-words` — não implementar isso por suposição foi decisão
explícita, dado o histórico de bugs reais de layout neste projeto.

Não implementado ainda / possíveis próximos passos:

- DANFE: campos fiscais específicos (série, natureza da operação, CFOP
  predominante, ICMS, frete) e tabela de itens por coordenadas — aguardando
  dump real (ver acima).
- Confiança por campo (`ResultadoExtracao.confiancas`) já é calculada em
  cada extrator mas ainda não está exposta em `ExtractionResult`/API nem
  na interface — etapas 4 e 5 do plano.
- Excel: as duas abas (Resumo/Itens) já existem, falta só formatação
  (cabeçalho em negrito, largura de coluna) — etapa 6 do plano.
- Correção de confusão de OCR cobre só CNPJ/CPF/valores — não
  numero_documento/Nosso Número nem texto livre (emissor/destinatario).
- Só há um fixture de teste automatizado com texto real (um boleto
  anonimizado). Uma fixture real de DANFE está a caminho (ver acima).

Plano original com todas as etapas em
`C:\Users\User\.claude\plans\jazzy-foraging-harbor.md`.
