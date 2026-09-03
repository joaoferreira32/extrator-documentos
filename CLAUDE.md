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
    danfe.py              # DanfeExtractor (deteccao, chave de acesso, valor_total,
                          #   delega tabela de itens pra danfe_tabela.py)
    danfe_tabela.py        # reconstrucao da tabela de itens da DANFE por coordenadas
                          #   x0/x1/top/bottom (nao regex sobre texto corrido)
    generico.py           # GenericExtractor (fallback: pedido_compra, relatorio, desconhecido)
  llm_extractor.py     # extração via Claude (modo "ia", opcional)
  excel_exporter.py    # ExtractionResult -> .xlsx
backend/tests/
  fixtures/boleto_real_anonimizado.txt     # texto bruto de boleto real (dados trocados)
  fixtures/danfe_real_anonimizado.txt      # texto bruto de DANFE real (destinatario ficticio)
  fixtures/danfe_palavras_anonimizado.json # palavras posicionadas (x0/x1/top/bottom) da mesma DANFE
  test_basic_extractor.py               # compat da funcao publica extrair()
  test_extractors_boleto.py             # BoletoExtractor isolado, com a fixture real acima
  test_extractors_danfe.py              # DanfeExtractor + montar_tabela_itens, fixtures reais acima
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

### DANFE (`danfe.py` + `danfe_tabela.py`)

Detecção por marcador (`MARCADORES_DETECCAO`: `danfe`, `documento auxiliar
da nota fiscal eletr[ônica]`, `chave de acesso`, `protocolo de autorização
de uso`, mais `nf-e`/`cfop` que já existiam). `DanfeExtractor` parte de
`GenericExtractor` para os campos universais, mas **sobrescreve
emissor/destinatário/valor_total com lógica própria** — ver "Layout real
de uma DANFE" abaixo pra entender por quê. Por cima disso:

- **Chave de acesso**: 44 dígitos + dígito verificador módulo 11, pesos
  cíclicos 2–9 — algoritmo público e padronizado pela SEFAZ, não depende
  de layout nenhum (`test_chave_acesso.py` tem vetores sintéticos
  calculados a mão). Confiança `"alta"` quando o DV bate, `"baixa"`
  quando acha 44 dígitos mas o DV não confere (não descarta, só avisa).
  **Bug real corrigido**: `_CANDIDATO_CHAVE_RE` usava `\s` (que também
  casa quebra de linha) como separador entre os blocos de dígitos — se a
  linha ANTERIOR à chave também terminasse em dígito (ex: um CNPJ
  `"...0010-01"`), o candidato "vazava" pra trás, pegando os últimos
  dígitos daquela linha e inflando a contagem pra mais de 44, descartando
  a chave de verdade. Trocado por `" "` (espaço literal) — os blocos da
  chave são sempre separados por espaço na mesma linha, nunca por quebra
  de linha.
- **`valor_total`**: sobrescrito via rótulo `"Valor Total da Nota"` (ou
  variantes) com confiança `"alta"` — o fallback genérico
  (`comum.valor_por_total_ou_ultimo`) exige um `"R$"` na frente do valor,
  mas uma DANFE real imprime só o número, sem o símbolo, então o fallback
  genérico deixava `valor_total` como `None`.
- **Campos fiscais específicos** (série, natureza da operação, data de
  saída, IE do emitente, ICMS/frete como campos de documento — não
  agregados da tabela) **ainda não implementados** — dependem de
  confirmar o rótulo exato num dump real antes de codar, pra não repetir
  o erro de supor layout sem material.

#### Layout real de uma DANFE: rótulo pode vir antes OU depois do valor

Confirmado com um dump real via `/debug/extract-text` (a mesma DANFE da
Dell): uma DANFE real **mistura duas convenções diferentes** de onde o
rótulo aparece em relação ao valor, dependendo da caixa:

- **Destinatário/datas** (`NOME/RAZÃO SOCIAL`, `DATA DA EMISSÃO`, ...):
  rótulo PRIMEIRO, valor na linha seguinte — o padrão "normal" que
  `comum.localizar_rotulo` já assumia.
- **Grade de totais** (`VALOR TOTAL DA NOTA`, `VALOR TOTAL DOS PRODUTOS`,
  `BASE DE CÁLCULO DO ICMS`, ...): valor PRIMEIRO, rótulo como legenda na
  linha seguinte (ex: `"229,00"` na linha logo ANTES de `"VALOR TOTAL DA
  NOTA"`). Isso é uma repetição de caixas visuais (valor em destaque,
  legenda pequena embaixo) — cada linha vizinha a um rótulo dessa grade
  tem *algum* valor monetário (o campo anterior ou o seguinte), então não
  dá pra simplesmente "tentar as duas direções e aceitar a que bater": as
  duas sempre batem, só uma é a certa.
- **Emitente**: o nome da empresa aparece **antes de qualquer rótulo**,
  logo no topo do documento — não tem "rótulo: valor" nenhum ali.

Isso gerou 4 bugs reais, só descobertos testando o PDF real no navegador
(a fixture de texto anterior era uma linearização escrita à mão, não uma
extração real, e não reproduzia nenhum deles):

1. **`emissor` saía `"DANFE"`**: `GenericExtractor` busca o rótulo
   genérico `"Emitente"`/`"Fornecedor"`, que bate como substring dentro da
   legenda `"Identificação do Emitente"`. Como essa legenda funciona como
   legenda-embaixo-do-valor (não "rótulo: valor"), a "próxima linha" que
   `localizar_rotulo` tentava como fallback pertencia à caixa vizinha (o
   título `"DANFE"`), não ao nome do emitente de verdade. **Corrigido**
   com `DanfeExtractor._extrair_emissor`: pega a primeira linha do
   documento que parece um nome (`comum.parece_nome`) e não é um dos
   `MARCADORES_DETECCAO` (evita recapturar `"DANFE"`/`"NF-e"`/etc. caso o
   nome não seja mesmo a primeira linha em algum layout) — estrutural,
   não depende de rótulo nenhum. O CNPJ do emitente fica até 7 linhas
   depois do nome (depois do bloco de endereço), então essa busca usa
   `janela=10` em `documento_fiscal_proximo` em vez do default (3).
2. **`destinatario` vazio**: `ROTULOS_DESTINATARIO` genérico
   (`"Destinatário"`, `"Cliente"`) não bate com o rótulo real
   `"NOME/RAZÃO SOCIAL"`. Corrigido com `ROTULOS_DESTINATARIO_DANFE`
   específico de DANFE (não adicionado ao genérico pra não arriscar
   ambiguidade com outro tipo de documento).
3. **`data_emissao` vazio**: dois bugs empilhados. (a) `ROTULOS_DATA_EMISSAO`
   só tinha `"Data de Emissão"` — o rótulo real usa `"da"`, não `"de"`
   (`"DATA DA EMISSÃO"`); variante adicionada em `generico.py` (genérica,
   não específica de DANFE — as duas preposições são válidas em
   português). (b) mesmo com o rótulo certo, `DATA_RE` exigia `\d{2}`
   pra dia E mês — a data real vem sem zero à esquerda no mês
   (`"15/4/2026"`), rejeitada pelo regex. Relaxado pra `\d{1,2}`.
4. **`valor_total` vazio** (e a validação de soma da tabela sempre
   avisava divergência): `comum.extrair_valor_rotulo` só olhava a MESMA
   linha do rótulo — sem fallback nenhum pra linha vizinha. Adicionado
   fallback bidirecional (linha seguinte E anterior), com um parâmetro
   `preferir_linha_anterior` pra controlar a ORDEM de tentativa: DANFE
   passa `True` (grade de totais, valor antes do rótulo — ver acima),
   boleto continua com o default `False` (linha seguinte primeiro, único
   padrão que boleto usa, comportamento já testado preservado).

Os 4 bugs (mais o item "fantasma" de linha vertical/coluna, já corrigido
antes) só apareceram com o PDF real — reforça a mesma lição já documentada
pro boleto: fixture escrita à mão não substitui um dump real.

#### Tabela de itens por coordenadas (`danfe_tabela.py`)

Usa `pagina.extract_words()` (x0/x1/top/bottom de cada palavra), não regex
sobre texto corrido — texto corrido perde a estrutura de colunas.
Estrutura confirmada com um dump real via `/debug/extract-words` de uma
DANFE real (nota da Dell); fixtures em `tests/fixtures/danfe_*` (dados do
destinatário são fictícios, dados do emitente/tabela são reais e
públicos — CNPJ de empresa, não dado pessoal).

- **Agrupamento de linha** (`_agrupar_linhas`): palavras cujo `top`
  difere em até `TOLERANCIA_LINHA` (3pt) formam a mesma linha visual.
- **Identificação de coluna pelo cabeçalho** (`_identificar_colunas`):
  cada coluna tem uma palavra-chave curta (ex: `"calc"` pra
  B.CALC.ICMS, `"i.p.i"` — com pontos — pra I.P.I., já que
  `pdfplumber` só quebra em espaço em branco e alguns rótulos saem
  colados sem espaço interno). A busca varre a linha de cabeçalho da
  esquerda pra direita e **nunca volta pra trás** (índice mínimo avança a
  cada rótulo casado) — evita que "ICMS" (que aparece dentro de
  "B.CALC.ICMS", "VALOR ICMS" e "ICMS/IPI") case com a coluna errada.
- **Fronteira de coluna** (`_limites_colunas`): ponto médio entre âncoras
  consecutivas. A primeira coluna não tem limite inferior
  (`float("-inf")`) e a última não tem limite superior (`float("inf")`)
  — **bug real corrigido**: a primeira versão usava o x0 da própria
  âncora do cabeçalho como limite inferior da primeira coluna, mas um
  token de dado pode começar antes do rótulo do cabeçalho (ex: código de
  produto "460-BCZS" em x0=85, rótulo "CÓDIGO" em x0=98.7) — sem o
  `-inf`, esse token ficava sem coluna nenhuma, a linha do item real era
  classificada como "fim de tabela" e a tabela inteira saía vazia.
  **Efeito colateral desse mesmo `-inf`, corrigido à parte** (ver
  `_linha_e_marcador_fim` abaixo): com a coluna CÓDIGO agora aceitando
  qualquer x0 pequeno, texto de margem esquerda de OUTRAS seções da
  página (não só rodapé) também passou a cair nela.
- **Fim da tabela por conteúdo, não posição** (`_linha_e_marcador_fim`):
  antes de classificar uma linha por coluna, checa se o texto da linha
  inteira contém um dos títulos de seção padronizados nacionalmente que
  vêm depois da tabela de itens em qualquer DANFE (`MARCADORES_FIM_TABELA`:
  "informações complementares", "dados adicionais") — se sim, `break`
  imediato, antes de qualquer lógica de coluna. **Bug real corrigido**:
  testando com uma DANFE real no navegador, o bloco "INFORMAÇÕES
  COMPLEMENTARES" (texto de observações livres, várias linhas) caía por
  coincidência de x na coluna CÓDIGO (efeito colateral do `-inf` acima) e
  virava uma sequência de itens fantasmas. Corte por CONTEÚDO da linha em
  vez de posição vertical fixa porque a posição varia de layout pra
  layout — os títulos de seção, não.
- **Classificação de linha** (`_tipo_linha`): `"item"` exige conteúdo na
  coluna CÓDIGO (todo item de verdade tem código de produto — sinal mais
  confiável do que "tem algum número em alguma coluna", que rodapé/seção
  seguinte também podem ter por coincidência de posição de x).
  `"continuação"` é conteúdo só em DESCRIÇÃO (quebra de linha dentro da
  célula). Qualquer outra coisa é `"fim"` e encerra a tabela. **Bug real
  corrigido**: a versão anterior só checava "a linha tem algum valor em
  ncm/cfop/quantidade/valor_unitario/valor_total" pra decidir se
  continuava — texto de rodapé como "Valor Total dos Produtos 215,03"
  caía por coincidência de x dentro da faixa da coluna NCM e virava um
  item fantasma vazio.
- **Valores colados** (`_separar_valores_colados`): uma coluna estreita
  pode fazer o pdfplumber juntar dois valores monetários num só token sem
  espaço (ex: `"13,9718,00"` = VALOR I.P.I. `13,97` + ALÍQUOTA ICMS
  `18,00`, confirmado no dump real). Separado por **padrão** de valor
  monetário (regex), não por posição de pixel — mais robusto do que
  adivinhar uma fronteira de sub-coluna que o cabeçalho nem rotula.
- **Texto vertical descartado** (`_eh_texto_vertical`): rótulos de seção
  impressos girados (ex: "SOTUDORP" = "PRODUTOS" ao contrário) ficam numa
  faixa estreita de x0 na margem esquerda (15–85) e têm caixa
  delimitadora alta e estreita (altura > 2× largura) — o oposto de uma
  palavra horizontal normal. Descartados antes de agrupar linhas.
- **Validação cruzada**: soma dos `valor_total` de todos os itens
  comparada com o rótulo nacional `"Valor Total dos Produtos"`
  (tolerância de 0.02 pra arredondamento). Se não bater, vira aviso em
  `resultado.avisos` em vez de devolver a tabela calada.
- Confiança `"media"` (heurística posicional) pra `itens` e pra `CFOP`
  (moda dos CFOPs das linhas da tabela, exposto em `campos_adicionais`).
- Só roda quando `contexto.paginas_palavras` existe — `None` quando a
  origem do texto é OCR (sem posição confiável de palavra); tabela de
  itens via OCR fica fora de escopo, documentado como limitação
  conhecida.

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

`tests/test_extractors_danfe.py` testa `DanfeExtractor` e
`danfe_tabela.montar_tabela_itens` contra as fixtures reais em
`tests/fixtures/danfe_*` (texto bruto + palavras posicionadas de uma DANFE
real, destinatário fictício). O texto bruto preserva a ORDEM DE LINHAS
exata que o pdfplumber produziu (via `/debug/extract-text`), não uma
linearização "rótulo: valor" escrita à mão — isso importa porque a
primeira versão da fixture era escrita à mão e não reproduziu nenhum dos
bugs reais de emissor/destinatário/data_emissão/valor_total/itens
fantasmas documentados em "Layout real de uma DANFE" acima, que só
apareceram testando o PDF real no navegador. Cobre também os bugs de
tabela já documentados (descarte de texto vertical, separação de valores
colados, corte por bloco de informações complementares) e a validação de
soma com aviso.

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

**Refatoração por padrão Strategy** (ver "Arquitetura de extratores"
acima) — etapas 1, 2 e 3 do plano concluídas: extratores separados por
tipo com boleto migrado sem regressão (mesma fixture real, mesmos
resultados), captura de palavras posicionadas, os dois endpoints de
debug, e a DANFE completa (chave de acesso, valor_total por rótulo, tabela
de itens por coordenadas com CFOP predominante e validação de soma) — tudo
construído a partir de um dump real via `/debug/extract-words` (nota da
Dell), não por suposição. `confiancas` (`ResultadoExtracao`/
`ExtractionResult`) já é calculada em cada extrator e exposta na API —
etapa 4 concluída.

Não implementado ainda / possíveis próximos passos:

- DANFE: campos fiscais adicionais que ainda dependem de confirmar o
  rótulo exato num dump real (série, natureza da operação, data de saída,
  IE do emitente, ICMS/frete como campos de documento — não agregados da
  tabela).
- Indicador de confiança na interface (marcador discreto pra campos
  `"media"`/`"baixa"` em `resultado.confiancas`) — etapa 5 do plano, ainda
  não iniciada.
- Excel: as duas abas (Resumo/Itens) já existem, falta só formatação
  (cabeçalho em negrito, largura de coluna) — etapa 6 do plano, ainda não
  iniciada.
- Correção de confusão de OCR cobre só CNPJ/CPF/valores — não
  numero_documento/Nosso Número nem texto livre (emissor/destinatario).
- Tabela de itens da DANFE não funciona com texto de origem OCR (só
  digital, que tem posição confiável de palavra) — limitação conhecida,
  não um bug.

Plano original com todas as etapas em
`C:\Users\User\.claude\plans\jazzy-foraging-harbor.md`.
