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
- Excel: openpyxl (escrita direta, sem pandas)
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
  excel_exporter.py    # lista de documentos -> .xlsx: Relatorio + 4 abas de dados (Tabelas nomeadas),
                        #   total/grafico de lote, listas suspensas, protecao (openpyxl) -- ver "Excel"
  confianca.py         # confianca geral ("X de Y alta"): a mesma regra da tela, em Python
  logging_config.py    # log estruturado (logger "app") + request-id em contextvar -- ver "Robustez e observabilidade"
  rate_limit.py        # limite por IP (middleware ASGI, janela deslizante em memoria) -- idem
backend/tests/
  fixtures/boleto_real_anonimizado.txt     # texto bruto de boleto real (dados trocados)
  fixtures/danfe_real_anonimizado.txt      # COMPOSTO de trechos reais (nao as 84 linhas) -- ver docstring do teste
  fixtures/danfe_palavras_anonimizado.json # palavras posicionadas (x0/x1/top/bottom) da mesma DANFE
  test_basic_extractor.py               # compat da funcao publica extrair()
  test_extractors_boleto.py             # BoletoExtractor isolado, com a fixture real acima
  test_extractors_danfe.py              # DanfeExtractor + montar_tabela_itens (fixture de texto e um TRECHO, nao a entrada real)
  test_debug_endpoints.py               # endpoints de debug mostram exatamente o que o extrator recebe
  test_comum.py                         # helpers compartilhados (normalizar_data)
  test_detector.py                      # selecionar_extrator() roteia pro extrator certo
  test_chave_acesso.py                  # digito verificador da chave de acesso (algoritmo publico)
  test_pdf_extractor.py                 # orquestracao digital->OCR (mocks) + 1 teste OCR real (skip se indisponivel)
  test_avisos.py                        # avisos (lista) x aviso (string juntada)
  test_excel_exporter.py                # Excel multi-aba lido de volta com openpyxl (ver "Excel")
  test_ooxml.py, ooxml.py, ooxml_sdk.ps1 # o .xlsx segue as regras do EXCEL (XML bruto) + SDK oficial da Microsoft (opt-in)
  test_confianca.py                     # regra da confianca geral em Python (a mesma da tela)
  test_concorrencia.py                  # trabalho pesado nao bloqueia o /health (run_in_threadpool)
  test_limites.py                       # tetos de tamanho/quantidade (upload, debug, /export-excel)
  test_logging_config.py, test_observabilidade.py # log estruturado, request-id e X-Request-ID
  test_erros_de_leitura.py              # PDF com senha e PDF sem paginas: mensagem propria, nao "corrompido"
  test_rate_limit.py                    # limite por IP: limitador (relogio falso), IP, middleware, app real, uvicorn real
  e2e/test_interface_limite.py          #   (e2e) o 429 aparece na tela com a mensagem do backend
  e2e/                                  # testes de interface no navegador (Playwright) -- opt-in, ver "Testes"
    helpers.py                          #   servidor uvicorn temporario + PDFs FICTICIOS gerados por PyMuPDF
    conftest.py, test_interface_confianca.py
    gerar_screenshot.py                 #   regenera docs/screenshot.png (README) com dados ficticios
backend/requirements-dev.txt            # so dev/teste de interface (playwright); requirements.txt fica so com o de producao
frontend/
  index.html, style.css, script.js   # resultado com chips de confianca, banner de avisos, edicao de campos
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
- **"Valor do Documento" vem na linha SEGUINTE ao rótulo** (cabeçalho de
  tabela): `"Valor Documento (-) desconto (-) outras deduções ..."` e, na
  linha de baixo, `"1.000,00"`. `comum.extrair_valor_rotulo` só lia a mesma
  linha, então esse campo não era extraído — o código ORIGINAL
  (pré-refatoração) também não o extraía desse texto (verificado rodando o
  `basic_extractor` de `432d42e` sobre a fixture: 6 campos, sem ele). Ele
  passou a sair (7 campos) por efeito colateral de um fallback bidirecional
  que adicionei pra DANFE em 0e804e7; ao **restaurar** a função "ao
  original" o campo sumiu de novo, e **nenhum teste travava o campo**, então
  a suíte passou com o boleto "quebrado". Hoje: `extrair_valor_rotulo(...,
  aceitar_linha_seguinte=True)`, **opt-in** e só pro Valor do Documento
  (`boleto.py`); o default continua só a mesma linha (usado pela DANFE, onde
  a linha vizinha é de OUTRO campo). `test_boleto_traz_os_7_campos_
  adicionais` trava os 7 campos com valor, ordem e confiança.
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
emissor/destinatário/valor_total com lógica própria** (e confere
número/emissor contra a chave de acesso) — ver "Layout real de uma DANFE"
abaixo pra entender por quê. Por cima disso:

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

#### Layout real de uma DANFE (confirmado com `/debug/extractor-input`)

Tudo abaixo vem da saída de `/debug/extractor-input` de uma DANFE real
(Dell), com dados pessoais trocados pelo usuário e a estrutura de linhas
preservada. **Histórico do que deu errado antes:** as correções anteriores
foram feitas em cima de um TRECHO colado à mão que o usuário disse ser o
texto completo — não era. O PDF real começava com texto girado ("FOLHA
1/", "e-FN", ...), e o bloco do destinatário saía com rótulos e valores
colados numa linha só. Resultado: 33 testes passando e o PDF real falhando
em emissor, destinatário e itens (commit 0e804e7). A hipótese que eu tinha
posto aqui ("o nome do emitente vem antes de qualquer rótulo, no topo") era
**falsa**. Regra que sai disso: não corrigir sem ter visto o texto exato
(ver "Diagnosticando um PDF real").

**1. Texto girado polui tudo (`pdf_extractor._sem_texto_girado`).** O
canhoto e os rótulos laterais são impressos girados (90/270°) e o
pdfplumber os extrai como linhas de letras soltas/invertidas ("FOLHA 1/",
"e-FN", "SERODATUPMOC", "SODATROPSNART"...) ANTES do conteúdo real (as 39
primeiras linhas da DANFE de teste). Descartados na leitura via `char["upright"]
== False` + `page.filter` — na raiz, vale pra todos os extratores e também
pras palavras posicionadas da tabela. Verificado empiricamente que
`upright` é `False` a 90°/270°. Salvaguardas/limites: (a) só filtra quando
os girados são MINORIA da página (página inteira girada por `/Rotate` não é
apagada); (b) texto de **cabeça pra baixo (180°) continua `upright=True`** e
não é descartado — a extração o devolve invertido. `TextoExtraido.
caracteres_girados_descartados` e `/debug/extractor-input` mostram quantos
foram descartados. O filtro por geometria em `danfe_tabela._eh_texto_
vertical` continua como rede de segurança.

**2. Emitente** (`danfe.extrair_emissor`). A legenda é `"Identificação do
emitente"` e o nome vem na **linha seguinte**, mas o pdfplumber cola texto da
caixa vizinha nas duas linhas:

    Identificação do emitente DANFE
    DELL COMPUTADORES DO BRASIL LTDA Documento Auxiliar da
    Nota Fiscal Eletrônica

Por isso o rótulo genérico `"Emitente"` (do `GenericExtractor`) bate na
primeira linha e captura `"DANFE"` como emissor — esse era o bug original.
Agora: ancora em `Identificação do emitente`, pega o resto da linha (se
sobrar algo além do título) ou a linha seguinte, e **corta no título da caixa
vizinha** (`documento auxiliar`, `danfe`). O emissor do `GenericExtractor` é
sempre descartado; se a legenda não existe (outro layout), o genérico só
sobrevive com confiança `"baixa"` e nunca se contiver o título.

**O CNPJ do emissor** (`_cnpjs_do_bloco_do_emitente`) fica varias linhas
depois do nome — no PDF real, **10 linhas** (nome na linha 40, CNPJ na 50:
`"INSCRIÇÃO ESTADUAL INSCR. ESTADUAL DO SUBST. TRIBUT. CNPJ"` e
`"748241245113 72.381.189/0010-01"`). A busca antiga usava a janela fixa de
10 de `documento_fiscal_proximo` (`linhas[40:50]`), que perdia o CNPJ **por
uma linha** — o emissor saía sem CNPJ e com confiança `"media"`. A fixture
antiga escondia isso porque tinha o CNPJ a ~7 linhas. Agora o bloco do
emitente vai do nome até o rótulo do destinatário (`NOME/RAZÃO SOCIAL`, onde
começa outra entidade cujo CNPJ/CPF não pode virar o do emissor), teto de 40
linhas; havendo mais de um CNPJ no bloco, prefere o que bate com a chave.

**3. Destinatário** (`danfe.extrair_destinatario`). Rótulos e valores vêm
COLADOS na mesma linha (pdfplumber junta a linha de rótulos e a de valores
quando estão próximas na vertical):

    NOME/RAZÃO SOCIAL CNPJ/CPF DATA DA EMISSÃO FULANO DE TAL SILVA 000.000.000-00 15/4/2026 ENDEREÇO ...

O nome é o que fica **entre o último rótulo da sequência inicial e o
CPF/CNPJ**. Pula rótulos consecutivos do bloco (`ROTULOS_BLOCO_DESTINATARIO`,
com limite de palavra: `"UF"` não casa dentro de `"UFRJ"`), corta no
primeiro CPF/CNPJ, data ou rótulo, e devolve `"NOME (CPF 000...)"` /
`"NOME (CNPJ ...)"` — **mesmo formato do emissor e do boleto**. Também aceita o layout de valores na linha seguinte. O destinatário do
`GenericExtractor` é descartado (buscava `"Destinatário"`, que bate no título
da seção e devolvia uma linha inteira de rótulos).

**4. Confiança por evidência, não por ter achado rótulo.** Antes,
emissor/destinatário saíam `"alta"` mesmo errados e o número da nota, certo,
saía `"baixa"`. Agora:
- **Chave de acesso** com DV válido serve de evidência independente
  (`danfe.partes_da_chave`: cUF, AAMM, CNPJ, modelo, série, nNF — layout
  padronizado pela SEFAZ). O CNPJ lido junto do emissor bate com o CNPJ da
  chave → `"alta"`; contradiz → `"baixa"` + aviso; sem chave válida →
  `"media"`. O número da nota lido bate com o nNF da chave → `"alta"`;
  contradiz → `"baixa"` + aviso.
- Destinatário: `"alta"` quando o nome fica encaixado antes do CPF/CNPJ
  (estrutura confirmada), `"media"` sem isso.

**5. Grade de totais: rótulo → valor de MESMA POSIÇÃO** (`danfe.
extrair_totais_grade`). Cada "linha" da grade sai do pdfplumber como **uma
linha de N rótulos seguida de uma linha de N valores**:

    BASE DE CÁLCULO DO ICMS VALOR DO ICMS BASE DE CÁLCULO ICMS ST VALOR DO ICMS SUBSTITUIÇÃO VALOR TOTAL DOS PRODUTOS
    229,00 41,22 0,00 0,00 215,03
    VALOR DO FRETE VALOR DO SEGURO DESCONTO OUTRAS DESPESAS ACESSÓRIAS VALOR TOTAL DO I.P.I. VALOR TOTAL DA NOTA
    0,00 0,00 0,00 0,00 13,97 229,00

O valor de um rótulo é o de mesma posição (5º rótulo → 5º valor), **nunca
"o número mais próximo"**. **Bug real corrigido**: a versão anterior pegava o
número mais próximo (e eu tinha documentado aqui que "o valor vem ANTES do
rótulo" — hipótese tirada de um trecho colado à mão, **refutada** pelas
linhas reais). Resultado: `Valor Total dos Produtos` saía 229,00 (a base do
ICMS, 1º valor da linha) em vez de 215,03 → aviso falso de soma; e
`valor_total = 229` estava certo **só por coincidência** (nota = produtos +
IPI = 215,03 + 13,97 = base do ICMS).

Regras (para não chutar): a linha só conta se, tirando os rótulos
conhecidos, sobram só valores/espaços; a quantidade de valores tem que ser
**igual** à de rótulos, senão os campos daquela linha não são devolvidos
(rótulo desconhecido, valor faltando...); valores na linha seguinte só valem
com 2+ rótulos (com 1 rótulo o número da próxima linha é ambíguo). A
confiança de `valor_total` é `"alta"` quando a fórmula do total fecha
(`totais_fecham`: produtos − desconto + ICMS ST + frete + seguro + outras
despesas + IPI = total da nota), `"media"` + aviso quando a grade foi lida e
não fecha. Sem grade (outro layout), tenta `"Valor Total da Nota <valor>"` na
mesma linha (`comum.extrair_valor_rotulo` com o **default**, só a mesma
linha — a DANFE não liga `aceitar_linha_seguinte`, ver "Boleto"), e por fim o
fallback genérico (exige `"R$"`). A soma dos itens é conferida contra
`valor_produtos` da grade.

Casamento é por **ordem**, não por coordenada x: não tenho coordenadas reais
dessa grade e não vou chutar; a checagem N = N e a fórmula dos totais fazem o
papel de validação. Se for preciso confirmar por x, pedir
`/debug/extract-words` da região.

**6. Datas.** O documento imprime `"15/4/2026"` (mês sem zero). `DATA_RE`
aceita 1–2 dígitos e `comum.normalizar_data` devolve sempre `dd/mm/aaaa`
(`"15/04/2026"`), em qualquer extrator (no boleto é no-op). O rótulo real é
`"DATA DA EMISSÃO"` ("da", não "de"); as duas preposições estão em
`generico.ROTULOS_DATA_EMISSAO`.

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
- **Fragmentos do cabeçalho abaixo dele** (`LIMITE_CABECALHO_MULTILINHA`):
  células de cabeçalho com várias linhas ("ALÍQUOTA" / "ICMS" / "IPI"
  empilhadas) saem como fragmentos soltos ("IC M S IP I") numa linha logo
  ABAIXO do cabeçalho (top 416.7 contra 413.6 no PDF real — 3.1pt, acima
  da tolerância de agrupamento, então é uma linha separada). Sem CÓDIGO e
  sem ser só descrição, era classificada `"fim"` e encerrava a tabela
  ANTES do primeiro item: `itens: []`. **Bug real corrigido.** Enquanto
  nenhum item começou, linhas `"fim"` a até 15pt do cabeçalho são puladas
  (`decisao: "pulada: fragmento do cabecalho..."` no `/debug/extractor-input`).
  Depois do primeiro item, uma linha `"fim"` continua encerrando a tabela.
  Verificado por mutação: com o limite em 0 o teste falha e a tabela sai
  vazia. Nota: `"13,9718,00"` (VALOR I.P.I. + ALÍQUOTA ICMS colados) não
  afeta o item — o schema só lê quantidade/valor unitário/valor total, de
  outras colunas.
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
- Confiança de `itens`: `"media"` (heurística posicional), que **sobe pra
  `"alta"` quando a soma dos itens fecha com o Valor Total dos Produtos**
  da grade de totais (evidência independente: o total vem de outra parte do
  documento) — e só se TODOS os itens tiverem total numérico (um item sem
  total contaria como 0 e poderia "fechar" a soma por acaso). Se a soma não
  fecha: continua `"media"` e gera aviso; sem Valor Total dos Produtos pra
  conferir: `"media"`. `CFOP` (moda dos CFOPs das linhas da tabela, em
  `campos_adicionais`) fica `"media"`.
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

## Interface: confiança e edição (etapa 5)

O frontend (`index.html`, `script.js`, `style.css`, sem framework) mostra a
confiança de cada campo (`resultado.confiancas`), lista os avisos e deixa o
usuário **corrigir qualquer campo antes de exportar**. Motivo de a confiança
não poder bloquear a edição: já saiu emissor errado com confiança `"alta"`.

**Indicador por campo** — sempre ícone (forma diferente) + TEXTO + cor, nunca
só cor:

| Estado | Ícone | Texto | Quando |
|---|---|---|---|
| alta | círculo ✔ | Alta | `confiancas[campo] == "alta"` |
| média | triângulo ▲ | Média | `"media"` |
| baixa | círculo ✖ | Baixa | `"baixa"` |
| vazio | círculo tracejado | Vazio | campo **obrigatório** sem valor |
| corrigido | lápis | Corrigido | o usuário mudou o valor |
| (nenhum) | — | — | opcional vazio (`—` neutro), tipo do documento, e tudo no modo IA |

Cores: `alta` usa o único token novo (`--color-success` `#1b6b3a` sobre
`#e6f4ea`, 5,8:1); média/baixa reaproveitam os tokens de warning/danger já
existentes (5,4:1 e 5,7:1). Todos os pares texto/fundo dos chips, dicas e
resumo medidos no navegador (estilos computados) ≥ 4,5:1 (AA).

**Estado derivado, não flags.** O estado de um campo é calculado a cada
render a partir de (valor original, valor atual, obrigatoriedade,
confiança) — `estadoDoCampo` em `script.js`. Ordem: vazio (obrigatório) /
opcional-vazio → `corrigido` (atual ≠ original) → confiança do backend. Se o
usuário desfaz a edição, o campo volta sozinho ao estado original.

**Edição.** Todos os campos são editáveis:
- **Já abertos como input:** `baixa` e `vazio` obrigatório (destacados: borda
  lateral vermelha / caixa tracejada + dica "Confira e corrija." /
  "Preencha este campo.").
- **Abrem ao clicar** (valor com ícone de lápis; é um `<button>`, então
  funciona por teclado): `alta`, `media`, opcional vazio, corrigido, e o tipo.
- **Enter** confirma, **Esc** descarta, sair do campo (blur) confirma. O valor
  vai para `ultimoResultado.documento` na hora (a cada tecla), e é esse objeto
  que o botão Excel envia — a correção sai no `.xlsx` sem mexer no backend.
- `valor_total` aceita formato BR (`1.500,50`, `R$ 1.500,50`, `1500.50`) e vai
  como **número**; texto não parseável fica como string (o schema aceita
  `float | str`). Exibido em pt-BR (`1.500,50`). Campo apagado vira `null`
  (`campos_adicionais` exige string: vira `""`).
- **Tipo do documento** é um `<select>`; mudar o tipo muda os campos
  obrigatórios na hora (re-renderiza todos).
- A tabela de itens NÃO é editável (só ganha o chip de confiança `itens`).

**Exibição formatada — só na tela; o valor interno não muda.** O JSON bruto e o
Excel continuam com o valor original (o `.xlsx` sai de `ultimoResultado`, que
nunca recebe o texto formatado):
- **Itens em pt-BR** (`215.03` → `215,03`): quantidade com até 4 casas, valor
  unitário 2–4 (preço unitário pode ter 4 casas), valor total 2. O `valor_total`
  do documento também é pt-BR (2 casas). Valor que não virou número fica como
  veio. No Excel os números continuam numéricos (somáveis).
- **Tipo do documento com rótulo amigável** (`ROTULOS_TIPO`): "Boleto", "Nota
  fiscal", "Pedido de compra", "Relatório", "Desconhecido" — na tela e nas
  opções do `<select>`. O `value` interno (`nota_fiscal`...) é o que vai pro JSON
  e pro Excel. Tipo fora da lista aparece cru.
- **Chave de Acesso em blocos de 4 dígitos** (padrão impresso da DANFE, 11
  blocos), também dentro do editor. Os espaços são só exibição: ao editar, o
  valor volta a ser só dígitos, e abrir/confirmar sem mudar NÃO conta como
  correção (o estado compara o valor sem espaços).

**Campos obrigatórios por tipo de documento** (`CAMPOS_OBRIGATORIOS` em
`script.js`). Obrigatório vazio → destaque, já aberto, e conta como "vazio" no
resumo. Opcional vazio → `—` neutro, editável ao clicar, fora do resumo.

| `tipo_documento` | Obrigatórios |
|---|---|
| `nota_fiscal` (inclui DANFE) | `emissor`, `destinatario`, `numero_documento`, `data_emissao`, `valor_total` |
| `boleto` | os mesmos + `data_vencimento` |
| `pedido_compra` | os mesmos da nota fiscal *(escolha minha, não pedida)* |
| `relatorio` | nenhum *(escolha minha)* |
| `desconhecido` e qualquer outro tipo (ex.: um que o modo IA invente) | nenhum |

`campos_adicionais` nunca são obrigatórios. Ex.: `data_vencimento` numa DANFE é
opcional — aparece `—` sem destaque e não entra no resumo.

**Resumo no topo** (`aria-live="polite"`, texto — não depende de cor):
"X de Y campos com alta confiança · N para revisar · M vazios · K corrigidos".
- Y = campos **com valor e com confiança** (documento sem o tipo,
  `campos_adicionais` e `itens`); X = os de confiança `alta`; "para revisar" =
  `media` + `baixa`.
- **Vazios e corrigidos ficam à parte** (não entram em X nem em Y): um campo
  obrigatório vazio não derruba a proporção, e um corrigido não conta como "alta
  confiança" só porque o usuário mexeu. Opcional vazio fica fora de tudo.
- **Modo IA** (`confiancas == {}`): sem chips de confiança nem "X de Y"; nota
  "Confiança por campo indisponível no modo IA"; vazios e corrigidos continuam
  aparecendo e a edição continua funcionando.

**Banner de avisos:** largura total no topo do card de resultado
(`role="status"`, "Atenção"). Lista `resultado.avisos` (uma linha por aviso);
cai para a string `aviso` se `avisos` não vier (compatibilidade). Backend:
`ExtractionResult.avisos: list[str]` foi adicionado, `aviso` foi mantido (mesmo
conteúdo juntado por espaço).

**Bug real achado na validação (reentrância).** Trocar o editor que está com
foco (`replaceWith`) faz o Chrome disparar `blur` **síncrono** no meio da
troca; o handler de `blur` chamava `finalizarEdicao` de novo, que redesenhava o
campo por dentro do redesenho, e o `replaceWith` externo lançava exceção — o
foco nunca voltava ao botão depois do Enter (o resumo saía certo só por
acaso). Corrigido com uma trava (`campo.finalizando`). Lição: erro de JS no
meio de um handler pode deixar a tela "quase certa"; a validação captura
`pageerror`.

**Como é validado:** testes e2e versionados em `backend/tests/e2e/`
(Playwright num Chromium real, 30 testes). PDFs **fictícios** gerados por
PyMuPDF (`helpers.gerar_pdfs`) — nunca documento real. Cobrem: chips e contagem
do resumo conferidos contra o JSON da resposta por uma conta independente em
Python, banner com 2 avisos, campos abertos/fechados por estado, clique e teclado
(Tab/Enter/Esc), correção refletida no JSON bruto e no `.xlsx` baixado (lido com
`openpyxl`), `valor_total` numérico, mudança de tipo, a exibição formatada (itens
em pt-BR, tipo amigável, chave em blocos — sempre conferindo que o valor interno
não mudou), modo IA (resposta simulada), nomes acessíveis (`aria-labelledby`) e
ids únicos, contraste WCAG AA computado dos estilos reais, celular de 390px sem
rolagem horizontal. **Todo teste falha se houver erro de console/JS** (fixture
`tela`): um erro no meio de um handler já deixou a tela "quase certa". Os testes
foram conferidos por mutação (desligar a formatação pt-BR, o rótulo do tipo ou a
trava de reentrância faz o teste correspondente falhar).

`docs/screenshot.png` (README) é gerado por `tests/e2e/gerar_screenshot.py` a
partir dos mesmos PDFs sintéticos — nunca de um documento real (uma captura de
tela com dado real já foi um vetor de vazamento neste projeto).

## Excel (etapas 6 a 10)

`POST /export-excel` recebe uma **lista** de documentos (a tela envia 1; a
estrutura já comporta lote) e devolve um `.xlsx` com **5 abas**: o **Relatório**
(1ª, a que abre; leitura) e 4 abas de dados. A estrutura é estável pra Power
Query/fórmulas: as 4 de dados existem sempre, mas **ficam ocultas (não
removidas) quando não têm nenhuma linha** (`sheet_state = "hidden"`, o nome da
aba continua valendo):

```
{ "documentos": [ { "arquivo": "x.pdf", "resultado": <ExtractionResult>,
                    "corrigidos": { "<chave>": <valor ORIGINAL extraído> } } ] }
```

| Aba | Colunas | Linhas |
|---|---|---|
| **Relatório** | (não tabular) blocos por documento: emitente, destinatário, valores, itens, avisos; rodapé com a legenda só das cores que aparecem | — |
| **Documentos** | ID, Arquivo, Documento, Tipo, Número, Data de emissão, Data de vencimento, Emissor, Emissor CNPJ/CPF, Destinatário, Destinatário CNPJ/CPF, Valor total, Confiança geral | 1 por documento |
| **Itens** | ID, Documento, Descrição, Quantidade, Valor unitário, Valor total | todos os itens |
| **Campos adicionais** | ID, Documento, Campo, Valor, Confiança | 1 por campo adicional |
| **Avisos** | ID, Documento, Aviso | 1 por aviso |

A aba "Documentos" se chamava "Resumo" até a etapa 9; identificadores internos
(`CABECALHOS_RESUMO`, `_linha_resumo`...) mantêm o nome antigo de propósito.

- **`ID`** (1, 2, 3… pela ordem do lote) está em todas as abas como **chave de
  ligação** — tipo + número pode colidir entre emissores diferentes.
  **`Documento`** é a descrição legível (`"Nota fiscal 000012345"`; sem número:
  `"Nota fiscal — arquivo.pdf"`), só pra leitura humana. **`Arquivo`** (só no
  Resumo) é o PDF de origem. A coluna `Tipo` do Resumo mantém o valor
  **interno** (`nota_fiscal`, bom pra filtrar), diferente do rótulo amigável
  da tela.
- **Emissor/Destinatário separados do documento fiscal:** `"NOME (CNPJ x)"`
  vira `Emissor = NOME` e `Emissor CNPJ/CPF = x` (idem destinatário), pra
  filtrar e cruzar. Pega o **último** parêntese, então nomes com parênteses
  próprios funcionam.
- **`corrigidos`** existe porque o backend só recebe o valor *atual*: a tela
  manda quais campos o usuário mudou e o valor original. A chave é o mesmo
  espaço de nomes de `confiancas` (`emissor`, `"Chave de Acesso"`...). A tela
  guarda `original` e `atual` por campo, então o custo é pequeno
  (`coletarCorrigidos` em `script.js`).

**Formatação (abas de dados; etapas 7 e 8):** cabeçalho escuro, zebra, bordas,
alinhamento por tipo, cabeçalho **e coluna ID** congelados (`B2`), largura
ajustada ao conteúdo (mín. 8, teto 60, com quebra de linha nos textos), cor da
guia, área de impressão em paisagem. Cada aba com dados vira uma **Tabela
nomeada** do Excel (`Documentos`, `Itens`, `CamposAdicionais`, `Avisos`) no
lugar do auto_filter solto, com o estilo próprio da Tabela **desligado** (o
listrado embutido brigaria com a zebra e os destaques). Metadados do arquivo:
`creator` é o nome do aplicativo, nunca uma pessoa. **Não existe linha de
título mesclada acima do cabeçalho** (cogitada e descartada: Power Query/pandas
assumem "linha 1 = cabeçalho" e leriam o título como cabeçalho).

**Bug real: o Excel pedia "reparar" o arquivo (etapa 8).** Passava em todos os
testes com openpyxl/pandas (as duas bibliotecas são tolerantes; o Excel não).
Duas causas: (1) **Tabela com `ref` só no cabeçalho** (aba sem dados); (2)
**célula `t="inlineStr"` sem `<is>`** (string vazia tipada como texto). Hoje:
só cria Tabela se há linhas (`if linhas:`) e célula vazia fica realmente vazia.
Detectado por `tests/ooxml.py`, que lê o `.xlsx` como **XML bruto** (não pelo
openpyxl) e checa regras que o Excel aplica além do schema. Três camadas de
validação: (a) o verificador próprio, sempre; (b) o **Open XML SDK** da
Microsoft (`--ooxml-sdk`, só Windows, opt-in; **não** pegou nenhum dos dois
bugs acima, só um falso positivo documentado sobre a ordem dos filhos de
`<font>`); (c) abrir no Excel de verdade (manual, feita pelo usuário a cada
etapa).

**Relatório (etapa 9):** aba de LEITURA, não tabular (retrato, sem Tabela, sem
proteção): o mesmo dado em blocos por documento. Campo obrigatório vazio aparece
como "não encontrado"; opcional vazio some. **Valor total** usa fonte grande em
célula mesclada, e o Excel **não calcula a altura da linha** de célula mesclada
com fonte grande (cortava o texto): a altura é explícita
(`ALTURA_LINHA_VALOR_PRINCIPAL`). A chave de acesso aparece em blocos de 4 dígitos
também em Campos adicionais (`_formatar_chave_acesso`).

**Regras que evitam bugs reais de planilha** (docstring de `excel_exporter.py`):
- **Formato monetário:** o código guardado no arquivo é `"R$" #,##0.00`
  (notação en-US); o Excel/LibreOffice **localizam na exibição** (pt-BR mostra
  `R$ 1.234,56`). Escrever `#.##0,00` literalmente quebraria o formato. Valor
  unitário usa `"R$" #,##0.00##` (preço pode ter até 4 casas).
- **Só valor monetário vira número:** `valor_total` e os valores dos itens
  (quando já são `float`) e 3 campos adicionais fixos (`CAMPOS_MONETARIOS`:
  Valor do Documento, Desconto, Valor a Pagar). **Todo o resto é texto** — a
  chave de acesso (44 dígitos; o Excel só guarda 15 de precisão), Nosso
  Número, CFOP, linha digitável, número do documento (zeros à esquerda).
  Nunca "parece número → converte". String em campo numérico (a conversão já
  tinha falhado) fica texto.
- **Datas** viram data de verdade (`dd/mm/aaaa`), convertidas de `15/04/2026`
  e de ISO (o modo IA pode devolver); dia inexistente ou texto livre fica
  texto, sem quebrar.
- **Texto é sempre texto — injeção de fórmula.** O `openpyxl` trata string
  que começa com `=` como **fórmula**, e o texto vem de PDF (não confiável):
  um emissor `=HYPERLINK(...)` viraria fórmula executável na máquina de quem
  abre. Forçamos `data_type = "s"` (o teste verifica que não existe nenhum
  `<f>` no XML e que, sem isso, o teste falha). Caracteres de controle (que o
  `openpyxl` rejeita) são removidos; texto acima de 32.767 caracteres é
  truncado.

**Destaques (mesmos tons da tela):** confiança **média** = fundo amarelo claro
(`FDF3E0`), **baixa** = vermelho claro (`FBECEB`), **corrigido pelo usuário** =
verde-petróleo claro (`E7F0EF`) + itálico. Cada célula destacada leva um
**comentário** com o texto (não depende só de cor): "Confiança média/baixa…" ou
"Corrigido pelo usuário. Valor extraído: X" (também quando o usuário
**esvaziou** o campo). Em Campos adicionais a coluna `Confiança` traz
Alta/Média/Baixa/Corrigido em texto. Tabela de itens com confiança média/baixa
pinta as linhas, com um comentário só (na 1ª linha de cada documento).

**Confiança geral** (`"6 de 7 alta"`, coluna do Resumo) usa a mesma regra da
tela — `app/confianca.py` (Y = campos com valor e confiança, sem o tipo;
corrigidos e vazios ficam fora; sem `confiancas` no modo IA → `"—"`). A regra
existe em **JS** (tela) e em **Python** (Excel); um teste e2e compara os dois
números (`test_excel_*`), então não divergem sem falhar.

**Lote:** a estrutura já é de lote (lista, `ID`, linhas por documento; testes
com 2 e 4 documentos). Falta a **interface** de lote (várias PDFs de uma vez,
estado de edição por documento, fila/progresso, erro por arquivo); a extração
continua por arquivo (o lote seriam N chamadas a `/extract-document`).

**Recursos nativos do Excel (etapa 10)** — cada um com o motivo, pra não
reabrir a discussão:
- **Total do lote** (`_escrever_linha_total_documentos`): linha `=SUM` logo abaixo
  da Tabela de Documentos (fora da `ref`), **só com 2+ documentos**. Rótulo "Total
  do lote" na coluna antes do Valor total, **sem mesclar** (mesclar A:K
  atravessaria a divisa do congelamento). Mesmo padrão da linha de total dos
  Itens; quem lê a aba crua vê uma linha a mais com ID vazio (`_linhas_sem_total`
  nos testes). Entra na área de impressão.
- **Gráfico** (`_grafico_valor_por_documento`): barras verticais de Valor total por
  Documento, abaixo do total (coluna **B**, não A: a A está congelada), **2+
  documentos com valor numérico**. Lê direto das células (editar/filtrar a
  planilha atualiza o gráfico). O Excel plota **texto como zero**, então o
  documento com `valor_total` em texto fica de fora via união de intervalos
  (`('Documentos'!$L$2,'Documentos'!$L$4:$L$5)`); no caso normal a referência é
  um intervalo simples. Fora da área de impressão (uma quebra de página cortaria
  o gráfico). **Pegadinhas do openpyxl 3.1** tratadas: não escreve
  `<c:delete val="0"/>` nos eixos (o Excel 365 os **esconde**), categorias de
  texto saem `numRef` (usa-se `strRef`), título sem `overlay` explícito, cantos
  arredondados por padrão, flags de `dLbls` omitidas (o Excel pode ligá-las).
- **Listas suspensas** (`_lista_suspensa`): só em **Documentos!Tipo** (valores
  **internos**, `nota_fiscal`, porque é o que a coluna guarda) e **Campos
  adicionais!Confiança** (Alta/Média/Baixa/Corrigido/"—": o domínio real da coluna;
  senão o arquivo violaria a própria regra). **"Confiança geral" não** (é uma
  proporção em texto, "6 de 7 alta"). Limites que fazem o Excel pedir reparo: lista
  literal ≤ 255 caracteres, título do erro ≤ 32, mensagem ≤ 255; `showDropDown=True`
  do openpyxl **esconde** a seta (semântica invertida). A validação só barra
  entrada NOVA: um tipo inventado pelo modo IA que já esteja na célula continua lá.
- **Formatação condicional nativa só na coluna Confiança de Campos adicionais**
  (`_colorir_coluna_confianca`), onde o texto da célula É o valor (a cor acompanha
  uma edição). **Trocar todos os destaques por regras foi descartado:** nas demais
  células (Emissor, Valor total, Itens) a cor documenta a **proveniência** da
  extração ("veio com confiança baixa") e não pode depender do valor atual;
  reescrever o campo apagaria o rastro que o comentário da célula preserva, e a
  linha nem carrega o dado de confiança que uma regra poderia ler. Nessa coluna
  não há fundo fixo (senão editar "Baixa" → "Alta" deixaria a célula vermelha).
- **Proteção de planilha sem senha** (`_proteger_planilha`) nas 4 abas de dados
  (Relatório fica de fora): cabeçalho, linhas de total e tudo fora da tabela
  **bloqueados**; células de dado **desbloqueadas** (`Protection(locked=False)`
  em `_escrever_aba`). Semântica **invertida** no formato: `autoFilter="0"` e
  `sort="0"` significam **liberado**. **Prioridade do usuário: filtro e
  ordenação da Tabela valem mais que a trava** (se a proteção os atrapalhasse,
  sai a proteção). Testado no Excel real: funcionam com a aba protegida. Gráfico
  fica editável (`objects` desligado). Consequência: a Tabela não cresce com a aba
  protegida (Revisão > Desproteger planilha).

**Como é validado:** `tests/test_excel_exporter.py` (lê o `.xlsx` de volta com
`openpyxl`: abas, formatos, datas, texto vs número, fórmula, fundos e
comentários, lote, endpoint, e os recursos da etapa 10, com invariantes tipo
"nenhuma fórmula fica desbloqueada" e "todo valor escrito nas colunas validadas
pertence à própria lista"), `tests/test_ooxml.py` (o verificador `ooxml.py` sobre
9 cenários, mais testes de que ele **acusa** arquivo quebrado: sem isso seria
decorativo; regras cobrem Tabelas, listas, regras condicionais, gráfico e
proteção; o SDK da Microsoft roda nos mesmos 9 cenários com `--ooxml-sdk`),
`tests/test_confianca.py` e os `test_excel_*` do e2e (Excel comparado com a
tela). **Conferido por mutação** a cada etapa (ex.: deixar o `openpyxl` inferir
fórmula, chave de acesso como número, gráfico sem `delete=False`, categorias como
`numRef`, filtro bloqueado pela proteção, fundo fixo de volta na Confiança): 33 de
34 mutantes da etapa 10 pegos; o que sobrou é equivalente (a função do gráfico já
exige 2+ documentos). Testes que adulteram o XML conferem que a adulteração
achou o trecho (`assert adulterado != conteudo`), senão o teste vira decorativo
quando o formato muda. **Renderização:** LibreOffice não está disponível no
ambiente; a conferência visual no Excel é manual e é feita pelo usuário com
arquivos de amostra (`Documents\extrator-docs-capturas\etapaN\`) **antes** de
qualquer commit de mudança no Excel. `pandas` deixou de ser dependência (só o
exportador o usava).

## Robustez e observabilidade (diagnóstico de maturidade, itens 1 a 5)

Medido antes de corrigir, cada item com teste que reproduz o problema:

- **Event loop não pode bloquear** (`tests/test_concorrencia.py`): ler o PDF,
  extrair campos e montar o Excel são síncronos e levam segundos; chamados direto
  numa rota `async def` travavam o servidor inteiro (até o `/health`). Tudo passa
  por `fastapi.concurrency.run_in_threadpool` (`main._ler_pdf`, extração básica e
  IA, `/export-excel`, `/debug/*`).
- **Tetos de tamanho** (`tests/test_limites.py`): upload sempre limitado a
  `TAMANHO_MAXIMO_BYTES` (20 MB), inclusive nos endpoints de debug;
  `/export-excel` limita `documentos` (100), `itens` por documento
  (`TETO_ITENS` = 1000), `campos_adicionais` (`TETO_CAMPOS_ADICIONAIS` = 100),
  `avisos` (50) e o total de itens do lote (`TETO_ITENS_TOTAL_DO_LOTE` = 5000), via
  Pydantic. Os testes conferem os valores absolutos (comparar `TETO + 1` com o
  próprio `TETO` é tautologia: uma mutação que sobe o teto passaria).
- **CI** (`.github/workflows/tests.yml`): `pytest tests` a cada push na `main` e em
  todo PR, Python 3.12, com badge no README. **Não** instala Playwright nem roda
  `--e2e`/`--ooxml-sdk` (esses se pulam sozinhos). Cuidado com o YAML 1.1: `on:`
  sem aspas vira booleano; está `"on":`.
- **Log estruturado** (`app/logging_config.py`, `main._logar_extracao`): só
  `logging`, sem Sentry. Configura só o logger `"app"` (`propagate = False`,
  `LOG_LEVEL` do ambiente), com `request_id` num `contextvars.ContextVar` que
  qualquer logger do projeto enxerga; o mesmo id volta no cabeçalho
  **`X-Request-ID`** (`RequestIdMiddleware`). Uma linha JSON por extração:
  arquivo, duração, páginas, origem do texto, modo, extrator escolhido, campos de
  confiança média/baixa e obrigatórios vazios. `caplog.at_level(logger=X)` dos
  testes continua funcionando (anexa direto ao logger nomeado).
- **Mensagens de erro por causa** (`tests/test_erros_de_leitura.py`): PDF com
  senha levanta `pdf_extractor.PDFProtegidoPorSenha` (o pdfplumber embrulha
  `PDFPasswordIncorrect` em `PdfminerException.args[0]`, não em `__cause__`) e
  vira 400 "protegido por senha"; PDF **sem nenhuma página** (`numero_paginas ==
  0`) ganha aviso próprio em vez de "parece ser uma imagem escaneada"; arquivo
  realmente corrompido continua "corrompido", e PDF escaneado de verdade continua
  com o aviso de imagem (testes de falso positivo).

## Limite de uso e desempenho (itens 6 a 8 do diagnóstico)

**Limite por IP** (`app/rate_limit.py`, `tests/test_rate_limit.py`,
`tests/e2e/test_interface_limite.py`):
- **10 por minuto, por IP, por grupo**: "extracao" (`/extract-document` e os 3
  `/debug/*`) e "exportacao" (`/export-excel`), cotas separadas. Critério: um
  humano faz 3–4 extrações por minuto; 10 dá folga pra quem grava vídeo e limita
  um script a 600/h. `RATE_LIMIT_POR_MINUTO` (0 desliga; inválido volta ao
  padrão). Um teste exige que **todo endpoint POST** esteja em `GRUPOS_POR_ROTA`.
- **Middleware ASGI**, não `Depends`: o 429 sai antes de o upload (até 20 MB)
  ser lido. Registrado **por dentro** do `RequestIdMiddleware` (o 429 tem
  `X-Request-ID`). Janela deslizante em memória; recusa não é registrada;
  memória limitada a 10 mil IPs (LRU); IPv6 conta por /64. Resposta JSON com
  `detail` em português ("Você atingiu o limite de 10 extrações por minuto da
  demonstração pública. Aguarde N segundos…"), `Retry-After`, `X-RateLimit-*`.
  A tela mostra o `detail` do 429 **sem** "Falha ao…" (não é falha); a
  exportação passou a mostrar o `detail` também (antes só "Erro N").
- **IP do cliente**: por padrão o da conexão. No Render, o primeiro IP do
  `X-Forwarded-For`, ligado automaticamente por `RENDER=true` (o Render sempre
  define): um serviço sem Blueprint ignora os envVars do `render.yaml`, e sem
  isso todos dividiriam a cota do IP do proxy. `CONFIAR_X_FORWARDED_FOR`
  explícita tem prioridade. **Conferido na demo (2026-09-26):** o proxy do
  Render **anexa** ao cabeçalho, não substitui: visitante normal é contado
  pelo IP real, mas forjar o cabeçalho dá cota nova (limitação documentada no
  README; barrar script de verdade pede limite na borda).
- Testes que sobem uvicorn: `helpers.iniciar_servidor()` sobe com o limite
  **desligado** (as suítes fazem dezenas de POSTs por minuto) e aceita
  `env_extra`. O uvicorn, por padrão, reescreve o IP do cliente a partir do
  `X-Forwarded-For` quando a conexão vem de 127.0.0.1 (`--forwarded-allow-ips`):
  os testes com o cabeçalho usam `FORWARDED_ALLOW_IPS` apontando pra outro
  endereço, pra testar a nossa lógica e não a do uvicorn. A fixture `tela` do
  e2e aceita `status_http_esperados` (o Chromium loga 429/500 provocados como
  erro de console).

**Desempenho** (README, seção "Desempenho"; medido em 2026-09-26, mediana de 5):
extração de DANFE sintética 1/10/50/100 páginas = 0,027/0,289/1,44/2,89 s
(~29 ms/página, linear); ponta a ponta via HTTP soma ~20 ms; OCR ~1,7 s/página;
Excel 23 ms no uso normal. **Achado:** no teto da API (5000 itens) o Excel leva
**8,5 s**, superlinear, por causa do `merge_cells` do openpyxl na aba Relatório
(`MultiCellRange.add` confere cada mesclagem contra todas as anteriores: 12,6
milhões de comparações). Veio da etapa 9 (2,9 s → 8,2 s no teto); a etapa 10
somou ~3%. **Não corrigido** (mudaria o layout aprovado do Relatório):
registrado como próximo passo.

**Decisões descartadas** estão no fim de "Decisões técnicas" do README (Render ×
túnel temporário, fila de tarefas, xlsxwriter, login, formatação condicional,
linha de título). Na hora do deploy **não** houve comparação formal com
Railway/Fly.io: o README não afirma uma.

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
real, destinatário fictício). **O texto bruto (`danfe_real_anonimizado.txt`)
é um COMPOSTO de trechos reais, não as 84 linhas que o extrator recebe**:
vem de `/debug/extractor-input` (linhas 00-02, 14-15, 38, 39-41, 49-50,
51-52, 53-56, 75-76 — texto girado, emitente, IE/CNPJ, destinatário colado,
grade de totais, cabeçalho e produto) mais um trecho anterior
(endereço/chave) com a adjacência entre blocos ASSUMIDA, e sem linha de
número da nota. **Cuidado com a adjacência assumida**: foi ela que escondeu o
bug do CNPJ (na fixture o CNPJ ficava a ~7 linhas do nome, no real a 10) —
por isso a distância real é coberta por teste com preenchimento sintético. O
docstring do teste diz isso; testes que montam texto à mão estão marcados
`SINTETICO`. **Deve ser
substituído pelas 84 linhas completas** quando disponíveis. Cobre também os
bugs de tabela documentados (texto vertical, valores colados, corte por
informações complementares, fragmentos do cabeçalho) e a validação de soma.

`tests/test_debug_endpoints.py` garante que `/extract-document`,
`/debug/extract-text` e `/debug/extractor-input` veem exatamente as mesmas
linhas (ver "Diagnosticando um PDF real"). Chama as funções dos endpoints
direto com um `UploadFile` montado na mão, porque `httpx` (TestClient) não
está nas dependências.

**Testes e2e (interface, opt-in).** `tests/e2e/` roda a interface num Chromium
real e **não** entra no `pytest tests` normal (que pula os 30 e2e sem subir
servidor nem navegador, e não exige Playwright). Dependências separadas:

```bash
cd backend
pip install -r requirements-dev.txt      # inclui requirements.txt + playwright
playwright install chromium
pytest tests/e2e --e2e -v
```

O servidor de teste sobe numa porta livre, com `ANTHROPIC_API_KEY` vazia (sempre
modo básico, mesmo que a máquina tenha chave). Se o Chromium não estiver
instalado, os testes são pulados com a instrução em vez de falhar.

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

### Diagnosticando um PDF real

Regra do projeto: **nunca ajustar um extrator sem ter visto o texto exato
que ele recebe.** Já custou 3 rodadas de correção por suposição (boleto e
DANFE). O caminho:

1. Rodar o PDF em `POST /debug/extractor-input` (`/docs` → "Try it out").
   Devolve, por página, as linhas **numeradas** exatamente como o extrator
   as vê (já SEM o texto girado descartado na leitura — o campo
   `caracteres_girados_descartados` diz quantos caracteres foram); qual
   extrator foi escolhido e as pontuações de cada um; o que ele
   extraiu (`resultado.documento`, `confiancas`, `avisos`); e, pra DANFE,
   `tabela_itens`: por página, se achou o cabeçalho da tabela, as âncoras de
   coluna, e **cada linha examinada com a decisão** (`item`,
   `continuacao`, `pulada: fragmento do cabecalho...`, `fim: ...`) — é o que responde "por que `itens` saiu
   vazio/curto".
2. Colar a saída anonimizada (trocar nome/CPF/endereço/telefone/e-mail por
   fictícios **sem juntar nem separar linhas** — a estrutura de linhas é o
   que importa) e reconstruir a fixture a partir dela.
3. Só então ajustar o extrator.

Todos os endpoints que leem PDF passam por `main._ler_pdf` e o contexto do
extrator é montado por `basic_extractor.montar_contexto` — não há "outro
jeito de ler" entre `/extract-document`, `/debug/extract-text` e
`/debug/extractor-input` (`tests/test_debug_endpoints.py` garante).
`exibicao_fiel_ao_texto_do_extrator` na resposta confere que a exibição por
página bate com o texto que o extrator recebe.

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
- (Etapa 5 — indicador de confiança na interface — **implementada**, ver
  "Interface: confiança e edição"; falta só a validação manual do usuário no
  navegador. A anotação de que `itens` da DANFE deve ser `"alta"` quando a
  soma fecha também já está implementada, ver "Tabela de itens".)
- Excel (etapas 6 a 10) **implementado** — ver "Excel".
- Interface de **lote** (várias PDFs de uma vez): o backend/Excel já comporta;
  falta a tela (estado por documento, fila/progresso, erro por arquivo).
- Correção de confusão de OCR cobre só CNPJ/CPF/valores — não
  numero_documento/Nosso Número nem texto livre (emissor/destinatario).
- Tabela de itens da DANFE não funciona com texto de origem OCR (só
  digital, que tem posição confiável de palavra) — limitação conhecida,
  não um bug.

Plano original com todas as etapas em
`C:\Users\User\.claude\plans\jazzy-foraging-harbor.md`.
