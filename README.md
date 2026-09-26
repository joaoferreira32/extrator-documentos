# Extrator Inteligente de Documentos

[![Testes](https://github.com/joaoferreira32/extrator-documentos/actions/workflows/tests.yml/badge.svg)](https://github.com/joaoferreira32/extrator-documentos/actions/workflows/tests.yml)

**[Demo online](https://extrator-docs.onrender.com)** — hospedada no plano gratuito
do Render: se ninguém acessou nos últimos minutos, o serviço hiberna e a primeira
visita pode demorar ~50s para responder enquanto ele acorda.

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
- **Excel com aba de leitura e abas de dados:** um Relatório em blocos (por documento,
  com hierarquia visual) ao lado de Documentos/Itens/Campos adicionais/Avisos (Tabelas
  nomeadas, prontas pra Tabela Dinâmica/Power Query), valores em R$, datas reais e chave
  de 44 dígitos preservada. Num lote, ganha total e gráfico de valor por documento; as
  abas de dados têm listas suspensas e proteção sem senha (filtro continua funcionando).
- **Rápido e medido:** cerca de 29 ms por página de DANFE, em escala linear (ver
  [Desempenho](#desempenho)).
- **Segurança e privacidade:** proteção contra injeção de fórmula no Excel; limite de uso
  por IP na demo pública; o servidor não guarda os documentos; nenhum dado pessoal real
  no repositório nem no histórico do Git.
- **358 testes:** 319 rodam por padrão (unitários, regressão sobre um boleto e uma DANFE
  reais anonimizados, e um verificador que lê o `.xlsx` gerado como XML bruto pra pegar
  erro que o Excel rejeitaria mas o openpyxl não veria); mais 30 de interface num
  navegador real (Playwright) e 9 que validam o Excel contra o SDK oficial da Microsoft.
- **CI no GitHub Actions** a cada push, e log estruturado por extração (tempo, extrator
  escolhido, campos vazios/de baixa confiança) com id de correlação por requisição.
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

O botão **Baixar Excel** gera um `.xlsx` com 5 abas:

| Aba | Conteúdo |
|---|---|
| **Relatório** | a mesma informação em blocos, por documento (emitente, destinatário, valores, itens…), pra quem só quer *ler*; é a aba que abre, com a legenda das cores que aparecem de fato |
| **Documentos** | uma linha por documento (tipo, número, datas, emissor e destinatário com CNPJ/CPF em colunas próprias, valor total, confiança geral) |
| **Itens** | descrição, quantidade, valor unitário e total, com linha de total |
| **Campos adicionais** | chave de acesso, CFOP, nosso número, linha digitável… com a confiança de cada um |
| **Avisos** | o que a extração pediu para conferir |

As quatro últimas são Tabelas nomeadas do Excel, com cabeçalho e coluna `ID` congelados, e
o `ID` liga as abas. Confiança média/baixa e campos corrigidos ficam destacados, com
comentário. Aba de dados sem nenhuma linha fica oculta (não removida: quem usa Power Query
pelo nome da aba não quebra). A estrutura (lista de documentos, `ID`) comporta vários
documentos no mesmo arquivo.

Num lote (2 ou mais documentos), a aba Documentos ganha uma linha de **total do lote** e um
**gráfico de barras** do valor total por documento, abaixo da tabela. O gráfico só aparece
com 2 ou mais valores numéricos; um documento cujo valor virou texto fica de fora, em vez
de aparecer como uma barra zero.

Nas abas de dados também há **listas suspensas** (Tipo, e a Confiança dos campos
adicionais, as únicas colunas com um conjunto fechado de valores), **cor condicional
nativa** na coluna Confiança (acompanha o texto) e **proteção sem senha**: cabeçalho e
totais travados, dados livres, filtro e ordenação funcionando. Com a aba protegida a Tabela
não cresce; para acrescentar linhas, Revisão > Desproteger planilha.

## Desempenho

Medido na minha máquina de desenvolvimento (Windows, Python 3.12, mediana de 5 execuções),
com a DANFE sintética dos testes, no modo básico:

| Páginas | Tempo de extração | Por página |
|---:|---:|---:|
| 1 | 0,027 s | 27 ms |
| 10 | 0,289 s | 29 ms |
| 50 | 1,44 s | 29 ms |
| 100 | 2,89 s | 29 ms |

A escala é linear: cerca de 29 ms por página, para ler o texto e a posição das palavras e
reconstruir a tabela de itens. Pela API, de ponta a ponta, somam-se uns 20 ms fixos (1
página: 0,046 s). OCR é outra ordem de grandeza: cerca de 1,7 s por página escaneada.

**O gargalo real não era CPU, era concorrência.** As rotas eram `async`, mas chamavam
código síncrono: uma extração de 9 s travava o servidor inteiro, e um `/health` chamado ao
mesmo tempo chegou a esperar mais de 5 s. Com o trabalho pesado numa thread
(`run_in_threadpool`), o `/health` responde em menos de 0,5 s durante a mesma extração
(de 343 a 466 ms no pior momento; o resto é a disputa pelo GIL do Python). Uma requisição
leve não fica mais presa atrás de uma pesada.

Uma exceção, também medida: gerar o Excel leva 23 ms no uso normal, mas no teto que a API
aceita (5 documentos com 1000 itens cada) leva 8,5 s, e cresce mais que linearmente. A
causa são as células mescladas da aba Relatório: o openpyxl confere cada mesclagem contra
todas as anteriores. Está nos [próximos passos](#limitações-conhecidas-e-próximos-passos).

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

**Limite de uso:** 10 extrações e 10 exportações por minuto, por IP, valendo também
localmente. `RATE_LIMIT_POR_MINUTO=0` no `backend\.env` desliga o limite.

### Testes

```powershell
cd backend
.venv\Scripts\python.exe -m pytest tests -v
```

São 319 testes rodando por padrão, entre unitários e de regressão. Os de regressão usam
o texto bruto de um boleto e de trechos de uma DANFE reais, com os dados pessoais
trocados por fictícios. Eles existem porque os cenários que escrevi à mão não
reproduziam os bugs que apareciam no documento de verdade.

O teste de OCR real só roda se o Tesseract e o idioma português estiverem instalados.
Sem eles, ele é pulado em vez de falhar. Os outros testes de OCR usam mocks.

Toda vez que dou push (ou abro um PR), o [GitHub Actions](.github/workflows/tests.yml)
roda essa suíte sozinho — é o badge que aparece no topo deste README.

Há mais 30 testes de interface, que rodam num Chromium de verdade via Playwright com
PDFs fictícios gerados na hora. Eles são opcionais e têm dependências à parte:

```powershell
.venv\Scripts\python.exe -m pip install -r requirements-dev.txt
.venv\Scripts\python.exe -m playwright install chromium
.venv\Scripts\python.exe -m pytest tests\e2e --e2e -v
```

Tem também 9 testes que validam o `.xlsx` exportado contra o SDK oficial da Microsoft
(Open XML), além do verificador próprio (que já roda nos 319 de sempre). São Windows-only
e opcionais, porque baixam esse SDK na primeira vez:

```powershell
.venv\Scripts\python.exe -m pytest tests\test_ooxml.py --ooxml-sdk -v
```

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

```powershell
winget install --id UB-Mannheim.TesseractOCR -e
```

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

**Recursos nativos do Excel, verificados no Excel de verdade.** Total do lote, gráfico,
listas suspensas e proteção sem senha. Um detalhe que o schema não pega: o openpyxl 3.1
não escreve o elemento que mantém os eixos do gráfico visíveis, e o Excel 365 os esconde;
o verificador do `.xlsx` acusa. E a proteção só entrou depois de conferir que filtro e
ordenação da Tabela continuam funcionando com a aba protegida.

**Frontend servido pelo próprio FastAPI**, em HTML, CSS e JavaScript puros. Um processo
só em `localhost:8000`, sem CORS e sem CDN externo.

**Trabalho pesado fora do event loop.** Ler o PDF, extrair e montar o Excel são síncronos
e rodam numa thread (`run_in_threadpool`), senão uma rota `async` travaria o servidor
inteiro. Medido antes de corrigir; os números estão em [Desempenho](#desempenho).

**Log estruturado sem infraestrutura pesada.** Uma linha em JSON por extração (tempo,
extrator escolhido, campos vazios ou de baixa confiança) e um id por requisição
(`X-Request-ID` no cabeçalho da resposta, correlacionado à mesma linha do log) — só
`logging` da biblioteca padrão, sem Sentry nem serviço pago.

**Limite de uso por IP, simples de propósito.** 10 extrações e 10 exportações por minuto,
por IP. Uma pessoa usando a tela faz 3 ou 4 por minuto; 10 dá folga para quem demonstra
vários documentos em sequência e ainda segura um script em 600 por hora. Um middleware
responde 429 antes de ler o upload, com uma mensagem que a tela mostra como está ("Aguarde
N segundos…"). O contador fica em memória, sem Redis: zera quando a instância reinicia, o
que basta para uma demo numa instância só.

**O que ficou de fora, e por quê.**

- **Túnel temporário no lugar de hospedagem.** Para gravar um vídeo bastaria um Cloudflare
  Tunnel apontando para o meu computador, mas a URL morre quando o notebook desliga. O
  Render foi a hospedagem gratuita mais direta para um serviço Python com Uvicorn: deploy a
  partir do GitHub com um `render.yaml`, HTTPS e URL permanente. O custo é a hibernação.
- **Fila de tarefas (Celery, RQ) para a extração.** Uma DANFE comum sai em menos de 0,1 s,
  e o problema real, travar o servidor, foi resolvido com `run_in_threadpool`. Uma fila
  pediria um broker (Redis), um processo worker e polling na tela, para um caso que os
  limites de 20 MB e de uso por IP já contêm. Volta a fazer sentido com lotes grandes ou
  OCR de documentos longos.
- **xlsxwriter no lugar do openpyxl.** O xlsxwriter escreve mais rápido e gera gráficos
  mais completos por padrão, mas não lê arquivos. O openpyxl já estava no projeto (era o
  motor do pandas, que o exportador usava no começo) e serve tanto para escrever quanto
  para os testes lerem a planilha de volta. O custo apareceu: os eixos do gráfico que
  somem no Excel 365 e as mesclagens quadráticas (ver [Desempenho](#desempenho)). O
  verificador que lê o `.xlsx` como XML bruto existe justamente para não depender da
  biblioteca que escreveu o arquivo.
- **Login na demo.** O servidor não guarda documentos nem o que foi extraído (o log
  registra só metadados, como nome do arquivo e tempo), então não há dado de um usuário
  para proteger de outro. Login só atrapalharia quem quer testar em 30 segundos. O abuso é
  contido pelos limites de tamanho e de uso, e a faixa no topo pede que ninguém envie
  documento real.
- **Formatação condicional no lugar das cores fixas de confiança.** A cor de um campo
  (Emissor, Valor total…) documenta *como* ele foi extraído ("veio com confiança baixa"),
  não o valor que está na célula. Se a cor seguisse o valor, corrigir o campo na planilha
  apagaria o rastro de que ele foi revisado, que o comentário da célula preserva. Só a
  coluna Confiança dos campos adicionais usa regra condicional, porque ali o texto da
  célula é o próprio valor.
- **Linha de título acima do cabeçalho do Excel.** Power Query e pandas assumem que a
  linha 1 é o cabeçalho e leriam o título no lugar dos nomes das colunas.

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
- Gerar o Excel no teto que a API aceita (5000 itens) leva 8,5 s, por causa das células
  mescladas da aba Relatório. Próximo passo: mesclar sem a checagem quadrática do
  openpyxl, ou listar no Relatório só os primeiros itens de cada documento.
- O limite de uso fica em memória e não seria compartilhado entre várias instâncias. E o
  proxy do Render acrescenta ao cabeçalho `X-Forwarded-For` em vez de substituí-lo
  (conferido na demo): quem forja esse cabeçalho ganha uma cota nova. Navegador não faz
  isso, então o limite vale para o uso normal; barrar um script de verdade pede um limite
  na borda (Cloudflare).

## Autor

João Pedro Ferreira — [LinkedIn](https://www.linkedin.com/in/joaopedroferreira-2824d9324/)

## Status

Funcional de ponta a ponta e no ar na [demo](https://extrator-docs.onrender.com): upload,
extração básica ou via IA, confiança por campo, correção na tela e download em Excel.
Falta a interface para processar vários PDFs de uma vez.
