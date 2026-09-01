# extrator-docs

Extrator inteligente de documentos: usuário sobe um PDF (nota fiscal, pedido de
compra ou relatório), o sistema extrai o texto, estrutura os dados em JSON,
exibe numa tabela e permite baixar em Excel. Projeto de portfólio (estudante
de Engenharia da Computação).

## Stack

- Backend: Python 3 + FastAPI
- Extração de PDF: pdfplumber
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
  pdf_extractor.py    # texto do PDF via pdfplumber + detecção de PDF escaneado
  basic_extractor.py  # extração heurística por regex (modo "basico")
  llm_extractor.py    # extração via Claude (modo "ia", opcional)
  excel_exporter.py   # ExtractionResult -> .xlsx
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
  também captura linha digitável, parcela, agência/código do beneficiário
  e os três valores (documento, desconto, valor a pagar) em
  `campos_adicionais`; `valor_total` prioriza "Valor a Pagar" quando existe.
  `itens` normalmente fica vazio nesse modo — reconstruir tabelas de itens
  por regex não é confiável. Essa é uma limitação conhecida do MVP, não um
  bug.
  - `data_emissao` e `data_vencimento` são campos separados — nunca se
    misturam porque cada um busca por um rótulo diferente
    (`ROTULOS_DATA_EMISSAO` vs `ROTULOS_VENCIMENTO`) e ambos exigem que o
    candidato tenha uma data reconhecível (`_parece_data`) antes de aceitar.
  - Regra de negócio importante: `numero_documento` só aceita valores
    curtos/poucos dígitos quando vêm de um rótulo conhecido (confiável);
    sem rótulo, o fallback por regex solto exige no mínimo 4 dígitos — é
    o que evita recapturar o bug do "02".
  - `_localizar_rotulo` busca por **prioridade do rótulo** (ordem da lista
    `ROTULOS_*`), não por posição no documento: primeiro procura o rótulo
    mais específico (ex: "Nosso Número") no documento inteiro; só cai para
    um rótulo mais genérico (ex: "Número") se o específico não aparecer em
    lugar nenhum. Existe porque um rótulo genérico que aparece mais cedo na
    página (ex: "Número do Banco") vencia um rótulo específico e confiável
    que aparecia mais tarde.
  - Candidatos a `emissor`/`destinatario` passam por `_parece_nome`
    (rejeita valores que são só dígitos/pontuação — carimbos de data/hora,
    números soltos). Sem essa validação, um "Sacado" seguido de um carimbo
    de data/hora no PDF virava destinatário.
  - "Nosso Número" às vezes vem como "02 / 10200000001-9" (prefixo de
    carteira / número real) — `_extrair_numero_documento` reconhece esse
    formato e usa a parte depois da barra.
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
- PDF escaneado (sem texto extraível) também não é erro: retorna aviso
  específico. OCR está fora do escopo do MVP (ver TODO em `pdf_extractor.py`).

`DocumentoExtraido` é o mesmo contrato de dados nos dois modos — é o que
permite tabela e exportação Excel funcionarem igual independente de como o
documento foi extraído.

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

- OCR para PDFs escaneados (`TODO` em `backend/app/pdf_extractor.py`).
- Modo básico não reconstrói tabela de itens (só o modo IA faz isso hoje).
- Testes automatizados (o projeto foi validado manualmente etapa por etapa
  durante a construção, mas não há suíte de testes ainda).

Plano original com todas as etapas em
`C:\Users\User\.claude\plans\jazzy-foraging-harbor.md`.
