# Extrator Inteligente de Documentos

MVP de um extrator de documentos: você sobe um PDF (nota fiscal, pedido de
compra ou relatório), o sistema extrai o texto, estrutura os dados em JSON,
exibe numa tabela e permite baixar o resultado em Excel.

![Tela do extrator com resultado extraído](docs/screenshot.png)

## Stack

- **Backend:** Python 3, FastAPI
- **Extração de PDF:** [pdfplumber](https://github.com/jsvine/pdfplumber)
- **Exportação para Excel:** pandas + openpyxl
- **Estruturação dos dados:** ver "Modo básico vs modo IA" abaixo
- **Frontend:** HTML/CSS/JS puro, sem framework, servido pelo próprio FastAPI

## Modo básico vs modo IA

A IA é um **upgrade opcional**, não um requisito para o sistema funcionar:

- **Modo básico** (padrão, sem nenhuma configuração): `pdfplumber` extrai o
  texto do PDF e uma extração por regex/heurística captura os campos óbvios
  — datas, valores monetários, número do documento, emissor/destinatário.
  Reconstruir a tabela de itens de forma confiável por regex não é viável a
  partir de texto de PDF sem estrutura, então nesse modo `itens` normalmente
  fica vazio. É uma limitação conhecida, não um bug.
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
- **PDF escaneado (sem texto extraível):** detectado antes de tentar
  qualquer extração, retornando um aviso claro em vez de um resultado vazio
  sem explicação. OCR está fora do escopo deste MVP (ver `TODO` em
  `backend/app/pdf_extractor.py`).
- **Frontend servido pelo próprio FastAPI** (`StaticFiles`, um único
  processo em `localhost:8000`) em vez de dois servidores separados —
  evita configurar CORS e simplifica rodar o projeto localmente.

## Limitações conhecidas / próximos passos

- Modo básico não extrai itens de tabela (linhas de nota fiscal/pedido) —
  apenas metadados de topo. O modo IA cobre esse caso.
- Sem OCR: PDFs escaneados (imagem, sem camada de texto) não são
  processados neste MVP.
- Heurísticas de regex do modo básico foram desenhadas para o formato de
  documento comercial brasileiro comum (datas `dd/mm/aaaa`, valores em
  `R$`) — podem não pegar todos os formatos.

## Status

MVP funcional: upload → extração (básica ou via IA) → tabela → download em
Excel. Estrutura completa e decisões de arquitetura em `CLAUDE.md`.
