const form = document.getElementById("upload-form");
const fileInput = document.getElementById("file-input");
const statusEl = document.getElementById("status");
const resultadoEl = document.getElementById("resultado");
const badgeModoEl = document.getElementById("badge-modo");
const avisoEl = document.getElementById("aviso");
const jsonBrutoEl = document.getElementById("json-bruto");
const tabelaMetadadosEl = document.getElementById("tabela-metadados");
const tabelaItensEl = document.getElementById("tabela-itens");
const btnExcelEl = document.getElementById("btn-excel");
const btnExtrairEl = document.getElementById("btn-extrair");

let ultimoResultado = null;

const ROTULOS_METADADOS = {
  tipo_documento: "Tipo de documento",
  numero_documento: "Número do documento",
  data_emissao: "Data de emissão",
  emissor: "Emissor",
  destinatario: "Destinatário",
  valor_total: "Valor total",
};

form.addEventListener("submit", async (event) => {
  event.preventDefault();

  const file = fileInput.files[0];
  if (!file) return;

  statusEl.textContent = "Extraindo...";
  resultadoEl.hidden = true;
  btnExtrairEl.disabled = true;

  const formData = new FormData();
  formData.append("file", file);

  try {
    const response = await fetch("/extract-document", {
      method: "POST",
      body: formData,
    });

    if (!response.ok) {
      const erro = await response.json().catch(() => null);
      throw new Error(erro?.detail || `Erro ${response.status}`);
    }

    const resultado = await response.json();
    mostrarResultado(resultado);
    statusEl.textContent = "";
  } catch (erro) {
    statusEl.textContent = `Falha ao extrair: ${erro.message}`;
  } finally {
    btnExtrairEl.disabled = false;
  }
});

btnExcelEl.addEventListener("click", async () => {
  if (!ultimoResultado) return;

  btnExcelEl.disabled = true;
  try {
    const response = await fetch("/export-excel", {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify(ultimoResultado),
    });

    if (!response.ok) {
      throw new Error(`Erro ${response.status}`);
    }

    const blob = await response.blob();
    const url = URL.createObjectURL(blob);
    const link = document.createElement("a");
    link.href = url;
    link.download = "documento_extraido.xlsx";
    link.click();
    URL.revokeObjectURL(url);
  } catch (erro) {
    statusEl.textContent = `Falha ao gerar Excel: ${erro.message}`;
  } finally {
    btnExcelEl.disabled = false;
  }
});

function mostrarResultado(resultado) {
  ultimoResultado = resultado;

  badgeModoEl.textContent =
    resultado.modo_extracao === "ia" ? "Modo: IA" : "Modo: Básico";
  badgeModoEl.className = `badge badge-${resultado.modo_extracao}`;

  avisoEl.textContent = resultado.aviso || "";

  preencherTabelaMetadados(resultado.documento);
  preencherTabelaItens(resultado.documento.itens);

  jsonBrutoEl.textContent = JSON.stringify(resultado, null, 2);
  resultadoEl.hidden = false;
}

function limparTabela(tabela) {
  tabela.innerHTML = "";
}

function criarLinha(celulas) {
  const tr = document.createElement("tr");
  for (const texto of celulas) {
    const td = document.createElement("td");
    td.textContent = texto ?? "—";
    tr.appendChild(td);
  }
  return tr;
}

function preencherTabelaMetadados(documento) {
  limparTabela(tabelaMetadadosEl);
  for (const [chave, rotulo] of Object.entries(ROTULOS_METADADOS)) {
    tabelaMetadadosEl.appendChild(criarLinha([rotulo, documento[chave]]));
  }
  for (const extra of documento.campos_adicionais || []) {
    tabelaMetadadosEl.appendChild(criarLinha([extra.campo, extra.valor]));
  }
}

function preencherTabelaItens(itens) {
  limparTabela(tabelaItensEl);

  const cabecalho = document.createElement("tr");
  for (const rotulo of ["Descrição", "Quantidade", "Valor unitário", "Valor total"]) {
    const th = document.createElement("th");
    th.textContent = rotulo;
    cabecalho.appendChild(th);
  }
  tabelaItensEl.appendChild(cabecalho);

  if (!itens || itens.length === 0) {
    const tr = document.createElement("tr");
    const td = document.createElement("td");
    td.colSpan = 4;
    td.textContent = "Nenhum item identificado.";
    tr.appendChild(td);
    tabelaItensEl.appendChild(tr);
    return;
  }

  for (const item of itens) {
    tabelaItensEl.appendChild(
      criarLinha([item.descricao, item.quantidade, item.valor_unitario, item.valor_total])
    );
  }
}
