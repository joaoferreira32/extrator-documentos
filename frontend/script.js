const form = document.getElementById("upload-form");
const fileInput = document.getElementById("file-input");
const dropzoneEl = document.getElementById("dropzone");
const dropzoneFilenameEl = document.getElementById("dropzone-filename");
const statusCardEl = document.getElementById("status-card");
const statusTextEl = document.getElementById("status-text");
const errorCardEl = document.getElementById("error-card");
const errorTextEl = document.getElementById("error-text");
const resultadoEl = document.getElementById("resultado");
const badgeModoEl = document.getElementById("badge-modo");
const avisoEl = document.getElementById("aviso");
const jsonBrutoEl = document.getElementById("json-bruto");
const tabelaMetadadosEl = document.getElementById("tabela-metadados");
const tabelaItensEl = document.getElementById("tabela-itens");
const btnExcelEl = document.getElementById("btn-excel");
const btnExtrairEl = document.getElementById("btn-extrair");

let ultimoResultado = null;
let arquivoSelecionado = null;

const ROTULOS_METADADOS = {
  tipo_documento: "Tipo de documento",
  numero_documento: "Número do documento",
  data_emissao: "Data de emissão",
  data_vencimento: "Data de vencimento",
  emissor: "Emissor",
  destinatario: "Destinatário",
  valor_total: "Valor total",
};

// ---------- Seleção de arquivo (clique ou arrastar) ----------

fileInput.addEventListener("change", () => {
  if (fileInput.files[0]) {
    definirArquivo(fileInput.files[0]);
  }
});

["dragenter", "dragover"].forEach((evento) => {
  dropzoneEl.addEventListener(evento, (event) => {
    event.preventDefault();
    dropzoneEl.classList.add("is-dragover");
  });
});

["dragleave", "dragend"].forEach((evento) => {
  dropzoneEl.addEventListener(evento, () => {
    dropzoneEl.classList.remove("is-dragover");
  });
});

dropzoneEl.addEventListener("drop", (event) => {
  event.preventDefault();
  dropzoneEl.classList.remove("is-dragover");
  const arquivo = event.dataTransfer.files[0];
  if (arquivo) {
    definirArquivo(arquivo);
  }
});

function definirArquivo(arquivo) {
  arquivoSelecionado = arquivo;
  dropzoneFilenameEl.textContent = arquivo.name;
  dropzoneFilenameEl.hidden = false;
  esconderErro();
}

// ---------- Envio para extração ----------

form.addEventListener("submit", async (event) => {
  event.preventDefault();

  if (!arquivoSelecionado) {
    mostrarErro("Selecione um arquivo PDF antes de extrair.");
    return;
  }

  mostrarCarregando("Extraindo dados do documento…");
  resultadoEl.hidden = true;
  esconderErro();
  btnExtrairEl.disabled = true;

  const formData = new FormData();
  formData.append("file", arquivoSelecionado);

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
  } catch (erro) {
    mostrarErro(`Falha ao extrair: ${erro.message}`);
  } finally {
    esconderCarregando();
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
    mostrarErro(`Falha ao gerar Excel: ${erro.message}`);
  } finally {
    btnExcelEl.disabled = false;
  }
});

// ---------- Estados visuais ----------

function mostrarCarregando(texto) {
  statusTextEl.textContent = texto;
  statusCardEl.hidden = false;
}

function esconderCarregando() {
  statusCardEl.hidden = true;
}

function mostrarErro(mensagem) {
  errorTextEl.textContent = mensagem;
  errorCardEl.hidden = false;
}

function esconderErro() {
  errorCardEl.hidden = true;
}

function mostrarResultado(resultado) {
  ultimoResultado = resultado;

  badgeModoEl.textContent =
    resultado.modo_extracao === "ia" ? "Modo: IA" : "Modo: Básico";
  badgeModoEl.className = `badge badge-${resultado.modo_extracao}`;

  if (resultado.aviso) {
    avisoEl.textContent = resultado.aviso;
    avisoEl.hidden = false;
  } else {
    avisoEl.hidden = true;
  }

  preencherTabelaMetadados(resultado.documento);
  preencherTabelaItens(resultado.documento.itens);

  jsonBrutoEl.textContent = JSON.stringify(resultado, null, 2);
  resultadoEl.hidden = false;
}

function preencherTabelaMetadados(documento) {
  tabelaMetadadosEl.innerHTML = "";
  for (const [chave, rotulo] of Object.entries(ROTULOS_METADADOS)) {
    adicionarLinhaMetadado(rotulo, documento[chave]);
  }
  for (const extra of documento.campos_adicionais || []) {
    adicionarLinhaMetadado(extra.campo, extra.valor);
  }
}

function adicionarLinhaMetadado(rotulo, valor) {
  const campo = document.createElement("div");
  campo.className = "meta-field";

  const dt = document.createElement("dt");
  dt.textContent = rotulo;
  const dd = document.createElement("dd");
  dd.textContent = valor ?? "—";

  campo.appendChild(dt);
  campo.appendChild(dd);
  tabelaMetadadosEl.appendChild(campo);
}

function preencherTabelaItens(itens) {
  tabelaItensEl.innerHTML = "";

  const thead = document.createElement("thead");
  const cabecalho = document.createElement("tr");
  for (const rotulo of ["Descrição", "Quantidade", "Valor unitário", "Valor total"]) {
    const th = document.createElement("th");
    th.textContent = rotulo;
    cabecalho.appendChild(th);
  }
  thead.appendChild(cabecalho);
  tabelaItensEl.appendChild(thead);

  const tbody = document.createElement("tbody");
  tabelaItensEl.appendChild(tbody);

  if (!itens || itens.length === 0) {
    const tr = document.createElement("tr");
    const td = document.createElement("td");
    td.colSpan = 4;
    td.className = "empty-state";
    td.textContent = "Nenhum item identificado.";
    tr.appendChild(td);
    tbody.appendChild(tr);
    return;
  }

  for (const item of itens) {
    const tr = document.createElement("tr");
    for (const texto of [item.descricao, item.quantidade, item.valor_unitario, item.valor_total]) {
      const td = document.createElement("td");
      td.textContent = texto ?? "—";
      tr.appendChild(td);
    }
    tbody.appendChild(tr);
  }
}
