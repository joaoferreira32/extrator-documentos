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
const badgeOrigemEl = document.getElementById("badge-origem");
const bannerAvisosEl = document.getElementById("banner-avisos");
const bannerCorpoEl = document.getElementById("banner-avisos-corpo");
const resumoEl = document.getElementById("resumo");
const notaConfiancaEl = document.getElementById("nota-confianca");
const legendaConfiancaEl = document.getElementById("legenda-confianca");
const chipItensEl = document.getElementById("chip-itens");
const jsonBrutoEl = document.getElementById("json-bruto");
const tabelaMetadadosEl = document.getElementById("tabela-metadados");
const tabelaItensEl = document.getElementById("tabela-itens");
const btnExcelEl = document.getElementById("btn-excel");
const btnExtrairEl = document.getElementById("btn-extrair");

let ultimoResultado = null; // o que o Excel exporta: recebe as correções do usuário
let arquivoSelecionado = null;
let campos = []; // estado de cada campo do documento (ver construirCampos)
let temConfiancas = false; // false no modo IA (confiancas vem vazio)

const ROTULOS_METADADOS = {
  tipo_documento: "Tipo de documento",
  numero_documento: "Número do documento",
  data_emissao: "Data de emissão",
  data_vencimento: "Data de vencimento",
  emissor: "Emissor",
  destinatario: "Destinatário",
  valor_total: "Valor total",
};

const TIPOS_DOCUMENTO = ["boleto", "nota_fiscal", "pedido_compra", "relatorio", "desconhecido"];

// Só a EXIBIÇÃO é amigável: o valor interno (nota_fiscal...) é o que vai pro
// JSON e pro Excel. Tipo fora da lista (ex: um que o modo IA invente) aparece cru.
const ROTULOS_TIPO = {
  boleto: "Boleto",
  nota_fiscal: "Nota fiscal",
  pedido_compra: "Pedido de compra",
  relatorio: "Relatório",
  desconhecido: "Desconhecido",
};

// Campo adicional exibido em blocos de 4 dígitos (padrão impresso da DANFE).
// Só exibição/edição: o valor guardado e exportado continua sem espaços.
const ROTULO_CHAVE_ACESSO = "Chave de Acesso";

// Campos OBRIGATÓRIOS por tipo de documento (ver CLAUDE.md, "Interface").
// Obrigatório vazio -> destaque, já aberto pra preencher, e conta como "vazio"
// no resumo. Opcional vazio -> "—" neutro, editável ao clicar, fora do resumo.
// Tipo fora da lista (ex: algo que o modo IA invente) não tem obrigatórios.
const OBRIGATORIOS_BASE = [
  "emissor",
  "destinatario",
  "numero_documento",
  "data_emissao",
  "valor_total",
];
const CAMPOS_OBRIGATORIOS = {
  nota_fiscal: OBRIGATORIOS_BASE, // inclui DANFE
  boleto: [...OBRIGATORIOS_BASE, "data_vencimento"],
  pedido_compra: OBRIGATORIOS_BASE,
  relatorio: [],
  desconhecido: [],
};

const CAMPOS_NUMERICOS = new Set(["valor_total"]);

// ---------- Ícones (SVG estático; forma diferente por estado, nunca só cor) ----------

const ICONES = {
  alta:
    '<svg viewBox="0 0 16 16" fill="none"><circle cx="8" cy="8" r="6.25" stroke="currentColor" stroke-width="1.5"/><path d="M5.2 8.2l2 2 3.6-4" stroke="currentColor" stroke-width="1.5" stroke-linecap="round" stroke-linejoin="round"/></svg>',
  media:
    '<svg viewBox="0 0 16 16" fill="none"><path d="M8 2L1.6 13.2h12.8L8 2z" stroke="currentColor" stroke-width="1.5" stroke-linejoin="round"/><path d="M8 6.6v3" stroke="currentColor" stroke-width="1.5" stroke-linecap="round"/><circle cx="8" cy="11.4" r="0.8" fill="currentColor"/></svg>',
  baixa:
    '<svg viewBox="0 0 16 16" fill="none"><circle cx="8" cy="8" r="6.25" stroke="currentColor" stroke-width="1.5"/><path d="M5.7 5.7l4.6 4.6M10.3 5.7l-4.6 4.6" stroke="currentColor" stroke-width="1.5" stroke-linecap="round"/></svg>',
  vazio:
    '<svg viewBox="0 0 16 16" fill="none"><circle cx="8" cy="8" r="6.25" stroke="currentColor" stroke-width="1.5" stroke-dasharray="2.4 2.4"/></svg>',
  corrigido:
    '<svg viewBox="0 0 16 16" fill="none"><path d="M10.6 2.6l2.8 2.8-7.6 7.6-3.4.6.6-3.4 7.6-7.6z" stroke="currentColor" stroke-width="1.5" stroke-linejoin="round"/></svg>',
  lapis:
    '<svg viewBox="0 0 16 16" fill="none"><path d="M10.6 2.6l2.8 2.8-7.6 7.6-3.4.6.6-3.4 7.6-7.6z" stroke="currentColor" stroke-width="1.4" stroke-linejoin="round"/></svg>',
};

function icone(nome) {
  const span = document.createElement("span");
  span.className = "icone";
  span.setAttribute("aria-hidden", "true");
  span.innerHTML = ICONES[nome]; // constantes acima, nunca dado do usuário
  return span;
}

const CHIPS = {
  alta: {
    texto: "Alta",
    icone: "alta",
    classe: "chip-alta",
    dica: "Confiança alta: lido por rótulo explícito ou conferido com outro dado do documento.",
  },
  media: {
    texto: "Média",
    icone: "media",
    classe: "chip-media",
    dica: "Confiança média: inferido pela posição no documento. Vale conferir.",
  },
  baixa: {
    texto: "Baixa",
    icone: "baixa",
    classe: "chip-baixa",
    dica: "Confiança baixa: sem rótulo que confirme. Confira e corrija.",
  },
  vazio: {
    texto: "Vazio",
    icone: "vazio",
    classe: "chip-vazio",
    dica: "Campo obrigatório não identificado. Preencha.",
  },
  corrigido: {
    texto: "Corrigido",
    icone: "corrigido",
    classe: "chip-corrigido",
    dica: "Você alterou este valor.",
  },
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
    // ultimoResultado já carrega as correções feitas na tela.
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

// ---------- Resultado ----------

function mostrarResultado(resultado) {
  ultimoResultado = resultado;
  temConfiancas = Object.keys(resultado.confiancas || {}).length > 0;

  badgeModoEl.textContent =
    resultado.modo_extracao === "ia" ? "Modo: IA" : "Modo: Básico";
  badgeModoEl.className = `badge badge-${resultado.modo_extracao}`;

  badgeOrigemEl.textContent =
    resultado.origem_texto === "ocr" ? "Leitura: OCR" : "Leitura: Digital";

  renderizarBanner(resultado);

  campos = construirCampos(resultado);
  renderizarTodosOsCampos();
  preencherTabelaItens(resultado.documento.itens);
  renderizarChipItens();
  atualizarResumo();
  atualizarJsonBruto();

  notaConfiancaEl.hidden = temConfiancas;
  legendaConfiancaEl.hidden = !temConfiancas;
  resultadoEl.hidden = false;
}

function atualizarJsonBruto() {
  jsonBrutoEl.textContent = JSON.stringify(ultimoResultado, null, 2);
}

// ---------- Banner de avisos ----------

function renderizarBanner(resultado) {
  // `avisos` (lista) tem prioridade; `aviso` (string) cobre respostas antigas.
  const avisos =
    Array.isArray(resultado.avisos) && resultado.avisos.length > 0
      ? resultado.avisos
      : resultado.aviso
        ? [resultado.aviso]
        : [];

  bannerCorpoEl.replaceChildren();
  if (avisos.length === 0) {
    bannerAvisosEl.hidden = true;
    return;
  }

  if (avisos.length === 1) {
    const p = document.createElement("p");
    p.textContent = avisos[0];
    bannerCorpoEl.appendChild(p);
  } else {
    const ul = document.createElement("ul");
    for (const aviso of avisos) {
      const li = document.createElement("li");
      li.textContent = aviso;
      ul.appendChild(li);
    }
    bannerCorpoEl.appendChild(ul);
  }
  bannerAvisosEl.hidden = false;
}

// ---------- Estado dos campos ----------

function construirCampos(resultado) {
  const documento = resultado.documento;
  const lista = [];

  for (const [chave, rotulo] of Object.entries(ROTULOS_METADADOS)) {
    const valor = documento[chave] ?? null;
    lista.push({
      id: chave,
      origem: "doc",
      chave,
      rotulo,
      editor: chave === "tipo_documento" ? "select" : CAMPOS_NUMERICOS.has(chave) ? "numero" : "texto",
      confChave: chave === "tipo_documento" ? null : chave, // tipo não tem confiança
      original: valor,
      atual: valor,
      emEdicao: false,
      el: null,
    });
  }

  (documento.campos_adicionais || []).forEach((extra, indice) => {
    lista.push({
      id: `extra-${indice}`,
      origem: "extra",
      indice,
      rotulo: extra.campo,
      editor: "texto",
      formato: extra.campo === ROTULO_CHAVE_ACESSO ? "chave" : null,
      confChave: extra.campo,
      original: extra.valor ?? "",
      atual: extra.valor ?? "",
      emEdicao: false,
      el: null,
    });
  });
  return lista;
}

function tipoAtualDoDocumento() {
  return campos.find((f) => f.id === "tipo_documento")?.atual;
}

function eObrigatorio(campo) {
  if (campo.origem !== "doc") return false;
  return (CAMPOS_OBRIGATORIOS[tipoAtualDoDocumento()] || []).includes(campo.chave);
}

function estaVazio(valor) {
  return valor === null || valor === undefined || String(valor).trim() === "";
}

function normalizarParaComparar(valor, editor) {
  if (estaVazio(valor)) return "";
  if (editor === "numero") {
    const n = typeof valor === "number" ? valor : parseNumeroBR(valor);
    if (n !== null) return `#${n}`;
  }
  return String(valor).trim();
}

function foiEditado(campo) {
  return (
    normalizarParaComparar(campo.original, campo.editor) !==
    normalizarParaComparar(campo.atual, campo.editor)
  );
}

// O estado é DERIVADO (valor original x atual, obrigatoriedade, confiança):
// nada de flags que possam ficar dessincronizadas.
//   vazio            obrigatório sem valor
//   opcional-vazio   opcional sem valor ("—" neutro, fora do resumo)
//   corrigido        o usuário mudou o valor
//   alta|media|baixa confiança devolvida pelo backend
//   sem-confianca    sem indicador (tipo do documento; tudo no modo IA)
function estadoDoCampo(campo) {
  if (estaVazio(campo.atual)) return eObrigatorio(campo) ? "vazio" : "opcional-vazio";
  if (foiEditado(campo)) return "corrigido";
  const conf =
    temConfiancas && campo.confChave ? ultimoResultado.confiancas[campo.confChave] : undefined;
  return ["alta", "media", "baixa"].includes(conf) ? conf : "sem-confianca";
}

// ---------- Números no formato brasileiro ----------

// "1.234,56" | "1234,56" | "1234.56" | "R$ 1.234,56" -> número; senão null.
function parseNumeroBR(texto) {
  const t = String(texto).trim().replace(/^R\$\s*/i, "").replace(/\s+/g, "");
  if (!t) return null;
  if (/^-?\d{1,3}(\.\d{3})+(,\d+)?$/.test(t)) return Number(t.replace(/\./g, "").replace(",", "."));
  if (/^-?\d+,\d+$/.test(t)) return Number(t.replace(",", "."));
  if (/^-?\d+(\.\d+)?$/.test(t)) return Number(t);
  return null;
}

function formatarNumero(n, minimo = 2, maximo = 2) {
  return n.toLocaleString("pt-BR", { minimumFractionDigits: minimo, maximumFractionDigits: maximo });
}

// "35260472..." -> "3526 0472 ..." (blocos de 4). Só agrupa se for só dígitos.
function formatarChave(texto) {
  return /^\d+$/.test(texto) ? texto.replace(/(\d{4})(?=\d)/g, "$1 ") : texto;
}

// Texto que vai DENTRO do editor (e da exibição do valor).
function textoDoValor(campo) {
  if (estaVazio(campo.atual)) return "";
  if (campo.editor === "numero" && typeof campo.atual === "number") {
    return formatarNumero(campo.atual);
  }
  if (campo.formato === "chave") return formatarChave(String(campo.atual));
  return String(campo.atual);
}

// Texto exibido quando o campo está fechado: igual ao do editor, exceto o
// tipo do documento, que mostra o rótulo amigável.
function textoExibido(campo) {
  if (campo.id === "tipo_documento" && !estaVazio(campo.atual)) {
    return ROTULOS_TIPO[campo.atual] ?? String(campo.atual);
  }
  return textoDoValor(campo);
}

// ---------- Renderização dos campos ----------

function renderizarTodosOsCampos() {
  for (const campo of campos) {
    campo.el = renderizarCampo(campo);
  }
  tabelaMetadadosEl.replaceChildren(...campos.map((campo) => campo.el));
  ajustarAlturas(tabelaMetadadosEl);
}

function rerenderizarCampo(campo) {
  const novo = renderizarCampo(campo);
  campo.el.replaceWith(novo);
  campo.el = novo;
  ajustarAlturas(novo);
}

function criarChip(estado, id) {
  const def = CHIPS[estado];
  if (!def) return null; // opcional-vazio e sem-confianca: sem indicador
  const chip = document.createElement("span");
  chip.className = `chip ${def.classe}`;
  chip.id = id;
  chip.title = def.dica;
  chip.appendChild(icone(def.icone));
  const texto = document.createElement("span");
  texto.textContent = def.texto;
  chip.appendChild(texto);
  return chip;
}

function renderizarCampo(campo) {
  const estado = estadoDoCampo(campo);
  // Baixa e vazio-obrigatório já vêm abertos; o resto abre ao clicar.
  const aberto = campo.emEdicao || estado === "baixa" || estado === "vazio";

  const raiz = document.createElement("div");
  raiz.className = `meta-field is-${estado}`;
  raiz.dataset.campo = campo.id;

  const dt = document.createElement("dt");
  const rotulo = document.createElement("span");
  rotulo.className = "meta-rotulo";
  rotulo.id = `rot-${campo.id}`;
  rotulo.textContent = campo.rotulo;
  dt.appendChild(rotulo);
  const chip = criarChip(estado, `chip-${campo.id}`);
  if (chip) dt.appendChild(chip);

  const dd = document.createElement("dd");
  if (aberto) {
    const dica = criarDica(estado, campo);
    dd.appendChild(criarEditor(campo, chip, dica));
    if (dica) dd.appendChild(dica);
  } else {
    dd.appendChild(criarBotaoValor(campo, estado));
  }

  raiz.append(dt, dd);
  return raiz;
}

function criarDica(estado, campo) {
  const texto =
    estado === "baixa" ? "Confira e corrija." : estado === "vazio" ? "Preencha este campo." : null;
  if (!texto) return null;
  const p = document.createElement("p");
  p.className = "dica";
  p.id = `dica-${campo.id}`;
  p.textContent = texto;
  return p;
}

function criarBotaoValor(campo, estado) {
  const botao = document.createElement("button");
  botao.type = "button";
  botao.className = "valor-btn";

  const texto = document.createElement("span");
  texto.className = "valor-texto";
  const vazio = estado === "opcional-vazio";
  texto.textContent = vazio ? "—" : textoExibido(campo);
  botao.append(texto, icone("lapis"));

  botao.setAttribute(
    "aria-label",
    vazio ? `Editar ${campo.rotulo} (vazio)` : `Editar ${campo.rotulo}: ${textoExibido(campo)}`
  );
  botao.addEventListener("click", () => {
    campo.emEdicao = true;
    rerenderizarCampo(campo);
    const editor = campo.el.querySelector(".valor-editor");
    editor.focus();
    if (typeof editor.select === "function" && campo.editor !== "select") editor.select();
  });
  return botao;
}

function criarEditor(campo, chip, dica) {
  let editor;
  if (campo.editor === "select") {
    editor = document.createElement("select");
    const valores = TIPOS_DOCUMENTO.includes(campo.atual)
      ? TIPOS_DOCUMENTO
      : [...TIPOS_DOCUMENTO, campo.atual];
    for (const valor of valores) {
      const opcao = document.createElement("option");
      opcao.value = valor;
      opcao.textContent = ROTULOS_TIPO[valor] ?? valor;
      editor.appendChild(opcao);
    }
    editor.value = campo.atual;
  } else {
    editor = document.createElement("textarea");
    editor.rows = 1;
    editor.spellcheck = false;
    editor.value = textoDoValor(campo);
  }
  editor.className = "valor-editor";
  editor.setAttribute("aria-labelledby", `rot-${campo.id}`);
  const descricao = [chip?.id, dica?.id].filter(Boolean).join(" ");
  if (descricao) editor.setAttribute("aria-describedby", descricao);

  const valorAoAbrir = campo.atual;

  editor.addEventListener("input", () => {
    aplicarTexto(campo, editor.value);
    if (campo.editor !== "select") ajustarAltura(editor);
  });
  editor.addEventListener("keydown", (evento) => {
    if (evento.key === "Enter" && campo.editor !== "select") {
      evento.preventDefault(); // valores são de uma linha só
      finalizarEdicao(campo, editor, true);
    } else if (evento.key === "Escape") {
      evento.preventDefault();
      campo.atual = valorAoAbrir;
      escreverNoResultado(campo);
      atualizarJsonBruto();
      finalizarEdicao(campo, editor, true);
    }
  });
  editor.addEventListener("blur", () => finalizarEdicao(campo, editor, false));
  if (campo.editor === "select") {
    editor.addEventListener("change", () => finalizarEdicao(campo, editor, true));
  }
  return editor;
}

function ajustarAltura(textarea) {
  textarea.style.height = "auto";
  textarea.style.height = `${textarea.scrollHeight}px`;
}

function ajustarAlturas(raiz) {
  raiz.querySelectorAll("textarea.valor-editor").forEach(ajustarAltura);
}

// Aplica o que foi digitado ao estado e ao resultado que o Excel exporta, sem
// redesenhar (redesenhar no meio da digitação faria o campo perder o foco).
function aplicarTexto(campo, texto) {
  const t = texto.trim();
  let valor;
  if (campo.editor === "select") {
    valor = texto;
  } else if (t === "") {
    valor = campo.origem === "extra" ? "" : null; // campo_adicional exige string
  } else if (campo.editor === "numero") {
    const numero = parseNumeroBR(t);
    valor = numero !== null ? numero : t; // não parseável: fica o texto (schema aceita)
  } else if (campo.formato === "chave") {
    valor = t.replace(/\s+/g, ""); // os blocos são só exibição: guarda sem espaços
  } else {
    valor = t;
  }
  campo.atual = valor;
  escreverNoResultado(campo);
  atualizarJsonBruto();
}

function escreverNoResultado(campo) {
  const documento = ultimoResultado.documento;
  if (campo.origem === "doc") {
    documento[campo.chave] = campo.atual;
  } else {
    documento.campos_adicionais[campo.indice].valor = campo.atual ?? "";
  }
}

function finalizarEdicao(campo, editor, devolverFoco) {
  // Trava de reentrância: substituir o editor que está com foco faz o Chrome
  // disparar `blur` SÍNCRONO no meio do replaceWith; esse blur chamaria
  // finalizarEdicao de novo, redesenharia o campo por dentro do redesenho e o
  // replaceWith externo lançaria exceção (foco nunca voltava ao botão).
  if (campo.finalizando || !editor.isConnected) return;
  campo.finalizando = true;
  try {
    campo.emEdicao = false;
    if (campo.id === "tipo_documento") {
      // mudar o tipo muda quais campos são obrigatórios
      renderizarTodosOsCampos();
    } else {
      rerenderizarCampo(campo);
    }
    atualizarResumo();
  } finally {
    campo.finalizando = false;
  }
  if (devolverFoco) {
    campo.el.querySelector(".valor-btn, .valor-editor")?.focus();
  }
}

// ---------- Resumo de confiança ----------

function confiancaDosItens() {
  const itens = ultimoResultado.documento.itens;
  const conf = temConfiancas ? ultimoResultado.confiancas.itens : undefined;
  return itens && itens.length > 0 && ["alta", "media", "baixa"].includes(conf) ? conf : null;
}

function renderizarChipItens() {
  chipItensEl.replaceChildren();
  const conf = confiancaDosItens();
  if (!conf) return;
  chipItensEl.appendChild(criarChip(conf, "chip-itens-valor"));
}

function contarEstados() {
  const contagem = { alta: 0, revisar: 0, vazio: 0, corrigido: 0 };
  for (const campo of campos) {
    if (campo.id === "tipo_documento") continue;
    const estado = estadoDoCampo(campo);
    if (estado === "alta") contagem.alta++;
    else if (estado === "media" || estado === "baixa") contagem.revisar++;
    else if (estado === "vazio") contagem.vazio++;
    else if (estado === "corrigido") contagem.corrigido++;
  }
  const itens = confiancaDosItens();
  if (itens === "alta") contagem.alta++;
  else if (itens) contagem.revisar++;
  contagem.total = contagem.alta + contagem.revisar;
  return contagem;
}

function plural(n, singular, pluralTexto) {
  return `${n} ${n === 1 ? singular : pluralTexto}`;
}

function atualizarResumo() {
  const c = contarEstados();
  const partes = [];

  if (temConfiancas) {
    partes.push({
      classe: "resumo-alta",
      icone: "alta",
      texto: `${c.alta} de ${c.total} campos com alta confiança`,
    });
    if (c.revisar > 0) {
      partes.push({ classe: "resumo-revisar", icone: "media", texto: `${c.revisar} para revisar` });
    }
  }
  if (c.vazio > 0) {
    partes.push({ classe: "resumo-vazio", icone: "vazio", texto: plural(c.vazio, "vazio", "vazios") });
  }
  if (c.corrigido > 0) {
    partes.push({
      classe: "resumo-corrigido",
      icone: "corrigido",
      texto: plural(c.corrigido, "corrigido", "corrigidos"),
    });
  }

  resumoEl.replaceChildren();
  if (partes.length === 0) {
    resumoEl.hidden = true;
    return;
  }
  partes.forEach((parte, i) => {
    if (i > 0) {
      const sep = document.createElement("span");
      sep.className = "resumo-sep";
      sep.setAttribute("aria-hidden", "true");
      sep.textContent = "·";
      const pausa = document.createElement("span");
      pausa.className = "visually-hidden";
      pausa.textContent = ", ";
      resumoEl.append(sep, pausa);
    }
    const item = document.createElement("span");
    item.className = `resumo-item ${parte.classe}`;
    item.appendChild(icone(parte.icone));
    const texto = document.createElement("span");
    texto.textContent = parte.texto;
    item.appendChild(texto);
    resumoEl.appendChild(item);
  });
  resumoEl.hidden = false;
}

// ---------- Itens ----------

// Número -> pt-BR (215.03 -> "215,03"); texto (valor que não virou número) fica como veio.
function formatarCelulaNumerica(valor, minimo, maximo) {
  if (typeof valor === "number") return formatarNumero(valor, minimo, maximo);
  return valor;
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
    const celulas = [
      item.descricao,
      formatarCelulaNumerica(item.quantidade, 0, 4),
      formatarCelulaNumerica(item.valor_unitario, 2, 4), // preço unitário pode ter 4 casas
      formatarCelulaNumerica(item.valor_total, 2, 2),
    ];
    for (const texto of celulas) {
      const td = document.createElement("td");
      td.textContent = texto ?? "—";
      tr.appendChild(td);
    }
    tbody.appendChild(tr);
  }
}
