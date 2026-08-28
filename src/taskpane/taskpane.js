const ui = {
  messages: document.querySelector("#messages"), prompt: document.querySelector("#prompt"), form: document.querySelector("#chat-form"), send: document.querySelector("#send"),
  settings: document.querySelector("#settings"), config: document.querySelector("#config"), refresh: document.querySelector("#refresh-context"), context: document.querySelector("#context-label"), actions: document.querySelector("#actions"), apply: document.querySelector("#apply"), discard: document.querySelector("#discard"), selectionContext: document.querySelector("#selection-context"), selectionPreview: document.querySelector("#selection-preview"), selectionState: document.querySelector("#selection-state"), toggleSelection: document.querySelector("#toggle-selection"), panelColor: document.querySelector("#panel-color"), accentColor: document.querySelector("#accent-color"), surfaceColor: document.querySelector("#surface-color"), textColor: document.querySelector("#text-color"), mutedColor: document.querySelector("#muted-color"), borderColor: document.querySelector("#border-color"), resetTheme: document.querySelector("#reset-theme"), webResearch: document.querySelector("#web-research")
};
let documentText = "", selectionText = "", activeSelectionText = "", includeSelection = true, selectionAnchorId = null, pendingPlan = null, history = [];

Office.onReady((info) => {
  if (info.host !== Office.HostType.Word) show("Este complemento debe abrirse desde Microsoft Word.", "error");
  else {
    loadTheme(); refreshContext();
    try { Office.context.document.addHandlerAsync(Office.EventType.DocumentSelectionChanged, refreshContext); } catch { /* El botón Actualizar contexto sigue disponible. */ }
  }
});

function show(content, type = "assistant") { const message = document.createElement("article"); message.className = type; message.textContent = content; ui.messages.append(message); ui.messages.scrollTop = ui.messages.scrollHeight; return message; }
function trim(text, limit = 24000) { return text.length > limit ? `${text.slice(0, limit)}\n\n[El resto del documento no se envió por longitud.]` : text; }
function planDescription(plan) { return `${plan.summary}\n\n${plan.operations.length} cambio${plan.operations.length === 1 ? "" : "s"} listo${plan.operations.length === 1 ? "" : "s"} para aplicar.`; }
function showResearchResult(message, data) {
  message.textContent = data.answer;
  if (!data.sources?.length) return;
  const title = document.createElement("strong"); title.textContent = "\n\nFuentes"; message.append(title);
  const list = document.createElement("ul");
  data.sources.forEach(({ title: sourceTitle, url }) => { const item = document.createElement("li"); const link = document.createElement("a"); link.href = url; link.target = "_blank"; link.rel = "noreferrer"; link.textContent = sourceTitle; item.append(link); list.append(item); });
  message.append(list);
}

function renderSelectionContext() {
  if (!activeSelectionText) { ui.selectionContext.classList.add("hidden"); return; }
  ui.selectionContext.classList.remove("hidden");
  ui.selectionPreview.textContent = activeSelectionText;
  ui.selectionState.textContent = includeSelection ? "Se enviará con tu instrucción" : "No se enviará con tu instrucción";
  ui.toggleSelection.textContent = includeSelection ? "Quitar del contexto" : "Usar como contexto";
}

const defaultTheme = { panel: "#17171c", accent: "#7456d8", surface: "#292831", text: "#e9e9ed", muted: "#a2a2ac", border: "#303039" };
const themeInputs = { panel: "panelColor", accent: "accentColor", surface: "surfaceColor", text: "textColor", muted: "mutedColor", border: "borderColor" };

function setTheme(theme) {
  for (const [name, color] of Object.entries(theme)) {
    document.documentElement.style.setProperty(`--${name === "panel" ? "panel-background" : name}`, color);
    ui[themeInputs[name]].value = color;
    localStorage.setItem(`word-gpt-${name}-color`, color);
  }
}

function colorsFromInputs() { return Object.fromEntries(Object.entries(themeInputs).map(([name, input]) => [name, ui[input].value])); }
function loadTheme() { setTheme(Object.fromEntries(Object.entries(defaultTheme).map(([name, value]) => [name, localStorage.getItem(`word-gpt-${name}-color`) || value]))); }

async function refreshContext() {
  try {
    await Word.run(async (context) => {
      const body = context.document.body; const selection = context.document.getSelection();
      body.load("text"); selection.load("text"); await context.sync();
      documentText = trim(body.text || "");
      const nextSelection = trim(selection.text || "", 10000);
      if (nextSelection !== activeSelectionText) includeSelection = true;
      activeSelectionText = nextSelection; selectionText = includeSelection ? activeSelectionText : "";
      renderSelectionContext();
      ui.context.textContent = activeSelectionText ? `Selección: ${activeSelectionText.length.toLocaleString()} caracteres${includeSelection ? "" : " (excluida)"}` : `Documento: ${documentText.length.toLocaleString()} caracteres`;
    });
  } catch { ui.context.textContent = "No se pudo leer el documento"; }
}

async function findFirstRange(context, find) {
  const results = context.document.body.search(find, { matchCase: true, matchWholeWord: false });
  results.load("items"); await context.sync();
  return results.items[0] || null;
}

function applyFont(range, font) {
  if (!font) return;
  if (typeof font.bold === "boolean") range.font.bold = font.bold;
  if (typeof font.italic === "boolean") range.font.italic = font.italic;
  if (typeof font.strikethrough === "boolean") range.font.strikethrough = font.strikethrough;
  if (typeof font.allCaps === "boolean") range.font.allCaps = font.allCaps;
  if (typeof font.smallCaps === "boolean") range.font.smallCaps = font.smallCaps;
  if (typeof font.superscript === "boolean") range.font.superscript = font.superscript;
  if (typeof font.subscript === "boolean") range.font.subscript = font.subscript;
  if (typeof font.underline === "string") {
    const underline = { none: Word.UnderlineType.none, single: Word.UnderlineType.single, double: Word.UnderlineType.double };
    range.font.underline = underline[font.underline.toLowerCase()];
  }
  if (typeof font.color === "string") range.font.color = font.color;
  if (typeof font.highlightColor === "string") range.font.highlightColor = font.highlightColor;
  if (typeof font.size === "number") range.font.size = font.size;
  if (typeof font.name === "string") range.font.name = font.name;
}

async function applyParagraphFormat(context, target, paragraphFormat) {
  if (!paragraphFormat) return;
  let targets = [target];
  if (target.paragraphs) {
    target.paragraphs.load("items"); await context.sync();
    targets = target.paragraphs.items;
  }
  const alignment = { left: "Left", centered: "Centered", right: "Right", justified: "Justified" };
  for (const item of targets) {
    const format = item.paragraphFormat || item;
    if (typeof paragraphFormat.alignment === "string") format.alignment = alignment[paragraphFormat.alignment.toLowerCase()];
    for (const key of ["leftIndent", "rightIndent", "firstLineIndent", "spaceBefore", "spaceAfter", "lineSpacing", "keepTogether", "keepWithNext", "widowControl"]) {
      if (paragraphFormat[key] !== undefined) format[key] = paragraphFormat[key];
    }
  }
}

async function paragraphAt(context, number) {
  const paragraphs = context.document.body.paragraphs;
  paragraphs.load("items"); await context.sync();
  return paragraphs.items[number - 1] || null;
}

async function applyOperation(operation) {
  return Word.run(async (context) => {
    if (operation.type === "insert_at_cursor") {
      context.document.body.insertText(operation.text, Word.InsertLocation.replace);
      await context.sync();
      return true;
    }
    if (operation.type === "format_paragraph") {
      const paragraph = await paragraphAt(context, operation.paragraph);
      if (!paragraph) return false;
      applyFont(paragraph, operation.font);
      await applyParagraphFormat(context, paragraph, operation.paragraphFormat);
      await context.sync();
      return true;
    }
    if (operation.type === "format_document") {
      const paragraphs = context.document.body.paragraphs;
      paragraphs.load("items"); await context.sync();
      for (const paragraph of paragraphs.items) { applyFont(paragraph, operation.font); await applyParagraphFormat(context, paragraph, operation.paragraphFormat); }
      await context.sync();
      return true;
    }
    let range;
    if (operation.anchorId) range = context.document.contentControls.getById(operation.anchorId).getRange();
    else if (operation.target === "selection" || ["replace_selection", "insert_at_selection"].includes(operation.type)) range = context.document.getSelection();
    else range = await findFirstRange(context, operation.find);
    if (!range) return false;
    if (operation.type === "replace" || operation.type === "replace_selection") range.insertText(operation.replacement ?? operation.text, Word.InsertLocation.replace);
    if (operation.type === "insert_after" || operation.type === "insert_at_selection") range.insertText(operation.text, Word.InsertLocation.after);
    if (operation.type === "insert_before") range.insertText(operation.text, Word.InsertLocation.before);
    if (operation.type === "format") { applyFont(range, operation.font); await applyParagraphFormat(context, range, operation.paragraphFormat); }
    await context.sync(); return true;
  });
}

async function removeSelectionAnchor() {
  if (!selectionAnchorId) return;
  const anchorId = selectionAnchorId; selectionAnchorId = null;
  try {
    await Word.run(async (context) => {
      context.document.contentControls.getById(anchorId).delete(false);
      await context.sync();
    });
  } catch { /* El ancla pudo haber sido eliminada por una edición. */ }
}

async function applyPlan() {
  if (!pendingPlan) return;
  ui.apply.disabled = true; ui.discard.disabled = true;
  const status = show("Aplicando cambios…"); let applied = 0, skipped = 0;
  try {
    const orderedOperations = [...pendingPlan.operations].sort((a, b) => (b.type === "format_document") - (a.type === "format_document"));
    for (const operation of orderedOperations) {
      if (await applyOperation(operation)) applied += 1;
      else skipped += 1;
    }
    status.textContent = `Cambios aplicados: ${applied}.${skipped ? ` No encontré ${skipped} fragmento${skipped === 1 ? "" : "s"}; no se modificaron.` : ""}`;
    pendingPlan = null; ui.actions.classList.add("hidden"); await removeSelectionAnchor(); await refreshContext();
  } catch (error) {
    status.className = "error"; status.textContent = `No se pudieron aplicar todos los cambios: ${error.message || "error de Word"}.`;
  } finally { ui.apply.disabled = false; ui.discard.disabled = false; }
}

ui.settings.onclick = () => ui.config.classList.toggle("hidden");
ui.refresh.onclick = refreshContext;
ui.apply.onclick = applyPlan;
ui.discard.onclick = async () => { pendingPlan = null; ui.actions.classList.add("hidden"); await removeSelectionAnchor(); show("Cambios descartados."); };
ui.toggleSelection.onclick = () => { includeSelection = !includeSelection; selectionText = includeSelection ? activeSelectionText : ""; renderSelectionContext(); ui.context.textContent = `Selección: ${activeSelectionText.length.toLocaleString()} caracteres${includeSelection ? "" : " (excluida)"}`; };
Object.values(themeInputs).forEach((input) => { ui[input].oninput = () => setTheme(colorsFromInputs()); });
ui.resetTheme.onclick = () => setTheme(defaultTheme);
ui.prompt.onkeydown = (event) => {
  if ((event.ctrlKey || event.metaKey) && event.key === "Enter") {
    event.preventDefault();
    if (!ui.send.disabled) ui.form.requestSubmit();
  }
};

async function anchorAndDeselectAfterSend() {
  if (!activeSelectionText) return;
  try {
    await Word.run(async (context) => {
      const selection = context.document.getSelection();
      const anchor = selection.insertContentControl(); anchor.load("id");
      selection.collapse(Word.InsertLocation.end); selection.select(); await context.sync();
      selectionAnchorId = anchor.id;
    });
  } catch { selectionAnchorId = null; }
  activeSelectionText = ""; selectionText = ""; includeSelection = true; renderSelectionContext();
  ui.context.textContent = `Documento: ${documentText.length.toLocaleString()} caracteres`;
}
ui.form.onsubmit = async (event) => {
  event.preventDefault(); const message = ui.prompt.value.trim(); if (!message) return;
  await refreshContext(); const selectionForRequest = selectionText; const useWebResearch = ui.webResearch.checked; await anchorAndDeselectAfterSend(); show(message, "user"); ui.prompt.value = ""; ui.send.disabled = true; const pending = show(useWebResearch ? "Buscando fuentes públicas…" : "Preparando cambios…");
  try {
    const response = await fetch(useWebResearch ? "/api/research" : "/api/edit", { method:"POST", headers:{"Content-Type":"application/json"}, body:JSON.stringify({ message, documentText, selectionText: selectionForRequest, history }) });
    const data = await response.json(); if (!response.ok) throw new Error(data.error || "Error inesperado.");
    if (useWebResearch) { await removeSelectionAnchor(); showResearchResult(pending, data); history = [...history, { role:"user", content:message }, { role:"assistant", content:data.answer }].slice(-8); return; }
    if (selectionAnchorId) data.operations = data.operations.map((operation) => operation.target === "selection" || ["replace_selection", "insert_at_selection"].includes(operation.type) ? { ...operation, anchorId: selectionAnchorId } : operation);
    pendingPlan = data; pending.textContent = planDescription(data);
    if (data.operations.length) ui.actions.classList.remove("hidden");
    else await removeSelectionAnchor();
    history = [...history, { role:"user", content:message }, { role:"assistant", content:data.summary }].slice(-8);
  } catch (error) { await removeSelectionAnchor(); pending.className = "error"; pending.textContent = error.message; }
  finally { ui.send.disabled = false; ui.prompt.focus(); }
};
