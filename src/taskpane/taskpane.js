const ui = {
  messages: document.querySelector("#messages"), prompt: document.querySelector("#prompt"), form: document.querySelector("#chat-form"), send: document.querySelector("#send"),
  settings: document.querySelector("#settings"), config: document.querySelector("#config"), refresh: document.querySelector("#refresh-context"), context: document.querySelector("#context-label"), actions: document.querySelector("#actions"), apply: document.querySelector("#apply"), discard: document.querySelector("#discard")
};
let documentText = "", selectionText = "", pendingPlan = null, history = [];

Office.onReady((info) => {
  if (info.host !== Office.HostType.Word) show("Este complemento debe abrirse desde Microsoft Word.", "error");
  else refreshContext();
});

function show(content, type = "assistant") { const message = document.createElement("article"); message.className = type; message.textContent = content; ui.messages.append(message); ui.messages.scrollTop = ui.messages.scrollHeight; return message; }
function trim(text, limit = 24000) { return text.length > limit ? `${text.slice(0, limit)}\n\n[El resto del documento no se envió por longitud.]` : text; }
function planDescription(plan) { return `${plan.summary}\n\n${plan.operations.length} cambio${plan.operations.length === 1 ? "" : "s"} listo${plan.operations.length === 1 ? "" : "s"} para aplicar.`; }

async function refreshContext() {
  try {
    await Word.run(async (context) => {
      const body = context.document.body; const selection = context.document.getSelection();
      body.load("text"); selection.load("text"); await context.sync();
      documentText = trim(body.text || ""); selectionText = trim(selection.text || "", 10000);
      ui.context.textContent = selectionText ? `Selección: ${selectionText.length.toLocaleString()} caracteres` : `Documento: ${documentText.length.toLocaleString()} caracteres`;
    });
  } catch { ui.context.textContent = "No se pudo leer el documento"; }
}

async function findFirstRange(context, find) {
  const results = context.document.body.search(find, { matchCase: true, matchWholeWord: false });
  results.load("items"); await context.sync();
  return results.items[0] || null;
}

function applyFont(range, font) {
  if (typeof font.bold === "boolean") range.font.bold = font.bold;
  if (typeof font.italic === "boolean") range.font.italic = font.italic;
  if (typeof font.color === "string") range.font.color = font.color;
  if (typeof font.highlightColor === "string") range.font.highlightColor = font.highlightColor;
  if (typeof font.size === "number") range.font.size = font.size;
  if (typeof font.name === "string") range.font.name = font.name;
}

async function applyOperation(operation) {
  return Word.run(async (context) => {
    if (operation.type === "insert_at_cursor") {
      context.document.body.insertText(operation.text, Word.InsertLocation.replace);
      await context.sync();
      return true;
    }
    if (operation.type === "format_paragraph") {
      const paragraph = context.document.body.paragraphs.getItemAt(operation.paragraph - 1);
      applyFont(paragraph, operation.font);
      await context.sync();
      return true;
    }
    let range;
    if (operation.target === "selection" || ["replace_selection", "insert_at_selection"].includes(operation.type)) range = context.document.getSelection();
    else range = await findFirstRange(context, operation.find);
    if (!range) return false;
    if (operation.type === "replace" || operation.type === "replace_selection") range.insertText(operation.replacement ?? operation.text, Word.InsertLocation.replace);
    if (operation.type === "insert_after" || operation.type === "insert_at_selection") range.insertText(operation.text, Word.InsertLocation.after);
    if (operation.type === "insert_before") range.insertText(operation.text, Word.InsertLocation.before);
    if (operation.type === "format") applyFont(range, operation.font);
    await context.sync(); return true;
  });
}

async function applyPlan() {
  if (!pendingPlan) return;
  ui.apply.disabled = true; ui.discard.disabled = true;
  const status = show("Aplicando cambios…"); let applied = 0, skipped = 0;
  try {
    for (const operation of pendingPlan.operations) {
      if (await applyOperation(operation)) applied += 1;
      else skipped += 1;
    }
    status.textContent = `Cambios aplicados: ${applied}.${skipped ? ` No encontré ${skipped} fragmento${skipped === 1 ? "" : "s"}; no se modificaron.` : ""}`;
    pendingPlan = null; ui.actions.classList.add("hidden"); await refreshContext();
  } catch (error) {
    status.className = "error"; status.textContent = `No se pudieron aplicar todos los cambios: ${error.message || "error de Word"}.`;
  } finally { ui.apply.disabled = false; ui.discard.disabled = false; }
}

ui.settings.onclick = () => ui.config.classList.toggle("hidden");
ui.refresh.onclick = refreshContext;
ui.apply.onclick = applyPlan;
ui.discard.onclick = () => { pendingPlan = null; ui.actions.classList.add("hidden"); show("Cambios descartados."); };
ui.form.onsubmit = async (event) => {
  event.preventDefault(); const message = ui.prompt.value.trim(); if (!message) return;
  await refreshContext(); show(message, "user"); ui.prompt.value = ""; ui.send.disabled = true; const pending = show("Preparando cambios…");
  try {
    const response = await fetch("/api/edit", { method:"POST", headers:{"Content-Type":"application/json"}, body:JSON.stringify({ message, documentText, selectionText, history }) });
    const data = await response.json(); if (!response.ok) throw new Error(data.error || "Error inesperado.");
    pendingPlan = data; pending.textContent = planDescription(data);
    if (data.operations.length) ui.actions.classList.remove("hidden");
    history = [...history, { role:"user", content:message }, { role:"assistant", content:data.summary }].slice(-8);
  } catch (error) { pending.className = "error"; pending.textContent = error.message; }
  finally { ui.send.disabled = false; ui.prompt.focus(); }
};
