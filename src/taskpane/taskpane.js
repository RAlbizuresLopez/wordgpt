const ui = {
  messages: document.querySelector("#messages"), prompt: document.querySelector("#prompt"), form: document.querySelector("#chat-form"), send: document.querySelector("#send"),
  settings: document.querySelector("#settings"), config: document.querySelector("#config"), refresh: document.querySelector("#refresh-context"), context: document.querySelector("#context-label"), actions: document.querySelector("#actions"), insert: document.querySelector("#insert"), replace: document.querySelector("#replace")
};
let documentText = "", selectionText = "", lastAnswer = "", history = [];

Office.onReady((info) => {
  if (info.host !== Office.HostType.Word) show("Este complemento debe abrirse desde Microsoft Word.", "error");
  else refreshContext();
});

function show(content, type = "assistant") { const message = document.createElement("article"); message.className = type; message.textContent = content; ui.messages.append(message); ui.messages.scrollTop = ui.messages.scrollHeight; return message; }
function trim(text, limit = 24000) { return text.length > limit ? `${text.slice(0, limit)}\n\n[El resto del documento no se envió por longitud.]` : text; }

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

async function writeToDocument(mode) {
  if (!lastAnswer) return;
  try {
    await Word.run(async (context) => {
      const range = context.document.getSelection();
      if (mode === "replace") range.insertText(lastAnswer, Word.InsertLocation.replace);
      else range.insertText(lastAnswer, Word.InsertLocation.after);
      await context.sync();
    });
    await refreshContext();
  } catch { show("No se pudo escribir en el documento. Verifica que no esté protegido.", "error"); }
}

ui.settings.onclick = () => ui.config.classList.toggle("hidden");
ui.refresh.onclick = refreshContext;
ui.insert.onclick = () => writeToDocument("insert");
ui.replace.onclick = () => writeToDocument("replace");
ui.form.onsubmit = async (event) => {
  event.preventDefault(); const message = ui.prompt.value.trim(); if (!message) return;
  await refreshContext(); show(message, "user"); ui.prompt.value = ""; ui.send.disabled = true; const pending = show("Pensando…");
  try {
    const response = await fetch("/api/chat", { method:"POST", headers:{"Content-Type":"application/json"}, body:JSON.stringify({ message, documentText, selectionText, history }) });
    const data = await response.json(); if (!response.ok) throw new Error(data.error || "Error inesperado.");
    pending.textContent = data.answer; lastAnswer = data.answer; ui.actions.classList.remove("hidden");
    history = [...history, { role:"user", content:message }, { role:"assistant", content:data.answer }].slice(-8);
  } catch (error) { pending.className = "error"; pending.textContent = error.message; }
  finally { ui.send.disabled = false; ui.prompt.focus(); }
};
