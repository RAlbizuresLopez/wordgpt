import "dotenv/config";
import express from "express";

const app = express();
const port = Number(process.env.PORT || 3001);
const bindHost = process.env.BIND_HOST || "127.0.0.1";
app.use(express.json({ limit: "1mb" }));
app.use(express.static("dist"));

const ollamaBaseUrl = (process.env.OLLAMA_BASE_URL || "http://127.0.0.1:11434").replace(/\/$/, "");
const ollamaModel = process.env.OLLAMA_MODEL || "qwen3:14b";
const contextLength = Number(process.env.OLLAMA_CONTEXT_LENGTH || 8192);
const allowedTailscaleLogins = new Set(
  (process.env.ALLOWED_TAILSCALE_USER_LOGINS || "")
    .split(",")
    .map(login => login.trim().toLowerCase())
    .filter(Boolean)
);
const systemInstructions = `Eres un asistente experto integrado en Microsoft Word. Responde en el idioma del usuario. Usa el contexto del documento únicamente para ayudar con su petición. Si propones texto para insertar, entrégalo listo para pegar y no inventes información que no esté sustentada por el documento. No proporciones diagnósticos, tratamientos ni recomendaciones clínicas.`;
const editInstructions = `Eres el motor de edición de un documento de Microsoft Word. Devuelve ÚNICAMENTE JSON válido, sin Markdown ni explicación exterior. La respuesta debe tener {"summary":"breve resumen","operations":[...]}. Operaciones permitidas: {"type":"replace","find":"fragmento exacto","replacement":"texto nuevo"}; {"type":"insert_after","find":"fragmento exacto","text":"texto"}; {"type":"insert_before","find":"fragmento exacto","text":"texto"}; {"type":"replace_selection","text":"texto nuevo"} (solo con selección); {"type":"insert_at_selection","text":"texto"} (solo con selección); y {"type":"format","target":"selection"} o {"type":"format","find":"fragmento exacto"}, con "font" que puede incluir bold, italic, color, highlightColor, size o name. Usa máximo 10 operaciones. "find" debe aparecer exactamente en el contexto y tener máximo 240 caracteres. Cada búsqueda se aplicará solo a la primera coincidencia. No inventes fragmentos. Si hay una selección, prioriza cambiarla. Para formato usa format; los colores pueden ser CSS sencillos o hexadecimales. Si el cambio no es seguro, devuelve operations vacía y explica el motivo. No hagas diagnósticos, tratamientos ni recomendaciones clínicas.`;

app.use("/api", (req, res, next) => {
  if (req.path === "/health") return next();
  if (!allowedTailscaleLogins.size) return next();
  const requester = req.get("Tailscale-User-Login")?.toLowerCase();
  if (requester && allowedTailscaleLogins.has(requester)) return next();
  res.status(403).json({ error: "Este usuario de Tailscale no tiene acceso al asistente." });
});

app.get("/api/health", async (_req, res) => {
  try {
    const response = await fetch(`${ollamaBaseUrl}/api/tags`);
    if (!response.ok) throw new Error(`Ollama respondió ${response.status}`);
    res.json({ ok: true, model: ollamaModel });
  } catch {
    res.status(503).json({ ok: false, error: "Ollama no está disponible en el Mac mini." });
  }
});

function makeContext(documentText, selectionText) {
  return [
    selectionText && `SELECCIÓN ACTUAL:\n${selectionText}`,
    documentText && `DOCUMENTO (puede estar truncado):\n${documentText}`
  ].filter(Boolean).join("\n\n");
}

function sanitizeEditPlan(value, { hasSelection, context }) {
  const allowedTypes = new Set(["replace", "insert_after", "insert_before", "replace_selection", "insert_at_selection", "format"]);
  const allowedFonts = new Set(["bold", "italic", "color", "highlightColor", "size", "name"]);
  const operations = Array.isArray(value?.operations) ? value.operations.slice(0, 10) : [];
  const clean = operations.flatMap((operation) => {
    if (!operation || !allowedTypes.has(operation.type)) return [];
    const { type } = operation;
    const find = typeof operation.find === "string" ? operation.find.trim().slice(0, 240) : "";
    const text = typeof operation.text === "string" ? operation.text.slice(0, 12000) : "";
    const replacement = typeof operation.replacement === "string" ? operation.replacement.slice(0, 12000) : "";
    if (["replace", "insert_after", "insert_before"].includes(type) && !find) return [];
    if (["replace_selection", "insert_at_selection"].includes(type) && (!text || !hasSelection)) return [];
    if (type === "replace" && typeof operation.replacement !== "string") return [];
    if (type === "format") {
      if ((operation.target === "selection" && !hasSelection) || (operation.target !== "selection" && !find)) return [];
      const font = Object.fromEntries(Object.entries(operation.font || {}).filter(([key, item]) => {
        if (!allowedFonts.has(key)) return false;
        if (["bold", "italic"].includes(key)) return typeof item === "boolean";
        if (key === "size") return typeof item === "number" && item >= 6 && item <= 96;
        return typeof item === "string" && item.length <= 80;
      }));
      return Object.keys(font).length && (!find || context.includes(find)) ? [{ type, ...(find ? { find } : { target: "selection" }), font }] : [];
    }
    if (find && !context.includes(find)) return [];
    return [{ type, ...(find ? { find } : {}), ...(type === "replace" ? { replacement } : { text }) }];
  });
  return { summary: typeof value?.summary === "string" ? value.summary.slice(0, 500) : "Plan de edición preparado.", operations: clean };
}

app.post("/api/chat", async (req, res) => {
  const { message, documentText = "", selectionText = "", history = [] } = req.body || {};
  if (typeof message !== "string" || !message.trim()) return res.status(400).json({ error: "Escribe un mensaje." });

  const context = makeContext(documentText, selectionText);
  const prior = Array.isArray(history) ? history.slice(-8).map(({ role, content }) => ({ role, content: String(content).slice(0, 6000) })) : [];

  try {
    const response = await fetch(`${ollamaBaseUrl}/api/chat`, {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify({
        model: ollamaModel,
        stream: false,
        think: false,
        keep_alive: "2m",
        options: { num_ctx: contextLength, temperature: 0.3 },
        messages: [
          { role: "system", content: systemInstructions },
          ...prior,
          { role: "user", content: `${context ? `${context}\n\n` : ""}PETICIÓN DEL USUARIO:\n${message.trim()}` }
        ]
      })
    });
    const data = await response.json();
    if (!response.ok) return res.status(response.status).json({ error: data?.error || "No se pudo contactar Ollama." });
    const answer = data.message?.content || "No se recibió texto.";
    res.json({ answer });
  } catch (error) {
    res.status(502).json({ error: "No se pudo conectar a Ollama en el Mac mini.", detail: error.message });
  }
});

app.post("/api/edit", async (req, res) => {
  const { message, documentText = "", selectionText = "", history = [] } = req.body || {};
  if (typeof message !== "string" || !message.trim()) return res.status(400).json({ error: "Describe el cambio que quieres hacer." });
  const context = makeContext(documentText, selectionText);
  const prior = Array.isArray(history) ? history.slice(-4).map(({ role, content }) => ({ role, content: String(content).slice(0, 3000) })) : [];
  try {
    const response = await fetch(`${ollamaBaseUrl}/api/chat`, {
      method: "POST", headers: { "Content-Type": "application/json" },
      body: JSON.stringify({ model: ollamaModel, stream: false, think: false, format: "json", keep_alive: "2m", options: { num_ctx: contextLength, temperature: 0.1 }, messages: [
        { role: "system", content: editInstructions }, ...prior,
        { role: "user", content: `${context ? `${context}\n\n` : ""}CAMBIO SOLICITADO:\n${message.trim()}` }
      ] })
    });
    const data = await response.json();
    if (!response.ok) return res.status(response.status).json({ error: data?.error || "No se pudo contactar Ollama." });
    let plan;
    try { plan = JSON.parse(data.message?.content || ""); }
    catch { return res.status(502).json({ error: "El modelo no devolvió un plan de edición válido. Inténtalo de nuevo." }); }
    res.json(sanitizeEditPlan(plan, { hasSelection: Boolean(selectionText.trim()), context }));
  } catch (error) { res.status(502).json({ error: "No se pudo conectar a Ollama en el Mac mini.", detail: error.message }); }
});

app.listen(port, bindHost, () => console.log(`Gateway local listo en http://${bindHost}:${port} (Ollama: ${ollamaModel})`));
