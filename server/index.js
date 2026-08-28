import "dotenv/config";
import express from "express";
import { lookup } from "node:dns/promises";

const app = express();
const port = Number(process.env.PORT || 3001);
const bindHost = process.env.BIND_HOST || "127.0.0.1";
app.use(express.json({ limit: "1mb" }));
app.use(express.static("dist"));

const ollamaBaseUrl = (process.env.OLLAMA_BASE_URL || "http://127.0.0.1:11434").replace(/\/$/, "");
const ollamaModel = process.env.OLLAMA_MODEL || "qwen3:14b";
const contextLength = Number(process.env.OLLAMA_CONTEXT_LENGTH || 8192);
const searxngBaseUrl = (process.env.SEARXNG_BASE_URL || "http://searxng:8080").replace(/\/$/, "");
const allowedTailscaleLogins = new Set(
  (process.env.ALLOWED_TAILSCALE_USER_LOGINS || "")
    .split(",")
    .map(login => login.trim().toLowerCase())
    .filter(Boolean)
);
const systemInstructions = `Eres un asistente experto integrado en Microsoft Word. Responde en el idioma del usuario. Usa el contexto del documento únicamente para ayudar con su petición. Si propones texto para insertar, entrégalo listo para pegar y no inventes información que no esté sustentada por el documento. No proporciones diagnósticos, tratamientos ni recomendaciones clínicas.`;
const editInstructions = `Eres el motor de edición de Word. Devuelve ÚNICAMENTE JSON válido: {"summary":"breve","operations":[...]}. Operaciones de texto: replace/find/replacement, insert_before o insert_after/find/text, replace_selection/text, insert_at_selection/text, insert_at_cursor/text para documento vacío. Operaciones de formato: format con target selection o find; format_paragraph con paragraph (1,2,3...); y format_document para todo el documento. Las operaciones de formato aceptan "font" y/o "paragraphFormat". font: bold, italic, underline (none|single|double), strikethrough, color, highlightColor, size, name, allCaps, smallCaps, superscript, subscript. paragraphFormat: alignment (left|centered|right|justified), leftIndent, rightIndent, firstLineIndent, spaceBefore, spaceAfter, lineSpacing, keepTogether, keepWithNext, widowControl. Herramientas estructurales: set_header/text/kind (primary|first_page|even_pages), set_footer/text/kind, insert_table/values (matriz de texto rectangular, máximo 20 filas x 12 columnas, location cursor|document_end), insert_page_break/location (cursor|document_end). set_header y set_footer sustituyen ese encabezado o pie en todas las secciones. Usa format_document para el formato base del resto del documento y ponlo ANTES de una selección especial, para que la selección la sobrescriba. Usa format_paragraph para "segundo párrafo". Máximo 10 operaciones. find debe aparecer exactamente en el contexto y medir máximo 240 caracteres. Si hay selección, priorízala. No inventes fragmentos. No hagas diagnósticos, tratamientos ni recomendaciones clínicas.`;

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

const privateIp = (address) => /^(127\.|10\.|0\.|169\.254\.|192\.168\.|172\.(1[6-9]|2\d|3[01])\.)/.test(address) || address === "::1" || address.startsWith("fc") || address.startsWith("fd") || address.startsWith("fe80:");
async function assertPublicHttps(url) {
  if (url.protocol !== "https:" || url.username || url.password || (url.port && url.port !== "443")) throw new Error("Solo se pueden consultar páginas HTTPS públicas.");
  const addresses = await lookup(url.hostname, { all: true });
  if (!addresses.length || addresses.some(({ address }) => privateIp(address))) throw new Error("La URL no apunta a una dirección pública.");
}
function htmlToText(value) { return value.replace(/<script\b[^>]*>[\s\S]*?<\/script>/gi, " ").replace(/<style\b[^>]*>[\s\S]*?<\/style>/gi, " ").replace(/<[^>]+>/g, " ").replace(/&nbsp;/gi, " ").replace(/&amp;/gi, "&").replace(/&quot;/gi, '"').replace(/\s+/g, " ").trim(); }
async function readPublicPage(rawUrl) {
  let url = new URL(rawUrl);
  for (let redirects = 0; redirects < 4; redirects += 1) {
    await assertPublicHttps(url);
    const response = await fetch(url, { redirect: "manual", signal: AbortSignal.timeout(8000), headers: { "User-Agent": "WordGPT-Research/1.0" } });
    if ([301, 302, 303, 307, 308].includes(response.status)) { url = new URL(response.headers.get("location"), url); continue; }
    if (!response.ok) throw new Error(`La página respondió ${response.status}.`);
    if (!/text\/(html|plain)/i.test(response.headers.get("content-type") || "")) throw new Error("La URL no contiene texto web.");
    const reader = response.body.getReader(); let bytes = 0, value = "";
    while (bytes < 180000) { const { done, value: chunk } = await reader.read(); if (done) break; bytes += chunk.byteLength; value += new TextDecoder().decode(chunk, { stream: true }); }
    reader.cancel().catch(() => {}); return htmlToText(value).slice(0, 12000);
  }
  throw new Error("Demasiadas redirecciones.");
}
async function searchWeb(query) {
  const response = await fetch(`${searxngBaseUrl}/search?${new URLSearchParams({ q: query, format: "json", language: "es-ES", safesearch: "1" })}`, { signal: AbortSignal.timeout(12000) });
  if (!response.ok) throw new Error("El servicio de búsqueda no está disponible.");
  const data = await response.json();
  return (data.results || []).filter(({ url }) => typeof url === "string" && url.startsWith("https://")).slice(0, 4).map(({ title, url, content }) => ({ title: String(title || "Fuente"), url, content: String(content || "") }));
}

function sanitizeEditPlan(value, { hasSelection, context }) {
  const allowedTypes = new Set(["replace", "insert_after", "insert_before", "replace_selection", "insert_at_selection", "insert_at_cursor", "format", "format_paragraph", "format_document", "set_header", "set_footer", "insert_table", "insert_page_break"]);
  const allowedFonts = new Set(["bold", "italic", "underline", "strikethrough", "color", "highlightColor", "size", "name", "allCaps", "smallCaps", "superscript", "subscript"]);
  const allowedParagraphs = new Set(["alignment", "leftIndent", "rightIndent", "firstLineIndent", "spaceBefore", "spaceAfter", "lineSpacing", "keepTogether", "keepWithNext", "widowControl"]);
  const operations = Array.isArray(value?.operations) ? value.operations.slice(0, 10) : [];
  const clean = operations.flatMap((operation) => {
    if (!operation) return [];
    const type = operation.type === "format" && ["document", "all", "body"].includes(String(operation.target).toLowerCase()) ? "format_document" : operation.type;
    if (!allowedTypes.has(type)) return [];
    const find = typeof operation.find === "string" ? operation.find.trim().slice(0, 240) : "";
    const text = typeof operation.text === "string" ? operation.text.slice(0, 12000) : "";
    const replacement = typeof operation.replacement === "string" ? operation.replacement.slice(0, 12000) : "";
    if (["set_header", "set_footer"].includes(type)) {
      if (!text) return [];
      const kind = ["primary", "first_page", "even_pages"].includes(String(operation.kind).toLowerCase()) ? String(operation.kind).toLowerCase() : "primary";
      return [{ type, text, kind }];
    }
    if (type === "insert_table") {
      const rawValues = Array.isArray(operation.values) ? operation.values : Array.isArray(operation.rows) ? operation.rows : [];
      const values = rawValues.slice(0, 20).map((row) => Array.isArray(row) ? row.slice(0, 12).map((cell) => String(cell ?? "").slice(0, 1000)) : []);
      const columns = Math.max(...values.map((row) => row.length), 0);
      if (!values.length || !columns) return [];
      const rectangularValues = values.map((row) => [...row, ...Array(Math.max(0, columns - row.length)).fill("")]);
      const location = ["cursor", "document_end"].includes(operation.location) ? operation.location : "cursor";
      return [{ type, values: rectangularValues, location }];
    }
    if (type === "insert_page_break") {
      const location = ["cursor", "document_end"].includes(operation.location) ? operation.location : "cursor";
      return [{ type, location }];
    }
    if (["replace", "insert_after", "insert_before"].includes(type) && !find) return [];
    if (["replace_selection", "insert_at_selection"].includes(type) && (!text || !hasSelection)) return [];
    if (type === "insert_at_cursor" && (!text || hasSelection)) return [];
    if (type === "replace" && typeof operation.replacement !== "string") return [];
    if (["format", "format_paragraph", "format_document"].includes(type)) {
      if (type === "format" && ((operation.target === "selection" && !hasSelection) || (operation.target !== "selection" && !find))) return [];
      const paragraph = Number(operation.paragraph);
      if (type === "format_paragraph" && (!Number.isInteger(paragraph) || paragraph < 1 || paragraph > 500)) return [];
      const font = Object.fromEntries(Object.entries(operation.font || {}).filter(([key, item]) => {
        if (!allowedFonts.has(key)) return false;
        if (["bold", "italic", "strikethrough", "allCaps", "smallCaps", "superscript", "subscript"].includes(key)) return typeof item === "boolean";
        if (key === "size") return typeof item === "number" && item >= 6 && item <= 96;
        if (key === "underline") return ["none", "single", "double"].includes(String(item).toLowerCase());
        return typeof item === "string" && item.length <= 80;
      }));
      const paragraphFormat = Object.fromEntries(Object.entries(operation.paragraphFormat || operation.paragraph_style || (typeof operation.paragraph === "object" ? operation.paragraph : {}) || {}).filter(([key, item]) => {
        if (!allowedParagraphs.has(key)) return false;
        if (["keepTogether", "keepWithNext", "widowControl"].includes(key)) return typeof item === "boolean";
        if (key === "alignment") return ["left", "centered", "right", "justified"].includes(String(item).toLowerCase());
        return typeof item === "number" && item >= -144 && item <= 144;
      }));
      if ((!Object.keys(font).length && !Object.keys(paragraphFormat).length) || (find && !context.includes(find))) return [];
      const changes = { ...(Object.keys(font).length ? { font } : {}), ...(Object.keys(paragraphFormat).length ? { paragraphFormat } : {}) };
      if (type === "format_paragraph") return [{ type, paragraph, ...changes }];
      if (type === "format_document") return [{ type, ...changes }];
      return [{ type, ...(find ? { find } : { target: "selection" }), ...changes }];
    }
    if (find && !context.includes(find)) return [];
    return [{ type, ...(find ? { find } : {}), ...(type === "replace" ? { replacement } : { text }) }];
  });
  return { summary: typeof value?.summary === "string" ? value.summary.slice(0, 500) : "Plan de edición preparado.", operations: clean };
}

function fallbackParagraphFormat(message) {
  const text = message.normalize("NFD").replace(/\p{Diacritic}/gu, "").toLowerCase();
  const positions = [[1, /(?:primer|primero|primera|1er|1ro|1)\s+parrafo/], [2, /(?:segundo|segunda|2do|2da|2)\s+parrafo/], [3, /(?:tercer|tercero|tercera|3er|3ro|3)\s+parrafo/]];
  const paragraph = positions.find(([, pattern]) => pattern.test(text))?.[0];
  if (!paragraph) return null;
  const font = {};
  if (/\bazul\b/.test(text)) font.color = "#0000FF";
  if (/\brojo\b/.test(text)) font.color = "#FF0000";
  if (/\bverde\b/.test(text)) font.color = "#008000";
  if (/\bnegrita\b/.test(text)) font.bold = !/(?:sin|quitar)\s+negrita/.test(text);
  if (/\bcursiva\b/.test(text)) font.italic = !/(?:sin|quitar)\s+cursiva/.test(text);
  if (/\bresalta(?:r|do)?\b/.test(text) && /amarillo/.test(text)) font.highlightColor = "yellow";
  return Object.keys(font).length ? { type: "format_paragraph", paragraph, font } : null;
}

function styleFromText(text) {
  const normalized = text.normalize("NFD").replace(/\p{Diacritic}/gu, "").toLowerCase();
  const font = {}, paragraphFormat = {};
  const colors = [["verde", "#008000"], ["rojo", "#FF0000"], ["azul", "#0000FF"], ["negro", "#000000"], ["morado", "#800080"], ["violeta", "#800080"]];
  const color = colors.find(([name]) => new RegExp(`\\b${name}\\b`).test(normalized));
  if (color) font.color = color[1];
  if (/\bnegrita\b/.test(normalized)) font.bold = true;
  if (/\bcursiva\b/.test(normalized)) font.italic = true;
  if (/\bsubrayad/.test(normalized)) font.underline = "single";
  if (/\btachad/.test(normalized)) font.strikethrough = true;
  if (/\bcentrad/.test(normalized)) paragraphFormat.alignment = "centered";
  if (/\bjustificad/.test(normalized)) paragraphFormat.alignment = "justified";
  if (/\balinead[oa]?\s+a la derecha|\bderecha\b/.test(normalized)) paragraphFormat.alignment = "right";
  if (/\balinead[oa]?\s+a la izquierda|\bizquierda\b/.test(normalized)) paragraphFormat.alignment = "left";
  if (/\binterlineado\b/.test(normalized)) {
    if (/doble|2\s*(?:x|veces)?/.test(normalized)) paragraphFormat.lineSpacing = 24;
    else if (/1[,.]5|uno y medio|media/.test(normalized)) paragraphFormat.lineSpacing = 18;
    else paragraphFormat.lineSpacing = 12;
  }
  if (/una pagina|compact|reduc/.test(normalized)) {
    paragraphFormat.lineSpacing = 11;
    paragraphFormat.spaceBefore = 0;
    paragraphFormat.spaceAfter = 0;
  }
  return { font, paragraphFormat };
}

function fallbackSelectionAndDocumentFormat(message, hasSelection) {
  const normalized = message.normalize("NFD").replace(/\p{Diacritic}/gu, "").toLowerCase();
  const restStart = normalized.search(/\b(resto|documento restante|todo el documento|demas texto)\b/);
  const selectionPart = restStart >= 0 ? message.slice(0, restStart) : message;
  const documentPart = restStart >= 0 ? message.slice(restStart) : "";
  const selectionStyle = styleFromText(selectionPart), documentStyle = styleFromText(documentPart);
  const hasStyle = (style) => Object.keys(style.font).length || Object.keys(style.paragraphFormat).length;
  const operations = [];
  if (hasStyle(documentStyle)) operations.push({ type: "format_document", ...documentStyle });
  if (hasSelection && hasStyle(selectionStyle)) operations.push({ type: "format", target: "selection", ...selectionStyle });
  if (!hasSelection && hasStyle(selectionStyle) && !operations.length) operations.push({ type: "format_document", ...selectionStyle });
  return operations;
}

function fallbackStructuralOperation(message) {
  const normalized = message.normalize("NFD").replace(/\p{Diacritic}/gu, "").toLowerCase();
  const location = /\b(final|terminar|ultimo)\b/.test(normalized) ? "document_end" : "cursor";
  if (/\btabla\b/.test(normalized)) {
    const columnsMatch = message.match(/columnas?\s+(.+?)(?:,?\s+con\s+(?:\d+|una|dos|tres|cuatro|cinco|seis|siete|ocho|nueve|diez)\s+filas?|\.|$)/i);
    const columns = (columnsMatch?.[1] || "Columna 1, Columna 2")
      .split(/\s*,\s*|\s+y\s+/i).map((cell) => cell.trim()).filter(Boolean).slice(0, 12);
    if (!columns.length) return null;
    const words = { una: 1, dos: 2, tres: 3, cuatro: 4, cinco: 5, seis: 6, siete: 7, ocho: 8, nueve: 9, diez: 10 };
    const rowsMatch = normalized.match(/\b(\d+|una|dos|tres|cuatro|cinco|seis|siete|ocho|nueve|diez)\s+filas?/);
    const emptyRows = rowsMatch ? (Number(rowsMatch[1]) || words[rowsMatch[1]] || 0) : 2;
    return { type: "insert_table", values: [columns, ...Array.from({ length: Math.min(emptyRows, 19) }, () => Array(columns.length).fill(""))], location };
  }
  if (/salto\s+de\s+pagina/.test(normalized)) return { type: "insert_page_break", location };
  const headerFooter = /\b(encabezado|pie(?:\s+de\s+pagina)?)\b/.exec(normalized)?.[1];
  const quoted = message.match(/[“"'`]([^“"'`]+)[”"'`]/)?.[1]?.trim();
  if (headerFooter && quoted) return { type: headerFooter.startsWith("pie") ? "set_footer" : "set_header", text: quoted, kind: "primary" };
  return null;
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
    const safePlan = sanitizeEditPlan(plan, { hasSelection: Boolean(selectionText.trim()), context });
    const structuralFallback = fallbackStructuralOperation(message);
    const requestedFallback = fallbackSelectionAndDocumentFormat(message, Boolean(selectionText.trim()));
    if (structuralFallback) {
      // Las estructuras de Word no deben degradarse a búsquedas de texto generadas por el modelo.
      safePlan.operations = [structuralFallback];
    } else if (!safePlan.operations.length) {
      const fallback = fallbackParagraphFormat(message);
      safePlan.operations = fallback ? [fallback] : requestedFallback;
    } else if (requestedFallback.length) {
      for (const requested of requestedFallback) {
        const existing = safePlan.operations.find((operation) => operation.type === requested.type && (requested.type !== "format" || operation.target === requested.target));
        if (existing) {
          existing.font = { ...existing.font, ...requested.font };
          existing.paragraphFormat = { ...existing.paragraphFormat, ...requested.paragraphFormat };
        } else if (requested.type === "format_document") safePlan.operations.unshift(requested);
        else safePlan.operations.push(requested);
      }
      safePlan.operations = safePlan.operations.slice(0, 10);
    }
    console.info("Word GPT edit plan", safePlan.operations.map(({ type, location, kind }) => ({ type, location, kind })));
    res.json(safePlan);
  } catch (error) { res.status(502).json({ error: "No se pudo conectar a Ollama en el Mac mini.", detail: error.message }); }
});

app.post("/api/research", async (req, res) => {
  const { message, documentText = "", selectionText = "" } = req.body || {};
  if (typeof message !== "string" || !message.trim()) return res.status(400).json({ error: "Escribe qué deseas investigar." });
  try {
    const results = await searchWeb(message.trim());
    if (!results.length) return res.status(404).json({ error: "No encontré fuentes públicas para esa consulta." });
    const researched = await Promise.all(results.slice(0, 3).map(async (result) => ({ ...result, text: await readPublicPage(result.url).catch(() => result.content) })));
    const sources = researched.map(({ title, url }) => ({ title, url }));
    const context = makeContext(documentText, selectionText);
    const sourceText = researched.map((source, index) => `[${index + 1}] ${source.title}\n${source.url}\n${source.text}`).join("\n\n");
    const response = await fetch(`${ollamaBaseUrl}/api/chat`, {
      method: "POST", headers: { "Content-Type": "application/json" },
      body: JSON.stringify({ model: ollamaModel, stream: false, think: false, keep_alive: "2m", options: { num_ctx: contextLength, temperature: 0.2 }, messages: [
        { role: "system", content: "Responde en español usando solo las fuentes proporcionadas. Incluye referencias [1], [2] junto a cada afirmación relevante. No inventes datos ni hagas diagnósticos, tratamientos o recomendaciones clínicas." },
        { role: "user", content: `${context ? `${context}\n\n` : ""}CONSULTA:\n${message.trim()}\n\nFUENTES:\n${sourceText}` }
      ] })
    });
    const data = await response.json();
    if (!response.ok) return res.status(response.status).json({ error: data?.error || "Ollama no pudo analizar las fuentes." });
    res.json({ answer: data.message?.content || "No se recibió respuesta.", sources });
  } catch (error) { res.status(502).json({ error: error.message || "No se pudo completar la investigación web." }); }
});

app.listen(port, bindHost, () => console.log(`Gateway local listo en http://${bindHost}:${port} (Ollama: ${ollamaModel})`));
