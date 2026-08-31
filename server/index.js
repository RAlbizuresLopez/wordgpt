import "dotenv/config";
import express from "express";
import https from "node:https";
import { getHttpsServerOptions } from "office-addin-dev-certs";
import { runCodexExec } from "./codexProvider.js";

const app = express();
const port = Number(process.env.PORT || 3001);
const bindHost = process.env.BIND_HOST || "127.0.0.1";
app.use(express.json({ limit: "1mb" }));
app.use(express.static("dist"));

const systemInstructions = `Eres un asistente experto integrado en Microsoft Word. Responde en el idioma del usuario. Usa el contexto del documento únicamente para ayudar con su petición. Si propones texto para insertar, entrégalo listo para pegar y no inventes información que no esté sustentada por el documento. No proporciones diagnósticos, tratamientos ni recomendaciones clínicas.`;
const editInstructions = `Eres el asistente de Word integrado en un panel de chat: el campo "summary" es tu respuesta al usuario y se muestra tal cual, así que contéstale ahí de forma completa y natural (como lo haría ChatGPT), no la resumas ni la recortes. Devuelve ÚNICAMENTE JSON válido: {"summary":"tu respuesta completa","operations":[...]}. Si el usuario solo pregunta o pide una opinión/análisis (nada que cambie el documento), responde completo en "summary" y deja "operations" vacío. Si pide agregar, escribir, generar o redactar contenido en el documento, SIEMPRE debes incluir la operación de inserción correspondiente con el texto completo (insert_at_cursor si el documento está vacío, insert_after/insert_before/replace usando un "find" del contexto como referencia, o replace_selection/insert_at_selection si hay selección) — no basta con describir o mostrar ese texto solo en "summary", tiene que quedar también en una operación para que se inserte de verdad en el documento. Operaciones de texto: replace/find/replacement, insert_before o insert_after/find/text, replace_selection/text, insert_at_selection/text, insert_at_cursor/text para documento vacío. Operaciones de formato: format con target selection o find; format_paragraph con paragraph (1,2,3...); y format_document para todo el documento. Las operaciones de formato aceptan "font" y/o "paragraphFormat". font: bold, italic, underline (none|single|double), strikethrough, color, highlightColor, size, name, allCaps, smallCaps, superscript, subscript. paragraphFormat: alignment (left|centered|right|justified), leftIndent, rightIndent, firstLineIndent, spaceBefore, spaceAfter, lineSpacing, keepTogether, keepWithNext, widowControl. Herramientas estructurales: set_header/text/kind (primary|first_page|even_pages), set_footer/text/kind, insert_table/values (matriz de texto rectangular, máximo 20 filas x 12 columnas, location cursor|document_end), insert_page_break/location (cursor|document_end). set_header y set_footer sustituyen ese encabezado o pie en todas las secciones. Usa format_document para el formato base del resto del documento y ponlo ANTES de una selección especial, para que la selección la sobrescriba. Usa format_paragraph para "segundo párrafo". Máximo 10 operaciones. find debe aparecer exactamente en el contexto y medir máximo 240 caracteres. Si hay selección, priorízala. No inventes fragmentos que no pediste generar. No hagas diagnósticos, tratamientos ni recomendaciones clínicas. Tú decides qué operaciones usar: aplica tú mismo cada cambio de formato o estilo que se te pida, con los valores correctos, sin dejar nada implícito para que otro sistema lo adivine.`;
// Codex exige el modo "strict" de OpenAI para --output-schema: additionalProperties:false en
// cada objeto anidado, y TODAS las propiedades listadas en required (los campos opcionales se
// expresan como ["tipo","null"], no como ausentes). sanitizeEditPlan ya trata null como ausente.
const editPlanSchema = {
  type: "object",
  properties: {
    summary: { type: "string" },
    operations: {
      type: "array",
      items: {
        type: "object",
        properties: {
          type: {
            type: "string",
            enum: ["replace", "insert_after", "insert_before", "replace_selection", "insert_at_selection", "insert_at_cursor", "format", "format_paragraph", "format_document", "set_header", "set_footer", "insert_table", "insert_page_break"]
          },
          find: { type: ["string", "null"] },
          replacement: { type: ["string", "null"] },
          text: { type: ["string", "null"] },
          target: { type: ["string", "null"] },
          paragraph: { type: ["integer", "null"] },
          location: { type: ["string", "null"] },
          kind: { type: ["string", "null"] },
          values: { type: ["array", "null"], items: { type: "array", items: { type: "string" } } },
          font: {
            type: ["object", "null"],
            properties: {
              bold: { type: ["boolean", "null"] }, italic: { type: ["boolean", "null"] }, underline: { type: ["string", "null"] }, strikethrough: { type: ["boolean", "null"] },
              color: { type: ["string", "null"] }, highlightColor: { type: ["string", "null"] }, size: { type: ["number", "null"] }, name: { type: ["string", "null"] },
              allCaps: { type: ["boolean", "null"] }, smallCaps: { type: ["boolean", "null"] }, superscript: { type: ["boolean", "null"] }, subscript: { type: ["boolean", "null"] }
            },
            required: ["bold", "italic", "underline", "strikethrough", "color", "highlightColor", "size", "name", "allCaps", "smallCaps", "superscript", "subscript"],
            additionalProperties: false
          },
          paragraphFormat: {
            type: ["object", "null"],
            properties: {
              alignment: { type: ["string", "null"] }, leftIndent: { type: ["number", "null"] }, rightIndent: { type: ["number", "null"] }, firstLineIndent: { type: ["number", "null"] },
              spaceBefore: { type: ["number", "null"] }, spaceAfter: { type: ["number", "null"] }, lineSpacing: { type: ["number", "null"] },
              keepTogether: { type: ["boolean", "null"] }, keepWithNext: { type: ["boolean", "null"] }, widowControl: { type: ["boolean", "null"] }
            },
            required: ["alignment", "leftIndent", "rightIndent", "firstLineIndent", "spaceBefore", "spaceAfter", "lineSpacing", "keepTogether", "keepWithNext", "widowControl"],
            additionalProperties: false
          }
        },
        required: ["type", "find", "replacement", "text", "target", "paragraph", "location", "kind", "values", "font", "paragraphFormat"],
        additionalProperties: false
      }
    }
  },
  required: ["summary", "operations"],
  additionalProperties: false
};

app.get("/api/health", async (_req, res) => {
  res.json({ ok: true, provider: "codex" });
});

function makeContext(documentText, selectionText) {
  return [
    selectionText && `SELECCIÓN ACTUAL:\n${selectionText}`,
    documentText && `DOCUMENTO (puede estar truncado):\n${documentText}`
  ].filter(Boolean).join("\n\n");
}

function flattenMessages(messages) {
  const [system, ...rest] = messages;
  const history = rest.slice(0, -1);
  const last = rest[rest.length - 1];
  const sections = [
    "No ejecutes comandos de shell ni edites archivos; responde solo con el texto/JSON pedido usando el contexto de abajo.",
    `[INSTRUCCIONES DEL SISTEMA]\n${system.content}`
  ];
  if (history.length) {
    sections.push(`[HISTORIAL RECIENTE]\n${history.map(({ role, content }) => `${role === "user" ? "Usuario" : "Asistente"}: ${content}`).join("\n")}`);
  }
  sections.push(`[PETICIÓN ACTUAL]\n${last.content}`);
  return sections.join("\n\n");
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
  return { summary: typeof value?.summary === "string" ? value.summary.slice(0, 8000) : "Plan de edición preparado.", operations: clean };
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
  if (/\btabla\b/.test(normalized) && !/tabla de conten/.test(normalized)) {
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
    const answer = await runCodexExec(flattenMessages([
      { role: "system", content: systemInstructions },
      ...prior,
      { role: "user", content: `${context ? `${context}\n\n` : ""}PETICIÓN DEL USUARIO:\n${message.trim()}` }
    ]));
    res.json({ answer });
  } catch (error) {
    res.status(502).json({ error: error.message || "No se pudo obtener respuesta de Codex." });
  }
});

app.post("/api/edit", async (req, res) => {
  const { message, documentText = "", selectionText = "", history = [] } = req.body || {};
  if (typeof message !== "string" || !message.trim()) return res.status(400).json({ error: "Describe el cambio que quieres hacer." });
  const context = makeContext(documentText, selectionText);
  const prior = Array.isArray(history) ? history.slice(-4).map(({ role, content }) => ({ role, content: String(content).slice(0, 3000) })) : [];
  const baseMessages = [
    { role: "system", content: editInstructions }, ...prior,
    { role: "user", content: `${context ? `${context}\n\n` : ""}CAMBIO SOLICITADO:\n${message.trim()}` }
  ];
  async function requestPlan(messages) {
    const text = await runCodexExec(flattenMessages(messages), { schema: editPlanSchema });
    return JSON.parse(text);
  }
  try {
    let plan;
    try { plan = await requestPlan(baseMessages); }
    catch {
      // Codex no devolvió JSON válido pese al esquema; le damos una segunda oportunidad
      // en vez de degradar directo a los heurísticos de texto.
      try {
        plan = await requestPlan([...baseMessages, { role: "user", content: "Tu respuesta anterior no era JSON válido. Devuelve SOLO el JSON del plan siguiendo el esquema, sin texto adicional." }]);
      } catch {
        return res.status(502).json({ error: "Codex no devolvió un plan de edición válido. Inténtalo de nuevo." });
      }
    }
    const safePlan = sanitizeEditPlan(plan, { hasSelection: Boolean(selectionText.trim()), context });
    const structuralFallback = fallbackStructuralOperation(message);
    const requestedFallback = fallbackSelectionAndDocumentFormat(message, Boolean(selectionText.trim()));
    if (structuralFallback) {
      // Las estructuras de Word (tabla/encabezado/pie/salto) no deben degradarse a búsquedas de
      // texto generadas por el modelo, pero el resto del plan (p. ej. formato pedido en la misma
      // instrucción) sí es responsabilidad del modelo y no debe descartarse.
      const structuralTypes = new Set(["insert_table", "set_header", "set_footer", "insert_page_break"]);
      safePlan.operations = safePlan.operations.filter((operation) => !structuralTypes.has(operation.type));
      safePlan.operations.push(structuralFallback);
    }
    if (!safePlan.operations.length) {
      const fallback = fallbackParagraphFormat(message);
      safePlan.operations = fallback ? [fallback] : requestedFallback;
    }
    // Si Codex ya produjo operaciones válidas, son definitivas: no se le agregan ni fusionan
    // adivinanzas de regex encima. Los heurísticos de arriba son solo una red de seguridad para
    // cuando el modelo no devolvió nada aplicable, no un "segundo opinión" sobre un plan ya bueno.
    console.info("Word GPT edit plan", safePlan.operations.map(({ type, location, kind }) => ({ type, location, kind })));
    res.json(safePlan);
  } catch (error) { res.status(502).json({ error: error.message || "No se pudo obtener un plan de edición de Codex." }); }
});

// El manifiesto siempre carga el panel por HTTPS. En "start" (uso diario, sin Vite) este mismo
// proceso sirve el panel directo, así que necesita el certificado local de confianza
// (`pnpm run certs`). En dev, Vite ya sirve el panel por HTTPS y solo llama a este proceso por
// HTTP interno para la API, así que no hace falta duplicar el certificado aquí.
const useHttps = process.env.HTTPS === "1" || process.env.HTTPS === "true";
if (useHttps) {
  const httpsOptions = await getHttpsServerOptions();
  https.createServer(httpsOptions, app).listen(port, bindHost, () => console.log(`Word GPT listo en https://${bindHost}:${port} (motor: Codex)`));
} else {
  app.listen(port, bindHost, () => console.log(`Word GPT listo en http://${bindHost}:${port} (motor: Codex)`));
}
