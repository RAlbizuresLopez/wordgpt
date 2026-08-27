import "dotenv/config";
import express from "express";

const app = express();
const port = Number(process.env.PORT || 3001);
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

app.post("/api/chat", async (req, res) => {
  const { message, documentText = "", selectionText = "", history = [] } = req.body || {};
  if (typeof message !== "string" || !message.trim()) return res.status(400).json({ error: "Escribe un mensaje." });

  const context = [
    selectionText && `SELECCIÓN ACTUAL:\n${selectionText}`,
    documentText && `DOCUMENTO (puede estar truncado):\n${documentText}`
  ].filter(Boolean).join("\n\n");
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

app.listen(port, "127.0.0.1", () => console.log(`Gateway local listo en http://127.0.0.1:${port} (Ollama: ${ollamaModel})`));
