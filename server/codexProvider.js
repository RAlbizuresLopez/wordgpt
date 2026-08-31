import { spawn } from "node:child_process";
import { mkdtemp, rm, writeFile } from "node:fs/promises";
import { tmpdir } from "node:os";
import path from "node:path";

const CODEX_TIMEOUT_MS = 60000;

function parseFinalAgentMessage(stdout) {
  let finalText = null;
  for (const line of stdout.split("\n")) {
    const trimmed = line.trim();
    if (!trimmed) continue;
    let event;
    try { event = JSON.parse(trimmed); } catch { continue; }
    if (event.type === "item.completed" && event.item?.type === "agent_message" && typeof event.item.text === "string") {
      finalText = event.item.text;
    }
  }
  return finalText;
}

export async function runCodexExec(prompt, { schema } = {}) {
  const scratchDir = await mkdtemp(path.join(tmpdir(), "word-gpt-codex-"));
  try {
    const args = ["exec", "-C", scratchDir, "--skip-git-repo-check", "--sandbox", "read-only", "--json"];
    if (schema) {
      const schemaFile = path.join(scratchDir, "schema.json");
      await writeFile(schemaFile, JSON.stringify(schema));
      args.push("--output-schema", schemaFile);
    }
    args.push("-");

    const result = await new Promise((resolve, reject) => {
      const child = spawn("codex", args, { stdio: ["pipe", "pipe", "pipe"] });
      let stdout = "", stderr = "";
      const timeout = setTimeout(() => {
        child.kill("SIGKILL");
        reject(new Error("Codex CLI no respondió a tiempo."));
      }, CODEX_TIMEOUT_MS);
      child.stdout.on("data", (chunk) => { stdout += chunk; });
      child.stderr.on("data", (chunk) => { stderr += chunk; });
      child.on("error", (error) => {
        clearTimeout(timeout);
        if (error.code === "ENOENT") reject(new Error("No se encontró Codex CLI. Instálalo y ejecuta `codex login` con tu cuenta de ChatGPT."));
        else reject(error);
      });
      child.on("close", (code) => {
        clearTimeout(timeout);
        resolve({ code, stdout, stderr });
      });
      child.stdin.write(prompt);
      child.stdin.end();
    });

    if (result.code !== 0) {
      throw new Error(`Codex CLI terminó con un error (código ${result.code}). ¿Está autenticado? Ejecuta \`codex login\`. Detalle: ${result.stderr.trim().slice(0, 500)}`);
    }
    const finalText = parseFinalAgentMessage(result.stdout);
    if (finalText == null) throw new Error("Codex CLI no devolvió una respuesta final.");
    return finalText;
  } finally {
    await rm(scratchDir, { recursive: true, force: true }).catch(() => {});
  }
}
