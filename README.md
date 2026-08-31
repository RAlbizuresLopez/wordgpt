# Word GPT

Complemento de Microsoft Word que lee el documento actual, entiende la selección y aplica cambios de contenido, formato y estilo directamente en Word — como si chatearas con ChatGPT, pero con tu documento siempre a la mano. Corre por completo en tu propia laptop, usando **Codex CLI** con tu propia suscripción de ChatGPT.

## Qué incluye

- Panel de edición en español con vista previa y botón explícito para aplicar o descartar cambios.
- Contexto del documento completo (limitado a 24 000 caracteres) y de la selección activa.
- Reemplazos e inserciones anclados a fragmentos exactos del documento; cada ancla se aplica solo a la primera coincidencia. En documentos vacíos puede crear el contenido en el cursor.
- Formato de texto y párrafo: negrita, cursiva, subrayado, fuente, tamaño, color, resaltado, alineación, sangrías, espaciado e interlineado.
- Herramientas estructurales del documento: crear o sustituir encabezados y pies en todas las secciones, insertar tablas y saltos de página.
- Todo corre localmente: un pequeño servidor en tu laptop llama a Codex CLI, que ya está autenticado con tu cuenta de ChatGPT. No hay servidores remotos, VPN ni infraestructura compartida que mantener.

## Puesta en marcha

Requiere `pnpm`, Node.js 20+, Word para Microsoft 365 y [Codex CLI](https://developers.openai.com/codex/cli) instalado en tu misma computadora.

```bash
pnpm install
# Autentica Codex CLI con tu cuenta de ChatGPT (una sola vez):
codex login
pnpm run certs
pnpm run dev
```

En otra terminal, valida e instala el manifiesto en Word:

```bash
pnpm run validate:manifest
pnpm run sideload
```

En macOS, si el sideload no abre Word automáticamente, usa **Insertar → Complementos → Mis complementos → Cargar mis complementos** y selecciona `manifest.xml`.

Para uso diario sin herramientas de desarrollo, genera y sirve el panel desde el mismo servidor:

```bash
pnpm run build
pnpm start
```

## Editar un documento

Selecciona una parte si quieres limitar el alcance y escribe una instrucción como: "reescribe esta selección con tono más claro", "después de la introducción añade un párrafo de conclusiones", o "resalta en amarillo las conclusiones y ponlas en negrita". El complemento presenta el número de cambios previstos. Revisa el resumen y pulsa **Aplicar cambios**; Word conserva su historial normal de deshacer.

El agente usa una capa interna de herramientas de Word: no requiere instalar un servidor MCP ni permisos adicionales. Ya puedes pedir, por ejemplo: "pon `Confidencial` como encabezado", "añade un pie con el número de expediente", "inserta una tabla con Nombre, Fecha y Observaciones", o "inserta un salto de página al final". Imágenes, formas, comentarios y controles avanzados se añadirán como herramientas específicas en la siguiente expansión; no todas las acciones de la cinta de Word están expuestas todavía.

## Privacidad

Cada consulta (el texto del documento, la selección activa y tu mensaje) se envía a OpenAI a través de Codex CLI, usando tu propia sesión de ChatGPT — igual que si escribieras directamente en ChatGPT. Nada se guarda en un servidor propio ni se comparte con nadie más; todo el procesamiento local ocurre en tu misma laptop. Revisa las obligaciones de privacidad y consentimiento aplicables a los documentos que edites; el modelo no debe producir diagnósticos, tratamientos ni recomendaciones clínicas.

## Producción

`server/index.js` sirve tanto el panel (`dist/`, tras `pnpm run build`) como los endpoints `/api/chat` y `/api/edit` desde un único proceso local por HTTPS (usando el mismo certificado de `pnpm run certs`), escuchando solo en `127.0.0.1`. No hace falta exponerlo a la red: cada persona lo corre en su propia máquina, junto a Word.

## Instalar en una computadora nueva (macOS)

Cada máquina necesita su propia confianza de certificado y su propia sesión de Codex — nada de esto se comparte entre computadoras ni viaja con el proyecto:

1. `pnpm install`
2. `codex login` (una vez, con la cuenta de ChatGPT de esa persona)
3. `pnpm run certs` (una vez; instala y confía el certificado local de `https://localhost` en el llavero de esa Mac)
4. `pnpm run build && pnpm start` para uso diario, o `pnpm run dev` si vas a seguir editando el código
5. `pnpm run sideload` (o carga manual del `manifest.xml` vía **Insertar → Complementos → Mis complementos**) para que Word registre el complemento

Si alguna vez el panel se ve desactualizado después de cambiar el código, primero prueba cerrando y reabriendo el panel; si sigue igual, sube el número de `<Version>` en `manifest.xml` — Word cachea el complemento por Id+Versión y a veces no vuelve a pedir el HTML/JS hasta que cambia la versión.

## Windows

Requiere Node.js 20+ ([nodejs.org](https://nodejs.org)), `pnpm` (`corepack enable` o `npm install -g pnpm`) y Codex CLI (`npm install -g @openai/codex`).

```powershell
codex login
powershell -ExecutionPolicy Bypass -File scripts\windows\setup.ps1
```

`setup.ps1` encadena todo lo automatizable: instala dependencias, instala y confía el certificado local, genera el panel, registra una **tarea programada** (`WordGPT-Gateway`) para que el gateway arranque solo — sin ventana visible — cada vez que inicias sesión en Windows (y se reintenta sola si el proceso se cae), y por último registra el complemento en Word. `codex login` no se puede automatizar porque necesita abrir el navegador para iniciar sesión.

Para quitar el arranque automático más adelante: `powershell -ExecutionPolicy Bypass -File scripts\windows\uninstall-startup-task.ps1`.

**Nota:** este flujo de Windows no se ha probado todavía en una máquina Windows real (el proyecto se validó hoy en macOS). Si algo falla en la primera instalación, revisa esto antes que nada:

- **Codex CLI no responde o se comporta raro**: el sandbox de Codex CLI en Windows todavía es experimental según OpenAI. Pruébalo aislado primero: `codex exec -C <una carpeta vacía cualquiera> --skip-git-repo-check --sandbox read-only --json "Responde OK"`. Si eso falla, avisa antes de asumir que el resto funcionará.
- **El panel muestra un error de certificado**: `pnpm run certs` instala el certificado solo para tu usuario de Windows. Si WebView2 (el motor de Word en Windows) no lo acepta, corre desde una PowerShell como Administrador: `pnpm exec office-addin-dev-certs install --machine`.
- **El complemento no aparece en Word**: si `pnpm run sideload` falla, carga `manifest.xml` manualmente desde **Insertar → Complementos → Mis complementos → Cargar mis complementos**.
