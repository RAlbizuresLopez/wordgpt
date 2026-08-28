# Word GPT

Complemento de Microsoft Word que lee el documento actual, entiende la selección y prepara cambios directos al estilo de un agente de código. El modelo se ejecuta con Ollama en un Mac mini privado.

## Qué incluye

- Panel de edición en español con vista previa y botón explícito para aplicar o descartar cambios.
- Contexto del documento completo (limitado a 24 000 caracteres) y de la selección activa.
- Reemplazos e inserciones anclados a fragmentos exactos del documento; cada ancla se aplica solo a la primera coincidencia. En documentos vacíos puede crear el contenido en el cursor.
- Formato de texto: negrita, cursiva, fuente, tamaño, color y resaltado.
- Búsqueda web opcional con SearXNG autoalojado, fuentes citables y extracción limitada a páginas HTTPS públicas.
- Gateway en el Mac mini que llama a Ollama por `localhost`; la laptop de la usuaria no ejecuta modelos ni guarda claves.
- Acceso privado mediante Tailscale Serve, sin puertos públicos.

## Puesta en marcha

Requiere `pnpm`, Node.js 20+, Word para Microsoft 365 y un Mac mini con Ollama. La usuaria solo necesita Word y Tailscale.

```bash
pnpm install
cp .env.example .env
# Instala Ollama en el Mac mini y descarga el modelo recomendado:
ollama pull qwen3:14b
pnpm run certs
pnpm run dev
```

Para ejecutar el Mac mini sin herramientas de desarrollo, genera y sirve el panel desde el mismo gateway:

```bash
pnpm run build
pnpm start
```

### Servicio permanente con Docker

En macOS, Ollama debe permanecer instalado **nativamente** para aprovechar la aceleración Metal del M4. Solo el gateway se ejecuta en Docker; se comunica con Ollama mediante `host.docker.internal` y publica el puerto únicamente en `127.0.0.1`. Dentro del contenedor escucha en su interfaz interna; Docker no lo expone a la LAN.

```bash
# Instala y abre Docker Desktop una vez; activa “Start Docker Desktop when you log in”.
docker compose up -d --build
docker compose ps
```

El contenedor usa `restart: unless-stopped`: se recupera tras una caída y después de reiniciar Docker. Tailscale Serve sigue en el host y debe apuntar al puerto local del contenedor:

```bash
tailscale serve --bg --https=443 http://127.0.0.1:3000
```

Para actualizarlo tras cambios del repositorio, ejecuta `docker compose up -d --build`. Consulta los registros con `docker compose logs -f` y detén el servicio con `docker compose down`.

### Consultar Internet

Marca **Buscar en Internet** antes de enviar una consulta. La consulta escrita se envía al metabuscador SearXNG que corre dentro del Mac mini; el contenido del documento no se envía a buscadores ni a sitios externos. El gateway consulta hasta tres fuentes HTTPS públicas, limita redirecciones, tiempo y tamaño de cada descarga, y muestra enlaces a las fuentes en el panel antes de que decidas incorporar algo al documento.

SearXNG no publica ningún puerto en el Mac ni en Tailscale: únicamente el gateway puede alcanzarlo por la red interna de Docker. Las búsquedas sí salen desde el Mac mini hacia los motores de búsqueda y las páginas seleccionadas.

En otra terminal, valida e instala el manifiesto en Word:

```bash
pnpm run validate:manifest
pnpm run sideload
```

En macOS, si el sideload no abre Word automáticamente, usa **Insertar → Complementos → Mis complementos → Cargar mis complementos** y selecciona `manifest.xml`.

## Editar un documento

Selecciona una parte si quieres limitar el alcance y escribe una instrucción como: “reescribe esta selección con tono más claro”, “después de la introducción añade un párrafo de conclusiones”, o “resalta en amarillo las conclusiones y ponlas en negrita”. El complemento presenta el número de cambios previstos. Revisa el resumen y pulsa **Aplicar cambios**; Word conserva su historial normal de deshacer.

## Mac mini y Tailscale

Ejecuta el gateway únicamente en el Mac mini. Ollama continúa en `127.0.0.1:11434`; no lo enlaces a `0.0.0.0` ni abras su puerto. Cuando el gateway esté activo en el puerto 3000, publícalo **solo dentro de tu tailnet** con:

```bash
tailscale serve --bg --https=443 http://127.0.0.1:3000
```

Este proyecto está configurado para `https://macmini4pro.tail97c777.ts.net`. Define también `ALLOWED_TAILSCALE_USER_LOGINS` con los correos autorizados, separados por comas, y crea una ACL que solo permita esas identidades/dispositivos al Mac mini en el puerto 443. Así hay autorización doble: la red y el gateway. No uses `tailscale funnel`.

El contenido se manda únicamente al Mac mini cuando se hace una consulta. Revisa las obligaciones de privacidad y consentimiento aplicables a los documentos clínicos; el modelo no debe producir diagnósticos ni decisiones de tratamiento.

## Producción

Cuando pruebes el flujo, sirve el frontend y los endpoints `/api/chat` y `/api/edit` desde un único proceso HTTP en el Mac mini y publícalo mediante Tailscale Serve. El manifiesto ya usa el dominio privado de Tailscale; no lo publiques fuera de tu tailnet.

La integración utiliza la [API de chat de Ollama](https://docs.ollama.com/api/chat).
