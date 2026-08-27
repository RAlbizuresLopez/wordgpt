# Word GPT

Complemento de Microsoft Word con un panel de chat que lee el documento actual, entiende la selección y puede insertar o sustituir texto generado. El modelo se ejecuta con Ollama en un Mac mini privado.

## Qué incluye

- Panel de conversación en español.
- Contexto del documento completo (limitado a 24 000 caracteres) y de la selección activa.
- Botones para insertar la última respuesta o reemplazar la selección.
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

En macOS, Ollama debe permanecer instalado **nativamente** para aprovechar la aceleración Metal del M4. Solo el gateway se ejecuta en Docker; se comunica con Ollama mediante `host.docker.internal` y publica el puerto únicamente en `127.0.0.1`.

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

En otra terminal, valida e instala el manifiesto en Word:

```bash
pnpm run validate:manifest
pnpm run sideload
```

En macOS, si el sideload no abre Word automáticamente, usa **Insertar → Complementos → Mis complementos → Cargar mis complementos** y selecciona `manifest.xml`.

## Mac mini y Tailscale

Ejecuta el gateway únicamente en el Mac mini. Ollama continúa en `127.0.0.1:11434`; no lo enlaces a `0.0.0.0` ni abras su puerto. Cuando el gateway esté activo en el puerto 3000, publícalo **solo dentro de tu tailnet** con:

```bash
tailscale serve --bg --https=443 http://127.0.0.1:3000
```

Este proyecto está configurado para `https://macmini4pro.tail97c777.ts.net`. Define también `ALLOWED_TAILSCALE_USER_LOGINS` con los correos autorizados, separados por comas, y crea una ACL que solo permita esas identidades/dispositivos al Mac mini en el puerto 443. Así hay autorización doble: la red y el gateway. No uses `tailscale funnel`.

El contenido se manda únicamente al Mac mini cuando se hace una consulta. Revisa las obligaciones de privacidad y consentimiento aplicables a los documentos clínicos; el modelo no debe producir diagnósticos ni decisiones de tratamiento.

## Producción

Cuando pruebes el flujo, sirve el frontend y el endpoint `/api/chat` desde un único proceso HTTP en el Mac mini y publícalo mediante Tailscale Serve. El manifiesto ya usa el dominio privado de Tailscale; no lo publiques fuera de tu tailnet.

La integración utiliza la [API de chat de Ollama](https://docs.ollama.com/api/chat).
