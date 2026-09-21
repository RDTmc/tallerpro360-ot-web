# TallerPro360 OT Web
Frontend MSAL (Auth Code + PKCE) contra API Gateway. Vanilla JS + CSS propio.

## Estructura
`Index.html` (vistas OT + modal cliente), `app.js` (MSAL, llamadas Bearer),
`utils.js` (funciones puras), `styles.css` (IBM Plex Sans + JetBrains Mono).

## Configuración (IDs públicos, sin secretos)
Arriba de `app.js`: `API_BASE_URL`, `TENANT_ID`, `CLIENT_ID`, `SCOPE`, `REDIRECT_URI`.

## Tests
`node --test test/` (utilidades puras) + `node --check app.js`.
