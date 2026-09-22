# TallerPro360 OT Web
Frontend MSAL (Auth Code + PKCE) contra API Gateway. Vanilla JS + CSS propio
(IBM Plex Sans + JetBrains Mono). Sin secretos: solo IDs públicos de Entra.

## Estructura
`Index.html` (lista con buscador, detalle, formulario OT, modal de cliente),
`app.js` (MSAL, Bearer auto, CRUD), `utils.js` (funciones puras),
`styles.css`, `test/utils.test.js`.

## Flujos
Login → nombre visible → Actualizar lista → buscador server-side (`?q`, debounce) →
detalle por click → buscar cliente (RUT/nombre/código) o modal + Nuevo cliente (6 campos) →
crear orden → Eliminar (con confirmación) → logout bloquea todo.
Sin sesión los buscadores siguen llamando al backend, que responde `401` (visible en DevTools).

## Configuración (arriba de `app.js`, IDs públicos)
`API_BASE_URL` (gateway), `TENANT_ID`, `CLIENT_ID`, `SCOPE`, `REDIRECT_URI` (debe calzar
con el redirect SPA registrado en Entra, barra final incluida).

## Deploy
Push a `main` → CI → Deploy: sync S3 → copia a `/var/www/ot/` (ojo: `Index.html` del repo
se publica como `index.html` minúscula, que es lo que pide Nginx).

## Tests
`node --check app.js && node --check utils.js` + `node --test test/utils.test.js`
(utilidades puras: formato CLP, patente, subtotal) + preflight `OPTIONS → 204` en CI.
