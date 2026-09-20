// TallerPro360 OT — frontend MSAL (Auth Code + PKCE) contra API Gateway.
const API_BASE_URL = "https://ny0gbkz69g.execute-api.us-east-1.amazonaws.com";
const TENANT_ID = "a4cc5fc6-a27b-43af-91af-bab64b4e97ce";
const CLIENT_ID = "94997922-80e9-4664-8743-85316d34ba1b";
const SCOPE = "api://94997922-80e9-4664-8743-85316d34ba1b/productos.read";
const REDIRECT_URI = "https://54-87-63-105.sslip.io/";
const msalConfig = {
  auth: {
    clientId: CLIENT_ID,
    authority: "https://login.microsoftonline.com/" + TENANT_ID + "/",
    redirectUri: REDIRECT_URI,
    postLogoutRedirectUri: REDIRECT_URI,
  },
  cache: { cacheLocation: "sessionStorage" },
};
const pca = new msal.PublicClientApplication(msalConfig);
const $ = (id) => document.getElementById(id);
const btnConsultar = $("btn-consultar"), listaOT = $("lista-ot"),
  formCrear = $("form-crear"), mensaje = $("mensaje"),
  btnLogin = $("btn-login"), btnSalir = $("btn-salir"),
  nombreUsuario = $("nombre-usuario"), itemsDiv = $("items"),
  detalle = $("detalle"), detalleTitulo = $("detalle-titulo"),
  detalleCuerpo = $("detalle-cuerpo");
const fmtCLP = (n) => "$" + Number(n || 0).toLocaleString("es-CL");

function cuentaActual() { return pca.getAllAccounts()[0] || null; }
async function obtenerToken() {
  const cuenta = cuentaActual();
  if (!cuenta) throw new Error("No hay sesión iniciada");
  try {
    const r = await pca.acquireTokenSilent({ scopes: [SCOPE], account: cuenta });
    return r.accessToken;
  } catch (e) {
    if (e.name === "InteractionRequiredAuthError") { await pca.acquireTokenRedirect({ scopes: [SCOPE] }); return null; }
    throw e;
  }
}
function mostrarSesion() {
  const cuenta = cuentaActual(), ok = Boolean(cuenta);
  nombreUsuario.textContent = ok ? cuenta.name : "";
  btnLogin.classList.toggle("oculto", ok);
  btnSalir.classList.toggle("oculto", !ok);
  btnConsultar.disabled = !ok;
  formCrear.querySelector('[type="submit"]').disabled = !ok;
}
async function api(path, opciones = {}) {
  const token = await obtenerToken();
  return fetch(API_BASE_URL + path, {
    ...opciones,
    headers: { "Content-Type": "application/json", Authorization: "Bearer " + token, ...(opciones.headers || {}) },
  });
}
function pintarLista(ots) {
  listaOT.innerHTML = "";
  if (!ots.length) { listaOT.innerHTML = "<li>Sin órdenes.</li>"; return; }
  ots.forEach((ot) => {
    const li = document.createElement("li");
    li.className = "ot";
    li.innerHTML = `<div><strong>${ot.ot_id}</strong> · ${ot.patente} · ${ot.cliente_id}<br>
      <span class="muted">${ot.descripcion || ""} — ${ot.n_items} ítems</span></div>
      <div class="total">${fmtCLP(ot.total)}</div>`;
    li.style.cursor = "pointer";
    li.addEventListener("click", () => verDetalle(ot.ot_id));
    listaOT.appendChild(li);
  });
}
async function obtenerOTs() {
  try {
    mensaje.textContent = "";
    const r = await api("/api/ot");
    if (r.status === 401) { mensaje.textContent = "Sesión expirada: inicia sesión de nuevo."; return; }
    if (!r.ok) throw new Error("HTTP " + r.status);
    pintarLista(await r.json());
  } catch (e) { mensaje.textContent = "Error al consultar: " + e.message; }
}
async function verDetalle(otId) {
  try {
    const r = await api("/api/ot/" + encodeURIComponent(otId));
    if (!r.ok) throw new Error("HTTP " + r.status);
    const ot = await r.json();
    detalleTitulo.textContent = ot.ot_id + " · " + ot.patente;
    detalleCuerpo.innerHTML =
      `<p><strong>Cliente:</strong> ${ot.cliente_id} · <strong>Total:</strong> ${fmtCLP(ot.total)}</p>
       <table><thead><tr><th>Concepto</th><th>Cant.</th><th>P. unit</th><th>Subtotal</th></tr></thead><tbody>` +
      ot.items.map((i) => `<tr><td>${i.concepto}</td><td>${i.cantidad}</td><td>${fmtCLP(i.precio_unit)}</td><td>${fmtCLP(i.subtotal)}</td></tr>`).join("") +
      `</tbody></table>`;
    detalle.classList.remove("oculto");
    detalle.scrollIntoView({ behavior: "smooth" });
  } catch (e) { mensaje.textContent = "Error al ver detalle: " + e.message; }
}
function agregarItem(concepto = "", cantidad = 1, precio = 0) {
  const row = document.createElement("div");
  row.className = "grid3 item-row";
  row.innerHTML = `<input placeholder="Concepto" value="${concepto}" required maxlength="40">
    <input type="number" placeholder="Cant." min="0.01" step="0.01" value="${cantidad}" required>
    <input type="number" placeholder="P. unit CLP" min="0" step="1" value="${precio}" required>`;
  itemsDiv.appendChild(row);
}
async function crearOT(ev) {
  ev.preventDefault();
  try {
    const items = [...itemsDiv.querySelectorAll(".item-row")].map((row) => {
      const [c, q, p] = row.querySelectorAll("input");
      return { concepto: c.value, cantidad: Number(q.value), precio_unit: Number(p.value) };
    });
    const body = {
      cliente_id: $("input-cliente").value, patente: $("input-patente").value,
      descripcion: $("input-descripcion").value, total: Number($("input-total").value || 0), items,
    };
    const r = await api("/api/ot", { method: "POST", body: JSON.stringify(body) });
    const datos = await r.json();
    if (!r.ok) throw new Error(datos.detail || datos.error || "HTTP " + r.status);
    mensaje.textContent = "Orden " + datos.ot_id + " creada.";
    formCrear.reset(); itemsDiv.innerHTML = ""; agregarItem();
    obtenerOTs();
  } catch (e) { mensaje.textContent = "Error: " + e.message; }
}
btnLogin.addEventListener("click", () => pca.loginRedirect({ scopes: [SCOPE] }));
btnSalir.addEventListener("click", () => pca.logoutRedirect());
btnConsultar.addEventListener("click", obtenerOTs);
$("btn-agregar-item").addEventListener("click", () => agregarItem());
$("btn-cerrar-detalle").addEventListener("click", () => detalle.classList.add("oculto"));
formCrear.addEventListener("submit", crearOT);
agregarItem();
pca.handleRedirectPromise()
  .then(() => mostrarSesion())
  .catch((e) => { mensaje.textContent = "Error de autenticación: " + e.message; mostrarSesion(); });
