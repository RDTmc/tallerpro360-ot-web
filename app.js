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
let temporizadorBusqueda = null, temporizadorCliente = null, detalleIdActual = null;

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
// ---------- órdenes ----------
function pintarLista(ots) {
  listaOT.innerHTML = "";
  if (!ots.length) { listaOT.innerHTML = "<li>Sin resultados.</li>"; return; }
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
async function obtenerOTs(q = "") {
  try {
    mensaje.textContent = "Cargando…";
    const r = await api("/api/ot" + (q ? "?q=" + encodeURIComponent(q) : ""));
    if (r.status === 401) { mensaje.textContent = "Sesión expirada: inicia sesión de nuevo."; return; }
    if (r.status === 403) { mensaje.textContent = "Sin permiso para ver órdenes."; return; }
    if (!r.ok) throw new Error("HTTP " + r.status);
    mensaje.textContent = "";
    pintarLista(await r.json());
  } catch (e) { mensaje.textContent = "Error al consultar: " + e.message; }
}
async function verDetalle(otId) {
  try {
    const r = await api("/api/ot/" + encodeURIComponent(otId));
    if (r.status === 404) { mensaje.textContent = "La orden ya no existe."; obtenerOTs(); return; }
    if (!r.ok) throw new Error("HTTP " + r.status);
    const ot = await r.json();
    detalleIdActual = ot.ot_id;
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
async function eliminarActual() {
  if (!detalleIdActual) return;
  if (!confirm("¿Eliminar la orden " + detalleIdActual + "?")) return;
  try {
    const r = await api("/api/ot/" + encodeURIComponent(detalleIdActual), { method: "DELETE" });
    if (r.status === 404) { mensaje.textContent = "La orden ya no existe."; }
    else if (!r.ok) throw new Error("HTTP " + r.status);
    else mensaje.textContent = "Orden " + detalleIdActual + " eliminada.";
    detalle.classList.add("oculto"); detalleIdActual = null;
    obtenerOTs();
  } catch (e) { mensaje.textContent = "Error al eliminar: " + e.message; }
}
// ---------- clientes: buscador RUT/nombre + alta ----------
async function buscarClientes(q) {
  const box = $("sugerencias-cliente");
  if (!q) { box.innerHTML = ""; return; }
  try {
    const r = await api("/api/clientes?q=" + encodeURIComponent(q));
    if (!r.ok) return;
    const rows = await r.json();
    box.innerHTML = "";
    rows.forEach((c) => {
      const b = document.createElement("button");
      b.type = "button"; b.className = "secundario sugerencia";
      b.textContent = `${c.rut} · ${c.nombres} ${c.apellidos} (${c.codigo})`;
      b.addEventListener("click", () => {
        $("input-cliente").value = c.codigo;
        $("input-buscar-cliente").value = `${c.nombres} ${c.apellidos} (${c.codigo})`;
        box.innerHTML = "";
      });
      box.appendChild(b);
    });
    if (!rows.length) box.innerHTML = '<span class="muted">Sin coincidencias: usa "+ Nuevo cliente".</span>';
  } catch (e) { /* reintenta al escribir */ }
}
function agregarItem(concepto = "", cantidad = 1, precio = 0) {
  const row = document.createElement("div");
  row.className = "grid3 item-row";
  row.innerHTML = `<input placeholder="Concepto" value="${concepto}" required maxlength="40">
    <input type="number" placeholder="Cant." min="0.01" step="0.01" value="${cantidad}" required>
    <input type="number" placeholder="P. unit CLP" min="0" step="1" value="${precio}" required>`;
  itemsDiv.appendChild(row);
}
function abrirModalCliente() { $("modal-cliente").classList.remove("oculto"); }
function cerrarModalCliente() { $("modal-cliente").classList.add("oculto"); }
async function guardarClienteModal() {
  try {
    const nuevo = {
      rut: $("cli-rut").value.trim(), codigo: $("cli-codigo").value.trim().toUpperCase(),
      nombres: $("cli-nombres").value.trim(), apellidos: $("cli-apellidos").value.trim(),
      fecha_nac: $("cli-fecha").value || null, correo: $("cli-correo").value.trim() || null,
      telefono: $("cli-telefono").value.trim() || null,
    };
    if (!nuevo.rut || !nuevo.codigo || !nuevo.nombres || !nuevo.apellidos)
      throw new Error("Completa RUT, código, nombres y apellidos.");
    const rc = await api("/api/clientes", { method: "POST", body: JSON.stringify(nuevo) });
    const dc = await rc.json();
    if (rc.status === 409) throw new Error("Cliente ya existe (RUT o código).");
    if (!rc.ok) throw new Error(dc.detail || "HTTP " + rc.status);
    $("input-cliente").value = dc.codigo;
    $("input-buscar-cliente").value = `${dc.nombres} ${dc.apellidos} (${dc.codigo})`;
    cerrarModalCliente();
    mensaje.textContent = "Cliente " + dc.codigo + " listo: completa la orden.";
  } catch (e) { mensaje.textContent = "Error: " + e.message; }
}
async function crearOT(ev) {
  ev.preventDefault();
  try {
    const codigo = $("input-cliente").value;
    if (!codigo) throw new Error("Busca y selecciona un cliente, o crea uno nuevo.");
    const items = [...itemsDiv.querySelectorAll(".item-row")].map((row) => {
      const [c, q, p] = row.querySelectorAll("input");
      return { concepto: c.value, cantidad: Number(q.value), precio_unit: Number(p.value) };
    });
    const body = {
      cliente_id: codigo, patente: $("input-patente").value,
      descripcion: $("input-descripcion").value, total: Number($("input-total").value || 0), items,
    };
    const r = await api("/api/ot", { method: "POST", body: JSON.stringify(body) });
    const datos = await r.json();
    if (!r.ok) throw new Error(datos.detail || datos.error || "HTTP " + r.status);
    mensaje.textContent = "Orden " + datos.ot_id + " creada.";
    formCrear.reset(); $("input-cliente").value = "";
    itemsDiv.innerHTML = ""; agregarItem();
    obtenerOTs();
  } catch (e) { mensaje.textContent = "Error: " + e.message; }
}
// ---------- eventos ----------
$("input-buscar").addEventListener("input", () => {
  clearTimeout(temporizadorBusqueda);
  temporizadorBusqueda = setTimeout(() => obtenerOTs($("input-buscar").value.trim()), 300);
});
$("input-buscar-cliente").addEventListener("input", () => {
  $("input-cliente").value = "";
  clearTimeout(temporizadorCliente);
  temporizadorCliente = setTimeout(() => buscarClientes($("input-buscar-cliente").value.trim()), 300);
});
$("btn-nuevo-cliente").addEventListener("click", abrirModalCliente);
$("btn-cerrar-modal").addEventListener("click", cerrarModalCliente);
$("modal-cliente").addEventListener("click", (e) => { if (e.target.id === "modal-cliente") cerrarModalCliente(); });
$("btn-guardar-cliente").addEventListener("click", guardarClienteModal);
btnLogin.addEventListener("click", () => pca.loginRedirect({ scopes: [SCOPE] }));
btnSalir.addEventListener("click", () => pca.logoutRedirect());
btnConsultar.addEventListener("click", () => { $("input-buscar").value = ""; obtenerOTs(); });
$("btn-agregar-item").addEventListener("click", () => agregarItem());
$("btn-eliminar").addEventListener("click", eliminarActual);
$("btn-cerrar-detalle").addEventListener("click", () => detalle.classList.add("oculto"));
formCrear.addEventListener("submit", crearOT);
agregarItem();
pca.handleRedirectPromise()
  .then(() => mostrarSesion())
  .catch((e) => { mensaje.textContent = "Error de autenticación: " + e.message; mostrarSesion(); });
