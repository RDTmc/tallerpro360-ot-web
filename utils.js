// --- Funciones compartidas (testeables en Node sin DOM).
function fmtCLP(n) { return "$" + Number(n || 0).toLocaleString("es-CL"); }
function normalizarPatente(p) { return (p || "").toUpperCase().replace(/[\s-]/g, ""); }
function subtotalItem(cantidad, precio) { return Math.round(Number(cantidad) * Number(precio)); }
if (typeof module !== "undefined") { module.exports = { fmtCLP, normalizarPatente, subtotalItem }; }
