const assert = require("node:assert/strict");
const { fmtCLP, normalizarPatente, subtotalItem } = require("../utils.js");
assert.equal(fmtCLP(80000), "$80.000");
assert.equal(normalizarPatente("zz xx99"), "ZZXX99");
assert.equal(subtotalItem(2, 25000), 50000);
console.log("web utils OK");
