// Genera ejemplos/bts_ejemplo.xlsx con BTS aleatorias agrupadas en núcleos urbanos.
// Uso: node tools/generar_ejemplo.js
const fs = require("fs");
const path = require("path");
const XLSX = require("../lib/xlsx.full.min.js");
const A = require("../js/algorithm.js");

const rng = A.createRng("ejemplo-excel");
const towns = [
  { x: 3000, y: 3500, r: 1500, n: 18 },
  { x: 12000, y: 4500, r: 2000, n: 24 },
  { x: 7500, y: 10000, r: 1300, n: 14 },
  { x: 15500, y: 11500, r: 1000, n: 8 },
];
const points = [];
towns.forEach((t) => {
  for (let i = 0; i < t.n; i++) {
    const ang = rng() * 2 * Math.PI;
    const rad = t.r * Math.sqrt(rng());
    points.push([t.x + rad * Math.cos(ang), t.y + rad * Math.sin(ang)]);
  }
});
for (let i = 0; i < 16; i++) points.push([rng() * 18000, rng() * 14000]);

const rows = [["Id", "X", "Y"]].concat(
  points.map((p, i) => ["BTS" + String(i + 1).padStart(3, "0"), Math.round(p[0]), Math.round(p[1])])
);
const wb = XLSX.utils.book_new();
XLSX.utils.book_append_sheet(wb, XLSX.utils.aoa_to_sheet(rows), "BTS");
const out = path.join(__dirname, "..", "ejemplos", "bts_ejemplo.xlsx");
fs.writeFileSync(out, XLSX.write(wb, { type: "buffer", bookType: "xlsx" }));
console.log("Generado " + out + " con " + points.length + " BTS");
