// consume_company_data.js
// Ejemplo simple para leer assets e inventario desde el paquete JSON.
// Uso:
//   node consume_company_data.js ./company_data Weston
//
// Requisitos: Node 16+

const fs = require("fs");
const path = require("path");

function readJson(fp) {
  return JSON.parse(fs.readFileSync(fp, "utf-8"));
}

function main() {
  const root = process.argv[2] || "./company_data";
  const company = process.argv[3] || "Weston";

  const assets = readJson(path.join(root, "master_data", "assets.json"));
  const inv = readJson(path.join(root, "snapshots", `inventory_${company}.json`));

  console.log("Assets:", assets.length);
  console.log("Inventory rows:", inv.length);

  const first = inv[0] || {};
  console.log("Ejemplo fila inventario:", first);
}

main();
