// verify_manifest.js
// Node.js script (sin dependencias) para validar hashes del paquete JSON.
//
// Uso:
//   node verify_manifest.js ./company_data
//
// Requisitos: Node 16+

const fs = require("fs");
const path = require("path");
const crypto = require("crypto");

function sha256File(fp) {
  const h = crypto.createHash("sha256");
  const data = fs.readFileSync(fp);
  h.update(data);
  return h.digest("hex");
}

function main() {
  const root = process.argv[2] || "./company_data";
  const manPath = path.join(root, "manifests", "manifest.json");
  if (!fs.existsSync(manPath)) {
    console.error("No existe manifest:", manPath);
    process.exit(2);
  }
  const manifest = JSON.parse(fs.readFileSync(manPath, "utf-8"));
  console.log("Package:", manifest.package_name);
  console.log("Generated:", manifest.generated_at);
  let ok = 0, bad = 0, missing = 0;

  for (const f of (manifest.files || [])) {
    const fp = path.join(root, f.path);
    if (!fs.existsSync(fp)) {
      console.log("[MISSING]", f.path);
      missing++;
      continue;
    }
    const got = sha256File(fp);
    if ((got || "").toLowerCase() === (f.sha256 || "").toLowerCase()) {
      ok++;
    } else {
      console.log("[BADHASH]", f.path, "expected=", f.sha256, "got=", got);
      bad++;
    }
  }

  console.log("OK:", ok, "BAD:", bad, "MISSING:", missing);
  process.exit((bad || missing) ? 1 : 0);
}

main();
