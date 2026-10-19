// Packt die Beispiel-Dateien aus /beispiele in src/examples.js, damit sie auch ohne Server (file://) laden.
const fs = require("fs"), path = require("path");
const dir = path.join(__dirname, "..", "beispiele");
const order = ["Stundenzettel_Oktober_2026.xlsx", "Material_Lager.xlsx", "Auftraege_2026.xlsx"];
const files = order.filter(f => fs.existsSync(path.join(dir, f)))
  .concat(fs.readdirSync(dir).filter(f => /\.xlsx$/i.test(f) && !order.includes(f)));
const out = {};
for (const f of files) out[f] = fs.readFileSync(path.join(dir, f)).toString("base64");
fs.writeFileSync(path.join(__dirname, "..", "src", "examples.js"),
  "/* automatisch erzeugt mit `npm run beispiele` */\nwindow.ZF_EXAMPLES = " + JSON.stringify(out) + ";\n");
console.log("src/examples.js:", files.join(", "));
