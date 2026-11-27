// Ordner mit mehreren Dateien: Blätter in eine Arbeitsmappe zusammenführen, Formeln rechnen weiter, Speichern und Wiederöffnen
const test = require("node:test"), assert = require("node:assert"), fs = require("fs"), path = require("path");
global.window = global;
const XLSX = require("../vendor/xlsx.full.min.js"), ExcelJS = require("../vendor/exceljs.min.js");
const ZF = require("../src/engine.js"), ZFX = require("../src/xlsxio.js");
const FILE = name => fs.readFileSync(path.join(__dirname, "..", "beispiele", name));
const open = async (bytes, n) => (await ZF.openXlsx(XLSX, ExcelJS, bytes, n));
const values = (m, sh) => { const c = new ZF.Calc(m), o = []; sh.rows.forEach((r, i) => { const rr = []; for (let k = 0; k < sh.ncol; k++) rr.push(c.value(sh, sh.hr + 2 + i, sh.c0 + k)); o.push(rr); }); return o; };
function single(name, sheet) { // Datei mit einem Blatt „Sheet1“ und einer Formel auf das eigene Blatt
  const ws = XLSX.utils.aoa_to_sheet([["Kampagne", "Mailings", "Klicks"], ["Erstkontakt", 4, 0], ["Qualifizierung", 4, 1]]);
  ws["D1"] = { t: "s", v: "Quote" }; ws["D2"] = { t: "n", v: 0, f: `'${sheet}'!C2/'${sheet}'!B2` }; ws["D3"] = { t: "n", v: 0.25, f: "C3/B3" }; ws["!ref"] = "A1:D3";
  const wb = XLSX.utils.book_new(); XLSX.utils.book_append_sheet(wb, ws, sheet); return new Uint8Array(XLSX.write(wb, { type: "array", bookType: "xlsx" }));
}

test("Zwei Dateien mit mehreren Blättern: Namen eindeutig, Formeln und Querbezüge rechnen gleich", async () => {
  const base = (await open(FILE("Material_Lager.xlsx"), "Material_Lager.xlsx")).model;
  const other = (await open(FILE("Krause_Betrieb_Oktober_2026.xlsx"), "Krause.xlsx")).model;
  const before = other.sheets.filter(s => s.hr >= 0).map(s => [s, JSON.stringify(values(other, s))]);
  const n0 = base.sheets.length, added = ZF.mergeModels(base, other, "Krause_Betrieb_Oktober_2026.xlsx");
  assert.strictEqual(base.sheets.length, n0 + other.sheets.length); assert.strictEqual(added.length, other.sheets.length);
  assert.strictEqual(new Set(base.sheets.map(s => s.name.toLowerCase())).size, base.sheets.length, "Blattnamen eindeutig");
  assert.ok(added.some(n => /\(Krause_Bet/.test(n)), "Lager wurde eindeutig benannt: " + added.join(", "));
  before.forEach(([sh, v]) => { assert.ok(base.sheets.includes(sh)); assert.strictEqual(JSON.stringify(values(base, sh)), v, "gleiche Werte in " + sh.name); });
});

test("Allgemeiner Blattname bekommt den Dateinamen, Formeln auf das eigene Blatt bleiben richtig", async () => {
  const base = (await open(FILE("Auftraege_2026.xlsx"), "a.xlsx")).model, other = ZF.loadWorkbook(XLSX, single("x", "Sheet1"), "Marketing-Automation.xlsx");
  ZF.mergeModels(base, other, "Marketing-Automation.xlsx");
  const sh = base.sheets.find(s => s.name === "Marketing-Automation"); assert.ok(sh, base.sheets.map(s => s.name).join(","));
  assert.ok(/Marketing-Automation'!C2/.test(sh.rows[0][3].f), sh.rows[0][3].f);
  const c = new ZF.Calc(base); assert.strictEqual(c.value(sh, 2, 3), 0); assert.strictEqual(c.value(sh, 3, 3), 0.25);
});

test("Zwei Dateien mit gleichem Blattnamen: zweite wird umbenannt", () => {
  const a = ZF.loadWorkbook(XLSX, single("x", "Daten"), "CRM.xlsx"), b = ZF.loadWorkbook(XLSX, single("y", "Daten"), "Kampagnen.xlsx");
  ZF.mergeModels(a, b, "Kampagnen.xlsx");
  assert.deepStrictEqual(a.sheets.map(s => s.name), ["Daten", "Daten (Kampagnen)"]);
  assert.ok(/'Daten \(Kampagnen\)'!C2/.test(a.sheets[1].rows[0][3].f), a.sheets[1].rows[0][3].f);
});

test("Speichern mit Formatierung der ersten Datei: alle Blätter und Werte kommen an", async () => {
  const bytes = FILE("Material_Lager.xlsx"), base = (await open(bytes, "m.xlsx")).model;
  ZF.mergeModels(base, ZF.loadWorkbook(XLSX, single("x", "Sheet1"), "Marketing-Automation.xlsx"), "Marketing-Automation.xlsx");
  const other = (await open(FILE("Krause_Betrieb_Oktober_2026.xlsx"), "k.xlsx")).model; ZF.mergeModels(base, other, "Krause_Betrieb_Oktober_2026.xlsx");
  const want = base.sheets.filter(s => s.hr >= 0).map(s => [s.name, JSON.stringify(values(base, s))]);
  const back = (await open(new Uint8Array(await ZFX.save(ExcelJS, ZF, bytes, base, {})), "x.xlsx")).model;
  want.forEach(([nm, v]) => { const sh = back.sheets.find(s => s.name === nm); assert.ok(sh, "Blatt " + nm + " gespeichert");
    assert.strictEqual(JSON.stringify(values(back, sh)), v, "Werte in " + nm); });
});
