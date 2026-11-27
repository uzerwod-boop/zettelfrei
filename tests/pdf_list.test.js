// Liste aus einem PDF (z. B. Odoo-Aktivitäten) als neues Blatt: Zusammenfügen, Datum/Beträge, Abfragen, Speichern
const test = require("node:test"), assert = require("node:assert"), fs = require("fs"), path = require("path");
global.window = global;
const XLSX = require("../vendor/xlsx.full.min.js"), ExcelJS = require("../vendor/exceljs.min.js");
const ZF = require("../src/engine.js"), ZFX = require("../src/xlsxio.js");
const FILE = name => fs.readFileSync(path.join(__dirname, "..", "beispiele", name));

const chunks = [
  { spalten: ["Verkaufschance", "Phase", "Betrag", "Zuständig", "Anruf", "To-do"], zeilen: [
    ["Holzhandel Nord GmbH", "2. Qualifizierung", "0,00 €", "O", "23.11.2026", ""],
    ["Pellets Süd KG", "1. Erstkontakt", "0,00 €", "A", "", "25.11.2026"],
    ["Brennholz West", "2. Qualifizierung", "0,00 €", "O", "24.11.2026", ""]] },
  { spalten: ["Verkaufschance", "Phase", "Betrag", "Zuständig", "Anruf", "To-do"], zeilen: [
    ["Brennholz West", "2. Qualifizierung", "0,00 €", "O", "24.11.2026", ""],      // an der Schnittkante doppelt gelesen
    ["Kaminholz Ost", "4. Treffen hat stattgefunden", "1.250,00 €", "A", "", "02.12.2026"]] },
  { spalten: ["Verkaufschance", "Phase", "Betrag", "Zuständig", "Anruf", "To-do", "Meeting"], zeilen: [
    ["Späne Mitte", "11. On Hold", "0,00 €", "O", "", "", "01.12.2026"]] }];

test("Teilstücke zusammenfügen: Kantendoppel weg, neue Spalten angehängt", () => {
  const j = ZF.joinChunks(chunks);
  assert.deepStrictEqual(j.headers, ["Verkaufschance", "Phase", "Betrag", "Zuständig", "Anruf", "To-do", "Meeting"]);
  assert.strictEqual(j.rows.length, 5); assert.strictEqual(j.dropped, 1);
  assert.strictEqual(j.rows[4][6], "01.12.2026");
});

test("Als Blatt: Datum als Datum, Beträge als Zahl, Kürzel bleibt Text, Abfrage auf überfällige Anrufe", () => {
  const j = ZF.joinChunks(chunks), sh = ZF.tableSheet(null, "Aktivitäten", j.headers, j.rows), m = { name: "Pipeline.xlsx", sheets: [sh] };
  assert.strictEqual(sh.name, "Aktivitäten");
  assert.deepStrictEqual(sh.rows[0][4], { v: "23.11.2026", d: 1 }); assert.ok(sh.dateCol[4] && sh.dateCol[5]);
  assert.strictEqual(sh.rows[3][2].v, 1250); assert.ok(/€/.test(sh.colFmt[2]));
  assert.strictEqual(sh.rows[0][3].v, "O"); assert.strictEqual(sh.rows[0][5], null);
  const r = ZF.runQueries(m, [{ id: "q1", blatt: "Aktivitäten", art: "liste", spalte: "Verkaufschance", filter: [{ spalte: "Anruf", op: "<", wert: "27.11.2026" }] }]).q1;
  assert.deepStrictEqual(r.list, ["Holzhandel Nord GmbH", "Brennholz West"]);
});

test("Neues Blatt in vorhandener Datei: eindeutiger Name, Speichern mit Formatierung und Wiederöffnen", async () => {
  const bytes = FILE("Krause_Betrieb_Oktober_2026.xlsx"), m = (await ZF.openXlsx(XLSX, ExcelJS, bytes, "k.xlsx")).model;
  const j = ZF.joinChunks(chunks), a = ZF.tableSheet(m, "Aktivitäten", j.headers, j.rows); m.sheets.push(a);
  const b = ZF.tableSheet(m, "Aktivitäten", j.headers, j.rows); assert.strictEqual(b.name, "Aktivitäten 2");
  const out = (await ZF.openXlsx(XLSX, ExcelJS, new Uint8Array(await ZFX.save(ExcelJS, ZF, bytes, m, {})), "x.xlsx")).model;
  const s2 = out.sheets.find(s => s.name === "Aktivitäten");
  assert.ok(s2, "Blatt gespeichert"); assert.strictEqual(ZF.dataCount(s2), 5);
  assert.strictEqual(s2.rows[3][0].v, "Kaminholz Ost"); assert.strictEqual(s2.rows[3][2].v, 1250);
  assert.strictEqual(String(s2.rows[0][4].v), "23.11.2026");
});

test("Ohne geöffnete Datei: Liste wird zur eigenen Excel-Datei", () => {
  const j = ZF.joinChunks(chunks), m = { name: "Pipeline.xlsx", sheets: [ZF.tableSheet(null, "Aktivitäten", j.headers, j.rows)] };
  const wb = XLSX.read(ZF.buildWorkbook(XLSX, m), { type: "array" }), back = ZF.loadWorkbook(XLSX, XLSX.write(wb, { type: "array", bookType: "xlsx" }), "p.xlsx");
  assert.strictEqual(ZF.dataCount(back.sheets[0]), 5); assert.strictEqual(back.sheets[0].headers[0], "Verkaufschance");
});
