// Excel-Dateien aus anderen Programmen (z. B. .NET OpenXML SDK) und ArbZG-Grenzfälle
const test = require("node:test"), assert = require("node:assert"), fs = require("fs"), path = require("path");
global.window = global;
const XLSX = require("../vendor/xlsx.full.min.js"), ExcelJS = require("../vendor/exceljs.min.js");
const ZF = require("../src/engine.js");
const FILE = path.join(__dirname, "..", "beispiele", "Material_Lager.xlsx");

// Datei so umschreiben, wie es das OpenXML SDK tut: Präfix x:, BOM, absolute Pfade
function foreignize(buf) {
  const z = XLSX.CFB.read(buf, { type: "buffer" }), dec = new TextDecoder(), enc = new TextEncoder();
  z.FileIndex.forEach((f, i) => {
    const p = z.FullPaths[i].replace(/^Root Entry\//, "");
    if (f.type !== 2 || !/\.(xml|rels)$/.test(p)) return;
    let t = dec.decode(f.content);
    if (/\.rels$/.test(p)) {
      const base = p.replace(/_rels\/[^/]*$/, "");
      t = "﻿" + t.replace(/Target="(?!\/|http)([^"]+)"/g, (_, tg) => 'Target="/' + path.posix.normalize(base + tg) + '"');
    } else if (/xmlns="http:\/\/schemas\.openxmlformats\.org\/spreadsheetml\/2006\/main"/.test(t)) {
      t = t.replace(/<(\/?)(?![?!])([A-Za-z])/g, "<$1x:$2").replace(/<(\/?)x:([A-Za-z]+):/g, "<$1$2:")
           .replace('xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"', 'xmlns:x="http://schemas.openxmlformats.org/spreadsheetml/2006/main"');
    }
    const b = enc.encode(t); f.content = b; f.size = b.length;
  });
  return new Uint8Array(XLSX.CFB.write(z, { fileType: "zip", type: "array" }));
}

test("Datei mit x:-Präfix, BOM und absoluten Pfaden öffnet mit ExcelJS", async () => {
  const orig = fs.readFileSync(FILE), foreign = foreignize(orig);
  await assert.rejects(ZF.loadWorkbookExcelJS(XLSX, ExcelJS, foreign, "f.xlsx"), "ExcelJS allein scheitert");
  const r = await ZF.openXlsx(XLSX, ExcelJS, foreign, "f.xlsx");
  assert.ok(r.bytes, "über angeglichene Datei mit ExcelJS gelesen (Formatierung bleibt beim Speichern)");
  const a = await ZF.loadWorkbookExcelJS(XLSX, ExcelJS, orig, "o.xlsx");
  assert.deepEqual(r.model.sheets.map(s => [s.name, s.headers, ZF.dataCount(s)]), a.sheets.map(s => [s.name, s.headers, ZF.dataCount(s)]));
  assert.equal(ZF.normalizeOOXML(XLSX, orig), null, "normale Dateien bleiben unverändert");
});

test("ArbZG: genau 6 Std. ohne Pause und genau 9 Std. mit 30 Min. sind erlaubt", () => {
  const csv = "Datum;Mitarbeiter;Beginn;Ende;Pause;Stunden\n15.10.2026;A;08:00;14:00;0;6\n15.10.2026;B;08:00;17:30;30;9\n15.10.2026;C;08:00;14:15;0;6,25\n15.10.2026;D;08:00;18:00;30;9,5";
  const m = ZF.loadCSV(XLSX, csv, "t.csv"), p = ZF.checks(m).filter(x => /Pause/.test(x.problem)).map(x => x.zeile);
  assert.deepEqual(p, [4, 5], "nur C (6,25 Std. ohne Pause) und D (9,5 Std. mit 30 Min.)");
});

test("Mehrere Zeilen per Filter: die Tabelle sucht die Zeilen, nicht die KI", () => {
  const csv = "Rechnungs-Nr;Kunde;Brutto EUR;Bezahlt EUR;Status\nR1;Schneider GmbH;100;0;Offen\nR2;Becker OHG;200;0;Offen\nR3;Schneider GmbH;300;0;Storniert\nR4;Schneider GmbH;400;100;Teilbezahlt\nR5;Schneider GmbH;500;500;Bezahlt";
  const m = ZF.loadCSV(XLSX, csv, "r.csv");
  const plan = ZF.planOps(m, [{ typ: "aendern", blatt: m.sheets[0].name, filter: [{ spalte: "Kunde", op: "=", wert: "Schneider GmbH" }, { spalte: "Status", op: "in", wert: ["Offen", "Teilbezahlt"] }],
    werte: { Status: "Bezahlt", "Bezahlt EUR": { aus: "Brutto EUR" } } }]);
  assert.deepEqual(plan.map(p => p.row), [2, 5], "nur R1 und R4");
  assert.match(plan[0].summary, /R1 · Schneider GmbH/, "Zeile ist in der Vorschau erkennbar");
  assert.deepEqual(plan[1].changes.map(c => [c.name, c.nv]), [["Status", "Bezahlt"], ["Bezahlt EUR", 400]], "Wert aus anderer Spalte derselben Zeile");
  const wrong = ZF.planOps(m, [{ typ: "aendern", blatt: m.sheets[0].name, zeile: 3, filter: [{ spalte: "Rechnungs-Nr", op: "=", wert: "R1" }], werte: { Status: "Bezahlt" } }]);
  assert.ok(wrong[0].error, "falsche Zeilennummer wird abgelehnt");
});

test("Summe meldet Text-Zahlen und unlesbare Werte", () => {
  const m = ZF.loadCSV(XLSX, "Beleg;Betrag\nA;100\nB;200\nC;300\nD;ca. 250\nE;\nF;400", "b.csv");
  m.sheets[0].rows[1][1] = { v: "1.234,56" };   // Zahl als Text, wie aus einem Import
  const r = ZF.query(m, { blatt: m.sheets[0].name, spalte: "Betrag", art: "summe" });
  assert.equal(r.value, 2034.56); assert.deepEqual(r.textNums, ["1.234,56"]); assert.deepEqual(r.unread, ["ca. 250"]);
  assert.equal(r.excelValue, 800, "so rechnet Excel (ohne Text)"); assert.equal(r.empty, 1);
});
