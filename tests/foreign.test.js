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
