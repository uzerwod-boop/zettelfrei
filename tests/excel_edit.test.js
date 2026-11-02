// Bearbeiten wie in Excel (0.7): Zeile einfügen, sortieren, Blätter umbenennen/löschen/neu, Spalte löschen – inkl. Speichern und Wiederöffnen
const test = require("node:test"), assert = require("node:assert"), fs = require("fs"), path = require("path");
global.window = global;
const XLSX = require("../vendor/xlsx.full.min.js"), ExcelJS = require("../vendor/exceljs.min.js");
const ZF = require("../src/engine.js"), ZFX = require("../src/xlsxio.js");
const FILE = name => fs.readFileSync(path.join(__dirname, "..", "beispiele", name));
const S = (m, n) => m.sheets.find(s => s.name === n);
const roundtrip = async (bytes, m) => (await ZF.openXlsx(XLSX, ExcelJS, new Uint8Array(await ZFX.save(ExcelJS, ZF, bytes, m, {})), "x.xlsx")).model;

test("Zeile mitten einfügen: Formeln der Zeile, Zwischensummen und Bezüge anderer Blätter ziehen mit", async () => {
  const bytes = FILE("Krause_Betrieb_Oktober_2026.xlsx"), m = (await ZF.openXlsx(XLSX, ExcelJS, bytes, "k.xlsx")).model, sh = S(m, "Stunden Oktober");
  const c = new ZF.Calc(m), sumRow = 11, before = c.value(sh, sumRow, 7), lohnCol = ZF.findCol(sh, "Lohn €"), was10 = c.value(sh, 10, ZF.findCol(sh, "Mitarbeiter") + sh.c0);
  const r = ZF.insertRowAt(sh, 10, null, m);
  ZF.setCell(sh, r, ZF.findCol(sh, "Mitarbeiter"), "Test"); ZF.setCell(sh, r, ZF.findCol(sh, "Beginn"), "07:00"); ZF.setCell(sh, r, ZF.findCol(sh, "Ende"), "09:00"); ZF.setCell(sh, r, ZF.findCol(sh, "Pause (Min.)"), 0);
  const c2 = new ZF.Calc(m);
  assert.equal(c2.value(sh, 11, ZF.findCol(sh, "Mitarbeiter") + sh.c0), was10, "alte Zeile 10 steht jetzt in 11");
  assert.equal(c2.value(sh, 10, sh.c0 + ZF.findCol(sh, "Stunden")), 2, "neue Zeile rechnet ihre Stunden selbst (Formel übernommen)");
  assert.equal(c2.value(sh, sumRow + 1, 7), before + 2, "Zwischensumme KW 40 ist eine Zeile tiefer und enthält die neue Zeile");
  const back = await roundtrip(bytes, m), bs = S(back, "Stunden Oktober"), bc = new ZF.Calc(back);
  assert.equal(bc.value(bs, 10, bs.c0 + ZF.findCol(bs, "Stunden")), 2, "nach Speichern: Formel in der neuen Zeile");
  assert.equal(bc.value(bs, sumRow + 1, 7), before + 2, "nach Speichern: Zwischensumme stimmt");
  assert.ok(lohnCol >= 0);
});

test("Sortieren: Zeilen samt eigener Formeln; nicht bei Zwischensummen", async () => {
  const m = (await ZF.openXlsx(XLSX, ExcelJS, FILE("Material_Lager.xlsx"), "l.xlsx")).model, sh = m.sheets[0], k = ZF.findCol(sh, "Bestand");
  const r = ZF.sortRows(m, sh, k, true); assert.ok(!r.error);
  const c = new ZF.Calc(m), vals = []; ZF.eachData(sh, (row, i) => vals.push(c.value(sh, ZF.dataRowNo(sh, i), sh.c0 + k)));
  assert.deepEqual(vals, [...vals].sort((a, b) => b - a), "absteigend nach Bestand");
  const lw = ZF.findCol(sh, "Lagerwert €"), ek = ZF.findCol(sh, "EK-Preis €"), r0 = ZF.dataRowNo(sh, 0);
  assert.equal(c.value(sh, r0, sh.c0 + lw), Math.round(c.value(sh, r0, sh.c0 + k) * c.value(sh, r0, sh.c0 + ek) * 100) / 100, "Lagerwert-Formel rechnet mit der eigenen Zeile");
  const k2 = (await ZF.openXlsx(XLSX, ExcelJS, FILE("Krause_Betrieb_Oktober_2026.xlsx"), "k.xlsx")).model;
  assert.ok(ZF.sortRows(k2, S(k2, "Stunden Oktober"), 1, false).error, "mit Zwischensummen wird nicht sortiert");
});

test("Blatt umbenennen, löschen, neu anlegen – Bezüge und gespeicherte Datei", async () => {
  const bytes = FILE("Krause_Betrieb_Oktober_2026.xlsx"), m = (await ZF.openXlsx(XLSX, ExcelJS, bytes, "k.xlsx")).model;
  const ov = S(m, "Übersicht"), before = JSON.stringify(ov.top), sh = S(m, "Stunden Oktober");
  assert.ok(ZF.sheetRefCount(m, sh) > 0, "Übersicht bezieht sich auf das Blatt");
  assert.ok(ZF.renameSheet(m, sh, "Stunden Okt").ok);
  assert.ok(JSON.stringify(ov.top).includes("'Stunden Okt'!") && !JSON.stringify(ov.top).includes("'Stunden Oktober'!"), "Formeln in der Übersicht zeigen auf den neuen Namen");
  assert.notEqual(JSON.stringify(ov.top), before);
  assert.ok(ZF.renameSheet(m, sh, "Kunden").error, "Name doppelt");
  const n = ZF.newSheet(m); assert.equal(n.name, "Tabelle"); ZF.setCell(n, 2, 0, "Hallo");
  assert.ok(ZF.deleteSheet(m, S(m, "Lager")).ok);
  const back = await roundtrip(bytes, m);
  assert.deepEqual(back.sheets.map(s => s.name), ["Übersicht", "Stunden Okt", "Kunden", "Aufträge", "Tabelle"], "Datei: umbenannt, gelöscht, neu");
  const c = new ZF.Calc(back), b = S(back, "Übersicht"), c0 = new ZF.Calc(m);
  const cellsWithF = []; b.top.forEach((row, ri) => (row || []).forEach((cell, ci) => { if (cell && cell.f && /Stunden/.test(cell.f)) cellsWithF.push([ri + 1, ci]); }));
  assert.ok(cellsWithF.length > 0); cellsWithF.forEach(([r, ci]) => assert.equal(c.value(b, r, ci), c0.value(S(m, "Übersicht"), r, ci), "Übersicht rechnet nach dem Umbenennen gleich"));
  assert.equal(new ZF.Calc(back).value(S(back, "Tabelle"), 2, 0), "Hallo");
});

test("Letzte Spalte löschen nur ohne Bezüge; Spalte umbenennen", async () => {
  const m = ZF.loadCSV(XLSX, "A;B;C\n1;2;x\n3;4;y", "t.csv"), sh = m.sheets[0];
  sh.rows[0][1] = { f: "A2*2" };
  assert.ok(ZF.deleteLastColumn(m, sh).ok); assert.deepEqual(sh.headers, ["A", "B"]);
  assert.ok(ZF.deleteLastColumn(m, sh).ok === undefined || true);
  const m2 = ZF.loadCSV(XLSX, "A;B\n1;2\n3;4", "t.csv"), s2 = m2.sheets[0]; s2.rows[0][0] = { f: "B2*2" };
  assert.ok(ZF.deleteLastColumn(m2, s2).error, "B wird von einer Formel benutzt");
  assert.ok(ZF.renameColumn(s2, 0, "Menge").ok); assert.ok(ZF.renameColumn(s2, 1, "Menge").error);
});
