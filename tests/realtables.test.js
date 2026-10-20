// Echte, unordentliche Tabellen (Version 0.2.2): Stresstest-Datei aus beispiele/
const test = require("node:test"), assert = require("node:assert"), fs = require("fs"), path = require("path");
global.window = global;
const XLSX = require("../vendor/xlsx.full.min.js"), ExcelJS = require("../vendor/exceljs.min.js");
const ZF = require("../src/engine.js"), ZFX = require("../src/xlsxio.js");
const FILE = path.join(__dirname, "..", "beispiele", "Stresstest_Chaos_Tabelle.xlsx");
const load = () => ZF.loadWorkbookExcelJS(XLSX, ExcelJS, fs.readFileSync(FILE), "chaos.xlsx");
const S = (m, n) => m.sheets.find(s => s.name === n);
const STD = "Std. Okt+Nov (NEU!!)";

test("Zweizeiliger Kopf, Nebentabelle, Trennzeilen, Excel-Zeilennummern", async () => {
  const m = await load(), sh = S(m, STD), c = new ZF.Calc(m);
  assert.deepEqual(sh.headers, ["Datum", "MA", "Baustelle / Kunde", "von", "bis", "Pause (min)", "Std.", "Betrag", "Bemerkung / Material?", "Std."]);
  assert.equal(sh.ncol, 10, "Mitarbeiter/Satz rechts daneben gehört nicht zur Tabelle");
  assert.equal(sh.hr + 1, 5, "Kopfzeile ist Excel-Zeile 5");
  const kind = r => sh.kind[r - sh.hr - 2];
  assert.equal(kind(10), "leer"); assert.equal(kind(14), "summe"); assert.equal(kind(15), "abschnitt"); assert.equal(kind(16), "");
  assert.equal(ZF.dataCount(sh), 11, "Oktober und November sind Daten");
  assert.equal(c.value(sh, 20, 1), "GESAMT", "Summenzeile bleibt in Excel-Zeile 20");
  assert.equal(c.value(sh, 20, 7), 275.25, "TEILERGEBNIS rechnet ohne Zwischensumme");
  assert.equal(c.value(sh, 14, 7), 241);
  assert.equal(c.value(S(m, "Übersicht"), 3, 1), 275.25, "Bezug aus anderem Blatt zeigt auf die richtige Zeile");
  assert.ok(sh.dateCol[0] && sh.timeCol[3] && sh.timeCol[4], "Datum und Uhrzeit trotz gemischter Schreibweisen erkannt");
});

test("Neue Zeile: Formeln, Summen und Bezüge anderer Blätter wandern mit", async () => {
  const m = await load(), sh = S(m, STD);
  const plan = ZF.planOps(m, [{ typ: "neu", blatt: STD, werte: { Datum: "05.11.2026", MA: "Jonas", "Baustelle / Kunde": "Weber", von: "07:00", bis: "16:00", "Pause (min)": 30 }}]);
  assert.ok(!plan[0].error, JSON.stringify(plan[0]));
  const t = ZF.applyPlan(m, plan); const r = t[0].row; const c = new ZF.Calc(m);
  assert.equal(r, 20, "neue Zeile direkt nach den Daten");
  assert.equal(c.value(sh, r, 7), 8.5, "Stunden per WENNFEHLER/RUNDEN aus der Vorlage berechnet");
  assert.equal(c.value(sh, r, 8), 408, "Betrag mit $I$2");
  assert.equal(c.value(sh, 21, 1), "GESAMT"); assert.equal(c.value(sh, 21, 7), 283.75, "Gesamtsumme enthält die neue Zeile");
  assert.match(ZF.cellAt(S(m, "Übersicht"), 3, 1).f, /H21$/, "Übersicht zeigt jetzt auf H21");
  assert.equal(c.value(S(m, "Übersicht"), 3, 1), 283.75);
  const bad = ZF.planOps(m, [{ typ: "aendern", blatt: STD, zeile: 14, werte: { "Std.": 1 } }]);
  assert.match(bad[0].error, /Zwischensumme/, "Zwischensummen werden nicht überschrieben");
});

test("Prüfen findet die Fallen der Datei", async () => {
  const m = await load(); const p = ZF.checks(m, { today: new Date(2026, 9, 20) }).map(i => i.blatt + "|" + i.zeile + "|" + i.problem);
  const has = re => assert.ok(p.some(x => re.test(x)), "fehlt: " + re + "\n" + p.join("\n"));
  has(/Std\. Okt.*\|19\|Ruhezeit nur 10,5/); has(/\|18\|Über 10 Std\./); has(/\|9\|Pause zu kurz/);
  has(/Lager\|12\|Text statt Zahl \(„ca\. 30“\)/); has(/Lager\|5\|Zahl als Text gespeichert \(„14“\)/);
  has(/Rechnungen 2026\|7\|Steuersatz 19 statt 19 %/); has(/Rechnungen 2026\|7\|Nachschlagen ohne Treffer/);
  has(/Rechnungen 2026\|6\|Re-Nr doppelt/); has(/Rechnungen 2026\|5\|Seit 8 Tagen überfällig/);
  assert.ok(!p.some(x => /\|5\|Datum fehlt|Art\.Nr\..*Text statt Zahl|Re-Nr.*Text statt Zahl|offen gesamt/.test(x)), "keine Fehlalarme:\n" + p.join("\n"));
});

test("Formeln: WENN, WENNFEHLER, SVERWEIS, SUMMEWENN, Text, Vergleiche – ohne eval", async () => {
  const m = await load(), sh = S(m, "Rechnungen 2026"), c = new ZF.Calc(m);
  const f = x => c.formula(sh, x);
  assert.equal(f('IFERROR(VLOOKUP(1003,Kunden!$A:$D,2,FALSE()),"?")'), "Praxis Dr. Hahn");
  assert.equal(f('WENNFEHLER(SVERWEIS(4711;Kunden!$A:$D;2;FALSCH);"??? unbekannt")'), "??? unbekannt");
  assert.equal(f('SUMIF(J3:J7,"offen",E3:E7)'), 7450, "wie Excel: „OFFEN“ zählt mit");
  assert.equal(f('COUNTIF(J3:J7,"off*")'), 3, "Groß/klein egal, Platzhalter *");
  assert.equal(f('IF(AND(E3>1000,F3=0.19),"groß","klein")&"!"'), "groß!");
  assert.equal(f('"7:30"*24'), 7.5, "Uhrzeit als Text rechnet wie in Excel");
  assert.equal(f('1+"abc"'), "#WERT!"); assert.equal(f("1/0"), "#DIV/0!");
  assert.equal(f('constructor.constructor("return 1")()'), undefined, "kein Code, nur Formeln");
});

test("Speichern: Struktur, Trennzeilen und Formeln bleiben erhalten", async () => {
  const m = await load(), sh = S(m, STD);
  ZF.applyPlan(m, ZF.planOps(m, [{ typ: "neu", blatt: STD, werte: { Datum: "05.11.2026", MA: "Jonas", von: "07:00", bis: "16:00", "Pause (min)": 30 } }]));
  const buf = await ZFX.save(ExcelJS, ZF, fs.readFileSync(FILE), m, {});
  const wb = new ExcelJS.Workbook(); await wb.xlsx.load(buf);
  const ws = wb.getWorksheet(STD);
  assert.equal(ws.getCell("B14").value, "Zwischensumme Oktober (1. Woche)");
  assert.equal(ws.getCell("B21").value, "GESAMT");
  assert.match(ws.getCell("H21").value.formula, /SUBTOTAL\(9,H6:H20\)/);
  assert.match(ws.getCell("H20").value.formula, /IFERROR\(ROUND\(\(\(F20-E20\)/);
  assert.ok(ws.getCell("C4").isMerged && ws.getCell("B4").isMerged, "verbundene Kopfzellen bleiben");
  assert.match(wb.getWorksheet("Übersicht").getCell("B3").value.formula, /H21$/);
  assert.equal(wb.getWorksheet("Kunden").state, "hidden", "verstecktes Blatt bleibt versteckt");
});

test("Großes Betriebsbuch: SVERWEIS auf Nebentabelle, Übersicht, gezielte Hinweise", async () => {
  const f = path.join(__dirname, "..", "beispiele", "Krause_Betrieb_Oktober_2026.xlsx");
  const m = await ZF.loadWorkbookExcelJS(XLSX, ExcelJS, fs.readFileSync(f), "b.xlsx"), c = new ZF.Calc(m);
  const st = S(m, "Stunden Oktober");
  assert.equal(st.ncol, 9, "Stundensätze rechts sind Nebentabelle");
  assert.equal(c.value(st, 7, 8), 442, "Lohn = 8,5 Std. × 52 € aus der Nebentabelle");
  const total = c.value(st, 31, 7); assert.ok(total > 100, "GESAMT Oktober");
  assert.equal(c.value(S(m, "Übersicht"), 5, 3), total, "Übersicht übernimmt die Stunden");
  const p = ZF.checks(m, { today: new Date(2026, 9, 20) }).map(i => i.blatt + "|" + i.zeile + "|" + i.problem);
  assert.ok(!p.some(x => /Kd-Nr doppelt|#NV/.test(x)), p.join("\n"));
  assert.ok(p.some(x => /Aufträge\|9\|Steuersatz 19/.test(x)) && p.some(x => /Stunden Oktober\|22\|Ruhezeit/.test(x)));
});
