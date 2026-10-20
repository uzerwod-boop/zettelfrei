// Sicherheits- und Robustheitstests (Version 0.2.1)
const test = require("node:test"), assert = require("node:assert"), fs = require("fs"), path = require("path");
global.window = global;
const XLSX = require("../vendor/xlsx.full.min.js"), ExcelJS = require("../vendor/exceljs.min.js");
const ZF = require("../src/engine.js"), ZFE = require("../src/einvoice.js"), ZFD = require("../src/docs.js");
const B = f => path.join(__dirname, "..", "beispiele", f);

function tiny() {
  const ws = XLSX.utils.aoa_to_sheet([["Artikel", "Menge"], ["A", 1], ["B", 2]]);
  const wb = XLSX.utils.book_new(); XLSX.utils.book_append_sheet(wb, ws, "T");
  return ZF.loadWorkbook(XLSX, XLSX.write(wb, { type: "array", bookType: "xlsx" }), "t.xlsx");
}

test("Formeln: nur Rechnen, kein JavaScript", () => {
  const m = tiny(), sh = m.sheets[0]; ZF.appendRow(sh, {});
  ZF.setCell(sh, 4, 0, "=[][[]]+[]", { formula: true });
  ZF.setCell(sh, 4, 1, "=SUMME(B2:B3)*2", { formula: true });
  const c = new ZF.Calc(m);
  assert.notEqual(c.value(sh, 4, 0), "", "JS-Ausdruck wird nicht ausgeführt");
  assert.equal(c.value(sh, 4, 1), 6);
});

test("Werte von KI oder E-Rechnung werden nie zur Formel (DDE)", () => {
  const sh = tiny().sheets[0];
  ZF.setCell(sh, 2, 0, "=cmd|' /C calc'!A0");
  assert.equal(sh.rows[0][0].f, undefined);
  assert.equal(sh.rows[0][0].v, "=cmd|' /C calc'!A0");
});

test("CSV-Export schützt vor Formel-Einschleusung", () => {
  const out = ZFD.csv([["=HYPERLINK(\"x\")", "+1+1", "@SUM(A1)", "-5", -5, "Becker; Thomas"]]);
  assert.match(out, /'=HYPERLINK/); assert.match(out, /;'\+1\+1;/); assert.match(out, /;'@SUM/);
  assert.match(out, /;-5;-5,00;/, "negative Zahlen bleiben Zahlen");
});

test("CSV lesen: Semikolon, Anführungszeichen, deutsche Zahlen und Daten", () => {
  const m = ZF.loadCSV(XLSX, "﻿Datum;Kunde;Netto €\n12.10.2026;\"Becker; Thomas\";1.234,50\n13.10.2026;\"Hahn \"\"Bau\"\"\";980\n", "x.csv");
  const sh = m.sheets[0];
  assert.deepEqual(sh.headers, ["Datum", "Kunde", "Netto €"]);
  assert.deepEqual(sh.rows.map(r => r.map(c => c && c.v)), [["12.10.2026", "Becker; Thomas", 1234.5], ["13.10.2026", "Hahn \"Bau\"", 980]]);
  assert.equal(ZF.loadCSV(XLSX, "a,b\n1,2\n", "y.csv").sheets[0].headers.length, 2, "Komma-CSV erkannt");
});

test("Excel lesen mit ExcelJS ergibt dieselben Werte wie bisher", async () => {
  for (const f of fs.readdirSync(path.dirname(B("x"))).filter(f => /\.xlsx$/.test(f))) {
    const a = ZF.loadWorkbook(XLSX, new Uint8Array(fs.readFileSync(B(f))), f);
    const b = await ZF.loadWorkbookExcelJS(XLSX, ExcelJS, fs.readFileSync(B(f)), f);
    assert.equal(a.sheets.length, b.sheets.length, f);
    a.sheets.forEach((sa, i) => {
      const sb = b.sheets[i], ca = new ZF.Calc(a), cb = new ZF.Calc(b);
      assert.deepEqual([sb.hr, sb.headers, sb.dateCol, sb.timeCol], [sa.hr, sa.headers, sa.dateCol, sa.timeCol], f + " Struktur");
      for (let r = 1; r <= sa.hr + 2 + sa.rows.length + sa.foot.length; r++) for (let c = 0; c < sa.c0 + sa.ncol; c++) {
        const va = ca.value(sa, r, c), vb = cb.value(sb, r, c);
        if (typeof va === "number") assert.ok(Math.abs(va - vb) < 1e-9, `${f} ${r}/${c}`); else assert.deepEqual(vb, va, `${f} ${r}/${c}`);
      }
    });
  }
});

test("Gutschrift: Beträge negativ, Lagerbestand sinkt", () => {
  const xml = fs.readFileSync(B("Rechnung_Weber_88213_XRechnung.xml"), "utf8");
  const inv = ZFE.read(xml); inv.type = "Gutschrift";
  const lager = ZF.loadWorkbook(XLSX, new Uint8Array(fs.readFileSync(B("Material_Lager.xlsx"))), "l.xlsx");
  const ops = ZFE.toOps(ZF, lager, lager.sheets[0], inv);
  assert.ok(ops.length && ops.every(o => o.typ === "aendern"), "keine neuen Artikel aus Gutschrift");
  assert.ok(ops.every(o => Object.values(o.werte)[0].plus < 0), "Bestand wird reduziert");
  const er = ZF.loadWorkbook(XLSX, new Uint8Array(fs.readFileSync(B("Eingangsrechnungen_2026.xlsx"))), "e.xlsx");
  const op = ZFE.toOps(ZF, er, er.sheets[0], inv)[0];
  assert.ok(Object.values(op.werte).some(v => v === -837.65), "Netto negativ");
});
