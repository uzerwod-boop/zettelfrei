const test = require("node:test"), assert = require("node:assert"), fs = require("fs"), path = require("path");
const XLSX = require("../vendor/xlsx.full.min.js"), ZF = require("../src/engine.js"), ZFE = require("../src/einvoice.js");
const B = f => path.join(__dirname, "..", "beispiele", f);

test("XRechnung lesen: Kopf, Positionen, Summen", () => {
  const inv = ZFE.read(fs.readFileSync(B("Rechnung_Weber_88213_XRechnung.xml"), "utf8"));
  assert.equal(inv.number, "88213");
  assert.equal(inv.date, "15.10.2026");
  assert.equal(inv.seller.name, "Fliesen-Großhandel Weber KG");
  assert.deepEqual([inv.net, inv.tax, inv.gross], [837.65, 159.15, 996.8]);
  assert.deepEqual(inv.lines.map(l => [l.qty, l.unit]), [[18.5, "m²"], [8, "Stk"], [6, "Sack"], [4, "Eimer"]]);
});

test("E-Rechnung ins Lager: Bestand erhöhen, unbekannten Artikel anlegen, ähnliche Artikel nicht verwechseln", () => {
  const inv = ZFE.read(fs.readFileSync(B("Rechnung_Weber_88213_XRechnung.xml"), "utf8"));
  const m = ZF.loadWorkbook(XLSX, new Uint8Array(fs.readFileSync(B("Material_Lager.xlsx"))), "Lager.xlsx");
  const plan = ZF.planOps(m, ZFE.toOps(ZF, m, m.sheets[0], inv));
  assert.deepEqual(plan.filter(p => p.typ === "aendern").map(p => [p.row, p.changes[0].nv]), [[13, 31], [11, 30], [3, 10]]);
  const neu = plan.find(p => p.typ === "neu");
  assert.equal(neu.changes.find(c => c.name === "Artikel").nv, "Fugenmörtel weiß 5 kg", "weiß ist nicht grau");
});

test("E-Rechnung in Eingangsrechnungen: eine Zeile pro Rechnung", () => {
  const inv = ZFE.read(fs.readFileSync(B("Rechnung_Weber_88213_XRechnung.xml"), "utf8"));
  const m = ZF.loadWorkbook(XLSX, new Uint8Array(fs.readFileSync(B("Eingangsrechnungen_2026.xlsx"))), "E.xlsx");
  const plan = ZF.planOps(m, ZFE.toOps(ZF, m, m.sheets[0], inv));
  assert.equal(plan.length, 1);
  const v = Object.fromEntries(plan[0].changes.map(c => [c.name, c.nv]));
  assert.deepEqual([v["Rechnungs-Nr"], v["Netto €"], v["MwSt-Satz"], v["Fällig am"]], ["88213", 837.65, 0.19, "29.10.2026"]);
});

test("XRechnung erzeugen: Pflichtangaben prüfen, Summen, wieder einlesen", () => {
  const inv = { number: "2026-029", date: "19.10.2026", due: "02.11.2026", buyerRef: "2026-029",
    seller: { name: "Fliesen Krause", street: "Gneisenaustr. 22", zip: "10961", city: "Berlin", vatId: "DE123456789", contactName: "Petra Krause", phone: "030 1234567", email: "buero@fliesen-krause.example", iban: "DE02120300000000202051" },
    buyer: { name: "Becker, Thomas", street: "Dorfstr. 5", zip: "14532", city: "Kleinmachnow", email: "t.becker@example.de" },
    lines: [{ name: "Terrasse Feinsteinzeug", qty: 1, unit: "pauschal", price: 7450, rate: 19 }] };
  assert.deepEqual(ZFE.checkData(inv), []);
  assert.ok(ZFE.checkData({ ...inv, seller: { ...inv.seller, iban: "" } }).includes("IBAN"));
  const { xml, totals } = ZFE.buildXRechnung(inv);
  assert.deepEqual([totals.net, totals.tax, totals.gross], [7450, 1415.5, 8865.5]);
  assert.match(xml, /xrechnung_3\.0/);
  const back = ZFE.read(xml);
  assert.deepEqual([back.number, back.buyer.name, back.gross, back.lines[0].unitCode], ["2026-029", "Becker, Thomas", 8865.5, "LS"]);
});
