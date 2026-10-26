// Formel-Interpreter: Excel-Funktionen (deutsch und englisch)
const test = require("node:test"), assert = require("node:assert");
const XLSX = require("../vendor/xlsx.full.min.js"), ZF = require("../src/engine.js");
function wb() {
  const aoa = [["Name", "Baustelle", "Stunden", "Datum", "Betrag"],
    ["Mehmet", "Weber", 8, "01.10.2026", 400], ["Kevin", "Hahn", 7.5, "02.10.2026", 337.5], ["Mehmet", "Hahn", 6, "05.10.2026", 312],
    ["Jonas", "Weber", 9, "06.10.2026", 405], ["Kevin", "Weber", 4, "07.10.2026", 180]];
  const w = XLSX.utils.book_new(); XLSX.utils.book_append_sheet(w, XLSX.utils.aoa_to_sheet(aoa), "T");
  const m = ZF.loadWorkbook(XLSX, XLSX.write(w, { type: "array", bookType: "xlsx" }), "t.xlsx"); m.today = "2026-10-26T12:00:00";
  return m;
}
test("Bedingte Summen, Anzahl, Mittelwert, Max/Min", () => {
  const m = wb(), sh = m.sheets[0], c = new ZF.Calc(m), f = x => c.formula(sh, x);
  assert.equal(f('SUMIFS(C2:C6,A2:A6,"Mehmet",B2:B6,"Weber")'), 8);
  assert.equal(f('SUMMEWENNS(C2:C6;B2:B6;"Weber")'), 21);
  assert.equal(f('COUNTIFS(B2:B6,"Weber",C2:C6,">=8")'), 2);
  assert.equal(f('AVERAGEIF(A2:A6,"Kevin",C2:C6)'), 5.75);
  assert.equal(f('MAXIFS(E2:E6,B2:B6,"Weber")'), 405);
  assert.equal(f('MINWENNS(E2:E6;A2:A6;"Kevin")'), 180);
  assert.equal(f('SUMPRODUCT(C2:C6,E2:E6)'), 8*400 + 7.5*337.5 + 6*312 + 9*405 + 4*180);
});
test("Nachschlagen: XVERWEIS, INDEX/VERGLEICH, WENNS, ERSTERWERT", () => {
  const m = wb(), sh = m.sheets[0], c = new ZF.Calc(m), f = x => c.formula(sh, x);
  assert.equal(f('XLOOKUP("Jonas",A2:A6,E2:E6)'), 405);
  assert.equal(f('XVERWEIS("Ayse";A2:A6;E2:E6;"nicht da")'), "nicht da");
  assert.equal(f('INDEX(B2:B6,MATCH("Jonas",A2:A6,0))'), "Weber");
  assert.equal(f('INDEX(A2:E6,3,5)'), 312);
  assert.equal(f('IFS(C2>8,"lang",C2>=8,"normal",TRUE,"kurz")'), "normal");
  assert.equal(f('SWITCH(B3,"Weber","W","Hahn","H","?")'), "H");
  assert.equal(f('IFNA(VLOOKUP("X",A2:E6,2,FALSE),"fehlt")'), "fehlt");
});
test("Datum, Text, Mathe", () => {
  const m = wb(), sh = m.sheets[0], c = new ZF.Calc(m), f = x => c.formula(sh, x);
  assert.equal(ZF.serialToDate(f("EDATE(D2,1)")), "01.11.2026");
  assert.equal(ZF.serialToDate(f("MONATSENDE(D2;0)")), "31.10.2026");
  assert.equal(f("ISOWEEKNUM(D2)"), 40);
  assert.equal(f("WEEKDAY(D2,2)"), 4, "01.10.2026 ist ein Donnerstag");
  assert.equal(f("NETWORKDAYS(D2,D6)"), 5);
  assert.equal(f('DATEDIF(D2,D6,"d")'), 6);
  assert.equal(f('TEXT(D2,"TT.MM.JJJJ")'), "01.10.2026");
  assert.equal(f('TEXT(C3,"0,00")'), "7,50");
  assert.equal(f('SUBSTITUTE("Bad Weber","Weber","Hahn")'), "Bad Hahn");
  assert.equal(f('TEXTJOIN(", ",TRUE,A2:A4)'), "Mehmet, Kevin, Mehmet");
  assert.equal(f('MID("Fliesenkleber",8,6)'), "kleber");
  assert.equal(f("MOD(-7,3)"), 2); assert.equal(f("INT(-2.5)"), -3);
  assert.equal(f("CEILING(7.2,0.5)"), 7.5); assert.equal(f("MEDIAN(C2:C6)"), 7.5);
});
