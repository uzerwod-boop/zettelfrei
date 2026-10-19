const test=require("node:test"), assert=require("node:assert");
const fs=require("fs"), path=require("path"), vm=require("vm");
const XLSX=require("../vendor/xlsx.full.min.js"); const ZF=require("../src/engine.js");
const load=f=>ZF.loadWorkbook(XLSX,new Uint8Array(fs.readFileSync(path.join(__dirname,"..","beispiele",f))),f);
const eq=(a,b,m)=>assert.deepStrictEqual(a,b,m);
test("Excel-Engine mit den drei Beispieldateien",()=>{

// --- Stundenzettel
let m=load("Stundenzettel_Oktober_2026.xlsx"); let sh=m.sheets[0];
eq([sh.hr,sh.c0,sh.ncol,sh.rows.length,sh.foot.length],[2,0,8,9,1],"Kopfzeile/Datenzeilen erkannt");
eq(sh.dateCol[0],true,"Datumsspalte erkannt");
eq(sh.rows[0][0].v,"12.10.2026","Datum als Text");
let c=new ZF.Calc(m);
eq(c.value(sh,4,6),384,"G4 = E4*F4");
eq(c.value(sh,13,4),55.5,"Summe Stunden");
eq(c.value(sh,13,6),2663,"Summe Betrag");
let r=ZF.appendRow(sh,["19.10.2026","Mehmet","Bad Schneider","Wand gefliest",7,52]);
eq(r,13,"neue Zeile 13"); c=new ZF.Calc(m);
eq(sh.rows[9][6].f,"E13*F13","Formel übernommen");
eq(sh.foot[0][4].f,"SUM(E4:E13)","Summenbereich erweitert");
eq(c.value(sh,14,4),62.5,"neue Summe Stunden");
ZF.deleteRow(sh,5); c=new ZF.Calc(m);
eq(sh.foot[0][4].f,"SUM(E4:E12)","Summenbereich nach Löschen");
eq(c.value(sh,13,4),55,"Summe nach Löschen (62,5-7,5)");
eq(sh.rows[3][6].f,"E7*F7","Zeilenformel nach Löschen verschoben");

// Abfragen
m=load("Stundenzettel_Oktober_2026.xlsx");
eq(ZF.query(m,{blatt:"Oktober 2026",spalte:"Stunden",art:"summe",filter:[{spalte:"Mitarbeiter",op:"=",wert:"Kevin"}]}).value,22,"Stunden Kevin");
eq(ZF.query(m,{blatt:"Oktober 2026",spalte:"Betrag €",art:"summe",filter:[{spalte:"Projekt",op:"=",wert:"bad schneider"}]}).value,1634,"Lohn Bad Schneider");
eq(ZF.query(m,{blatt:"Oktober 2026",spalte:"Stunden",art:"summe",filter:[{spalte:"Datum",op:">=",wert:"14.10.2026"}]}).value,31,"Stunden ab 14.10.");

// Prüfen
eq(ZF.checks(m).map(x=>x.zeile+":"+x.problem),["11:Projekt fehlt"],"Prüfung Stundenzettel");

// --- Lager: plus-Änderung + Export
m=load("Material_Lager.xlsx"); sh=m.sheets[0];
eq([sh.hr,sh.rows.length,sh.foot.length],[0,14,1],"Lager erkannt");
let plan=ZF.planOps(m,[{typ:"aendern",blatt:"Lager",zeile:2,werte:{"Bestand":{plus:-3}}},{typ:"neu",blatt:"Lager",werte:{"Artikel-Nr":"A-1012","Artikel":"Fugenmörtel weiß 5 kg","Bestand":6,"EK-Preis €":"11,40","Lagerwert €":999}}]);
eq(plan[0].changes[0].nv,11,"Bestand 14-3");
eq(plan[1].changes.length,4,"berechnete Spalte ignoriert");
ZF.applyPlan(m,plan); c=new ZF.Calc(m);
eq(c.value(sh,2,7),163.9,"Lagerwert neu 11*14,9");
eq(c.value(sh,16,7),68.4,"Lagerwert neue Zeile");
const buf=ZF.buildWorkbook(XLSX,m); const back=XLSX.read(buf,{type:"array",cellNF:true}); const ws=back.Sheets["Lager"];
eq([ws.H16.f,ws.H17.f,ws.G16.z],["D16*G16","SUM(H2:H16)",'#,##0.00" €"'],"Export: Formeln + Format");

// --- Aufträge: Duplikat, fehlender Kunde, Datum
m=load("Auftraege_2026.xlsx"); sh=m.sheets[0];
eq([sh.hr,sh.rows.length],[3,8],"Aufträge erkannt (Titel + Stand oben)");
c=new ZF.Calc(m); eq(c.value(sh,5,6),2812,"ROUND MwSt");
eq(ZF.checks(m).map(x=>x.zeile+":"+x.problem),["12:Auftrags-Nr doppelt (auch Zeile 7)","11:Kunde fehlt"],"Prüfung Aufträge");
eq(ZF.query(m,{blatt:"Aufträge 2026",spalte:"Brutto €",art:"summe",filter:[{spalte:"Status",op:"=",wert:"Fertig"},{spalte:"Bezahlt am",op:"leer"}]}).value,11221.7,"offene Rechnungen brutto");
plan=ZF.planOps(m,[{typ:"aendern",blatt:"Aufträge 2026",zeile:9,werte:{"Status":"Bezahlt","Bezahlt am":"19.10.2026"}}]);
ZF.applyPlan(m,plan);
const b2=XLSX.read(ZF.buildWorkbook(XLSX,m),{type:"array",cellNF:true}).Sheets["Aufträge 2026"];
eq([b2.J9.v,b2.J9.z,b2.A1.v],[46314,"dd\\.mm\\.yyyy","Auftragsliste 2026"],"Datum als echtes Excel-Datum, Titel erhalten");
assert.match(ZF.describe(m),/G „MwSt €“ \(berechnet, nicht setzen\)/);
});
