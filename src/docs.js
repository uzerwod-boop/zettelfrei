/* Zettelfrei – Dokumente: Mahnwesen, Rechnungsansicht, Abnahmeprotokoll, Steuerberater-Export, Wochenbericht.
 * PDFs mit jsPDF (Standardschrift Helvetica). Texte sind Vorlagen, keine Rechtsberatung.
 */
var ZFD = (function(){
"use strict";
function de(n,d){ return Number(n).toLocaleString("de-DE",{minimumFractionDigits:d==null?2:d,maximumFractionDigits:d==null?2:d}); }
function eur(n){ return de(n)+" €"; }

/* ---------- Brief im DIN-5008-Stil ---------- */
function letterBase(jsPDF, firma, empf, ort){
  var doc=new jsPDF({unit:"mm",format:"a4"});
  doc.setFont("helvetica","bold"); doc.setFontSize(15); doc.text(firma.name||"",20,20);
  doc.setFont("helvetica","normal"); doc.setFontSize(8.5); doc.setTextColor(90);
  doc.text([firma.street,(firma.zip||"")+" "+(firma.city||""),firma.phone?"Tel. "+firma.phone:"",firma.email||""].filter(Boolean).join(" · "),20,26);
  doc.setFontSize(7); doc.text((firma.name||"")+" · "+(firma.street||"")+" · "+(firma.zip||"")+" "+(firma.city||""),20,47);
  doc.setTextColor(0); doc.setFontSize(10.5);
  doc.text([empf.name||"",empf.street||"",((empf.zip||"")+" "+(empf.city||"")).trim()].filter(Boolean),20,53);
  doc.text(ort,190,98,{align:"right"});
  var foot=[firma.name, firma.iban?"IBAN "+firma.iban:"", firma.bic?"BIC "+firma.bic:"", firma.vatId?"USt-IdNr. "+firma.vatId:(firma.taxNo?"St.-Nr. "+firma.taxNo:"")].filter(Boolean).join(" · ");
  doc.setFontSize(7.5); doc.setTextColor(110); doc.text(foot,105,287,{align:"center"}); doc.setTextColor(0);
  return doc;
}
function paragraphs(doc,y,parts,width){
  doc.setFontSize(10.5);
  parts.forEach(function(p){ var lines=doc.splitTextToSize(p,width||170); if(y+lines.length*5>270){ doc.addPage(); y=25; } doc.text(lines,20,y); y+=lines.length*5+3.5; });
  return y;
}

/* ---------- Mahnwesen ---------- */
var STUFEN=["Zahlungserinnerung","1. Mahnung","2. Mahnung"];
function mahnText(stufe, d){
  // d: {kunde, nr, datum, betrag, faellig, frist, leistung, firma}
  var titel=STUFEN[Math.min(stufe,2)];
  var bez="Rechnung Nr. "+d.nr+" vom "+d.datum+(d.leistung?" ("+d.leistung+")":"")+" über "+eur(d.betrag);
  var person=d.kunde&&/^[^,]+, [^,]+$/.test(d.kunde.trim()) && !/gmbh|eg$|\bag\b|kg|praxis|gbr|ug\b/i.test(d.kunde);
  var anrede=person?"Guten Tag "+d.kunde.replace(/^(.*), (.*)$/,"$2 $1")+",":"Sehr geehrte Damen und Herren,";
  var body;
  if(stufe<=0) body=[anrede,
    "sicher ist es Ihrer Aufmerksamkeit entgangen: Für unsere "+bez+" konnten wir bis heute keinen Zahlungseingang feststellen. Die Zahlung war am "+d.faellig+" fällig.",
    "Bitte überweisen Sie den offenen Betrag bis zum "+d.frist+". Sollten Sie die Zahlung inzwischen veranlasst haben, betrachten Sie dieses Schreiben bitte als gegenstandslos.",
    "Vielen Dank und freundliche Grüße"];
  else if(stufe===1) body=[anrede,
    "trotz unserer Zahlungserinnerung ist die "+bez+" weiterhin offen. Die Zahlung war am "+d.faellig+" fällig.",
    "Wir bitten Sie, den Betrag von "+eur(d.betrag)+" bis spätestens "+d.frist+" auf das unten genannte Konto zu überweisen.",
    "Falls es Rückfragen zur Rechnung gibt, melden Sie sich gern direkt bei uns.",
    "Freundliche Grüße"];
  else body=[anrede,
    "leider ist die "+bez+" trotz Erinnerung und Mahnung noch immer nicht beglichen.",
    "Wir fordern Sie hiermit letztmalig auf, den Betrag von "+eur(d.betrag)+" bis zum "+d.frist+" zu überweisen. Nach Ablauf dieser Frist behalten wir uns vor, weitere Schritte einzuleiten und die entstandenen Verzugskosten geltend zu machen.",
    "Freundliche Grüße"];
  body.push(d.firma&&d.firma.contactName?d.firma.contactName:(d.firma?d.firma.name:""));
  return {titel:titel, betreff:titel+": "+"Rechnung Nr. "+d.nr, absaetze:body};
}
function mahnPDF(jsPDF, firma, empf, text, ortDatum){
  var doc=letterBase(jsPDF, firma, empf, ortDatum);
  doc.setFont("helvetica","bold"); doc.setFontSize(12); doc.text(text.betreff,20,108); doc.setFont("helvetica","normal");
  paragraphs(doc,120,text.absaetze);
  if(firma.iban){ doc.setFontSize(10); doc.text("Bankverbindung: IBAN "+firma.iban+(firma.bic?" · BIC "+firma.bic:""),20,262); }
  return doc;
}

/* ---------- Rechnungsansicht (Begleit-PDF zur XRechnung) ---------- */
function invoicePDF(jsPDF, inv, totals){
  var s=inv.seller, b=inv.buyer, doc=letterBase(jsPDF, s, b, s.city+", "+inv.date);
  doc.setFont("helvetica","bold"); doc.setFontSize(13); doc.text("Rechnung Nr. "+inv.number,20,108);
  doc.setFont("helvetica","normal"); doc.setFontSize(9.5);
  doc.text(["Rechnungsdatum: "+inv.date, inv.deliveryDate?"Leistungsdatum: "+inv.deliveryDate:"", "Ihre Referenz: "+(inv.buyerRef||inv.number)].filter(Boolean),20,116);
  var y=134, cols=[20,30,118,140,165,190];
  doc.setFont("helvetica","bold"); ["Pos.","Leistung","Menge","Einzelpreis","Gesamt"].forEach(function(h,i){ doc.text(h,i>=2?cols[i+1]:cols[i],y,{align:i>=2?"right":"left"}); });
  doc.setDrawColor(180); doc.line(20,y+2,190,y+2); y+=8; doc.setFont("helvetica","normal");
  inv.lines.forEach(function(l,i){
    var name=doc.splitTextToSize(l.name,82);
    doc.text(String(i+1),20,y); doc.text(name,30,y);
    doc.text(de(l.qty,Number(l.qty)%1?2:0)+" "+(l.unit||""),140,y,{align:"right"}); doc.text(eur(l.price),165,y,{align:"right"}); doc.text(eur(l.total!=null?l.total:l.qty*l.price),190,y,{align:"right"});
    y+=name.length*4.6+2.5;
  });
  doc.line(120,y,190,y); y+=6;
  [["Nettobetrag",totals.net]].concat(totals.groups.map(function(g){ return ["MwSt "+de(g.rate,0)+" % auf "+eur(g.base),g.tax]; })).concat([["Rechnungsbetrag",totals.gross]]).forEach(function(r,i,a){
    if(i===a.length-1) doc.setFont("helvetica","bold"); doc.text(r[0],120,y); doc.text(eur(r[1]),190,y,{align:"right"}); y+=6; });
  doc.setFont("helvetica","normal"); y+=6;
  y=paragraphs(doc,y,[inv.paymentTerms||("Bitte überweisen Sie den Betrag bis zum "+inv.due+" ohne Abzug."), inv.note||""].filter(Boolean));
  doc.setFontSize(8); doc.setTextColor(110); doc.text("Die rechtsverbindliche Rechnung ist die beigefügte XRechnung (XML). Dieses PDF dient der Ansicht.",20,278); doc.setTextColor(0);
  return doc;
}

/* ---------- Abnahmeprotokoll ---------- */
function abnahmePDF(jsPDF, d){
  // d: {firma, kunde, adresse, auftrag, leistung, datum, ort, teilnehmer, maengel[], vorbehalt, fotos[{data,w,h}], sigKunde, sigFirma}
  var doc=new jsPDF({unit:"mm",format:"a4"}), y=22;
  doc.setFont("helvetica","bold"); doc.setFontSize(16); doc.text("Abnahmeprotokoll",20,y);
  doc.setFontSize(9); doc.setFont("helvetica","normal"); doc.setTextColor(90); doc.text((d.firma.name||"")+(d.firma.phone?" · Tel. "+d.firma.phone:""),190,y,{align:"right"}); doc.setTextColor(0); y+=12;
  var kv=[["Auftrag",d.auftrag],["Auftraggeber",d.kunde],["Ort der Leistung",d.adresse],["Leistung",d.leistung],["Datum der Abnahme",d.datum+(d.ort?", "+d.ort:"")],["Anwesend",d.teilnehmer]];
  doc.setFontSize(10.5);
  kv.forEach(function(r){ if(!r[1]) return; doc.setFont("helvetica","bold"); doc.text(r[0]+":",20,y); doc.setFont("helvetica","normal"); var l=doc.splitTextToSize(String(r[1]),120); doc.text(l,70,y); y+=l.length*5+2; });
  y+=4; doc.setFont("helvetica","bold"); doc.text("Ergebnis",20,y); y+=6; doc.setFont("helvetica","normal");
  var m=(d.maengel||[]).filter(function(x){ return String(x).trim(); });
  if(!m.length) y=paragraphs(doc,y,["Die Leistung wurde ohne Mängel abgenommen."]);
  else { y=paragraphs(doc,y,["Die Leistung wurde abgenommen. Folgende Mängel wurden festgestellt und werden nachgebessert:"]); m.forEach(function(x,i){ y=paragraphs(doc,y,[(i+1)+". "+x]); }); }
  if(d.vorbehalt) y=paragraphs(doc,y,["Vorbehalte: "+d.vorbehalt]);
  if(d.fotos&&d.fotos.length){
    y+=2; doc.setFont("helvetica","bold"); doc.text("Fotos",20,y); y+=4; doc.setFont("helvetica","normal");
    var x=20; d.fotos.forEach(function(f){ var w=54, h=Math.min(54,w*f.h/f.w); if(x+w>191){ x=20; y+=58; } if(y+h>250){ doc.addPage(); y=25; x=20; } doc.addImage(f.data,"JPEG",x,y,w,h); x+=58; });
    y+=60;
  }
  if(y>225){ doc.addPage(); y=30; }
  y=Math.max(y+8,230);
  [["Auftraggeber",d.sigKunde,20],["Auftragnehmer",d.sigFirma,115]].forEach(function(s){
    if(s[1]) doc.addImage(s[1],"PNG",s[2],y-22,70,20);
    doc.setDrawColor(120); doc.line(s[2],y,s[2]+75,y); doc.setFontSize(9); doc.text("Unterschrift "+s[0],s[2],y+5); });
  doc.setFontSize(8); doc.setTextColor(110); doc.text("Erstellt mit Zettelfrei · Vorlage, keine Rechtsberatung",105,287,{align:"center"}); doc.setTextColor(0);
  return doc;
}

/* ---------- Export für den Steuerberater (CSV, Excel-tauglich) ---------- */
function csv(rows){
  var cell=function(v){ if(v==null) return ""; if(typeof v==="number") return de(v,2).replace(/\./g,""); var s=String(v); return /[;"\n]/.test(s)?'"'+s.replace(/"/g,'""')+'"':s; };
  return "﻿"+rows.map(function(r){ return r.map(cell).join(";"); }).join("\r\n")+"\r\n";
}
function belegliste(ZF, model, sheet, art, monat){
  // art: "Ausgangsrechnung" | "Eingangsrechnung"; monat: "MM.JJJJ" oder ""
  var sh=sheet, calc=new ZF.Calc(model), H=function(re){ return ZF.col(sh,re); };
  var kDat=H(/rechnung am|rechnungsdatum|^datum$/i), kNr=H(/rechnungs-?nr|auftrags-?nr|beleg-?nr|^nr/i), kPar=H(/kunde|lieferant|firma|name/i),
      kNet=H(/netto|angebot/i), kSatz=H(/satz/i), kTax=H(/mwst|ust/i), kBr=H(/brutto/i), kBez=H(/bezahlt am|zahlung am/i), kPro=H(/projekt|leistung|kommission/i);
  if(kSatz===kTax) kTax=-1;
  var out=[["Belegdatum","Belegnummer","Art","Geschäftspartner","Netto","USt-Satz","USt","Brutto","Bezahlt am","Projekt/Leistung"]];
  sh.rows.forEach(function(r,i){
    var row=ZF.dataRowNo(sh,i), v=function(k){ return k<0?null:calc.value(sh,row,sh.c0+k); };
    var dat=v(kDat); if(!ZF.isDateStr(dat)) return;
    if(monat && dat.slice(3)!==monat) return;
    var net=calc.num(v(kNet)), tax=kTax>=0?calc.num(v(kTax)):NaN, satz=kSatz>=0?calc.num(v(kSatz)):NaN, br=kBr>=0?calc.num(v(kBr)):NaN;
    if(isNaN(satz)) satz=(!isNaN(tax)&&net)?Math.round(tax/net*100)/100:0.19; if(satz>1) satz=satz/100;
    if(isNaN(tax)) tax=Math.round(net*satz*100)/100; if(isNaN(br)) br=Math.round((net+tax)*100)/100;
    out.push([dat, v(kNr), art, v(kPar), net, Math.round(satz*100)+" %", tax, br, v(kBez)||"", v(kPro)||""]);
  });
  return out;
}

/* ---------- Wochenbericht ---------- */
function weekRange(today){
  var d=new Date(today.getFullYear(),today.getMonth(),today.getDate()), wd=(d.getDay()+6)%7, mon=new Date(d); mon.setDate(d.getDate()-wd);
  var sun=new Date(mon); sun.setDate(mon.getDate()+6);
  var f=function(x){ return String(x.getDate()).padStart(2,"0")+"."+String(x.getMonth()+1).padStart(2,"0")+"."+x.getFullYear(); };
  return {from:f(mon), to:f(sun)};
}
function weekReport(ZF, model, opts){
  opts=opts||{}; var calc=new ZF.Calc(model), lines=[], from=ZF.dateToSerial(opts.from), to=ZF.dateToSerial(opts.to);
  lines.push("Wochenbericht "+opts.from+" – "+opts.to+" · "+model.name);
  model.sheets.forEach(function(sh){
    if(sh.hr<0||!sh.rows.length) return;
    var H=function(re){ return ZF.col(sh,re); }, V=function(i,k){ return k<0?null:calc.value(sh,ZF.dataRowNo(sh,i),sh.c0+k); };
    var kDat=H(/^datum$|^tag$/i), kStd=H(/^stunden|^std/i), kPro=H(/projekt|kommission|baustelle/i), kMa=H(/mitarbeiter/i);
    if(kDat>=0 && kStd>=0){
      var byP={}, byM={}, sum=0;
      sh.rows.forEach(function(r,i){ var d=V(i,kDat); if(!ZF.isDateStr(d)) return; var s=ZF.dateToSerial(d); if(s<from||s>to) return;
        var h=calc.num(V(i,kStd)); if(isNaN(h)) return; sum+=h; var p=V(i,kPro)||"ohne Projekt", m=V(i,kMa)||"?"; byP[p]=(byP[p]||0)+h; byM[m]=(byM[m]||0)+h; });
      if(sum){ lines.push("","Stunden („"+sh.name+"“): "+de(sum,1)+" Std.");
        Object.keys(byP).sort(function(a,b){ return byP[b]-byP[a]; }).forEach(function(p){ lines.push("  · "+p+": "+de(byP[p],1)+" Std."); });
        lines.push("  Mitarbeiter: "+Object.keys(byM).map(function(m){ return m+" "+de(byM[m],1); }).join(", ")); }
    }
    var kBest=H(/bestand/i), kMin=H(/mindest/i), kArt=H(/artikel$|^artikel |bezeichnung/i);
    if(kBest>=0 && kMin>=0){
      var low=[]; sh.rows.forEach(function(r,i){ var b=calc.num(V(i,kBest)), mi=calc.num(V(i,kMin)); if(!isNaN(b)&&!isNaN(mi)&&mi>0&&b<mi) low.push(V(i,kArt>=0?kArt:0)+" ("+de(b,b%1?1:0)+" von min. "+de(mi,0)+")"); });
      if(low.length){ lines.push("","Nachbestellen („"+sh.name+"“):"); low.forEach(function(x){ lines.push("  · "+x); }); }
    }
    var kRe=H(/rechnung am|rechnungsdatum/i), kBez=H(/bezahlt am/i), kBr=H(/brutto/i), kPar=H(/kunde|lieferant/i), kFae=H(/fällig/i);
    if(kRe>=0 && kBez>=0 && kBr>=0){
      var open=[], total=0;
      sh.rows.forEach(function(r,i){ if(!ZF.isDateStr(V(i,kRe))) return; var b=V(i,kBez); if(b!=null&&b!=="") return; var br=calc.num(V(i,kBr)); total+=br; open.push(V(i,kPar)+" "+de(br)+" €"+(kFae>=0&&ZF.isDateStr(V(i,kFae))?" (fällig "+V(i,kFae)+")":"")); });
      if(open.length){ lines.push("",(kFae>=0?"Offene Eingangsrechnungen":"Offene Rechnungen an Kunden")+" („"+sh.name+"“): "+de(total)+" €"); open.forEach(function(x){ lines.push("  · "+x); }); }
    }
  });
  var iss=ZF.checks(model,{today:opts.today}).filter(function(x){ return /ArbZG|überfällig|offen/.test(x.problem); });
  if(iss.length){ lines.push("","Bitte prüfen:"); iss.slice(0,8).forEach(function(x){ lines.push("  · "+x.blatt+", Zeile "+x.zeile+": "+x.problem); }); }
  if(lines.length===1) lines.push("","Für diese Woche gibt es keine Einträge.");
  return lines.join("\n");
}

return {mahnText:mahnText, mahnPDF:mahnPDF, invoicePDF:invoicePDF, abnahmePDF:abnahmePDF, csv:csv, belegliste:belegliste, weekReport:weekReport, weekRange:weekRange, STUFEN:STUFEN, eur:eur, de:de};
})();
if(typeof module!=="undefined") module.exports=ZFD;
