/* Zettelfrei – Bereich „Dokumente“: E-Rechnung schreiben, Mahnwesen, Abnahmeprotokoll,
 * Export für den Steuerberater, Wochenbericht, Firmendaten. */
var ZFDOC = (function(){
"use strict";
var A=function(){ return window.ZFAPP; }, $=function(id){ return document.getElementById(id); };
var esc=function(s){ return A().esc(s); }, de=function(n,d){ return ZFD.de(n,d); };
function M(){ return A().model(); }
function today(){ return A().today(); }
function jsPDF(){ return window.jspdf&&window.jspdf.jsPDF; }

/* ---------- Firmendaten ---------- */
var FIELDS=[["fName","name"],["fContact","contactName"],["fStreet","street"],["fZip","zip"],["fCity","city"],["fPhone","phone"],["fEmail","email"],
  ["fVat","vatId"],["fTaxNo","taxNo"],["fReg","regNo"],["fIban","iban"],["fBic","bic"],["fBank","bank"],["fZiel","zahlungsziel"],["fBearbeiter","bearbeiter"]];
function firma(){ var f=A().firma(); f.country="DE"; return f; }
function openFirma(){ var f=A().firma(); FIELDS.forEach(function(x){ $(x[0]).value=f[x[1]]||(x[1]==="zahlungsziel"?"14":""); }); var d=$("firmaDlg"); if(d.showModal) d.showModal(); else d.setAttribute("open",""); }
function saveFirma(){ var f={}; FIELDS.forEach(function(x){ f[x[1]]=$(x[0]).value.trim(); }); try{ localStorage.setItem("zf-firma",JSON.stringify(f)); }catch(e){} A().toast("Firmendaten gespeichert"); render(); }
function firmaMissing(){ var f=firma(), m=[];
  if(!f.name) m.push("Firmenname"); if(!f.street||!f.zip||!f.city) m.push("Adresse"); if(!f.vatId&&!f.taxNo) m.push("USt-IdNr. oder Steuernummer");
  if(!f.contactName||!f.phone||!f.email) m.push("Ansprechpartner, Telefon, E-Mail"); if(!f.iban) m.push("IBAN"); return m; }

/* ---------- Hilfen für Tabellen ---------- */
function orderSheets(){ return M().sheets.filter(function(sh){ return sh.hr>=0 && ZF.col(sh,/kunde|auftraggeber/i)>=0 && ZF.col(sh,/netto|angebot|betrag/i)>=0; }); }
function rowObj(sh,i){
  var calc=new ZF.Calc(M()), r=ZF.dataRowNo(sh,i), V=function(re){ var k=ZF.col(sh,re); return k<0?null:calc.value(sh,r,sh.c0+k); };
  var net=calc.num(V(/netto|angebot|betrag/i)), br=V(/brutto/i);
  return { sheet:sh.name, row:r, nr:V(/auftrags-?nr|rechnungs-?nr|^nr/i)||String(r), kunde:V(/kunde|auftraggeber/i)||"", adresse:V(/adresse|anschrift|ort/i)||"",
    leistung:V(/leistung|beschreibung|projekt/i)||"", status:V(/status/i)||"", netto:isNaN(net)?0:net, brutto:typeof br==="number"?br:Math.round((isNaN(net)?0:net)*119)/100,
    reAm:V(/rechnung am|rechnungsdatum/i), bezahlt:V(/bezahlt am|zahlung am/i), mahn:V(/mahnstufe/i), email:V(/e-?mail/i), abnahme:V(/abnahme am/i) };
}
function splitAddr(a){ var m=String(a||"").match(/^(.*?),\s*(\d{5})\s+(.+)$/); return m?{street:m[1].trim(),zip:m[2],city:m[3].trim()}:{street:String(a||""),zip:"",city:""}; }
function mailBook(){ try{ return JSON.parse(localStorage.getItem("zf-kunden-mail")||"{}"); }catch(e){ return {}; } }
function remember(k,v){ var b=mailBook(); b[k]=v; try{ localStorage.setItem("zf-kunden-mail",JSON.stringify(b)); }catch(e){} }
function apply(ops,label){ var m=M(), plan=ZF.planOps(m,ops).filter(function(p){ return !p.error; }); if(!plan.length) return; var t=ZF.applyPlan(m,plan); A().mark(t.map(function(x){ return x.row; }),t[0]?t[0].sheet:""); A().commit(label); }
function ensureCols(sheetName,names){ var sh=M().sheets.find(function(s){ return s.name===sheetName; }); names.forEach(function(n){ ZF.addColumn(sh,n); }); }
function rowOptions(list,sel){ return list.map(function(o,i){ return '<option value="'+i+'"'+(i===sel?" selected":"")+'>'+esc(o.nr+" · "+o.kunde+(o.status?" · "+o.status:"")+(o.reAm?" · Rechnung "+o.reAm:""))+"</option>"; }).join(""); }
function allOrders(){ var out=[]; orderSheets().forEach(function(sh){ sh.rows.forEach(function(_,i){ if(ZF.isData(sh,i)) out.push(rowObj(sh,i)); }); }); return out; }

/* ---------- Oberfläche ---------- */
var state={inv:0,abn:0};
function render(){
  if(!M()) return;
  var orders=allOrders(), miss=firmaMissing(), f=firma();
  var html='';
  html+='<section class="doc"><div class="doc-h"><h3>Firmendaten</h3><button class="btn quiet small" type="button" data-act="firma">Bearbeiten</button></div>'+
    (miss.length?'<p class="hint warnc">Für E-Rechnungen fehlen noch: '+esc(miss.join(", "))+'.</p>':'<p class="hint">'+esc(f.name)+' · '+esc(f.city)+' · '+(f.vatId?"USt-IdNr. "+esc(f.vatId):"St.-Nr. "+esc(f.taxNo))+'</p>')+'</section>';
  // E-Rechnung
  html+='<section class="doc"><h3>E-Rechnung schreiben</h3><p class="hint">XRechnung 3.0 aus einer Auftragszeile, geprüft gegen die Regeln EN 16931 und XRechnung. Dazu ein PDF zur Ansicht.</p>';
  if(!orders.length) html+='<p class="hint">Dafür braucht es eine Auftragsliste mit Kunde und Nettobetrag, z. B. das Beispiel „Auftraege 2026“.</p>';
  else { var o=orders[Math.min(state.inv,orders.length-1)];
    html+='<label class="fl">Auftrag<select id="invRow">'+rowOptions(orders,Math.min(state.inv,orders.length-1))+'</select></label>'+
      '<div class="grid2"><label class="fl">Rechnungs-Nr.<input id="invNr" value="'+esc(o.nr)+'"></label><label class="fl">Leistungsdatum<input id="invLd" value="'+esc(today())+'"></label>'+
      '<label class="fl">E-Mail des Kunden<input id="invMail" type="email" value="'+esc(o.email||mailBook()[o.kunde]||"")+'" placeholder="kunde@beispiel.de"></label>'+
      '<label class="fl">MwSt<select id="invRate"><option value="19">19 %</option><option value="7">7 %</option></select></label></div>'+
      '<label class="fl">Position<input id="invPos" value="'+esc(o.leistung)+'"></label>'+
      '<p class="hint">Netto '+de(o.netto)+' € · Zahlungsziel '+esc(f.zahlungsziel||"14")+' Tage'+(o.reAm?' · <b class="warnc">Rechnung wurde am '+esc(o.reAm)+' schon gestellt</b>':'')+'</p>'+
      '<div class="row"><button class="btn" type="button" data-act="xr">XRechnung + PDF erstellen</button></div><div id="invMsg"></div>'; }
  html+='</section>';
  // Mahnwesen
  var ziel=+(f.zahlungsziel||14), now=ZF.todaySerial(), over=orders.filter(function(o){ return ZF.isDateStr(o.reAm) && (o.bezahlt==null||o.bezahlt==="") && now-ZF.dateToSerial(o.reAm)>ziel; });
  html+='<section class="doc"><h3>Mahnwesen</h3><p class="hint">Rechnungen, die länger als '+ziel+' Tage offen sind. Ein Klick erstellt das passende Schreiben als PDF und trägt die Mahnstufe ein.</p>';
  if(!over.length) html+='<p class="hint">Keine überfälligen Rechnungen.</p>';
  else html+='<ul class="dl">'+over.map(function(o){ var lvl=ZFD.STUFEN.indexOf(o.mahn)+1, days=now-ZF.dateToSerial(o.reAm)-ziel, next=ZFD.STUFEN[Math.min(lvl,2)];
    return '<li><div><b>'+esc(o.kunde)+'</b> · '+esc(o.nr)+'<br><small>'+ZFD.eur(o.brutto)+' · '+days+' Tage überfällig'+(o.mahn?' · zuletzt: '+esc(o.mahn):'')+'</small></div>'+
      '<button class="btn quiet small" type="button" data-act="mahn" data-sheet="'+esc(o.sheet)+'" data-row="'+o.row+'">'+esc(next)+'</button></li>'; }).join("")+'</ul>';
  html+='</section>';
  // Abnahme
  html+='<section class="doc"><h3>Abnahmeprotokoll</h3><p class="hint">Auf dem Handy direkt beim Kunden: Mängel, Fotos, zwei Unterschriften. Ergibt ein PDF und das Abnahmedatum in der Tabelle.</p>';
  if(orders.length) html+='<label class="fl">Auftrag<select id="abnRow">'+rowOptions(orders,Math.min(state.abn,orders.length-1))+'</select></label><div class="row"><button class="btn quiet" type="button" data-act="abn">Abnahme aufnehmen</button></div>';
  else html+='<p class="hint">Dafür braucht es eine Auftragsliste.</p>';
  html+='</section>';
  // Export
  var exp=M().sheets.filter(function(sh){ return sh.hr>=0 && ZF.col(sh,/netto|angebot/i)>=0 && ZF.col(sh,/rechnung am|rechnungsdatum|^datum$/i)>=0; });
  html+='<section class="doc"><h3>Für den Steuerberater</h3><p class="hint">Belegliste als CSV (Excel-tauglich): Datum, Nummer, Partner, Netto, USt, Brutto, bezahlt am. Kein DATEV-Format.</p>';
  if(!exp.length) html+='<p class="hint">Dafür braucht es ein Blatt mit Rechnungsdatum und Nettobetrag.</p>';
  else { var months={}; exp.forEach(function(sh){ var calc=new ZF.Calc(M()), k=ZF.col(sh,/rechnung am|rechnungsdatum|^datum$/i); sh.rows.forEach(function(_,i){ if(!ZF.isData(sh,i)) return; var d=calc.value(sh,ZF.dataRowNo(sh,i),sh.c0+k); if(ZF.isDateStr(d)) months[d.slice(3)]=1; }); });
    html+='<div class="grid2"><label class="fl">Blatt<select id="expSheet">'+exp.map(function(sh){ return '<option>'+esc(sh.name)+'</option>'; }).join("")+'</select></label>'+
      '<label class="fl">Monat<select id="expMonth"><option value="">alle</option>'+Object.keys(months).sort(function(a,b){ return a.slice(3)+a.slice(0,2)<b.slice(3)+b.slice(0,2)?1:-1; }).map(function(m){ return '<option>'+m+'</option>'; }).join("")+'</select></label></div>'+
      '<div class="row"><button class="btn quiet" type="button" data-act="csv">CSV herunterladen</button></div>'; }
  html+='</section>';
  // Wochenbericht
  html+='<section class="doc"><h3>Wochenbericht</h3><p class="hint">Stunden je Projekt, Nachbestellungen, offene Rechnungen und Hinweise zur Arbeitszeit – als Nachricht an den Chef.</p>'+
    '<div class="row"><select id="wkSel"><option value="0">diese Woche</option><option value="1">letzte Woche</option></select><button class="btn quiet" type="button" data-act="week">Bericht erstellen</button></div><div id="wkOut"></div></section>';
  $("docs").innerHTML=html;
}

/* ---------- Aktionen ---------- */
function makeInvoice(){
  var orders=allOrders(), o=orders[+$("invRow").value], f=firma(), msg=$("invMsg");
  var miss=firmaMissing(); if(miss.length){ msg.innerHTML='<p class="hint warnc">Bitte zuerst die Firmendaten ergänzen: '+esc(miss.join(", "))+'. <button class="linkbtn" type="button" data-act="firma">Jetzt ergänzen</button></p>'; return; }
  var ad=splitAddr(o.adresse), mail=$("invMail").value.trim(), rate=+$("invRate").value, ziel=+(f.zahlungsziel||14);
  if(mail) remember(o.kunde,mail);
  var inv={ number:$("invNr").value.trim(), date:today(), due:ZFE.addDays(today(),ziel), deliveryDate:$("invLd").value.trim()||today(), buyerRef:o.nr, project:o.leistung,
    paymentTerms:"Zahlbar innerhalb von "+ziel+" Tagen ohne Abzug bis "+ZFE.addDays(today(),ziel)+".",
    seller:f, buyer:{name:o.kunde,street:ad.street,zip:ad.zip,city:ad.city,email:mail},
    lines:[{name:$("invPos").value.trim()||o.leistung||"Leistung",qty:1,unit:"pauschal",price:o.netto,rate:rate}] };
  var m=ZFE.checkData(inv); if(m.length){ msg.innerHTML='<p class="hint warnc">Es fehlt noch: '+esc(m.join(", "))+'.</p>'; return; }
  if(!ad.zip){ msg.innerHTML='<p class="hint warnc">Die Adresse des Kunden braucht das Format „Straße Nr., PLZ Ort“.</p>'; return; }
  var res=ZFE.buildXRechnung(inv);
  A().saveFile("XRechnung_"+inv.number+".xml",res.xml,"application/xml").then(function(){
    var J=jsPDF(); if(J) return A().saveFile("Rechnung_"+inv.number+".pdf",ZFD.invoicePDF(J,inv,res.totals).output("blob"),"application/pdf"); });
  var ops=[{typ:"aendern",blatt:o.sheet,zeile:o.row,werte:{}}]; ensureCols(o.sheet,["Rechnung am"]); ops[0].werte["Rechnung am"]=today();
  apply(ops,"E-Rechnung "+inv.number+" an "+o.kunde+" erstellt ("+ZFD.eur(res.totals.gross)+")");
  msg.innerHTML='<p class="hint okc">XRechnung '+esc(inv.number)+' über '+ZFD.eur(res.totals.gross)+' erstellt. „Rechnung am“ ist eingetragen.</p>';
}
function makeMahnung(sheet,row){
  var o=allOrders().find(function(x){ return x.sheet===sheet&&x.row===row; }); if(!o) return;
  var f=firma(), lvl=Math.min(ZFD.STUFEN.indexOf(o.mahn)+1,2), ziel=+(f.zahlungsziel||14), ad=splitAddr(o.adresse);
  var text=ZFD.mahnText(lvl,{kunde:o.kunde,nr:o.nr,datum:o.reAm,betrag:o.brutto,faellig:ZFE.addDays(o.reAm,ziel),frist:ZFE.addDays(today(),10),leistung:o.leistung,firma:f});
  var J=jsPDF(); if(!J){ A().toast("PDF-Erzeugung nicht verfügbar."); return; }
  var doc=ZFD.mahnPDF(J,f,{name:o.kunde,street:ad.street,zip:ad.zip,city:ad.city},text,(f.city||"")+", "+today());
  A().saveFile(text.titel.replace(/[^\wäöüÄÖÜß-]+/g,"_")+"_"+o.nr+".pdf",doc.output("blob"),"application/pdf");
  ensureCols(o.sheet,["Mahnstufe","Gemahnt am"]);
  apply([{typ:"aendern",blatt:o.sheet,zeile:o.row,werte:{"Mahnstufe":ZFD.STUFEN[lvl],"Gemahnt am":today()}}],text.titel+" an "+o.kunde+" ("+o.nr+")");
}
function csvExport(){
  var sh=M().sheets.find(function(s){ return s.name===$("expSheet").value; }), mon=$("expMonth").value;
  var art=ZF.col(sh,/lieferant|kreditor/i)>=0?"Eingangsrechnung":"Ausgangsrechnung", rows=ZFD.belegliste(ZF,M(),sh,art,mon);
  if(rows.length<2){ A().toast("Keine Belege in diesem Zeitraum."); return; }
  A().saveFile("Belege_"+sh.name.replace(/[^\wäöüÄÖÜß-]+/g,"_")+(mon?"_"+mon.replace(".","-"):"")+".csv",ZFD.csv(rows),"text/csv");
}
function week(){
  var back=+$("wkSel").value, d=new Date(); d.setDate(d.getDate()-7*back);
  var r=ZFD.weekRange(d), txt=ZFD.weekReport(ZF,M(),{from:r.from,to:r.to,today:new Date()});
  $("wkOut").innerHTML='<pre class="report">'+esc(txt)+'</pre><div class="row"><button class="btn quiet small" type="button" data-act="wkcopy">Kopieren</button>'+
    '<a class="btn quiet small" target="_blank" rel="noopener" href="https://wa.me/?text='+encodeURIComponent(txt)+'">Per WhatsApp teilen</a>'+
    '<button class="btn quiet small" type="button" data-act="wktxt">Als Text speichern</button></div>';
  state.week=txt;
}

/* ---------- Abnahme-Dialog mit Unterschrift ---------- */
function pad(canvas){
  var ctx=canvas.getContext("2d"), drawing=false, ink=false;
  var size=function(){ var r=canvas.getBoundingClientRect(), dpr=window.devicePixelRatio||1; canvas.width=Math.max(10,r.width*dpr); canvas.height=Math.max(10,r.height*dpr); ctx.setTransform(dpr,0,0,dpr,0,0); ctx.lineWidth=2.2; ctx.lineCap="round"; ctx.lineJoin="round"; ctx.strokeStyle="#1d2528"; ink=false; };
  var pos=function(e){ var r=canvas.getBoundingClientRect(); return [e.clientX-r.left,e.clientY-r.top]; };
  canvas.addEventListener("pointerdown",function(e){ drawing=true; canvas.setPointerCapture(e.pointerId); var p=pos(e); ctx.beginPath(); ctx.moveTo(p[0],p[1]); e.preventDefault(); });
  canvas.addEventListener("pointermove",function(e){ if(!drawing) return; var p=pos(e); ctx.lineTo(p[0],p[1]); ctx.stroke(); ink=true; });
  ["pointerup","pointercancel","pointerleave"].forEach(function(t){ canvas.addEventListener(t,function(){ drawing=false; }); });
  return {reset:size, has:function(){ return ink; }, data:function(){ return canvas.toDataURL("image/png"); }};
}
var pads=null, fotos=[];
function openAbnahme(){
  var o=allOrders()[+$("abnRow").value]; if(!o) return;
  $("abnTitle").textContent="Abnahme: "+o.nr+" · "+o.kunde; $("abnDatum").value=today(); $("abnAnw").value=(firma().contactName?firma().contactName+", ":"")+o.kunde;
  $("abnMaengel").value=""; $("abnVorb").value=""; fotos=[]; $("abnFotos").innerHTML=""; $("abnMsg").textContent="";
  var d=$("abnDlg"); if(d.showModal) d.showModal(); else d.setAttribute("open","");
  if(!pads) pads={k:pad($("sigKunde")),f:pad($("sigFirma"))};
  requestAnimationFrame(function(){ pads.k.reset(); pads.f.reset(); });
  d.dataset.idx=$("abnRow").value;
}
function addFotos(files){
  [].slice.call(files,0,6).forEach(function(file){
    var url=URL.createObjectURL(file), img=new Image();
    img.onload=function(){ var s=Math.min(1,1200/Math.max(img.width,img.height)), c=document.createElement("canvas"); c.width=img.width*s; c.height=img.height*s; c.getContext("2d").drawImage(img,0,0,c.width,c.height);
      var data=c.toDataURL("image/jpeg",0.82); fotos.push({data:data,w:c.width,h:c.height}); URL.revokeObjectURL(url);
      var t=document.createElement("img"); t.src=data; t.alt="Foto "+fotos.length; $("abnFotos").appendChild(t); };
    img.src=url; });
}
function finishAbnahme(){
  var o=allOrders()[+$("abnDlg").dataset.idx], J=jsPDF();
  if(!pads.k.has()){ $("abnMsg").textContent="Bitte unterschreiben lassen: Auftraggeber fehlt."; return; }
  if(!J){ $("abnMsg").textContent="PDF-Erzeugung nicht verfügbar."; return; }
  var datum=$("abnDatum").value.trim()||today(), maengel=$("abnMaengel").value.split(/\n+/).map(function(s){ return s.trim(); }).filter(Boolean);
  var doc=ZFD.abnahmePDF(J,{firma:firma(),kunde:o.kunde,adresse:o.adresse,auftrag:o.nr,leistung:o.leistung,datum:datum,teilnehmer:$("abnAnw").value.trim(),
    maengel:maengel,vorbehalt:$("abnVorb").value.trim(),fotos:fotos,sigKunde:pads.k.data(),sigFirma:pads.f.has()?pads.f.data():null});
  A().saveFile("Abnahmeprotokoll_"+o.nr+".pdf",doc.output("blob"),"application/pdf");
  ensureCols(o.sheet,["Abnahme am"]);
  var w={"Abnahme am":datum}; if(maengel.length){ ensureCols(o.sheet,["Mängel"]); w["Mängel"]=maengel.join("; "); }
  apply([{typ:"aendern",blatt:o.sheet,zeile:o.row,werte:w}],"Abnahme "+o.nr+" ("+o.kunde+")"+(maengel.length?" mit "+maengel.length+" Mängeln":" ohne Mängel"));
  $("abnDlg").close();
}

/* ---------- Ereignisse ---------- */
function init(){
  $("docs").addEventListener("click",function(e){ var b=e.target.closest("[data-act]"); if(!b) return; var a=b.dataset.act;
    if(a==="firma") openFirma(); else if(a==="xr") makeInvoice(); else if(a==="mahn") makeMahnung(b.dataset.sheet,+b.dataset.row);
    else if(a==="abn") openAbnahme(); else if(a==="csv") csvExport(); else if(a==="week") week();
    else if(a==="wkcopy"){ navigator.clipboard.writeText(state.week).then(function(){ A().toast("Bericht kopiert"); },function(){ A().toast("Kopieren nicht möglich, bitte markieren."); }); }
    else if(a==="wktxt") A().saveFile("Wochenbericht.txt",state.week,"text/plain"); });
  $("docs").addEventListener("change",function(e){ if(e.target.id==="invRow"){ state.inv=+e.target.value; render(); } if(e.target.id==="abnRow") state.abn=+e.target.value; });
  $("firmaForm").addEventListener("submit",function(e){ if(e.submitter&&e.submitter.value==="save") saveFirma(); });
  $("abnFotoIn").addEventListener("change",function(e){ addFotos(e.target.files); e.target.value=""; });
  $("abnOk").addEventListener("click",finishAbnahme);
  $("abnCancel").addEventListener("click",function(){ $("abnDlg").close(); });
  document.querySelectorAll("[data-clear]").forEach(function(b){ b.addEventListener("click",function(){ pads[b.dataset.clear].reset(); }); });
}
return {init:init, render:render, openFirma:openFirma};
})();
ZFDOC.init();
