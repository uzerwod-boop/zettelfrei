/* Zettelfrei – Oberfläche */
(function(){
"use strict";
var EXAMPLES = window.ZF_EXAMPLES || {};
var NOW = new Date();
var TODAY = String(NOW.getDate()).padStart(2,"0")+"."+String(NOW.getMonth()+1).padStart(2,"0")+"."+NOW.getFullYear();
var WEEKDAY = ["Sonntag","Montag","Dienstag","Mittwoch","Donnerstag","Freitag","Samstag"][NOW.getDay()];
var SAMPLE_INVOICE = "Fliesen-Großhandel Weber KG · Industriestr. 14 · 04129 Leipzig\nRECHNUNG Nr. 88213        Datum: 15.10.2026\nKunde: Fliesen Krause, Kommission: Küche Yilmaz\nPos  Artikel                                Menge   Einh.   Einzelpreis   Gesamt\n1    Feinsteinzeug 60x60 anthrazit          18,5    m²      32,90 €       608,65 €\n2    Abschlussprofil Alu 2,5 m              8       Stk     9,80 €        78,40 €\n3    Flexkleber S1 25 kg                    6       Sack    17,50 €       105,00 €\n4    Fugenmörtel weiß 5 kg                  4       Eimer   11,40 €       45,60 €\nNetto 837,65 €   MwSt 19 % 159,15 €   Brutto 996,80 €\nZahlbar bis 29.10.2026 ohne Abzug.";

var model=null, history=[], cur=0, fresh={}, hits={}, aiIssues=[], downloads=null;
var $=function(id){ return document.getElementById(id); };
var esc=function(s){ return String(s==null?"":s).replace(/[&<>"]/g,function(c){ return {"&":"&amp;","<":"&lt;",">":"&gt;","\"":"&quot;"}[c]; }); };
function toast(m){ var t=$("toast"); t.textContent=m; t.hidden=false; clearTimeout(toast._t); toast._t=setTimeout(function(){ t.hidden=true; },2800); }
function de(n,d){ return Number(n).toLocaleString("de-DE",{minimumFractionDigits:d||0,maximumFractionDigits:d==null?2:d}); }
function sheetObj(){ return model.sheets[cur]; }

/* ---------- Speichern im Browser ---------- */
function persist(){ try{ localStorage.setItem("zf2",JSON.stringify({model:model,history:history.slice(-15),cur:cur})); }catch(e){} }
function restore(){ try{ var s=JSON.parse(localStorage.getItem("zf2")||"null"); if(s&&s.model&&s.model.sheets){ model=s.model; history=s.history||[]; cur=Math.min(s.cur||0,model.sheets.length-1); return true; } }catch(e){} return false; }
function commit(label){ history.push({t:Date.now(),label:label,snap:JSON.stringify(model)}); if(history.length>40) history=history.slice(-40); persist(); renderAll(); }

/* ---------- Datei öffnen / speichern ---------- */
function b64(s){ var bin=atob(s), a=new Uint8Array(bin.length); for(var i=0;i<bin.length;i++) a[i]=bin.charCodeAt(i); return a; }
function openData(data,name){
  if(!window.XLSX){ toast("Der Excel-Leser konnte nicht geladen werden. Bitte Seite neu laden."); return; }
  try{ var m=ZF.loadWorkbook(XLSX,data,name); }catch(e){ toast("Diese Datei lässt sich nicht lesen."); return; }
  if(!m.sheets.length){ toast("Die Datei ist leer."); return; }
  model=m; cur=Math.max(0,m.sheets.findIndex(function(s){ return s.hr>=0; })); fresh={}; hits={}; aiIssues=[];
  history=[{t:Date.now(),label:"Geöffnet: "+name,snap:JSON.stringify(model)}];
  $("cmdOut").innerHTML=""; $("pdfOut").innerHTML=""; $("fxAddr").textContent="–"; $("fxVal").textContent="";
  persist(); renderAll(); renderExamples();
  var parsed=m.sheets.filter(function(s){ return s.hr>=0; });
  toast(parsed.length? name+" geöffnet · "+parsed.map(function(s){ return s.rows.length+" Zeilen in „"+s.name+"“"; }).join(", ") : "Keine Tabelle mit Überschriften gefunden.");
}
Object.keys(EXAMPLES).forEach(function(n){ var o=document.createElement("option"); o.value=n; o.textContent=n.replace(/_/g," ").replace(/\.xlsx$/,""); $("exampleSel").appendChild(o); });
$("exampleSel").addEventListener("change",function(){ var n=this.value; if(n) openData(b64(EXAMPLES[n]),n); this.value=""; });
$("openBtn").addEventListener("click",function(){ $("openInput").click(); });
$("openInput").addEventListener("change",function(e){ var f=e.target.files[0]; e.target.value=""; if(f) f.arrayBuffer().then(function(b){ openData(new Uint8Array(b),f.name); }); });
var sw=document.querySelector(".sheetwrap");
sw.addEventListener("dragover",function(e){ if([].some.call(e.dataTransfer.items||[],function(i){ return i.kind==="file"; })) e.preventDefault(); });
sw.addEventListener("drop",function(e){ var f=e.dataTransfer.files[0]; if(!f) return; e.preventDefault();
  if(/\.(xlsx|xlsm|xls|csv|ods)$/i.test(f.name)) f.arrayBuffer().then(function(b){ openData(new Uint8Array(b),f.name); });
  else { selectTab("p-pdf"); handleFile(f); } });
$("saveBtn").addEventListener("click",function(){
  if(!model||!window.XLSX) return;
  var buf; try{ buf=ZF.buildWorkbook(XLSX,model); }catch(e){ toast("Speichern fehlgeschlagen."); return; }
  var name=model.name.replace(/\.(xls|csv|ods|xlsm)$/i,".xlsx"); if(!/\.xlsx$/i.test(name)) name+=".xlsx";
  var blob=new Blob([buf],{type:"application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"});
  if(downloads){ downloads.save({filename:name,data:blob}).then(function(){ toast(name+" gespeichert"); },function(e){ if(e&&e.code!=="declined") toast("Speichern ist hier nicht möglich."); }); return; }
  var a=document.createElement("a"); a.href=URL.createObjectURL(blob); a.download=name; document.body.appendChild(a); a.click();
  setTimeout(function(){ URL.revokeObjectURL(a.href); a.remove(); },1000); toast(name+" gespeichert");
});

/* ---------- Tabelle zeichnen ---------- */
function fmtVal(sh,k,v,cell){
  if(v==null||v==="") return "";
  if(typeof v==="number"){
    var z=(cell&&cell.z)||(k>=0?sh.colFmt[k]:null)||"", h=k>=0?sh.headers[k]:"";
    var money=/€|EUR/.test(z)||/€/.test(h), two=money||/0\.00/.test(z);
    return de(v,two?2:null)+(money?" €":"");
  }
  return String(v);
}
function rawOf(cell){ if(!cell) return ""; if(cell.f) return "="+cell.f; if(typeof cell.v==="number") return String(cell.v).replace(".",","); return cell.v==null?"":String(cell.v); }
function issuesFor(){ var all=ZF.checks(model).concat(aiIssues); return all; }
function renderTabs(){
  $("sheetTabs").innerHTML=model.sheets.map(function(s,i){ return '<button role="tab" aria-selected="'+(i===cur)+'" data-i="'+i+'" title="'+esc(s.name)+'">'+esc(s.name)+'</button>'; }).join("");
  var opts=model.sheets.map(function(s,i){ return s.hr>=0?'<option value="'+i+'">'+esc(s.name)+'</option>':""; }).join("");
  var sel=$("pdfTarget"), keep=sel.value; sel.innerHTML=opts; if(keep&&sel.querySelector('option[value="'+keep+'"]')) sel.value=keep; else sel.value=String(cur);
}
$("sheetTabs").addEventListener("click",function(e){ var b=e.target.closest("button"); if(!b) return; cur=+b.dataset.i; persist(); renderAll(); });
function renderGrid(){
  var sh=sheetObj(), calc=new ZF.Calc(model), issues=issuesFor().filter(function(x){ return x.blatt===sh.name; });
  var flagMap={}; issues.forEach(function(x){ var key=x.zeile+":"+(x.k!=null?x.k:ZF.findCol(sh,x.spalte)); flagMap[key]=(flagMap[key]?flagMap[key]+"; ":"")+x.problem; });
  $("fileName").textContent=model.name;
  if(sh.hr<0){
    var w=0; sh.top.forEach(function(r){ w=Math.max(w,(r||[]).length); });
    var h='<thead><tr><th class="rn"></th>'; for(var c=0;c<w;c++) h+="<th>"+ZF.colLetter(c)+"</th>"; h+="</tr></thead><tbody>";
    sh.top.forEach(function(r,ri){ h+='<tr class="top"><th class="rn">'+(ri+1)+"</th>"; for(var c=0;c<w;c++){ var v=calc.value(sh,ri+1,c); h+="<td>"+esc(fmtVal(sh,-1,v,r[c]))+"</td>"; } h+="</tr>"; });
    $("grid").innerHTML=h+"</tbody>";
    $("sumbar").innerHTML="<span>In diesem Blatt wurde keine Tabelle mit Überschriften erkannt. Es wird beim Speichern unverändert übernommen.</span>";
    return;
  }
  var C=sh.c0+sh.ncol, comp=ZF.computedCols(sh), numc=sh.headers.map(function(_,k){ return comp[k]||ZF.numericCol(sh,k); });
  var hset=hits[sh.name]||{}, fset=fresh[sh.name]||{};
  var h='<thead><tr><th class="rn"></th>'; for(c=0;c<C;c++) h+="<th>"+ZF.colLetter(c)+"</th>"; h+="<th></th></tr></thead><tbody>";
  sh.top.forEach(function(r,ri){ h+='<tr class="top"><th class="rn">'+(ri+1)+"</th>";
    var filledIdx=[]; (r||[]).forEach(function(cell,ci){ if(!ZF.blank(cell)) filledIdx.push(ci); });
    if(filledIdx.length<=1){ var ci=filledIdx.length?filledIdx[0]:0, v0=calc.value(sh,ri+1,ci);
      h+=(ci>0?'<td colspan="'+ci+'"></td>':"")+'<td colspan="'+(C-ci+1)+'" class="'+(ri===0?"big":"")+'" style="max-width:none">'+esc(fmtVal(sh,-1,v0,(r||[])[ci]))+"</td></tr>"; return; }
    for(var c=0;c<C;c++){ var v=calc.value(sh,ri+1,c); h+="<td>"+esc(fmtVal(sh,-1,v,(r||[])[c]))+"</td>"; } h+="<td></td></tr>"; });
  h+='<tr class="hdr"><th class="rn">'+(sh.hr+1)+"</th>"; for(c=0;c<C;c++){ var k=c-sh.c0; h+="<td>"+(k>=0?esc(sh.headers[k]):"")+"</td>"; } h+="<td></td></tr>";
  var cellsHtml=function(row,r,editable){
    var out=""; for(var c=0;c<C;c++){ var k=c-sh.c0; if(k<0){ out+="<td></td>"; continue; }
      var cell=row[k], v=calc.value(sh,r,c), fl=flagMap[r+":"+k];
      var cls=[(typeof v==="number"||numc[k])?"num":"", cell&&cell.f?"fxc":"", fl?"flag":""].join(" ").trim();
      out+='<td'+(editable?' contenteditable="true" spellcheck="false"':"")+' data-r="'+r+'" data-k="'+k+'" class="'+cls+'"'+(fl?' title="'+esc(fl)+'"':"")+'>'+esc(fmtVal(sh,k,v,cell))+"</td>"; }
    return out; };
  var limit=sh._showAll?sh.rows.length:Math.min(sh.rows.length,400);
  for(var i=0;i<limit;i++){ var r=ZF.dataRowNo(sh,i);
    h+='<tr class="'+(fset[r]?"fresh":hset[r]?"hit":"")+'" data-row="'+r+'"><th class="rn">'+r+"</th>"+cellsHtml(sh.rows[i],r,true)+'<td class="del"><button type="button" data-del="'+r+'" aria-label="Zeile '+r+' löschen" title="Zeile löschen">×</button></td></tr>'; }
  if(limit<sh.rows.length) h+='<tr class="add"><th class="rn"></th><td colspan="'+(C+1)+'"><button type="button" id="showAll">Alle '+sh.rows.length+' Zeilen anzeigen</button></td></tr>';
  h+='<tr class="add"><th class="rn"></th><td colspan="'+(C+1)+'"><button type="button" id="addRow">+ Zeile hinzufügen</button></td></tr>';
  sh.foot.forEach(function(row,j){ var r=ZF.footRowNo(sh,j); h+='<tr class="foot"><th class="rn">'+r+"</th>"+cellsHtml(row,r,true)+"<td></td></tr>"; });
  $("grid").innerHTML=h+"</tbody>";
  var nIss=issues.length;
  var s="<span>Zeilen <b>"+sh.rows.length+"</b></span><span>Spalten <b>"+sh.ncol+"</b></span>";
  var fcount=comp.filter(Boolean).length; if(fcount) s+="<span>Berechnete Spalten <b>"+fcount+"</b></span>";
  s+= nIss? '<span class="warnc">'+nIss+" Hinweis"+(nIss===1?"":"e")+" zum Prüfen</span>" : "<span>Keine Auffälligkeiten</span>";
  $("sumbar").innerHTML=s;
  var sa=$("showAll"); if(sa) sa.onclick=function(){ sh._showAll=true; renderGrid(); };
  $("addRow").onclick=function(){ var r=ZF.appendRow(sh,[]); mark(fresh,sh.name,[r]); commit("Leere Zeile "+r+" hinzugefügt"); var td=document.querySelector('td[data-r="'+r+'"]:not(.fxc)'); if(td) td.focus(); };
}
var grid=$("grid");
grid.addEventListener("focusin",function(e){
  var td=e.target.closest("td[contenteditable]"); if(!td) return;
  var sh=sheetObj(), r=+td.dataset.r, k=+td.dataset.k, cell=ZF.cellAt(sh,r,sh.c0+k);
  td.dataset.shown=td.textContent; td.textContent=rawOf(cell); td.dataset.raw=td.textContent;
  $("fxAddr").textContent=ZF.colLetter(sh.c0+k)+r; $("fxVal").textContent=rawOf(cell);
  var sel=window.getSelection(), rg=document.createRange(); rg.selectNodeContents(td); sel.removeAllRanges(); sel.addRange(rg);
});
grid.addEventListener("focusout",function(e){
  var td=e.target.closest("td[contenteditable]"); if(!td) return;
  var sh=sheetObj(), r=+td.dataset.r, k=+td.dataset.k, txt=td.textContent.trim();
  if(td.dataset.cancel){ delete td.dataset.cancel; td.textContent=td.dataset.shown; return; }
  if(txt===td.dataset.raw){ td.textContent=td.dataset.shown; return; }
  ZF.setCell(sh,r,k,txt);
  var nx=e.relatedTarget&&e.relatedTarget.closest&&e.relatedTarget.closest("td[contenteditable]");
  var nr=nx?nx.dataset.r:null, nk=nx?nx.dataset.k:null;
  commit(ZF.colLetter(sh.c0+k)+r+" geändert: „"+(txt.length>30?txt.slice(0,30)+"…":txt)+"“");
  if(nr!=null){ var again=document.querySelector('td[data-r="'+nr+'"][data-k="'+nk+'"]'); if(again) again.focus(); }
});
grid.addEventListener("keydown",function(e){
  var td=e.target.closest("td[contenteditable]"); if(!td) return;
  if(e.key==="Enter"){ e.preventDefault(); var r=+td.dataset.r, k=td.dataset.k; td.blur();
    var next=document.querySelector('td[data-r="'+(r+1)+'"][data-k="'+k+'"]'); if(next) next.focus(); }
  if(e.key==="Escape"){ td.dataset.cancel="1"; td.blur(); }
});
grid.addEventListener("click",function(e){ var b=e.target.closest("[data-del]"); if(!b) return; var r=+b.dataset.del;
  ZF.deleteRow(sheetObj(),r); fresh={}; hits={}; commit("Zeile "+r+" gelöscht"); toast("Zeile "+r+" gelöscht. Im Verlauf zurückholen."); });

function mark(store,sheetName,rows){ store[sheetName]={}; rows.forEach(function(r){ store[sheetName][r]=1; }); }
function goTo(sheetName,row,k){
  var i=model.sheets.findIndex(function(s){ return s.name===sheetName; }); if(i<0) return;
  if(cur!==i){ cur=i; renderAll(); }
  var tr=document.querySelector('tr[data-row="'+row+'"]'); if(!tr) return;
  tr.scrollIntoView({block:"center",behavior:"smooth"});
  var td=k!=null?tr.querySelector('td[data-k="'+k+'"]'):null; if(td) td.focus();
}

/* ---------- Seitenleiste ---------- */
function selectTab(id){ document.querySelectorAll(".tabs button").forEach(function(x){ x.setAttribute("aria-selected",x.dataset.pane===id); $(x.dataset.pane).hidden=x.dataset.pane!==id; }); if(id==="p-hist") renderHist(); if(id==="p-check") renderIssues(); }
document.querySelectorAll(".tabs button").forEach(function(b){ b.addEventListener("click",function(){ selectTab(b.dataset.pane); }); });

function renderExamples(){
  var heads=model.sheets.map(function(s){ return s.headers.join(" "); }).join(" ").toLowerCase(), ex;
  if(/mitarbeiter/.test(heads)&&/stunden/.test(heads)) ex=["Mehmet heute 7 Stunden Bad Schneider, Wand gefliest","Wie viele Stunden hatte Kevin im Oktober?","Bei der Fahrt am 15.10. gehört das zu Küche Yilmaz","Was kostet uns Bad Schneider bisher an Lohn?"];
  else if(/bestand/.test(heads)) ex=["Heute 3 Sack Fliesenkleber C2TE verbraucht","Welche Artikel liegen unter dem Mindestbestand?","Neuer Artikel A-1012: Fugenmörtel weiß 5 kg, 6 Eimer, 11,40 €, Baustoff Kaiser","Wie hoch ist der Lagerwert bei Weber?"];
  else if(/status/.test(heads)&&/kunde/.test(heads)) ex=["Becker hat heute bezahlt","Welche fertigen Aufträge sind noch nicht bezahlt und wie viel ist offen?","Neuer Auftrag 2026-036 für Lange, Martina: Bad 4.200 € netto, Status Angebot","Setz Neumann auf Beauftragt"];
  else ex=["Wie viele Zeilen hat die Tabelle?","Füge eine neue Zeile hinzu mit …","Welche Zeilen sind unvollständig?"];
  $("cmdEx").innerHTML=ex.map(function(t){ return '<button type="button">'+esc(t)+"</button>"; }).join("");
}
$("cmdEx").addEventListener("click",function(e){ var b=e.target.closest("button"); if(b){ $("cmdText").value=b.textContent; $("cmdText").focus(); } });

/* ---------- KI ---------- */
if(window.claude && window.claude.use) window.claude.use("downloads").then(function(d){ downloads=d; },function(){});
function setAi(st){ $("aiChip").className="chip click "+(st.on?"on":"off"); $("aiText").textContent=st.label; }
AI.onChange(setAi);
AI.init();
var NO_AI='Dafür braucht Zettelfrei eine KI. <button type="button" class="linkbtn" data-open-settings>KI einrichten</button>';
document.addEventListener("click",function(e){ if(e.target.closest("[data-open-settings]")) openSettings(); });

/* ---------- Einstellungen ---------- */
function openSettings(){
  var c=AI.status().cfg, d=$("settings");
  draft={keys:Object.assign({},c.keys), models:Object.assign({},c.models)};
  $("setProvider").value=c.provider||"auto"; $("setUrl").value=c.url||"http://localhost:11434/v1";
  $("setStatus").hidden=true; lastProv=null; updateSettingsForm();
  if(d.showModal) d.showModal(); else d.setAttribute("open","");
}
var draft={keys:{},models:{}}, lastProv=null;
function keyProv(p){ var st=AI.status(); if(p==="auto") return st.server?null:"openrouter"; return (p==="claude"||p==="openrouter"||p==="local")?p:null; }
function stash(){ if(lastProv){ draft.keys[lastProv]=$("setKey").value.trim(); draft.models[lastProv]=$("setModel").value.trim(); } }
function updateSettingsForm(){
  stash();
  var p=$("setProvider").value, st=AI.status(), kp=keyProv(p);
  $("rowKey").hidden=!kp; $("rowModel").hidden=!kp; $("rowUrl").hidden=p!=="local";
  if(kp){ $("setKey").value=draft.keys[kp]||""; $("setModel").value=draft.models[kp]||""; $("setModel").placeholder=st.defaults[kp];
    $("setKey").placeholder=kp==="openrouter"?"sk-or-…":kp==="claude"?"sk-ant-…":"optional"; }
  $("keyLabel").textContent=kp==="local"?"API-Schlüssel (meist leer lassen)":kp==="openrouter"?"OpenRouter-Schlüssel":"Claude-API-Schlüssel";
  lastProv=kp;
  var help={
    auto:"Nutzt den eigenen Server, wenn Zettelfrei mit „npm start“ läuft"+(st.server?" (gefunden)":" (nicht gefunden)")+", sonst OpenRouter mit dem Schlüssel unten.",
    server:"Schlüssel liegt nur auf dem Server: <code>OPENROUTER_API_KEY=… npm start</code> oder <code>ANTHROPIC_API_KEY=… npm start</code>. Am besten für Betriebe mit mehreren Rechnern."+(st.server?" Server gefunden.":" Kein Server gefunden."),
    openrouter:"Ein Schlüssel für viele Modelle: Claude, GPT, Gemini, Mistral. Schlüssel auf <code>openrouter.ai/keys</code> anlegen und mit Guthaben aufladen. Modellnamen stehen auf <code>openrouter.ai/models</code>. Der Schlüssel bleibt in diesem Browser.",
    claude:"Schlüssel aus <code>console.anthropic.com</code>. Er wird nur in diesem Browser gespeichert und direkt an Anthropic geschickt. Nur für den eigenen Rechner.",
    local:"Daten verlassen den Rechner nicht. Für Ollama: <code>OLLAMA_ORIGINS=* ollama serve</code>, Adresse <code>http://localhost:11434/v1</code>. Für Fotos und Scans ein Modell mit Bildverständnis wählen.",
    off:"Tabelle öffnen, bearbeiten, prüfen und speichern funktioniert ohne KI."
  };
  $("setHelp").innerHTML=help[p]||"";
}
$("aiChip").addEventListener("click",openSettings);
$("setProvider").addEventListener("change",updateSettingsForm);
function readSettings(){ stash(); return {provider:$("setProvider").value,url:$("setUrl").value.trim(),keys:draft.keys,models:draft.models}; }
$("settingsForm").addEventListener("submit",function(e){ if(e.submitter&&e.submitter.value==="save"){ AI.save(readSettings()); toast("KI-Einstellungen gespeichert"); } });
$("setTest").addEventListener("click",function(){
  var before=AI.status().cfg; AI.save(readSettings());
  var box=$("setStatus"); box.hidden=false; box.className="status"; box.textContent="Teste die Verbindung…";
  AI.test().then(function(ok){ box.className="status "+(ok?"ok":"bad"); box.textContent=ok?"Verbindung steht. Die KI antwortet.":"Die KI antwortet, aber nicht im erwarteten Format. Anderes Modell versuchen."; },
    function(e){ box.className="status bad"; box.textContent=aiError(e)||"Abgebrochen."; });
});
function aiError(e){
  var c=e&&e.code;
  if(c==="cancelled") return "";
  if(c==="no_provider") return "Keine KI eingerichtet. Oben rechts auf den KI-Status klicken.";
  if(c==="auth") return "Der API-Schlüssel wurde abgelehnt. Bitte in den KI-Einstellungen prüfen.";
  if(c==="credits") return "Beim KI-Anbieter ist kein Guthaben mehr. Bitte aufladen (bei OpenRouter unter openrouter.ai/credits).";
  if(c==="model") return "Dieses Modell kennt der Anbieter nicht. Modellnamen in den KI-Einstellungen prüfen.";
  if(c==="network") return "Keine Verbindung zur KI. Internet, Server oder lokales Modell prüfen.";
  if(c==="http") return "Die KI hat einen Fehler gemeldet ("+esc(e.message)+"). Modellname in den Einstellungen prüfen.";
  if(c==="not_granted"||c==="sampling_disabled"||c==="not_declared"||c==="capability_disabled"){ return "Die KI wurde für diese Seite nicht freigegeben. Tabelle bearbeiten, öffnen und speichern funktioniert weiter."; }
  if(c==="rate_limited") return "Gerade zu viele Anfragen. Bitte in einer Minute noch einmal.";
  if(c==="image_rejected") return "Diese Datei lässt sich nicht lesen. Bitte ein anderes Foto oder PDF wählen.";
  if(c==="images_unavailable") return "Bilder kann die KI hier nicht lesen. Ein PDF mit Text funktioniert.";
  if(c==="prompt_too_large") return "Die Datei ist für eine Anfrage zu groß. Bitte ein kleineres Blatt öffnen.";
  if(c==="invalid_json") return "Die Antwort war unvollständig. Bitte noch einmal versuchen oder einfacher formulieren.";
  return "Die KI hat nicht geantwortet. Bitte noch einmal versuchen.";
}
var RULES = 'Antworte NUR mit JSON in genau dieser Form:\n{"aenderungen":[ {"typ":"neu","blatt":"…","werte":{"Spaltenname":Wert}} | {"typ":"aendern","blatt":"…","zeile":Excel-Zeilennummer,"werte":{"Spaltenname":Wert}} | {"typ":"loeschen","blatt":"…","zeile":Excel-Zeilennummer} ],\n "abfragen":[{"id":"q1","blatt":"…","spalte":"…","art":"summe|anzahl|mittelwert|min|max|liste","filter":[{"spalte":"…","op":"=|!=|enthält|>|<|>=|<=|leer|nicht leer","wert":"…"}]}],\n "antwort":"kurzer deutscher Text, Ergebnisse von Abfragen als {q1} einsetzen",\n "rueckfrage":null}\nRegeln:\n- Rechne Summen, Anzahlen und Mittelwerte NIE selbst. Lege dafür immer eine Abfrage an und setze {id} in die Antwort ein. Für Listen von Namen/Artikeln nutze art "liste".\n- Zu- und Abgänge (z. B. Bestand) als {"plus": Zahl} angeben, z. B. {"Bestand":{"plus":-3}}.\n- Spalten mit „berechnet“ niemals setzen, die Tabelle rechnet sie.\n- Spaltennamen und Blattnamen exakt wie angegeben. Datum als TT.MM.JJJJ, Zahlen als Zahl.\n- Zeilennummern sind Excel-Zeilennummern aus der Spalte „Zeile“.\n- Ändere nur, was verlangt ist. Erfinde keine Werte. Ist unklar, welche Zeile gemeint ist, oder fehlen Pflichtangaben, stelle eine kurze rueckfrage und ändere nichts.\n- Leere Listen weglassen ist erlaubt.';
function context(){ return "Heute ist "+TODAY+" ("+WEEKDAY+"). Du arbeitest in der Excel-Datei eines kleinen deutschen Handwerksbetriebs.\n\n"+ZF.describe(model,180000); }
function fmtQuery(res,q){
  if(res.error) return "(?)";
  if(res.list){ var l=res.list.map(function(v){ return typeof v==="number"?de(v):String(v); }); if(!l.length) return "keine"; return l.length>15? l.slice(0,15).join(", ")+" und "+(l.length-15)+" weitere" : l.join(", "); }
  if(res.value==null) return "–";
  return de(res.value,res.money?2:null);
}
function handleResult(res,box,label,stillRunningLabel){
  box.innerHTML="";
  if(!res||typeof res!=="object"){ box.innerHTML='<p class="hint">Keine verwertbare Antwort.</p>'; return; }
  var sheetHits={}, txt=res.antwort||"";
  (res.abfragen||[]).forEach(function(q){
    var r=ZF.query(model,q); txt=txt.split("{"+q.id+"}").join(fmtQuery(r,q));
    if(r.rows&&r.rows.length){ var sh=(model.sheets.find(function(s){ return s.name.toLowerCase()===String(q.blatt).toLowerCase(); })||sheetObj()).name; sheetHits[sh]=(sheetHits[sh]||[]).concat(r.rows); }
  });
  if(txt){ var a=document.createElement("div"); a.className="answer"; a.textContent=txt;
    var names=Object.keys(sheetHits);
    if(names.length){ var sp=document.createElement("span"); sp.className="rows";
      var all=[]; names.forEach(function(n){ all=all.concat(sheetHits[n]); });
      sp.innerHTML='Berechnet aus '+all.length+' Zeile'+(all.length===1?"":"n")+' · <button type="button">in der Tabelle zeigen</button>';
      sp.querySelector("button").onclick=function(){ hits={}; names.forEach(function(n){ mark(hits,n,sheetHits[n]); }); fresh={};
        var si=model.sheets.findIndex(function(s){ return s.name===names[0]; }); if(si>=0) cur=si; renderAll(); goTo(names[0],sheetHits[names[0]][0]); };
      a.appendChild(sp); }
    box.appendChild(a); }
  if(res.rueckfrage){ var q=document.createElement("div"); q.className="ask"; q.textContent=res.rueckfrage; box.appendChild(q); }
  var plan=ZF.planOps(model,res.aenderungen||[]);
  if(plan.length) box.appendChild(planView(plan,label));
  if(!txt&&!res.rueckfrage&&!plan.length) box.innerHTML='<p class="hint">Dazu gab es nichts zu tun. Versuchen Sie es mit einer genaueren Anweisung.</p>';
}
function fv(v){ if(v==null||v==="") return "leer"; return typeof v==="number"?de(v):String(v); }
function planView(plan,label){
  var div=document.createElement("div"); div.className="plan";
  var ok=plan.filter(function(p){ return !p.error; });
  div.innerHTML='<div class="ph"><span>'+ok.length+" Änderung"+(ok.length===1?"":"en")+' vorgeschlagen</span><span class="row"><button class="btn small" data-ok>Übernehmen</button><button class="btn quiet small" data-no>Verwerfen</button></span></div><ul></ul>';
  var ul=div.querySelector("ul");
  plan.forEach(function(p,i){
    var li=document.createElement("li");
    if(p.error){ li.innerHTML='<span></span><span class="err">'+esc(p.error)+"</span>"; ul.appendChild(li); return; }
    var head = p.typ==="neu"?"Neue Zeile in „"+p.sheet+"“" : p.typ==="loeschen"?"Zeile "+p.row+" löschen" : "Zeile "+p.row+" ändern";
    var body = p.typ==="loeschen" ? '<div class="hint">'+esc(p.summary)+"</div>" :
      '<dl class="diff">'+p.changes.map(function(c){ return c.error? "<dt>"+esc(c.name)+'</dt><dd class="err">'+esc(c.error)+"</dd>" :
        "<dt>"+esc(c.name)+"</dt><dd>"+(p.typ==="neu"?"<ins>"+esc(fv(c.nv))+"</ins>":"<del>"+esc(fv(c.old))+"</del> → <ins>"+esc(fv(c.nv))+"</ins>")+"</dd>"; }).join("")+"</dl>";
    li.innerHTML='<input type="checkbox" checked id="pl'+i+'_'+Date.now()+'" data-i="'+i+'" aria-label="'+esc(head)+'"><div><div class="what">'+esc(head)+"</div>"+body+"</div>";
    ul.appendChild(li);
  });
  div.querySelector("[data-no]").onclick=function(){ div.remove(); };
  div.querySelector("[data-ok]").onclick=function(){
    var chosen=[].map.call(div.querySelectorAll("input[type=checkbox]:checked"),function(x){ return plan[+x.dataset.i]; });
    if(!chosen.length){ div.remove(); return; }
    var touched=ZF.applyPlan(model,chosen); fresh={}; hits={};
    touched.forEach(function(t){ fresh[t.sheet]=fresh[t.sheet]||{}; fresh[t.sheet][t.row]=1; });
    var first=touched[0]; if(first){ var si=model.sheets.findIndex(function(s){ return s.name===first.sheet; }); if(si>=0) cur=si; }
    commit(label); div.remove(); toast(chosen.length+" Änderung"+(chosen.length===1?"":"en")+" übernommen");
    if(first) goTo(first.sheet,first.row);
    setTimeout(function(){ fresh={}; renderGrid(); },7000);
  };
  return div;
}
function busy(el,text){ el.innerHTML=text?'<span class="busy">'+esc(text)+"</span>":""; }

/* Assistent */
var cmdCtl;
$("cmdStop").addEventListener("click",function(){ if(cmdCtl) cmdCtl.abort(); });
$("cmdText").addEventListener("keydown",function(e){ if(e.key==="Enter"&&(e.metaKey||e.ctrlKey)) runCmd(); });
$("cmdBtn").addEventListener("click",runCmd);
function runCmd(){
  var text=$("cmdText").value.trim(); if(!text||!model) return;
  var out=$("cmdOut");
  if(!AI.active()){ out.innerHTML='<p class="hint">'+NO_AI+'</p>'; return; }
  cmdCtl=new AbortController(); $("cmdBtn").disabled=true; $("cmdStop").hidden=false; busy($("cmdBusy"),"Liest die Tabelle…"); out.innerHTML="";
  var prompt=context()+"\n\nAnweisung oder Frage des Nutzers: \"\"\""+text+"\"\"\"\n\n"+RULES;
  AI.json(prompt,{signal:cmdCtl.signal}).then(function(res){
    handleResult(res,out,"Assistent: „"+(text.length>50?text.slice(0,50)+"…":text)+"“");
  },function(e){ var m=aiError(e); if(m) out.innerHTML='<p class="hint" style="color:var(--warn)">'+esc(m)+"</p>"; })
  .then(function(){ $("cmdBtn").disabled=false; $("cmdStop").hidden=true; busy($("cmdBusy"),""); });
}

/* PDF & Foto */
var drop=$("drop");
["dragenter","dragover"].forEach(function(ev){ drop.addEventListener(ev,function(e){ e.preventDefault(); e.stopPropagation(); drop.classList.add("over"); }); });
["dragleave","drop"].forEach(function(ev){ drop.addEventListener(ev,function(e){ e.preventDefault(); e.stopPropagation(); drop.classList.remove("over"); }); });
drop.addEventListener("drop",function(e){ var f=e.dataTransfer.files[0]; if(f) handleFile(f); });
$("pdfInput").addEventListener("change",function(e){ var f=e.target.files[0]; e.target.value=""; if(f) handleFile(f); });
$("pdfSample").addEventListener("click",function(){ extract(SAMPLE_INVOICE,[],"Beispiel-Rechnung Weber KG"); });
if(window.pdfjsLib) pdfjsLib.GlobalWorkerOptions.workerSrc="vendor/pdf.worker.min.js";
function handleFile(f){
  var st=$("pdfBusy"); $("pdfOut").innerHTML="";
  if(f.type.indexOf("image/")===0){ return extract("",[f],f.name); }
  if(!/pdf$/i.test(f.type)&&!/\.pdf$/i.test(f.name)){ toast("Bitte ein PDF oder Foto wählen."); return; }
  if(!window.pdfjsLib){ $("pdfOut").innerHTML='<p class="hint">Der PDF-Leser konnte nicht geladen werden.</p>'; return; }
  busy(st,"PDF wird geöffnet…");
  f.arrayBuffer().then(function(buf){ return pdfjsLib.getDocument({data:new Uint8Array(buf)}).promise; }).then(async function(pdf){
    var text="";
    for(var p=1;p<=Math.min(pdf.numPages,6);p++){ var pg=await pdf.getPage(p), tc=await pg.getTextContent(), lastY=null;
      tc.items.forEach(function(it){ var y=it.transform[5]; text+=(lastY!==null&&Math.abs(y-lastY)>3)?"\n":(lastY!==null?" ":""); text+=it.str; lastY=y; }); text+="\n\n"; }
    if(text.replace(/\s/g,"").length>40) return extract(text.slice(0,30000),[],f.name);
    busy(st,"Gescanntes PDF, Seiten werden als Bild gelesen…");
    var imgs=[];
    for(p=1;p<=Math.min(pdf.numPages,3);p++){ var pg2=await pdf.getPage(p), vp=pg2.getViewport({scale:1.6}), c=document.createElement("canvas");
      c.width=vp.width; c.height=vp.height; await pg2.render({canvasContext:c.getContext("2d"),viewport:vp}).promise;
      imgs.push(await new Promise(function(r){ c.toBlob(r,"image/jpeg",0.85); })); }
    return extract("",imgs,f.name);
  }).catch(function(){ busy(st,""); $("pdfOut").innerHTML='<p class="hint" style="color:var(--warn)">Diese Datei lässt sich nicht öffnen. Ist es ein gültiges PDF?</p>'; });
}
function extract(text,images,name){
  var out=$("pdfOut"), st=$("pdfBusy");
  if(!AI.active()){ busy(st,""); out.innerHTML='<p class="hint">'+NO_AI+'</p>'; return; }
  var target=model.sheets[+$("pdfTarget").value]||sheetObj();
  busy(st,"KI liest das Dokument…");
  var prompt=context()+"\n\nDer Nutzer hat ein Dokument hochgeladen ("+name+")"+(images.length?", siehe Bild":"")+". Übertrage den Inhalt in das Blatt „"+target.name+"“.\n"+
    "Überlege, was das Dokument für dieses Blatt bedeutet: z. B. bei einer Lieferantenrechnung im Lager den Bestand passender vorhandener Artikel per {\"plus\": Menge} erhöhen und unbekannte Artikel neu anlegen; bei einer Kundenbestellung einen neuen Auftrag anlegen; bei einem Stundenzettel neue Zeilen pro Person und Tag. Ordne Kommission oder Baustelle einem vorhandenen Projekt zu, wenn eindeutig.\n"+
    "In \"antwort\" in einem Satz zusammenfassen, was im Dokument steht und was du überträgst. Unleserliches nicht raten, sondern in rueckfrage nennen.\n"+
    (text?"Text des Dokuments:\n\"\"\""+text+"\"\"\"\n\n":"")+RULES;
  AI.json(prompt,{images:images}).then(function(res){ handleResult(res,out,"Dokument übernommen: "+name); },
    function(e){ var m=aiError(e); if(m) out.innerHTML='<p class="hint" style="color:var(--warn)">'+esc(m)+"</p>"; })
  .then(function(){ busy(st,""); });
}

/* Prüfen */
function renderIssues(){
  var all=issuesFor(); $("issueN").hidden=!all.length; $("issueN").textContent=all.length;
  if($("p-check").hidden) return;
  $("issues").innerHTML=all.length? all.map(function(x,i){ var sh=model.sheets.find(function(s){ return s.name===x.blatt; });
    var k=x.k!=null?x.k:(sh?ZF.findCol(sh,x.spalte):-1); var addr=sh&&k>=0?ZF.colLetter(sh.c0+k)+x.zeile:"Z. "+x.zeile;
    return '<li><button type="button" data-i="'+i+'"><span class="where">'+esc(x.blatt)+" · "+addr+"</span><span>"+esc(x.problem)+(x.ki?'<span class="ki">KI</span>':"")+"</span></button></li>"; }).join("")
    : '<li class="hint" style="padding:8px 0">Keine Auffälligkeiten gefunden.</li>';
}
$("issues").addEventListener("click",function(e){ var b=e.target.closest("button[data-i]"); if(!b) return; var x=issuesFor()[+b.dataset.i];
  var sh=model.sheets.find(function(s){ return s.name===x.blatt; }); goTo(x.blatt,x.zeile,x.k!=null?x.k:(sh?ZF.findCol(sh,x.spalte):null)); });
$("aiCheckBtn").addEventListener("click",function(){
  if(!AI.active()){ openSettings(); return; }
  $("aiCheckBtn").disabled=true; busy($("checkBusy"),"KI prüft die Datei…");
  var prompt=context()+"\n\nPrüfe die Datei fachlich wie eine erfahrene Bürokraft im Handwerk. Finde höchstens 12 echte Auffälligkeiten, die Geld, Termine oder Ordnung betreffen: z. B. Bestand unter Mindestbestand, fertige Aufträge ohne Rechnung, Rechnungen lange unbezahlt, unplausible Stunden, Tippfehler in Namen, fehlende Zuordnungen. Keine Kleinigkeiten zur Formatierung.\nAntworte NUR mit JSON: {\"probleme\":[{\"blatt\":\"…\",\"zeile\":Excel-Zeilennummer,\"spalte\":\"Spaltenname\",\"problem\":\"kurz, deutsch\"}]}";
  AI.json(prompt).then(function(res){
    aiIssues=((res&&res.probleme)||[]).filter(function(p){ return p&&p.blatt&&p.zeile; }).map(function(p){ return {blatt:p.blatt,zeile:+p.zeile,spalte:p.spalte,problem:String(p.problem||""),ki:1}; });
    renderAll(); toast(aiIssues.length?aiIssues.length+" fachliche Hinweise gefunden":"Fachlich nichts Auffälliges gefunden");
  },function(e){ var m=aiError(e); if(m) toast(m); }).then(function(){ $("aiCheckBtn").disabled=false; busy($("checkBusy"),""); });
});

/* Verlauf */
function renderHist(){
  var items=history.map(function(h,i){ return {h:h,i:i}; }).reverse();
  $("hist").innerHTML=items.map(function(x,j){ var d=new Date(x.h.t);
    return "<li><time>"+String(d.getHours()).padStart(2,"0")+":"+String(d.getMinutes()).padStart(2,"0")+"</time><span>"+esc(x.h.label)+"</span>"+(j===0?'<span class="hint">aktuell</span>':'<button class="btn quiet small" data-restore="'+x.i+'">Zurückholen</button>')+"</li>"; }).join("");
}
$("hist").addEventListener("click",function(e){ var b=e.target.closest("[data-restore]"); if(!b) return; var h=history[+b.dataset.restore];
  model=JSON.parse(h.snap); cur=Math.min(cur,model.sheets.length-1); fresh={}; hits={}; commit("Zurückgeholt: Stand „"+h.label+"“"); renderHist(); toast("Früherer Stand zurückgeholt"); });

function renderAll(){ if(!model) return; renderTabs(); renderGrid(); renderIssues(); if(!$("p-hist").hidden) renderHist(); }

/* Start */
if(restore()){ renderAll(); renderExamples(); }
else { var first=Object.keys(EXAMPLES)[0]; openData(b64(EXAMPLES[first]),first); }
})();
