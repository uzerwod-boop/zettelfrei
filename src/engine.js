/* Zettelfrei Excel-Engine: liest beliebige Tabellen, rechnet Formeln, ändert Zeilen, schreibt xlsx. */
var ZF = (function(){
"use strict";

/* ---------- Hilfen ---------- */
function colLetter(i){ var s=""; i=i+1; while(i>0){ var m=(i-1)%26; s=String.fromCharCode(65+m)+s; i=Math.floor((i-1)/26); } return s; }
function colIndex(s){ var n=0; for(var i=0;i<s.length;i++) n=n*26+(s.charCodeAt(i)-64); return n-1; }
function parseNum(s){
  if(typeof s==="number") return s; if(s==null) return null;
  s=String(s).trim().replace(/\s|€|EUR/gi,""); if(!s) return null;
  if(/^-?\d{1,3}(\.\d{3})+(,\d+)?$/.test(s)) s=s.replace(/\./g,"").replace(",",".");
  else if(/,/.test(s)) s=s.replace(/\./g,"").replace(",",".");
  if(!/^-?\d*\.?\d+$/.test(s)) return null;
  var n=parseFloat(s); return isNaN(n)?null:n;
}
function isDateStr(s){ return typeof s==="string" && /^\d{1,2}\.\d{1,2}\.\d{4}$/.test(s.trim()); }
function dateToSerial(s){ var m=String(s).trim().match(/^(\d{1,2})\.(\d{1,2})\.(\d{4})$/); if(!m) return null;
  return (Date.UTC(+m[3],+m[2]-1,+m[1])-Date.UTC(1899,11,30))/86400000; }
function serialToDate(n){ var d=new Date(Date.UTC(1899,11,30)+Math.round(n)*86400000);
  return String(d.getUTCDate()).padStart(2,"0")+"."+String(d.getUTCMonth()+1).padStart(2,"0")+"."+d.getUTCFullYear(); }
function normDate(s){ var m=String(s).trim().match(/^(\d{1,2})\.(\d{1,2})\.(\d{2,4})$/); if(!m) return null;
  var y=+m[3]; if(y<100) y+=2000; return String(+m[1]).padStart(2,"0")+"."+String(+m[2]).padStart(2,"0")+"."+y; }
function isTimeFmt(z){ z=String(z||"").toLowerCase().replace(/"[^"]*"|\\./g,""); return /h/.test(z) && !/[dy]/.test(z); }
function fmtTime(v){ if(typeof v!=="number"||!isFinite(v)) return v; var m=Math.round((v%1+1)%1*1440); if(m===1440) m=0; return String(Math.floor(m/60)).padStart(2,"0")+":"+String(m%60).padStart(2,"0"); }
function parseTime(s){ var m=String(s).trim().match(/^(\d{1,2})[:.h](\d{2})(?:\s*uhr)?$/i)||String(s).trim().match(/^(\d{1,2})\s*uhr$/i); if(!m) return null; var h=+m[1], mi=m[2]?+m[2]:0; if(h>24||mi>59) return null; return (h*60+mi)/1440; }
function todaySerial(d){ d=d||new Date(); return (Date.UTC(d.getFullYear(),d.getMonth(),d.getDate())-Date.UTC(1899,11,30))/86400000; }
function clone(o){ return JSON.parse(JSON.stringify(o)); }
function blank(c){ return !c || ((c.v==null || c.v==="") && !c.f); }

/* ---------- Formel-Referenzen ---------- */
/* Gleiche-Blatt-Referenzen: kein "!" oder Buchstabe/Ziffer davor, kein "(" / Ziffer danach. */
var REF_RE=/(^|[^A-Za-z0-9_!.'"$])(\$?)([A-Z]{1,3})(\$?)(\d+)(?::(\$?)([A-Z]{1,3})(\$?)(\d+))?(?![\d(A-Za-z_])/g;
function transformRefs(f, fn){
  // fn(row, colIdx, part, absRow) -> new row or null (#BEZUG!)
  return splitStrings(f).map(function(seg){
    if(seg.str) return seg.t;
    return seg.t.replace(REF_RE,function(all,pre,ca,c,ra,r,ca2,c2,ra2,r2){
      if(c2){
        var a=fn(+r,colIndex(c),"start",!!ra), b=fn(+r2,colIndex(c2),"end",!!ra2);
        if(a==null||b==null||a>b) return pre+"#BEZUG!";
        return pre+ca+c+ra+a+":"+ca2+c2+ra2+b;
      }
      var n=fn(+r,colIndex(c),"single",!!ra); return n==null? pre+"#BEZUG!" : pre+ca+c+ra+n;
    });
  }).join("");
}
function splitStrings(f){ var out=[],re=/"[^"]*"/g,last=0,m; while((m=re.exec(f))){ if(m.index>last) out.push({t:f.slice(last,m.index)}); out.push({t:m[0],str:true}); last=m.index+m[0].length; } if(last<f.length) out.push({t:f.slice(last)}); return out; }

/* ---------- Laden ---------- */
function cellFrom(c, XLSX){
  if(!c) return null;
  var o={};
  if(c.f) o.f=String(c.f);
  if(c.t==="n"){
    if(c.z && isTimeFmt(c.z)) { o.v=c.v; o.tm=1; }
    else if(c.z && XLSX.SSF.is_date(c.z)) { o.v=serialToDate(c.v); o.d=1; }
    else o.v=c.v; }
  else if(c.t==="d"){ var dt=new Date(c.v); o.v=String(dt.getDate()).padStart(2,"0")+"."+String(dt.getMonth()+1).padStart(2,"0")+"."+dt.getFullYear(); o.d=1; }
  else if(c.t==="b") o.v=c.v?"WAHR":"FALSCH";
  else if(c.t==="e") o.v=c.w||"#FEHLER";
  else if(c.v!=null) o.v=String(c.v);
  else o.v=null;
  if(c.z && c.z!=="General") o.z=c.z;
  return o;
}
function loadWorkbook(XLSX, data, name){
  var wb=XLSX.read(data,{type:"array",cellNF:true,cellFormula:true,cellDates:false});
  var sheets=wb.SheetNames.map(function(sn){ return parseSheet(XLSX, wb.Sheets[sn], sn); });
  return {name:name||"Tabelle.xlsx", sheets:sheets};
}
// Datum in typischen Schreibweisen erkennen (auch als Text): 01.10.2026, 2.10.26, 03/11/2026, 2026-10-01
var MON={jan:1,feb:2,"mär":3,mrz:3,mar:3,apr:4,mai:5,may:5,jun:6,jul:7,aug:8,sep:9,okt:10,oct:10,nov:11,dez:12,dec:12};
function anyDateSerial(v){
  if(v==null||v==="") return null;
  if(typeof v==="number") return v>20000&&v<80000?Math.floor(v):null;
  var t=String(v).trim(), m;
  if((m=t.match(/^(\d{1,2})\.(\d{1,2})\.(\d{2}|\d{4})$/))||(m=t.match(/^(\d{1,2})\/(\d{1,2})\/(\d{4})$/))){ var y=+m[3]; if(y<100) y+=2000; if(+m[2]<1||+m[2]>12||+m[1]<1||+m[1]>31) return null; return (Date.UTC(y,+m[2]-1,+m[1])-Date.UTC(1899,11,30))/86400000; }
  if((m=t.match(/^(\d{4})-(\d{2})-(\d{2})$/))) return (Date.UTC(+m[1],+m[2]-1,+m[3])-Date.UTC(1899,11,30))/86400000;
  return null;
}
function looksTime(v){ return typeof v==="string" && /^\s*\d{1,2}:\d{2}(:\d{2})?\s*$/.test(v); }
function timeNum(v){ if(typeof v==="number") return v; if(looksTime(v)){ var p=v.trim().split(":"); return (+p[0]*60+ +p[1])/1440; } return NaN; }
var SUM_RE=/(^|\s)(summe|gesamt|zwischensumme|total|ergebnis|übertrag)\b/i;
function rowKind(part, ncol){
  var filled=part.filter(function(c){ return !blank(c); });
  if(!filled.length) return "leer";
  var texts=filled.filter(function(c){ return typeof c.v==="string" && !c.f; });
  if(texts.some(function(c){ return SUM_RE.test(c.v.trim()); }) && filled.length<=Math.max(2,Math.ceil(ncol*0.6)) && filled.some(function(c){ return c.f||typeof c.v==="number"; })) return "summe";
  if(filled.length<=2 && texts.length>=1 && /:\s*$/.test(String(filled[0].v||"")) && typeof filled[0].v==="string") return "notiz";
  if(filled.length===1 && texts.length===1 && ncol>=3 && !isDateStr(texts[0].v) && anyDateSerial(texts[0].v)==null) return "abschnitt";
  return "";
}
function isData(sh,i){ return !(sh.kind && sh.kind[i]); }
function eachData(sh,fn){ sh.rows.forEach(function(r,i){ if(isData(sh,i)) fn(r,i); }); }
function dataCount(sh){ var n=0; eachData(sh,function(){ n++; }); return n; }

function parseSheet(XLSX, ws, name){
  var ref=ws["!ref"]; var grid=[];
  if(ref){ var rg=XLSX.utils.decode_range(ref);
    for(var r=0;r<=rg.e.r;r++){ var row=[]; for(var c=0;c<=rg.e.c;c++){ row.push(cellFrom(ws[XLSX.utils.encode_cell({r:r,c:c})],XLSX)); } grid.push(row); } }
  var widths=(ws["!cols"]||[]).map(function(c){ return c&&(c.wch||c.width)?Math.round(c.wch||c.width):null; });
  var merges=(ws["!merges"]||[]).map(function(m){ return [m.s.r,m.s.c,m.e.r,m.e.c]; });
  // Kopfzeile suchen
  var hr=-1, best=0;
  var isLabel=function(c){ return c && !c.f && typeof c.v==="string" && c.v.trim() && !isDateStr(c.v) && anyDateSerial(c.v)==null && !looksTime(c.v) && parseNum(c.v)==null; };
  var score=function(row){ return row?row.filter(isLabel).length:0; };
  var filled=function(row){ return row?row.filter(function(c){ return !blank(c); }).length:0; };
  for(var i=0;i<Math.min(grid.length,20);i++) best=Math.max(best,score(grid[i]));
  if(best>=2){ for(i=0;i<Math.min(grid.length,20);i++){ var s=score(grid[i]); if(s>=2 && s>=0.7*best && (i+1>=grid.length || filled(grid[i+1])>=1)){ hr=i; break; } } }
  if(hr<0) return {name:name, hr:-1, top:grid, headers:[], rows:[], kind:[], foot:[], c0:0, ncol:0, widths:widths, merges:merges, colFmt:[], dateCol:[], timeCol:[]};
  var up=grid[hr], low=grid[hr+1]||[];
  // Zweizeiliger Kopf: verbundene Zellen in der Kopfzeile und darunter nur Beschriftungen
  var mergedHead=merges.some(function(m){ return m[0]===hr && (m[2]===hr+1 || m[3]>m[1]); });
  // nur den Bereich unter der oberen Kopfzeile betrachten (Nebentabellen rechts ignorieren)
  var covered=function(c){ return (up[c]&&!blank(up[c])) || merges.some(function(m){ return m[0]===hr && m[1]<=c && m[3]>=c; }); };
  var u0=up.findIndex(function(c){ return c&&!blank(c); }), u1=u0; while(u1+1<Math.max(up.length,low.length) && covered(u1+1)) u1++;
  var lowPart=low.slice(Math.max(0,u0),u1+1);
  var lowLabels=lowPart.filter(isLabel).length, lowOther=lowPart.filter(function(c){ return !blank(c) && !isLabel(c); }).length;
  var two=mergedHead && lowLabels>=2 && lowOther===0;
  var headCell=function(c){ var a=two&&low[c]&&!blank(low[c])?low[c]:up[c]; if(!a||blank(a)){ // Wert der verbundenen Zelle darüber/links
      var m=merges.find(function(m){ return m[0]<=hr && m[2]>=hr && m[1]<=c && m[3]>=c; }); if(m && !(two && low[c] && !blank(low[c]))) a=grid[m[0]][m[1]]; }
    return a&&!blank(a)?String(a.v).replace(/\s+/g," ").trim():""; };
  var hrow=two?hr+1:hr;
  var width=Math.max(up.length,low.length), names=[]; for(var c=0;c<width;c++) names.push(headCell(c));
  var c0=names.findIndex(function(x){ return x!==""; }), cEnd=c0;
  // Ende der Tabelle: erste Spalte ohne Überschrift, in der auch kaum Daten stehen (Nebentabellen rechts nicht mitnehmen)
  var sample=grid.slice(hrow+1,hrow+61);
  for(c=c0+1;c<width;c++){
    if(names[c]!==""){ cEnd=c; continue; }
    var n=sample.filter(function(rw){ return rw && !blank(rw[c]); }).length, rowsN=sample.filter(function(rw){ return filled(rw)>0; }).length;
    if(n>=Math.max(1,rowsN*0.5)){ cEnd=c; continue; }
    break;
  }
  hr=hrow;
  var ncol=cEnd-c0+1;
  var headers=[]; for(var k=0;k<ncol;k++) headers.push(names[c0+k]||colLetter(c0+k));
  // Zeilen: Excel-Zeilennummern bleiben erhalten. Leer-, Abschnitts- und Zwischensummenzeilen werden markiert, nicht entfernt.
  var parts=[], kinds=[];
  for(r=hr+1;r<grid.length;r++){ var g=grid[r]||[], part=[]; for(k=0;k<ncol;k++) part.push(g[c0+k]||null); parts.push(part); kinds.push(rowKind(part,ncol)); }
  var lastData=-1; for(i=0;i<kinds.length;i++) if(kinds[i]==="") lastData=i;
  var rows=parts.slice(0,lastData+1), kind=kinds.slice(0,lastData+1), foot=parts.slice(lastData+1);
  while(foot.length && foot[foot.length-1].every(blank)) foot.pop();
  var oldRows=rows.map(function(_,i){ return hr+2+i; }), footOld=foot.map(function(_,j){ return hr+2+rows.length+j; });
  // Zellen links/rechts der Tabelle (z. B. Stundensätze) bleiben für Formeln erreichbar
  var extra={}; for(r=hr;r<grid.length;r++){ var gr=grid[r]||[]; for(c=0;c<gr.length;c++){ if((c<c0||c>cEnd) && !blank(gr[c])) extra[(r+1)+":"+c]=gr[c]; } }
  // Spaltentypen (nur echte Datenzeilen)
  var colFmt=[], dateCol=[], timeCol=[];
  for(k=0;k<ncol;k++){
    var z=null, d=0, n2=0, t=0;
    rows.forEach(function(rw,i){ if(kind[i]) return; var c=rw[k]; if(c&&!blank(c)&&!c.f){ n2++; if(c.d||(typeof c.v==="string"&&anyDateSerial(c.v)!=null)) d++; if(c.tm||looksTime(c.v)) t++; if(!z&&c.z) z=c.z; } else if(c&&c.f&&!z&&c.z) z=c.z; });
    colFmt.push(z); dateCol.push(n2>0 && d/n2>=0.6); timeCol.push(n2>0 && t/n2>=0.6);
  }
  return {name:name, hr:hr, top:grid.slice(0,hr), headers:headers, rows:rows, kind:kind, foot:foot, c0:c0, ncol:ncol, widths:widths, merges:merges.filter(function(m){ return m[2]<hr; }), colFmt:colFmt, dateCol:dateCol, timeCol:timeCol, twoRowHeader:two||undefined, extra:extra, orig:{data:oldRows, foot:footOld}};
}

/* ---------- Lesen mit ExcelJS (sicherer als SheetJS 0.18.5 für fremde Dateien) ---------- */
function excelSerial(d){ return (d.getTime()-Date.UTC(1899,11,30))/86400000; }
function fromExcelJSCell(cell){
  // liefert eine Zelle im SheetJS-Format {t,v,z,f}
  var v=cell.value, z=cell.numFmt&&cell.numFmt!=="General"?cell.numFmt:undefined, o={};
  if(v==null) return null;
  if(cell.type===1) return null; // verbundene Zelle (Teil einer Zusammenführung)
  if(typeof v==="object" && !(v instanceof Date)){
    if(v.formula!=null || v.sharedFormula!=null){ o.f=String(cell.formula||v.formula||""); v=v.result; if(v&&typeof v==="object"&&v.error) v=String(v.error); }
    else if(v.richText) v=v.richText.map(function(t){ return t.text; }).join("");
    else if(v.text!=null) v=typeof v.text==="string"?v.text:(v.text.richText||[]).map(function(t){ return t.text; }).join("");
    else if(v.error) { o.t="e"; o.w=String(v.error); o.v=0; return o; }
    else v=null;
  }
  if(v instanceof Date){ o.t="n"; o.v=excelSerial(v); if(!z) z="dd.mm.yyyy"; }
  else if(typeof v==="number"){ o.t="n"; o.v=v; }
  else if(typeof v==="boolean"){ o.t="b"; o.v=v; }
  else if(v==null){ if(!o.f) return null; o.t="s"; o.v=""; }
  else { o.t="s"; o.v=String(v); }
  if(z) o.z=z;
  return o;
}
function loadWorkbookExcelJS(XLSX, ExcelJS, data, name){
  var wb=new ExcelJS.Workbook();
  return wb.xlsx.load(data).then(function(){
    var sheets=[];
    wb.worksheets.forEach(function(ws){
      if(ws.state==="veryHidden") return;
      var sj={}, maxR=0, maxC=0;
      ws.eachRow({includeEmpty:false},function(row,r){ row.eachCell({includeEmpty:false},function(cell,c){
        var o=fromExcelJSCell(cell); if(!o) return; sj[XLSX.utils.encode_cell({r:r-1,c:c-1})]=o; maxR=Math.max(maxR,r); maxC=Math.max(maxC,c); }); });
      if(maxR) sj["!ref"]=XLSX.utils.encode_range({s:{r:0,c:0},e:{r:maxR-1,c:maxC-1}});
      var cols=[]; for(var c=1;c<=maxC;c++){ var w=ws.getColumn(c).width; cols.push(w?{wch:w}:{}); } sj["!cols"]=cols;
      sj["!merges"]=((ws.model&&ws.model.merges)||[]).map(function(m){ return XLSX.utils.decode_range(m); });
      sheets.push(parseSheet(XLSX, sj, ws.name));
    });
    return {name:name||"Tabelle.xlsx", sheets:sheets};
  });
}

/* ---------- CSV (auch deutsches Format mit Semikolon) ---------- */
function parseCSV(text){
  text=String(text).replace(/^\uFEFF/,"");
  var first=text.split(/\r?\n/)[0]||"", count=function(ch){ return first.split(ch).length-1; };
  var d=[";",",","\t"].sort(function(a,b){ return count(b)-count(a); })[0];
  var rows=[], row=[], cell="", q=false;
  for(var i=0;i<text.length;i++){
    var ch=text[i];
    if(q){ if(ch==='"'){ if(text[i+1]==='"'){ cell+='"'; i++; } else q=false; } else cell+=ch; }
    else if(ch==='"') q=true;
    else if(ch===d){ row.push(cell); cell=""; }
    else if(ch==="\n"||ch==="\r"){ if(ch==="\r"&&text[i+1]==="\n") i++; row.push(cell); rows.push(row); row=[]; cell=""; }
    else cell+=ch;
  }
  if(cell!==""||row.length){ row.push(cell); rows.push(row); }
  return rows.filter(function(r){ return r.some(function(c){ return String(c).trim()!==""; }); });
}
function loadCSV(XLSX, text, name){
  var rows=parseCSV(text), sj={}, maxC=0;
  rows.forEach(function(r,ri){ r.forEach(function(raw,ci){
    var t=String(raw).trim(); if(t==="") return; var o;
    var dm=t.match(/^(\d{1,2})\.(\d{1,2})\.(\d{4})$/), n=/^-?[\d.]*\d(,\d+)?\s*(€|%)?$/.test(t)||/^-?\d+(\.\d+)?$/.test(t)?parseNum(t.replace(/%$/,"")):null;
    if(dm) o={t:"n",v:dateToSerial(normDate(t)),z:"dd.mm.yyyy"};
    else if(n!=null) o={t:"n",v:/%$/.test(t)?n/100:n};
    else o={t:"s",v:t};
    sj[XLSX.utils.encode_cell({r:ri,c:ci})]=o; maxC=Math.max(maxC,ci+1); }); });
  if(rows.length) sj["!ref"]=XLSX.utils.encode_range({s:{r:0,c:0},e:{r:rows.length-1,c:Math.max(0,maxC-1)}});
  return {name:name, sheets:[parseSheet(XLSX, sj, String(name).replace(/\.[^.]+$/,"").slice(0,31)||"Tabelle")]};
}

/* ---------- Adressen ---------- */
function dataRowNo(sh,i){ return sh.hr+2+i; }
function footRowNo(sh,j){ return sh.hr+2+sh.rows.length+j; }
function cellAt(sh,r,c){ // r 1-basiert, c absolute Spalte
  if(sh.hr<0){ var g=sh.top[r-1]; return g?g[c]||null:null; }
  if(r<=sh.hr){ var t=sh.top[r-1]; return t?t[c]||null:null; }
  var k=c-sh.c0; if(k<0||k>=sh.ncol) return sh.extra?sh.extra[r+":"+c]||null:null;
  if(r===sh.hr+1) return {v:sh.headers[k]};
  var i=r-sh.hr-2; if(i<sh.rows.length) return sh.rows[i][k]||null;
  var j=i-sh.rows.length; return sh.foot[j]?sh.foot[j][k]||null:null;
}
function eachFormulaCell(sh,fn){
  sh.top.forEach(function(row){ row.forEach(function(c){ if(c&&c.f) fn(c); }); });
  sh.rows.forEach(function(row){ row.forEach(function(c){ if(c&&c.f) fn(c); }); });
  sh.foot.forEach(function(row){ row.forEach(function(c){ if(c&&c.f) fn(c); }); });
}

/* ---------- Rechnen ---------- */
function Calc(model){ this.m=model; this.memo={}; this.stack={}; }
Calc.prototype.sheet=function(name){ var s=this.m.sheets; for(var i=0;i<s.length;i++) if(s[i].name.toLowerCase()===String(name).toLowerCase()) return s[i]; return null; };
Calc.prototype.value=function(sh,r,c){
  var key=sh.name+"!"+r+"!"+c; if(key in this.memo) return this.memo[key];
  var cell=cellAt(sh,r,c), v;
  if(!cell) v=null;
  else if(cell.f){
    if(this.stack[key]) return "#ZIRKEL";
    this.stack[key]=1; v=this.formula(sh,cell.f); delete this.stack[key];
    if(v===undefined) v=(cell.v!=null?cell.v:"#?");
  } else v=cell.v;
  this.memo[key]=v; return v;
};
Calc.prototype.num=function(v){ if(typeof v==="number") return v; if(v==null||v==="") return 0; if(isDateStr(v)) return dateToSerial(v); var n=parseNum(v); return n==null?NaN:n; };
/* Formeln: eigener Parser und Interpreter, kein eval/new Function.
   Unbekannte Funktionen -> undefined (dann wird der in Excel gespeicherte Wert angezeigt). */
var ERRS=/^#(NV|N\/A|BEZUG!|REF!|WERT!|VALUE!|DIV\/0!|NAME\?|ZAHL!|NUM!|NULL!|ZIRKEL|\?)$/;
function FErr(t){ this.err=t; }
function isErr(v){ return v instanceof FErr; }
function Unsupported(){}
function fTokens(src){
  var out=[], i=0, m, re;
  var SH="(?:'((?:[^']|'')+)'|([A-Za-zÄÖÜäöüß_][\\wÄÖÜäöüß.]*))!";
  var CELL=new RegExp("^(?:"+SH+")?\\$?([A-Z]{1,3})\\$?(\\d+)(?::\\$?([A-Z]{1,3})\\$?(\\d+))?(?![\\w(])");
  var COLS=new RegExp("^(?:"+SH+")?\\$?([A-Z]{1,3}):\\$?([A-Z]{1,3})(?![\\w(])");
  while(i<src.length){
    var ch=src[i], rest=src.slice(i);
    if(/\s/.test(ch)){ i++; continue; }
    if(ch==='"'){ var j=i+1, str=""; while(j<src.length){ if(src[j]==='"'){ if(src[j+1]==='"'){ str+='"'; j+=2; continue; } break; } str+=src[j++]; } if(j>=src.length) throw new Unsupported(); out.push({t:"str",v:str}); i=j+1; continue; }
    if((m=rest.match(/^(\d+\.?\d*|\.\d+)([eE][+-]?\d+)?/))){ out.push({t:"num",v:parseFloat(m[0])}); i+=m[0].length; continue; }
    if((m=rest.match(/^#(NV|N\/A|BEZUG!|REF!|WERT!|VALUE!|DIV\/0!|NAME\?|ZAHL!|NUM!|NULL!)/i))){ out.push({t:"err",v:m[0].toUpperCase()}); i+=m[0].length; continue; }
    if((m=rest.match(CELL))){ out.push({t:"ref",sh:m[1]?m[1].replace(/''/g,"'"):m[2]||null,c1:m[3],r1:+m[4],c2:m[5]||m[3],r2:m[6]?+m[6]:+m[4]}); i+=m[0].length; continue; }
    if((m=rest.match(COLS))){ out.push({t:"ref",sh:m[1]?m[1].replace(/''/g,"'"):m[2]||null,c1:m[3],r1:1,c2:m[4],r2:-1}); i+=m[0].length; continue; }
    if((m=rest.match(/^([A-Za-zÄÖÜäöüß_][\wÄÖÜäöüß.]*)\s*\(/))){ out.push({t:"fn",v:m[1].toUpperCase().replace(/^_XLFN\./,"")}); i+=m[0].length; continue; }
    if((m=rest.match(/^(WAHR|TRUE|FALSCH|FALSE)(?![\w(])/i))){ out.push({t:"bool",v:/^(WAHR|TRUE)$/i.test(m[1])}); i+=m[0].length; continue; }
    if((m=rest.match(/^(<=|>=|<>|[-+*/^&=<>%(),;])/))){ out.push({t:"op",v:m[0]===";"?",":m[0]}); i+=m[0].length; continue; }
    throw new Unsupported();
  }
  return out;
}
var FN_ALIAS={SUMME:"SUM",MITTELWERT:"AVERAGE",ANZAHL:"COUNT",ANZAHL2:"COUNTA",RUNDEN:"ROUND",AUFRUNDEN:"ROUNDUP",ABRUNDEN:"ROUNDDOWN",
  WENN:"IF",WENNFEHLER:"IFERROR",ISTZAHL:"ISNUMBER",ISTLEER:"ISBLANK",ISTFEHLER:"ISERROR",UND:"AND",ODER:"OR",NICHT:"NOT",
  TEILERGEBNIS:"SUBTOTAL",SUMMEWENN:"SUMIF",ZÄHLENWENN:"COUNTIF",SVERWEIS:"VLOOKUP",HEUTE:"TODAY",DATUM:"DATE",VERKETTEN:"CONCATENATE",
  LINKS:"LEFT",RECHTS:"RIGHT",LÄNGE:"LEN",GLÄTTEN:"TRIM",GROSS:"UPPER",KLEIN:"LOWER",JAHR:"YEAR",MONAT:"MONTH",TAG:"DAY",WOCHENTAG:"WEEKDAY"};
Calc.prototype.formula=function(sh,f){
  var self=this, toks, pos=0;
  try{ toks=fTokens(String(f).replace(/^=/,"")); }catch(e){ return undefined; }
  var peek=function(){ return toks[pos]; }, isOp=function(v){ var t=toks[pos]; return t&&t.t==="op"&&t.v===v; };
  var expect=function(v){ if(!isOp(v)) throw new Unsupported(); pos++; };
  function sheetOf(name){ if(!name) return sh; var s2=self.sheet(name); if(!s2) throw new FErr("#BEZUG!"); return s2; }
  function lastRow(s2){ return s2.hr<0?s2.top.length:s2.hr+1+s2.rows.length+s2.foot.length; }
  function cellVal(s2,r,c){ var v=self.value(s2,r,c); if(typeof v==="string"&&ERRS.test(v)) return new FErr(v==="#?"?"#NAME?":v); return v; }
  function rangeCells(rg){ var out=[], r2=rg.r2<0?lastRow(rg.s):rg.r2; for(var r=Math.min(rg.r1,r2);r<=Math.max(rg.r1,r2);r++) for(var c=Math.min(rg.c1,rg.c2);c<=Math.max(rg.c1,rg.c2);c++) out.push({r:r,c:c,v:cellVal(rg.s,r,c)}); return out; }
  function scalar(v){ if(v&&v.range){ if(v.r1===v.r2&&v.c1===v.c2) return cellVal(v.s,v.r1,v.c1); throw new FErr("#WERT!"); } return v; }
  function toNum(v){ v=scalar(v); if(isErr(v)) throw v; if(typeof v==="number") return v; if(typeof v==="boolean") return v?1:0; if(v==null||v==="") return 0;
    if(isDateStr(v)) return dateToSerial(v); if(looksTime(v)) return timeNum(v); var n=parseNum(v); if(n==null) throw new FErr("#WERT!"); return n; }
  function toStr(v){ v=scalar(v); if(isErr(v)) throw v; if(v==null) return ""; if(typeof v==="boolean") return v?"WAHR":"FALSCH"; if(typeof v==="number") return String(Math.round(v*1e9)/1e9).replace(".",","); return String(v); }
  function toBool(v){ v=scalar(v); if(isErr(v)) throw v; if(typeof v==="boolean") return v; if(typeof v==="number") return v!==0; if(v==null||v==="") return false;
    if(/^(wahr|true)$/i.test(v)) return true; if(/^(falsch|false)$/i.test(v)) return false; throw new FErr("#WERT!"); }
  function nums(args,opts){ var out=[]; args.forEach(function(a){
      if(a&&a.range){ rangeCells(a).forEach(function(x){ if(opts&&opts.skipSub){ var cl=cellAt(a.s,x.r,x.c); if(cl&&cl.f&&/(SUBTOTAL|TEILERGEBNIS)\s*\(/i.test(cl.f)) return; }
        if(isErr(x.v)) throw x.v; if(typeof x.v==="number") out.push(x.v); }); }
      else if(a!==null&&a!==undefined) out.push(toNum(a)); }); return out; }
  function cmp(a,b){ // Excel-Vergleich: Zahlen < Text < Wahrheitswerte
    a=scalar(a); b=scalar(b); if(isErr(a)) throw a; if(isErr(b)) throw b;
    if((a==null||a==="")&&(b==null||b==="")) return 0;
    var na=a==null||a===""?(typeof b==="string"?"":0):a, nb=b==null||b===""?(typeof a==="string"?"":0):b;
    if(typeof na==="string"&&isDateStr(na)&&typeof nb==="number") na=dateToSerial(na); if(typeof nb==="string"&&isDateStr(nb)&&typeof na==="number") nb=dateToSerial(nb);
    var ta=typeof na, tb=typeof nb, rank={number:0,string:1,boolean:2};
    if(ta!==tb) return rank[ta]-rank[tb];
    if(ta==="string"){ na=na.toLowerCase(); nb=nb.toLowerCase(); }
    return na<nb?-1:na>nb?1:0; }
  function crit(c){ // Kriterium für SUMIF/COUNTIF
    c=scalar(c); var op="=", w=c; if(typeof c==="string"){ var mm=c.match(/^(<=|>=|<>|<|>|=)?(.*)$/); op=mm[1]||"="; w=mm[2]; var pn=parseNum(w); if(pn!=null&&w.trim()!=="") w=pn; }
    var re=typeof w==="string"&&/[*?]/.test(w)?new RegExp("^"+w.replace(/[.+^${}()|[\]\\]/g,"\\$&").replace(/\*/g,".*").replace(/\?/g,".")+"$","i"):null;
    return function(v){ if(isErr(v)) return false; if(re) return op==="<>"?!re.test(String(v||"")):re.test(String(v||""));
      var d=cmp(typeof w==="number"&&typeof v==="string"&&parseNum(v)!=null?parseNum(v):v,w);
      return op==="="?d===0:op==="<>"?d!==0:op==="<"?d<0:op===">"?d>0:op==="<="?d<=0:d>=0; }; }
  var FN={
    SUM:function(a){ return nums(a).reduce(function(x,y){ return x+y; },0); },
    AVERAGE:function(a){ var n=nums(a); if(!n.length) throw new FErr("#DIV/0!"); return n.reduce(function(x,y){ return x+y; },0)/n.length; },
    MIN:function(a){ var n=nums(a); return n.length?Math.min.apply(null,n):0; },
    MAX:function(a){ var n=nums(a); return n.length?Math.max.apply(null,n):0; },
    COUNT:function(a){ var n=0; a.forEach(function(x){ if(x&&x.range) rangeCells(x).forEach(function(c){ if(typeof c.v==="number") n++; }); else if(typeof scalar(x)==="number") n++; }); return n; },
    COUNTA:function(a){ var n=0; a.forEach(function(x){ if(x&&x.range) rangeCells(x).forEach(function(c){ if(c.v!=null&&c.v!=="") n++; }); else if(x!=null) n++; }); return n; },
    ROUND:function(a){ var x=toNum(a[0]), d=a[1]==null?0:toNum(a[1]), p=Math.pow(10,d); return Math.round(x*p+(x>=0?1e-9:-1e-9))/p; },
    ROUNDUP:function(a){ var x=toNum(a[0]), p=Math.pow(10,a[1]==null?0:toNum(a[1])); return (x<0?-1:1)*Math.ceil(Math.abs(x)*p-1e-9)/p; },
    ROUNDDOWN:function(a){ var x=toNum(a[0]), p=Math.pow(10,a[1]==null?0:toNum(a[1])); return (x<0?-1:1)*Math.floor(Math.abs(x)*p+1e-9)/p; },
    ABS:function(a){ return Math.abs(toNum(a[0])); },
    IF:"lazy", IFERROR:"lazy", AND:function(a){ return a.every(function(x){ return toBool(x); }); }, OR:function(a){ return a.some(function(x){ return toBool(x); }); },
    NOT:function(a){ return !toBool(a[0]); },
    ISNUMBER:function(a){ try{ return typeof scalar(a[0])==="number"; }catch(e){ return false; } },
    ISBLANK:function(a){ try{ var v=scalar(a[0]); return v==null||v===""; }catch(e){ return false; } },
    ISERROR:function(a){ try{ return isErr(scalar(a[0])); }catch(e){ return true; } },
    SUBTOTAL:function(a){ var k=toNum(a[0])%100, n=nums(a.slice(1),{skipSub:true});
      if(k===9) return n.reduce(function(x,y){ return x+y; },0); if(k===1){ if(!n.length) throw new FErr("#DIV/0!"); return n.reduce(function(x,y){ return x+y; },0)/n.length; }
      if(k===2) return n.length; if(k===4) return n.length?Math.max.apply(null,n):0; if(k===5) return n.length?Math.min.apply(null,n):0;
      if(k===3){ var c=0; a.slice(1).forEach(function(x){ if(x&&x.range) rangeCells(x).forEach(function(cc){ var cl=cellAt(x.s,cc.r,cc.c); if(cc.v!=null&&cc.v!==""&&!(cl&&cl.f&&/(SUBTOTAL|TEILERGEBNIS)/i.test(cl.f))) c++; }); }); return c; }
      throw new Unsupported(); },
    SUMIF:function(a){ var test=crit(a[1]), src=rangeCells(a[0]), sum=a[2]&&a[2].range?rangeCells({s:a[2].s,r1:a[2].r1,c1:a[2].c1,r2:a[2].r1+(src.length?src[src.length-1].r-src[0].r:0),c2:a[2].c1}):src, t=0;
      src.forEach(function(x,i){ if(test(x.v)){ var v=sum[i]&&sum[i].v; if(typeof v==="number") t+=v; else if(typeof v==="string"&&parseNum(v)!=null) t+=parseNum(v); } }); return t; },
    COUNTIF:function(a){ var test=crit(a[1]); return rangeCells(a[0]).filter(function(x){ return test(x.v); }).length; },
    VLOOKUP:function(a){ var key=scalar(a[0]); if(isErr(key)) throw key; var rg=a[1]; if(!rg||!rg.range) throw new FErr("#BEZUG!"); var col=toNum(a[2]), exact=a.length<4||a[3]==null?false:!toBool(a[3]);
      var r2=rg.r2<0?lastRow(rg.s):rg.r2, hit=null;
      for(var r=rg.r1;r<=r2;r++){ var v=cellVal(rg.s,r,rg.c1); if(v==null||v==="") continue; if(cmp(v,key)===0){ hit=r; break; } if(!exact){ if(cmp(v,key)<0) hit=r; else break; } }
      if(hit==null) throw new FErr("#NV"); var out=cellVal(rg.s,hit,rg.c1+col-1); if(isErr(out)) throw out; return out==null?0:out; },
    TRUE:function(){ return true; }, FALSE:function(){ return false; }, WAHR:function(){ return true; }, FALSCH:function(){ return false; },
    TODAY:function(){ return todaySerial(self.m.today?new Date(self.m.today):undefined); },
    DATE:function(a){ return (Date.UTC(toNum(a[0]),toNum(a[1])-1,toNum(a[2]))-Date.UTC(1899,11,30))/86400000; },
    YEAR:function(a){ return +serialToDate(toNum(a[0])).slice(6); }, MONTH:function(a){ return +serialToDate(toNum(a[0])).slice(3,5); }, DAY:function(a){ return +serialToDate(toNum(a[0])).slice(0,2); },
    CONCATENATE:function(a){ return a.map(toStr).join(""); }, LEFT:function(a){ return toStr(a[0]).slice(0,a[1]==null?1:toNum(a[1])); },
    RIGHT:function(a){ var s=toStr(a[0]), n=a[1]==null?1:toNum(a[1]); return n?s.slice(-n):""; }, LEN:function(a){ return toStr(a[0]).length; },
    TRIM:function(a){ return toStr(a[0]).trim().replace(/\s+/g," "); }, UPPER:function(a){ return toStr(a[0]).toUpperCase(); }, LOWER:function(a){ return toStr(a[0]).toLowerCase(); }
  };
  function args(){ var out=[]; if(isOp(")")){ pos++; return out; }
    while(true){ if(isOp(",")){ out.push(null); pos++; continue; } out.push({start:pos}); skipExpr(); if(isOp(",")){ pos++; if(isOp(")")){ out.push(null); pos++; return out; } continue; } expect(")"); return out; } }
  function skipExpr(){ var depth=0; while(pos<toks.length){ var t=toks[pos]; if(t.t==="fn"){ depth++; pos++; continue; } if(t.t==="op"&&t.v==="("){ depth++; pos++; continue; }
      if(t.t==="op"&&t.v===")"){ if(!depth) return; depth--; pos++; continue; } if(t.t==="op"&&t.v===","&&!depth) return; pos++; } }
  function evalAt(start){ var save=pos; pos=start; var v; try{ v=comparison(); } catch(e){ pos=save; if(isErr(e)) return e; throw e; } pos=save; return v; }
  function call(name){
    name=FN_ALIAS[name]||name; var fn=FN[name]; if(!fn) throw new Unsupported();
    var a=args();
    if(name==="IF"){ var c=a[0]?evalAt(a[0].start):false; if(isErr(c)) throw c; var cv=toBool(c);
      var b=cv?a[1]:a[2]; if(!b) return cv?true:false; var r=evalAt(b.start); if(isErr(r)) throw r; return r; }
    if(name==="IFERROR"){ var v=a[0]?evalAt(a[0].start):null; try{ v=scalar(v); }catch(e){ if(isErr(e)) v=e; else throw e; }
      if(isErr(v)){ if(!a[1]) return ""; var alt=evalAt(a[1].start); if(isErr(alt)) throw alt; return alt; } return v; }
    var vals=a.map(function(x){ if(!x) return null; var v=evalAt(x.start); if(isErr(v)) throw v; return v; });
    return fn(vals);
  }
  function primary(){
    var t=peek(); if(!t) throw new Unsupported(); pos++;
    if(t.t==="num") return t.v; if(t.t==="str") return t.v; if(t.t==="bool") return t.v; if(t.t==="err") throw new FErr(t.v);
    if(t.t==="ref"){ var s2=sheetOf(t.sh); return {range:true,s:s2,r1:t.r1,c1:colIndex(t.c1),r2:t.r2,c2:colIndex(t.c2)}; }
    if(t.t==="fn") return call(t.v);
    if(t.t==="op"&&t.v==="("){ var v=comparison(); expect(")"); return v; }
    if(t.t==="op"&&(t.v==="-"||t.v==="+")){ var u=postfix(); return t.v==="-"?-toNum(u):toNum(u); }
    throw new Unsupported();
  }
  function postfix(){ var v=primary(); while(isOp("%")){ pos++; v=toNum(v)/100; } return v; }
  function power(){ var v=postfix(); while(isOp("^")){ pos++; v=Math.pow(toNum(v),toNum(postfix())); } return v; }
  function mult(){ var v=power(); while(isOp("*")||isOp("/")){ var o=toks[pos++].v, b=toNum(power()), a=toNum(v); if(o==="/"&&b===0) throw new FErr("#DIV/0!"); v=o==="*"?a*b:a/b; } return v; }
  function add(){ var v=mult(); while(isOp("+")||isOp("-")){ var o=toks[pos++].v, b=toNum(mult()); v=o==="+"?toNum(v)+b:toNum(v)-b; } return v; }
  function concat(){ var v=add(); while(isOp("&")){ pos++; v=toStr(v)+toStr(add()); } return v; }
  function comparison(){ var v=concat(); while(isOp("=")||isOp("<>")||isOp("<")||isOp(">")||isOp("<=")||isOp(">=")){ var o=toks[pos++].v, d=cmp(v,concat());
      v=o==="="?d===0:o==="<>"?d!==0:o==="<"?d<0:o===">"?d>0:o==="<="?d<=0:d>=0; } return v; }
  try{
    var res=comparison(); if(pos<toks.length) return undefined;
    res=scalar(res); if(isErr(res)) return res.err;
    if(typeof res==="boolean") return res?"WAHR":"FALSCH";
    if(typeof res==="number"){ if(!isFinite(res)) return "#DIV/0!"; return Math.round(res*1e9)/1e9; }
    return res==null?0:res;
  }catch(e){ if(isErr(e)) return e.err; return undefined; }
};
function display(model, sh, r, c, calc){
  calc=calc||new Calc(model);
  return calc.value(sh,r,c);
}

/* ---------- Zeilen einfügen / löschen ---------- */
// Bezüge anderer Blätter auf dieses Blatt ('Blatt'!A1, Blatt!A1:B5) mitziehen
function transformExtRefs(f, sheetName, fn){
  var low=String(sheetName).toLowerCase();
  return splitStrings(f).map(function(seg){
    if(seg.str) return seg.t;
    return seg.t.replace(/('((?:[^']|'')+)'|([A-Za-zÄÖÜäöüß_][\wÄÖÜäöüß.]*))!(\$?)([A-Z]{1,3})(\$?)(\d+)(?::(\$?)([A-Z]{1,3})(\$?)(\d+))?(?![\d(A-Za-z_])/g,
      function(all,pfx,q,nm,ca,c,ra,r,ca2,c2,ra2,r2){
        var n=(q!=null?q.replace(/''/g,"'"):nm).toLowerCase(); if(n!==low) return all;
        if(c2){ var a=fn(+r,colIndex(c),"start"), b=fn(+r2,colIndex(c2),"end"); if(a==null||b==null||a>b) return "#BEZUG!"; return pfx+"!"+ca+c+ra+a+":"+ca2+c2+ra2+b; }
        var x=fn(+r,colIndex(c),"single"); return x==null?"#BEZUG!":pfx+"!"+ca+c+ra+x;
      });
  }).join("");
}
function shiftOthers(model, sh, fn){ model.sheets.forEach(function(o){ if(o===sh) return; eachFormulaCell(o,function(c){ c.f=transformExtRefs(c.f,sh.name,fn); }); }); }
function shiftSheet(sh, fn){ eachFormulaCell(sh,function(c){ c.f=transformRefs(c.f,fn); }); }
function templateRow(sh){
  var tpl=[]; for(var k=0;k<sh.ncol;k++){ tpl.push(null);
    for(var i=sh.rows.length-1;i>=0;i--){ if(!isData(sh,i)) continue; var c=sh.rows[i][k]; if(c&&c.f){ tpl[k]={f:c.f,from:dataRowNo(sh,i)}; break; } if(c&&!blank(c)) break; } }
  return tpl;
}
function computedCols(sh){
  var out=[]; for(var k=0;k<sh.ncol;k++){ var nf=0,n=0; eachData(sh,function(r){ var c=r[k]; if(c&&!blank(c)){ n++; if(c.f) nf++; } }); out.push(n>0 && nf/n>=0.6); } return out;
}
function appendRow(sh, values, model){
  var tpl=templateRow(sh), comp=computedCols(sh);
  var newRow=dataRowNo(sh,sh.rows.length), lastData=newRow-1;
  // vorhandene Formeln: Zeilen ab newRow verschieben, Bereiche die an der letzten Datenzeile enden erweitern
  shiftSheet(sh,function(row,col,part,abs){
    if(row>=newRow) return row+1;
    if(part==="end" && row===lastData) return row+1;
    return row;
  });
  var cells=[];
  for(var k=0;k<sh.ncol;k++){
    if(comp[k] && tpl[k]){ var d=newRow-tpl[k].from; cells.push({f:transformRefs(tpl[k].f,function(row,col,part,abs){ return abs?row:row+d; })}); }
    else cells.push(values && values[k]!=null && values[k]!=="" ? {v:values[k]} : null);
  }
  sh.rows.push(cells); if(sh.kind) sh.kind.push("");
  if(model) shiftOthers(model,sh,function(row,col,part){ if(row>=newRow) return row+1; if(part==="end"&&row===lastData) return row+1; return row; });
  return newRow;
}
function deleteRow(sh, excelRow, model){
  var i=excelRow-sh.hr-2; if(i<0||i>=sh.rows.length) return false;
  sh.rows.splice(i,1); if(sh.kind) sh.kind.splice(i,1);
  var fn=function(row,col,part){
    if(row<excelRow) return row;
    if(row===excelRow){ if(part==="single") return null; if(part==="start") return row; return row-1; }
    return row-1;
  };
  shiftSheet(sh,fn); if(model) shiftOthers(model,sh,fn);
  return true;
}

/* ---------- Spalte anhängen ---------- */
function addColumn(sh, name){
  var k=findCol(sh,name); if(k>=0 && String(sh.headers[k]).toLowerCase()===String(name).toLowerCase()) return k;
  sh.headers.push(name); sh.ncol++; sh.colFmt.push(null); sh.dateCol.push(/datum|am$|\bam\b/i.test(name)); if(sh.timeCol) sh.timeCol.push(false);
  sh.rows.forEach(function(r){ r.push(null); }); sh.foot.forEach(function(r){ r.push(null); });
  if(sh.widths) sh.widths[sh.c0+sh.ncol-1]=Math.max(12,name.length+2);
  return sh.ncol-1;
}

/* ---------- Werte setzen ---------- */
function findCol(sh, name){
  if(name==null) return -1; var n=String(name).toLowerCase().replace(/[^a-zäöüß0-9]/g,"");
  var h=sh.headers.map(function(x){ return String(x).toLowerCase().replace(/[^a-zäöüß0-9]/g,""); });
  var i=h.indexOf(n); if(i>=0) return i;
  i=h.findIndex(function(x){ return x && (x.indexOf(n)===0 || n.indexOf(x)===0); }); if(i>=0) return i;
  i=h.findIndex(function(x){ return x && n.length>=3 && x.indexOf(n)>=0; });
  if(i<0 && /^[A-Z]{1,3}$/.test(String(name))) { var ci=colIndex(String(name))-sh.c0; if(ci>=0&&ci<sh.ncol) return ci; }
  return i;
}
function numericCol(sh,k){ var n=0,num=0; eachData(sh,function(r){ var c=r[k]; if(c&&!blank(c)&&!c.f){ n++; if(typeof c.v==="number") num++; } }); return n>0&&num/n>=0.6; }
function coerce(sh,k,val){
  if(val==null||val==="") return null;
  if(sh.timeCol&&sh.timeCol[k]){ if(typeof val==="number") return val; var tt=parseTime(val); if(tt!=null) return tt; }
  if(sh.dateCol[k]){ var d=normDate(val); if(d) return d; }
  if(typeof val==="string" && /^\s*-?[\d.,]+\s*%\s*$/.test(val)){ var pc=parseNum(val.replace("%","")); if(pc!=null) return Math.round(pc*1e6)/1e8; }  // 19 % -> 0,19
  if(typeof val==="string" && /%/.test(sh.colFmt[k]||"") ){ var pv=parseNum(val); if(pv!=null) return pv>1?pv/100:pv; }
  if(numericCol(sh,k)){ var n=typeof val==="number"?val:parseNum(val); if(n!=null) return n; }
  if(typeof val==="number") return val;
  return String(val);
}
function setCell(sh, excelRow, k, raw, opts){
  var i=excelRow-sh.hr-2, row;
  if(i>=0&&i<sh.rows.length) row=sh.rows[i]; else { var j=i-sh.rows.length; row=sh.foot[j]; }
  if(!row) return;
  // Formeln nur, wenn ein Mensch sie in die Zelle tippt; Werte von KI oder E-Rechnung bleiben immer Text.
  if(opts&&opts.formula && typeof raw==="string" && raw.trim().charAt(0)==="=") row[k]={f:splitStrings(raw.trim().slice(1)).map(function(p){ return p.str?p.t:p.t.toUpperCase(); }).join("")};
  else { var v=coerce(sh,k,raw); row[k]= v==null?null:{v:v}; }
}

/* ---------- Abfragen (exakt gerechnet) ---------- */
function rowValues(model, sh, i, calc){ var r=dataRowNo(sh,i), o=[]; for(var k=0;k<sh.ncol;k++) o.push(calc.value(sh,r,sh.c0+k)); return o; }
function cmp(a,b){ var da=typeof a==="string"?anyDateSerial(a):null, db=typeof b==="string"?anyDateSerial(b):null;
  if(da!=null&&db!=null) return da-db;
  if(da!=null&&typeof b==="number"&&b>20000) return da-b; if(db!=null&&typeof a==="number"&&a>20000) return a-db;
  var na=typeof a==="number"?a:parseNum(a), nb=typeof b==="number"?b:parseNum(b);
  if(na!=null&&nb!=null) return na-nb;
  var ns=function(x){ return String(x==null?"":x).trim().replace(/\s+/g," ").toLowerCase(); };
  return ns(a).localeCompare(ns(b),"de"); }
function matches(v, op, w){
  var e=(v==null||v==="");
  switch(op){
    case "leer": return e; case "nicht leer": return !e;
    case "enthält": return !e && String(v).toLowerCase().indexOf(String(w).toLowerCase())>=0;
    case "!=": return e ? (w!=null&&w!=="") : cmp(v,w)!==0;
    case ">": return !e && cmp(v,w)>0; case "<": return !e && cmp(v,w)<0;
    case ">=": return !e && cmp(v,w)>=0; case "<=": return !e && cmp(v,w)<=0;
    default: return e ? (w==null||w==="") : cmp(v,w)===0;
  }
}
function query(model, q){
  var calc=new Calc(model), sh=calc.sheet(q.blatt)||model.sheets.find(function(s){ return s.hr>=0; });
  if(!sh||sh.hr<0) return {error:"Blatt nicht gefunden"};
  var k=findCol(sh,q.spalte); var filters=(q.filter||[]).map(function(f){ return {k:findCol(sh,f.spalte),op:f.op||"=",w:f.wert}; });
  if(filters.some(function(f){ return f.k<0; })) return {error:"Spalte nicht gefunden"};
  var hits=[];
  eachData(sh,function(r,i){ var vals=rowValues(model,sh,i,calc); if(filters.every(function(f){ return matches(vals[f.k],f.op,f.w); })) hits.push({i:i,vals:vals}); });
  var art=(q.art||"summe").toLowerCase();
  if(art==="anzahl") return {value:hits.length, rows:hits.map(function(h){ return dataRowNo(sh,h.i); })};
  if(k<0) return {error:"Spalte nicht gefunden"};
  if(art==="liste") return {list:hits.map(function(h){ return h.vals[k]; }).filter(function(v){ return v!=null&&v!==""; }), rows:hits.map(function(h){ return dataRowNo(sh,h.i); }), col:k};
  var nums=hits.map(function(h){ return calc.num(h.vals[k]); }).filter(function(n){ return !isNaN(n); });
  var v = art==="mittelwert" ? (nums.length?nums.reduce(function(a,b){return a+b;},0)/nums.length:0)
        : art==="min" ? (nums.length?Math.min.apply(null,nums):null)
        : art==="max" ? (nums.length?Math.max.apply(null,nums):null)
        : nums.reduce(function(a,b){ return a+b; },0);
  return {value:v==null?null:Math.round(v*100)/100, rows:hits.map(function(h){ return dataRowNo(sh,h.i); }), col:k, money:/€|eur|preis|betrag|wert|netto|brutto|summe|mwst/i.test(sh.headers[k])};
}

/* ---------- Prüfen (allgemein) ---------- */
function checks(model,opts){
  var calc=new Calc(model), out=[];
  model.sheets.forEach(function(sh){
    if(sh.hr<0||!sh.rows.length) return;
    var comp=computedCols(sh);
    for(var k=0;k<sh.ncol;k++){
      var n=0, filledC=0, nums=0, numLike=0; eachData(sh,function(r){ n++; var c=r[k]; if(c&&!blank(c)){ filledC++; if(typeof c.v==="number"||c.f) nums++; if(typeof c.v==="number"||c.f||(typeof c.v==="string"&&/\d/.test(c.v)&&!/[a-zäöü]{3,}/i.test(c.v.replace(/€|eur|stk|ca\.?/gi,""))&&!/^\d{1,2}[.:/]\d{1,2}([.:/]\d{2,4})?$/.test(c.v.trim()))) numLike++; } });
      var keyRe=/(^|[^a-z])(nr|nummer|id|artikel-?nr|auftrags-?nr|rechnungs-?nr)([^a-z]|$)/i, key=keyRe.test(sh.headers[k]);
      var firstKey=sh.headers.findIndex(function(h){ return keyRe.test(h); }), uniq=key && k===firstKey;
      var seen={};
      eachData(sh,function(r,i){
        var row=dataRowNo(sh,i), c=r[k], v=calc.value(sh,row,sh.c0+k);
        if(typeof v==="string" && /^#/.test(v)) out.push({blatt:sh.name,zeile:row,k:k,problem:"Formel ergibt "+v});
        if(!comp[k] && filledC/n>=0.75 && filledC<n && blank(c)) out.push({blatt:sh.name,zeile:row,k:k,problem:sh.headers[k]+" fehlt"});
        if(!comp[k] && !key && !/tel|telefon|handy|mobil|fax|plz|iban|bic|e-?mail/i.test(sh.headers[k]) && numLike/Math.max(1,filledC)>=0.6 && c && !c.f && typeof c.v==="string" && c.v.trim() && !sh.dateCol[k] && !(sh.timeCol&&sh.timeCol[k]) && !/^#/.test(c.v)){
          if(parseNum(c.v)==null) out.push({blatt:sh.name,zeile:row,k:k,problem:"Text statt Zahl („"+c.v.trim()+"“) – Excel rechnet damit nicht"});
          else if(nums>0 && !/^0\d/.test(c.v.trim()) && !/tel|telefon|handy|mobil|fax|plz|iban|bic/i.test(sh.headers[k])) out.push({blatt:sh.name,zeile:row,k:k,problem:"Zahl als Text gespeichert („"+c.v.trim()+"“) – Summen in Excel lassen sie weg"}); }
        if(c && c.f && typeof v==="string" && /\?\?\?|unbekannt|nicht gefunden/i.test(v)) out.push({blatt:sh.name,zeile:row,k:k,problem:"Nachschlagen ohne Treffer: „"+v+"“"});
        if(/(ust|mwst|steuer).*(%|satz)|steuersatz|^satz$/i.test(sh.headers[k]) && typeof v==="number" && v>1 && v<=100) out.push({blatt:sh.name,zeile:row,k:k,problem:"Steuersatz "+deNum(v)+" statt "+deNum(v)+" % – Excel rechnet mit "+deNum(v*100)+" %"});
        if(uniq && v!=null && v!==""){ var s=String(v).trim().toLowerCase(); if(seen[s]) out.push({blatt:sh.name,zeile:row,k:k,problem:sh.headers[k]+" doppelt (auch Zeile "+seen[s]+")"}); else seen[s]=row; }
      });
    }
    lawChecks(sh,calc,out,opts);
  });
  return out;
}

/* Fachliche Prüfungen: Arbeitszeit (ArbZG), MwSt, offene Rechnungen */
function deNum(n){ return String(n).replace(".",","); }
function col(sh,re){ for(var k=0;k<sh.ncol;k++) if(re.test(sh.headers[k])) return k; return -1; }
function lawChecks(sh,calc,out,opts){
  var today=todaySerial(opts&&opts.today), add=function(i,k,p){ out.push({blatt:sh.name,zeile:dataRowNo(sh,i),k:k,problem:p}); };
  var V=function(i,k){ return k<0?null:calc.value(sh,dataRowNo(sh,i),sh.c0+k); };
  var dser=function(v,k){ if(typeof v==="number") return (sh.dateCol[k]||/[dy]/i.test(sh.colFmt[k]||""))&&v>20000?Math.floor(v):null; return anyDateSerial(v); };
  var who=function(v){ return String(v==null?"":v).trim().replace(/\s+/g," ").toLowerCase(); };
  var pauseMin=function(v){ if(typeof v==="number") return v; if(v==null||v==="") return 0; var t=String(v).toLowerCase().replace(",",".");
    var m=t.match(/(\d+)\s*\/\s*(\d+)/); var x=m?(+m[1]/+m[2]):parseFloat(t); if(isNaN(x)) return 0; return /std|h\b|stunde/.test(t)?x*60:x; };
  var cDat=col(sh,/^datum$|^tag$/i), cMa=col(sh,/mitarbeiter|^name$|person|^ma$/i), cBeg=col(sh,/beginn|von|start|anfang/i), cEnd=col(sh,/ende|bis/i),
      cPau=col(sh,/pause/i), cStd=col(sh,/^stunden|^std|arbeitszeit/i);
  // Arbeitszeit pro Person und Tag
  if(cDat>=0 && cMa>=0 && cStd>=0){
    var day={};
    eachData(sh,function(r,i){ var ds=dser(V(i,cDat),cDat), key=who(V(i,cMa))+"|"+(ds!=null?ds:V(i,cDat)), h=calc.num(V(i,cStd)); if(!isNaN(h)){ (day[key]=day[key]||{h:0,last:i}).h+=h; day[key].last=i; } });
    Object.keys(day).forEach(function(key){ if(day[key].h>10) add(day[key].last,cStd,"Über 10 Std. Arbeitszeit am Tag ("+deNum(Math.round(day[key].h*100)/100)+" Std., § 3 ArbZG)"); });
  }
  if(cBeg>=0 && cEnd>=0 && sh.timeCol && (sh.timeCol[cBeg]||sh.timeCol[cEnd])){
    var shifts=[];
    eachData(sh,function(r,i){
      var b=timeNum(V(i,cBeg)), e=timeNum(V(i,cEnd)); if(isNaN(b)||isNaN(e)) return;
      var dur=(e-b)*24; if(dur<0) dur+=24;
      var p=cPau>=0?pauseMin(V(i,cPau)):0; // Minuten
      var work=dur-p/60;
      if(work>9 && p<45) add(i,cPau>=0?cPau:cEnd,"Pause zu kurz: über 9 Std. Arbeit braucht 45 Min. (§ 4 ArbZG)");
      else if(work>6 && p<30) add(i,cPau>=0?cPau:cEnd,"Pause zu kurz: über 6 Std. Arbeit braucht 30 Min. (§ 4 ArbZG)");
      if(cDat>=0 && cMa>=0){ var d=dser(V(i,cDat),cDat); if(d!=null) shifts.push({i:i,ma:who(V(i,cMa)),start:d+b,end:d+e+(e<b?1:0)}); }
    });
    var byMa={}; shifts.forEach(function(s){ (byMa[s.ma]=byMa[s.ma]||[]).push(s); });
    Object.keys(byMa).forEach(function(m){ var l=byMa[m].sort(function(a,b){ return a.start-b.start; });
      for(var j=1;j<l.length;j++){ var gap=(l[j].start-l[j-1].end)*24; if(gap>0 && gap<11 && Math.floor(l[j].start)!==Math.floor(l[j-1].start)) add(l[j].i,cBeg,"Ruhezeit nur "+deNum(Math.round(gap*10)/10)+" Std. seit letzter Schicht, mindestens 11 Std. (§ 5 ArbZG)"); } });
  }
  // MwSt-Satz plausibel?
  var cNet=col(sh,/netto/i), cTax=col(sh,/mwst|ust|steuer/i);
  if(cNet>=0 && cTax>=0 && cTax!==cNet && !/satz|%/i.test(sh.headers[cTax])){
    eachData(sh,function(r,i){ var n=calc.num(V(i,cNet)), t=calc.num(V(i,cTax)); if(n>0 && !isNaN(t) && t>0){ var q=t/n; if(Math.abs(q-0.19)>0.002 && Math.abs(q-0.07)>0.002) add(i,cTax,"MwSt passt weder zu 19 % noch zu 7 % ("+deNum(Math.round(q*1000)/10)+" %)"); } });
  }
  // Offene Rechnungen
  var cSt=col(sh,/^status$/i), cRe=col(sh,/rechnung am|rechnungsdatum|re-datum/i), cBez=col(sh,/bezahlt am|zahlung am|bezahlt$/i), cFae=col(sh,/fällig|faellig/i);
  if((cRe>=0||cFae>=0) && cBez>=0){
    var limit=(opts&&opts.zahlungsziel)||30;
    eachData(sh,function(r,i){ var b=V(i,cBez); if(b!=null&&b!=="") return;
      var due=cFae>=0?V(i,cFae):null;
      var st=cSt>=0?String(V(i,cSt)||"").toLowerCase():""; if(/bezahlt|erledigt|storniert/.test(st)&&!/teil/.test(st)) return;
      var dueS=cFae>=0?dser(due,cFae):null;
      if(dueS!=null){ var over=today-dueS; if(over>0) add(i,cBez,"Seit "+over+" Tag"+(over===1?"":"en")+" überfällig"); return; }
      var d=cRe>=0?dser(V(i,cRe),cRe):null; if(d!=null){ var days=today-d; if(days>limit) add(i,cBez,"Rechnung seit "+days+" Tagen offen"); } });
  }
}

/* ---------- Für die KI beschreiben ---------- */
var KIND_LABEL={leer:"Leerzeile",summe:"Zwischensumme",abschnitt:"Abschnitt",notiz:"Notiz"};
function describe(model, maxChars){
  maxChars=maxChars||120000; var calc=new Calc(model), out=["Datei: "+model.name];
  model.sheets.forEach(function(sh){
    if(sh.hr<0){ out.push("\nBlatt „"+sh.name+"“: keine Tabelle erkannt."); return; }
    var comp=computedCols(sh), tpl=templateRow(sh);
    out.push("\nBlatt „"+sh.name+"“ — Kopfzeile in Excel-Zeile "+(sh.hr+1)+", Daten in Zeilen "+dataRowNo(sh,0)+"–"+dataRowNo(sh,sh.rows.length-1)+".");
    out.push("Spalten: "+sh.headers.map(function(h,k){ return colLetter(sh.c0+k)+" „"+h+"“ ("+(comp[k]?"berechnet, nicht setzen":sh.dateCol[k]?"Datum":sh.timeCol&&sh.timeCol[k]?"Uhrzeit HH:MM":numericCol(sh,k)?"Zahl":"Text")+")"; }).join("; "));
    out.push("Zeile;"+sh.headers.join(";"));
    if(sh.twoRowHeader) out.push("(Überschrift über zwei Zeilen, Excel-Zeilen "+sh.hr+"–"+(sh.hr+1)+")");
    sh.rows.forEach(function(r,i){ var row=dataRowNo(sh,i);
      if(!isData(sh,i)){ if(sh.kind[i]!=="leer") out.push(row+" ("+KIND_LABEL[sh.kind[i]]+", nicht ändern);"+rowValues(model,sh,i,calc).map(function(v){ return v==null?"":String(v).replace(/;/g,","); }).join(";")); return; }
      out.push(row+";"+rowValues(model,sh,i,calc).map(function(v,k){ if(sh.timeCol&&sh.timeCol[k]) v=fmtTime(v); return v==null?"":String(v).replace(/;/g,","); }).join(";")); });
    sh.foot.forEach(function(r,j){ var row=footRowNo(sh,j); var vals=[]; for(var k=0;k<sh.ncol;k++) vals.push(calc.value(sh,row,sh.c0+k)); out.push(row+" (Summenzeile);"+vals.map(function(v){ return v==null?"":v; }).join(";")); });
  });
  var s=out.join("\n"); return s.length>maxChars? s.slice(0,maxChars)+"\n[gekürzt]" : s;
}

/* ---------- KI-Änderungen anwenden ---------- */
function planOps(model, ops){
  // prüft und beschreibt Änderungen, ohne sie anzuwenden
  var calc=new Calc(model), out=[];
  (ops||[]).forEach(function(op){
    var sh=calc.sheet(op.blatt)||model.sheets.find(function(s){ return s.hr>=0; });
    if(!sh||sh.hr<0){ out.push({op:op,error:"Blatt „"+op.blatt+"“ nicht gefunden"}); return; }
    var comp=computedCols(sh), typ=String(op.typ||"").toLowerCase();
    var changes=[];
    if(typ==="neu"||typ==="aendern"||typ==="ändern"){
      var row=typ==="neu"?null:+op.zeile;
      if(row!=null){ var i=row-sh.hr-2; if(!(i>=0&&i<sh.rows.length)){ out.push({op:op,error:"Zeile "+op.zeile+" gibt es nicht"}); return; }
        if(!isData(sh,i)){ out.push({op:op,error:"Zeile "+op.zeile+" ist eine "+KIND_LABEL[sh.kind[i]]+" und wird nicht geändert"}); return; } }
      Object.keys(op.werte||{}).forEach(function(key){
        var k=findCol(sh,key); if(k<0){ changes.push({k:-1,name:key,error:"Spalte unbekannt"}); return; }
        if(comp[k]) return;
        var w=op.werte[key], old=row!=null?calc.value(sh,row,sh.c0+k):null, nv;
        if(w && typeof w==="object" && "plus" in w){ var base=calc.num(old); nv=Math.round(((isNaN(base)?0:base)+Number(w.plus))*1000)/1000; }
        else nv=coerce(sh,k,w);
        changes.push({k:k,name:sh.headers[k],old:old,nv:nv});
      });
      out.push({op:op, sheet:sh.name, typ:typ==="neu"?"neu":"aendern", row:row, changes:changes});
    } else if(typ==="loeschen"||typ==="löschen"){
      var r=+op.zeile, ii=r-sh.hr-2; if(!(ii>=0&&ii<sh.rows.length)){ out.push({op:op,error:"Zeile "+op.zeile+" gibt es nicht"}); return; }
      if(!isData(sh,ii)){ out.push({op:op,error:"Zeile "+op.zeile+" ist eine "+KIND_LABEL[sh.kind[ii]]+" und wird nicht gelöscht"}); return; }
      out.push({op:op, sheet:sh.name, typ:"loeschen", row:r, summary:rowValues(model,sh,ii,calc).filter(function(v){ return v!=null&&v!==""; }).slice(0,4).join(" · ")});
    } else out.push({op:op,error:"Unbekannte Änderung"});
  });
  return out;
}
function applyPlan(model, plan){
  var touched=[];
  // erst ändern/neu, Löschungen zuletzt von unten nach oben
  plan.filter(function(p){ return !p.error && p.typ!=="loeschen"; }).forEach(function(p){
    var sh=model.sheets.find(function(s){ return s.name===p.sheet; });
    if(p.typ==="neu"){ var vals=[]; p.changes.forEach(function(c){ if(c.k>=0) vals[c.k]=c.nv; }); var r=appendRow(sh,vals,model); touched.push({sheet:sh.name,row:r}); }
    else { p.changes.forEach(function(c){ if(c.k>=0) setCell(sh,p.row,c.k,c.nv); }); touched.push({sheet:sh.name,row:p.row}); }
  });
  plan.filter(function(p){ return !p.error && p.typ==="loeschen"; }).sort(function(a,b){ return b.row-a.row; }).forEach(function(p){
    var sh=model.sheets.find(function(s){ return s.name===p.sheet; }); deleteRow(sh,p.row,model);
  });
  return touched;
}

/* ---------- Schreiben ---------- */
function toXlsxCell(cell, calc, sh, r, c, colFmt, isDate){
  if(!cell) return null;
  var o={};
  if(cell.f){ o.f=cell.f; var v=calc.value(sh,r,c); if(typeof v==="number"){ o.t="n"; o.v=v; } else if(v!=null){ o.t="s"; o.v=String(v); } else { o.t="n"; o.v=0; } }
  else if(cell.v==null||cell.v==="") return null;
  else if(isDate && isDateStr(cell.v)){ o.t="n"; o.v=dateToSerial(cell.v); o.z=cell.z&&cell.z!=="General"?cell.z:(colFmt||"dd.mm.yyyy"); return o; }
  else if(typeof cell.v==="number"){ o.t="n"; o.v=cell.v; }
  else { o.t="s"; o.v=String(cell.v); }
  var z=cell.z||colFmt; if(z && o.t==="n") o.z=z;
  return o;
}
function buildWorkbook(XLSX, model){
  var wb=XLSX.utils.book_new(), calc=new Calc(model);
  model.sheets.forEach(function(sh){
    var ws={}, maxR=0, maxC=0;
    var put=function(r,c,cell){ if(!cell) return; ws[XLSX.utils.encode_cell({r:r-1,c:c})]=cell; maxR=Math.max(maxR,r); maxC=Math.max(maxC,c); };
    sh.top.forEach(function(row,ri){ (row||[]).forEach(function(cell,ci){ put(ri+1,ci,toXlsxCell(cell,calc,sh,ri+1,ci,null,cell&&cell.d)); }); });
    if(sh.hr>=0){
      sh.headers.forEach(function(h,k){ put(sh.hr+1,sh.c0+k,{t:"s",v:h}); });
      sh.rows.forEach(function(row,i){ var r=dataRowNo(sh,i); row.forEach(function(cell,k){ put(r,sh.c0+k,toXlsxCell(cell,calc,sh,r,sh.c0+k,sh.colFmt[k],sh.dateCol[k])); }); });
      sh.foot.forEach(function(row,j){ var r=footRowNo(sh,j); row.forEach(function(cell,k){ put(r,sh.c0+k,toXlsxCell(cell,calc,sh,r,sh.c0+k,sh.colFmt[k],false)); }); });
    }
    ws["!ref"]=XLSX.utils.encode_range({s:{r:0,c:0},e:{r:Math.max(0,maxR-1),c:maxC}});
    if(sh.widths&&sh.widths.length) ws["!cols"]=sh.widths.map(function(w){ return w?{wch:w}:{}; });
    if(sh.merges&&sh.merges.length) ws["!merges"]=sh.merges.map(function(m){ return {s:{r:m[0],c:m[1]},e:{r:m[2],c:m[3]}}; });
    XLSX.utils.book_append_sheet(wb,ws,sh.name.slice(0,31));
  });
  return XLSX.write(wb,{bookType:"xlsx",type:"array"});
}

return {colLetter:colLetter,colIndex:colIndex,parseNum:parseNum,isDateStr:isDateStr,normDate:normDate,clone:clone,blank:blank,
  isTimeFmt:isTimeFmt,fmtTime:fmtTime,parseTime:parseTime,dateToSerial:dateToSerial,serialToDate:serialToDate,todaySerial:todaySerial,addColumn:addColumn,col:col,
  loadWorkbook:loadWorkbook,isData:isData,eachData:eachData,dataCount:dataCount,anyDateSerial:anyDateSerial,timeNum:timeNum,looksTime:looksTime,loadWorkbookExcelJS:loadWorkbookExcelJS,loadCSV:loadCSV,parseCSV:parseCSV,parseSheet:parseSheet,Calc:Calc,display:display,dataRowNo:dataRowNo,footRowNo:footRowNo,cellAt:cellAt,
  appendRow:appendRow,deleteRow:deleteRow,setCell:setCell,findCol:findCol,computedCols:computedCols,numericCol:numericCol,
  query:query,checks:checks,describe:describe,planOps:planOps,applyPlan:applyPlan,buildWorkbook:buildWorkbook,transformRefs:transformRefs};
})();
if(typeof module!=="undefined") module.exports=ZF;
