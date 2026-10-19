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
  if(c.t==="n"){ if(c.z && XLSX.SSF.is_date(c.z)) { o.v=serialToDate(c.v); o.d=1; } else o.v=c.v; }
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
function parseSheet(XLSX, ws, name){
  var ref=ws["!ref"]; var grid=[];
  if(ref){ var rg=XLSX.utils.decode_range(ref);
    for(var r=0;r<=rg.e.r;r++){ var row=[]; for(var c=0;c<=rg.e.c;c++){ row.push(cellFrom(ws[XLSX.utils.encode_cell({r:r,c:c})],XLSX)); } grid.push(row); } }
  var widths=(ws["!cols"]||[]).map(function(c){ return c&&(c.wch||c.width)?Math.round(c.wch||c.width):null; });
  var merges=(ws["!merges"]||[]).map(function(m){ return [m.s.r,m.s.c,m.e.r,m.e.c]; });
  // Kopfzeile suchen
  var hr=-1, best=0;
  var score=function(row){ return row?row.filter(function(c){ return c && !c.f && typeof c.v==="string" && c.v.trim() && !isDateStr(c.v); }).length:0; };
  var filled=function(row){ return row?row.filter(function(c){ return !blank(c); }).length:0; };
  for(var i=0;i<Math.min(grid.length,20);i++) best=Math.max(best,score(grid[i]));
  if(best>=2){ for(i=0;i<Math.min(grid.length,20);i++){ var s=score(grid[i]); if(s>=2 && s>=0.7*best && (i+1>=grid.length || filled(grid[i+1])>=1)){ hr=i; break; } } }
  if(hr<0) return {name:name, hr:-1, top:grid, headers:[], rows:[], foot:[], c0:0, ncol:0, widths:widths, merges:merges, colFmt:[], dateCol:[]};
  var hrow=grid[hr], c0=hrow.findIndex(function(c){ return !blank(c); }), cEnd=hrow.length-1;
  while(cEnd>c0 && blank(hrow[cEnd])) cEnd--;
  var ncol=cEnd-c0+1;
  var headers=[]; for(var k=0;k<ncol;k++){ var hc=hrow[c0+k]; headers.push(hc&&hc.v!=null?String(hc.v).trim():colLetter(c0+k)); }
  // Datenzeilen / Fußzeilen, Leerzeilen weglassen, Zeilennummern neu vergeben
  var rows=[], foot=[], oldRows=[], footOld=[], inFoot=false;
  for(r=hr+1;r<grid.length;r++){
    var g=grid[r]||[]; var part=[]; for(k=0;k<ncol;k++) part.push(g[c0+k]||null);
    if(part.every(blank)) continue;
    var first=part.find(function(c){ return !blank(c); });
    if(!inFoot && first && typeof first.v==="string" && /^(summe|gesamt|total|zwischensumme|ergebnis)\b/i.test(first.v.trim()) ) inFoot=true;
    if(inFoot){ foot.push(part); footOld.push(r+1); } else { rows.push(part); oldRows.push(r+1); }
  }
  var keep=oldRows.concat(footOld); // alte Excel-Zeilen (1-basiert) in neuer Reihenfolge
  var newOf={}; keep.forEach(function(o,idx){ newOf[o]=hr+2+idx; });
  for(r=1;r<=hr+1;r++) newOf[r]=r;
  var mapRow=function(row,part){
    if(newOf[row]!=null) return newOf[row];
    if(part==="end"){ for(var x=row;x>=1;x--) if(newOf[x]!=null) return newOf[x]; return null; }
    for(var y=row;y<=row+5000;y++) if(newOf[y]!=null) return newOf[y];
    return part==="start"?null:null;
  };
  var remap=function(cells){ cells.forEach(function(c){ if(c&&c.f) c.f=transformRefs(c.f,function(row,col,part){ return mapRow(row,part); }); }); };
  rows.forEach(remap); foot.forEach(remap); grid.slice(0,hr).forEach(remap);
  // Spaltentypen
  var colFmt=[], dateCol=[];
  for(k=0;k<ncol;k++){
    var z=null, d=0, n=0;
    rows.forEach(function(rw){ var c=rw[k]; if(c&&!blank(c)){ n++; if(c.d) d++; if(!z&&c.z) z=c.z; } });
    colFmt.push(z); dateCol.push(n>0 && d/n>=0.6);
  }
  return {name:name, hr:hr, top:grid.slice(0,hr), headers:headers, rows:rows, foot:foot, c0:c0, ncol:ncol, widths:widths, merges:merges.filter(function(m){ return m[2]<hr; }), colFmt:colFmt, dateCol:dateCol};
}

/* ---------- Adressen ---------- */
function dataRowNo(sh,i){ return sh.hr+2+i; }
function footRowNo(sh,j){ return sh.hr+2+sh.rows.length+j; }
function cellAt(sh,r,c){ // r 1-basiert, c absolute Spalte
  if(sh.hr<0){ var g=sh.top[r-1]; return g?g[c]||null:null; }
  if(r<=sh.hr){ var t=sh.top[r-1]; return t?t[c]||null:null; }
  var k=c-sh.c0; if(k<0||k>=sh.ncol) return null;
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
Calc.prototype.formula=function(sh,f){
  var self=this;
  try{
    var src=f.replace(/^=/,"");
    if(/#BEZUG!/.test(src)) return "#BEZUG!";
    // Blatt-Präfix + Bereich/Zelle
    var REFX=/(?<![A-Za-z0-9_.$ÄÖÜäöüß])(?:(?:'([^']+)'|([A-Za-zÄÖÜäöüß_][A-Za-z0-9ÄÖÜäöüß_]*))!)?\$?([A-Z]{1,3})\$?(\d+)(?::\$?([A-Z]{1,3})\$?(\d+))?(?![\d(A-Za-z_])/g;
    var parts=splitStrings(src), out="", strOnly=null;
    var single=src.match(/^\s*(?:'([^']+)'!|([A-Za-zÄÖÜäöüß_][\wÄÖÜäöüß]*)!)?\$?([A-Z]{1,3})\$?(\d+)\s*$/);
    if(single){ var s0=single[1]||single[2]?self.sheet(single[1]||single[2]):sh; if(!s0) return "#BEZUG!"; return self.value(s0,+single[4],colIndex(single[3])); }
    for(var p=0;p<parts.length;p++){
      if(parts[p].str){ strOnly=parts[p].t; out+="0"; continue; }
      out+=parts[p].t.replace(REFX,function(all,q,nm,c,r,c2,r2){
        var tsh=(q||nm)?self.sheet(q||nm):sh; if(!tsh) throw new Error("ref");
        if(!c2) { var v=self.num(self.value(tsh,+r,colIndex(c))); if(isNaN(v)) throw new Error("text"); return "("+v+")"; }
        var vals=[], a=Math.min(+r,+r2), b=Math.max(+r,+r2), ca=Math.min(colIndex(c),colIndex(c2)), cb=Math.max(colIndex(c),colIndex(c2));
        for(var rr=a;rr<=b;rr++) for(var cc=ca;cc<=cb;cc++){ var x=self.value(tsh,rr,cc); if(typeof x==="number") vals.push(x); else if(typeof x==="string" && x!=="" && !isDateStr(x)){ var pn=parseNum(x); if(pn!=null) vals.push(pn); } }
        return "["+vals.join(",")+"]";
      });
    }
    var expr=out.replace(/;/g,",").replace(/\^/g,"**").replace(/(\d)%/g,"($1/100)")
      .replace(/\b(SUMME|SUM)\(/gi,"_s(").replace(/\b(MITTELWERT|AVERAGE)\(/gi,"_av(").replace(/\bMIN\(/gi,"_mi(").replace(/\bMAX\(/gi,"_ma(")
      .replace(/\b(RUNDEN|ROUND)\(/gi,"_r(").replace(/\bABS\(/gi,"_ab(").replace(/\b(ANZAHL|COUNT)\(/gi,"_n(");
    var check=expr.replace(/_(s|av|mi|ma|r|ab|n)\(/g,"(");
    if(!/^[\d\s.,+\-*/()\[\]eE]*$/.test(check)) return undefined;
    var flat=function(a){ var o=[]; a.forEach(function(x){ if(Array.isArray(x)) o.push.apply(o,x); else o.push(x); }); return o; };
    var fn=new Function("_s","_av","_mi","_ma","_r","_ab","_n","return ("+expr+");");
    var res=fn(
      function(){ return flat([].slice.call(arguments)).reduce(function(a,b){ return a+b; },0); },
      function(){ var a=flat([].slice.call(arguments)); return a.length?a.reduce(function(x,y){ return x+y; },0)/a.length:0; },
      function(){ var a=flat([].slice.call(arguments)); return a.length?Math.min.apply(null,a):0; },
      function(){ var a=flat([].slice.call(arguments)); return a.length?Math.max.apply(null,a):0; },
      function(x,d){ var p=Math.pow(10,d||0); return Math.round((Array.isArray(x)?x[0]:x)*p+(x>=0?1e-9:-1e-9))/p; },
      function(x){ return Math.abs(Array.isArray(x)?x[0]:x); },
      function(){ return flat([].slice.call(arguments)).length; });
    if(Array.isArray(res)) res=res[0];
    if(typeof res!=="number"||!isFinite(res)) return "#DIV/0!";
    return Math.round(res*1e9)/1e9;
  }catch(e){ return undefined; }
};
function display(model, sh, r, c, calc){
  calc=calc||new Calc(model);
  return calc.value(sh,r,c);
}

/* ---------- Zeilen einfügen / löschen ---------- */
function shiftSheet(sh, fn){ eachFormulaCell(sh,function(c){ c.f=transformRefs(c.f,fn); }); }
function templateRow(sh){
  var tpl=[]; for(var k=0;k<sh.ncol;k++){ tpl.push(null);
    for(var i=sh.rows.length-1;i>=0;i--){ var c=sh.rows[i][k]; if(c&&c.f){ tpl[k]={f:c.f,from:dataRowNo(sh,i)}; break; } if(c&&!blank(c)) break; } }
  return tpl;
}
function computedCols(sh){
  var out=[]; for(var k=0;k<sh.ncol;k++){ var nf=0,n=0; sh.rows.forEach(function(r){ var c=r[k]; if(c&&!blank(c)){ n++; if(c.f) nf++; } }); out.push(n>0 && nf/n>=0.6); } return out;
}
function appendRow(sh, values){
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
  sh.rows.push(cells);
  return newRow;
}
function deleteRow(sh, excelRow){
  var i=excelRow-sh.hr-2; if(i<0||i>=sh.rows.length) return false;
  sh.rows.splice(i,1);
  shiftSheet(sh,function(row,col,part){
    if(row<excelRow) return row;
    if(row===excelRow){ if(part==="single") return null; if(part==="start") return row; return row-1; }
    return row-1;
  });
  return true;
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
function numericCol(sh,k){ var n=0,num=0; sh.rows.forEach(function(r){ var c=r[k]; if(c&&!blank(c)&&!c.f){ n++; if(typeof c.v==="number") num++; } }); return n>0&&num/n>=0.6; }
function coerce(sh,k,val){
  if(val==null||val==="") return null;
  if(sh.dateCol[k]){ var d=normDate(val); if(d) return d; }
  if(numericCol(sh,k)){ var n=typeof val==="number"?val:parseNum(val); if(n!=null) return n; }
  if(typeof val==="number") return val;
  return String(val);
}
function setCell(sh, excelRow, k, raw){
  var i=excelRow-sh.hr-2, row;
  if(i>=0&&i<sh.rows.length) row=sh.rows[i]; else { var j=i-sh.rows.length; row=sh.foot[j]; }
  if(!row) return;
  if(typeof raw==="string" && raw.trim().charAt(0)==="=") row[k]={f:raw.trim().slice(1).toUpperCase().replace(/"[^"]*"/g,function(s){ return s; })};
  else { var v=coerce(sh,k,raw); row[k]= v==null?null:{v:v}; }
}

/* ---------- Abfragen (exakt gerechnet) ---------- */
function rowValues(model, sh, i, calc){ var r=dataRowNo(sh,i), o=[]; for(var k=0;k<sh.ncol;k++) o.push(calc.value(sh,r,sh.c0+k)); return o; }
function cmp(a,b){ var da=isDateStr(a)?dateToSerial(a):null, db=isDateStr(b)?dateToSerial(normDate(b)||b):null;
  if(da!=null&&db!=null) return da-db;
  var na=typeof a==="number"?a:parseNum(a), nb=typeof b==="number"?b:parseNum(b);
  if(na!=null&&nb!=null) return na-nb;
  return String(a==null?"":a).toLowerCase().localeCompare(String(b==null?"":b).toLowerCase(),"de"); }
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
  sh.rows.forEach(function(r,i){ var vals=rowValues(model,sh,i,calc); if(filters.every(function(f){ return matches(vals[f.k],f.op,f.w); })) hits.push({i:i,vals:vals}); });
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
function checks(model){
  var calc=new Calc(model), out=[];
  model.sheets.forEach(function(sh){
    if(sh.hr<0||!sh.rows.length) return;
    var comp=computedCols(sh);
    for(var k=0;k<sh.ncol;k++){
      var n=0, filledC=0, nums=0; sh.rows.forEach(function(r){ n++; var c=r[k]; if(c&&!blank(c)){ filledC++; if(typeof c.v==="number"||c.f) nums++; } });
      var key=/(^|[^a-z])(nr|nummer|id|artikel-?nr|auftrags-?nr|rechnungs-?nr)([^a-z]|$)/i.test(sh.headers[k]);
      var seen={};
      sh.rows.forEach(function(r,i){
        var row=dataRowNo(sh,i), c=r[k], v=calc.value(sh,row,sh.c0+k);
        if(typeof v==="string" && /^#/.test(v)) out.push({blatt:sh.name,zeile:row,k:k,problem:"Formel ergibt "+v});
        if(!comp[k] && filledC/n>=0.75 && filledC<n && blank(c)) out.push({blatt:sh.name,zeile:row,k:k,problem:sh.headers[k]+" fehlt"});
        if(!comp[k] && nums/Math.max(1,filledC)>=0.8 && c && !c.f && typeof c.v==="string" && c.v.trim() && parseNum(c.v)==null && !sh.dateCol[k]) out.push({blatt:sh.name,zeile:row,k:k,problem:"Text statt Zahl"});
        if(key && v!=null && v!==""){ var s=String(v).trim().toLowerCase(); if(seen[s]) out.push({blatt:sh.name,zeile:row,k:k,problem:sh.headers[k]+" doppelt (auch Zeile "+seen[s]+")"}); else seen[s]=row; }
      });
    }
  });
  return out;
}

/* ---------- Für die KI beschreiben ---------- */
function describe(model, maxChars){
  maxChars=maxChars||120000; var calc=new Calc(model), out=["Datei: "+model.name];
  model.sheets.forEach(function(sh){
    if(sh.hr<0){ out.push("\nBlatt „"+sh.name+"“: keine Tabelle erkannt."); return; }
    var comp=computedCols(sh), tpl=templateRow(sh);
    out.push("\nBlatt „"+sh.name+"“ — Kopfzeile in Excel-Zeile "+(sh.hr+1)+", Daten in Zeilen "+dataRowNo(sh,0)+"–"+dataRowNo(sh,sh.rows.length-1)+".");
    out.push("Spalten: "+sh.headers.map(function(h,k){ return colLetter(sh.c0+k)+" „"+h+"“ ("+(comp[k]?"berechnet, nicht setzen":sh.dateCol[k]?"Datum":numericCol(sh,k)?"Zahl":"Text")+")"; }).join("; "));
    out.push("Zeile;"+sh.headers.join(";"));
    sh.rows.forEach(function(r,i){ var row=dataRowNo(sh,i); out.push(row+";"+rowValues(model,sh,i,calc).map(function(v){ return v==null?"":String(v).replace(/;/g,","); }).join(";")); });
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
      if(row!=null){ var i=row-sh.hr-2; if(!(i>=0&&i<sh.rows.length)){ out.push({op:op,error:"Zeile "+op.zeile+" gibt es nicht"}); return; } }
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
    if(p.typ==="neu"){ var vals=[]; p.changes.forEach(function(c){ if(c.k>=0) vals[c.k]=c.nv; }); var r=appendRow(sh,vals); touched.push({sheet:sh.name,row:r}); }
    else { p.changes.forEach(function(c){ if(c.k>=0) setCell(sh,p.row,c.k,c.nv); }); touched.push({sheet:sh.name,row:p.row}); }
  });
  plan.filter(function(p){ return !p.error && p.typ==="loeschen"; }).sort(function(a,b){ return b.row-a.row; }).forEach(function(p){
    var sh=model.sheets.find(function(s){ return s.name===p.sheet; }); deleteRow(sh,p.row);
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
  loadWorkbook:loadWorkbook,parseSheet:parseSheet,Calc:Calc,display:display,dataRowNo:dataRowNo,footRowNo:footRowNo,cellAt:cellAt,
  appendRow:appendRow,deleteRow:deleteRow,setCell:setCell,findCol:findCol,computedCols:computedCols,numericCol:numericCol,
  query:query,checks:checks,describe:describe,planOps:planOps,applyPlan:applyPlan,buildWorkbook:buildWorkbook,transformRefs:transformRefs};
})();
if(typeof module!=="undefined") module.exports=ZF;
