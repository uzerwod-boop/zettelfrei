/* Zettelfrei – Speichern mit Formatierung.
 * Öffnet die Originaldatei mit ExcelJS, schreibt nur die erkannten Tabellen neu
 * und übernimmt Farben, Schriften, Rahmen und Zahlenformate der Originalzeilen.
 * Neue Zeilen bekommen das Format der letzten Datenzeile.
 */
var ZFX = (function(){
"use strict";
function cloneStyle(st){ return st?JSON.parse(JSON.stringify(st)):{}; }
function serialToDate(n){ return new Date(Date.UTC(1899,11,30)+Math.round(n*86400000)); }
function deToDate(s){ var m=String(s).match(/^(\d{1,2})\.(\d{1,2})\.(\d{4})$/); return m?new Date(Date.UTC(+m[3],+m[2]-1,+m[1])):null; }

function save(ExcelJS, ZF, origBytes, model, opts){
  opts=opts||{};
  var wb=new ExcelJS.Workbook(), calc=new ZF.Calc(model);
  return wb.xlsx.load(origBytes).then(function(){
    // Blätter, die in Zettelfrei gelöscht oder umbenannt wurden
    (model.deleted||[]).forEach(function(n){ var w=wb.getWorksheet(n); if(w) wb.removeWorksheet(w.id); });
    var esc=function(x){ return String(x).replace(/[.*+?^${}()|[\]\\]/g,"\\$&"); };
    model.sheets.forEach(function(sh){
      if(!sh.origName||sh.origName===sh.name) return; var w=wb.getWorksheet(sh.origName); if(!w) return;
      var re=new RegExp("('"+esc(sh.origName).replace(/'/g,"''")+"'|\\b"+esc(sh.origName)+")!","g"), q=/^[A-Za-zÄÖÜäöüß_][\wÄÖÜäöüß.]*$/.test(sh.name)?sh.name:"'"+sh.name.replace(/'/g,"''")+"'";
      w.name=sh.name;
      wb.eachSheet(function(o){ o.eachRow(function(row){ row.eachCell(function(c){ var v=c.value; if(v&&typeof v==="object"&&typeof v.formula==="string"&&re.test(v.formula)){ re.lastIndex=0; c.value={formula:v.formula.replace(re,q+"!"),result:v.result}; } re.lastIndex=0; }); }); });
    });
    var wsOf=function(sh){ return wb.getWorksheet(sh.name); };
    // Formeln außerhalb der Tabellen (Titelzeilen, Übersichtsblätter) aktualisieren, z. B. verschobene Bezüge auf andere Blätter
    model.sheets.forEach(function(sh){
      var ws=wsOf(sh); if(!ws) return;
      (sh.top||[]).forEach(function(row,ri){ (row||[]).forEach(function(cell,ci){
        if(!cell||!cell.f) return; var xc=ws.getCell(ri+1,ci+1), cur=xc.value&&xc.value.formula;
        if(cur!==undefined&&String(cur).replace(/^=/,"").toUpperCase()===String(cell.f).toUpperCase()) return;
        var v=calc.value(sh,ri+1,ci); xc.value={formula:cell.f,result:typeof v==="number"?v:(v==null?undefined:String(v))}; }); });
    });
    model.sheets.forEach(function(sh){
      if(sh.hr<0) return;
      var ws=wsOf(sh);
      if(!ws){ if(!sh.isNew) return;   // neue Auswertungsblätter anlegen
        ws=wb.addWorksheet(sh.name);
        sh.headers.forEach(function(h,k){ var c=ws.getCell(sh.hr+1,sh.c0+1+k); c.value=h; c.font={bold:true}; c.fill={type:"pattern",pattern:"solid",fgColor:{argb:"FFDDE7E3"}}; ws.getColumn(sh.c0+1+k).width=(sh.widths&&sh.widths[k])||16; }); }
      var C0=sh.c0+1, N=sh.ncol, orig=sh.orig||{data:[],foot:[]};
      // Stile der Originalzeilen sichern, bevor überschrieben wird
      var grab=function(r){ var a=[]; for(var k=0;k<N;k++) a.push(cloneStyle(ws.getCell(r,C0+k).style)); return a; };
      var headerStyle=grab(sh.hr+1);
      // Stil je Zeile von ihrer Herkunftszeile (bleibt beim Einfügen, Löschen und Sortieren an der Zeile); neue Zeilen erben vom Nachbarn
      var known=orig.data.filter(function(r){ return r; }).concat(orig.foot.filter(function(r){ return r; }));
      var cache={}, styleOf=function(r){ return cache[r]||(cache[r]=grab(r)); };
      var footStyles=orig.foot.map(function(r){ return r?styleOf(r):null; });
      var lastOrig=Math.max.apply(null,[sh.hr+1].concat(known)), maxOrig=sh.orig&&sh.orig.maxRow||0; if(maxOrig>lastOrig) lastOrig=maxOrig;
      var heights={}; known.forEach(function(r){ var h=ws.getRow(r).height; if(h) heights[r]=h; });
      var plain=headerStyle.map(function(){ return {}; }), dataStyles=[], prev=null;
      sh.rows.forEach(function(row,i){ var o=orig.data[i]; if(o){ prev=styleOf(o); dataStyles.push(prev); } else dataStyles.push(null); });
      for(var i2=0;i2<dataStyles.length;i2++){ if(dataStyles[i2]) continue; var near=null;
        for(var a2=i2-1;a2>=0&&!near;a2--) if(dataStyles[a2]&&(!sh.kind||!sh.kind[a2])) near=dataStyles[a2];
        for(var b2=i2+1;b2<dataStyles.length&&!near;b2++) if(dataStyles[b2]&&orig.data[b2]&&(!sh.kind||!sh.kind[b2])) near=dataStyles[b2];
        dataStyles[i2]=near||plain; }
      var fallback=dataStyles.length?dataStyles[dataStyles.length-1]:plain;
      // neue Spalten (z. B. „Mahnstufe“) bekommen den Stil der Kopfzeile davor
      for(var k=0;k<N;k++){ var hc=ws.getCell(sh.hr+1,C0+k); if(String(hc.value||"")!==String(sh.headers[k])){ hc.value=sh.headers[k]; if(!hc.style||!hc.style.font) hc.style=cloneStyle(headerStyle[Math.max(0,k-1)]); } }
      var clearRow=function(r){ for(var k=0;k<N;k++){ var c=ws.getCell(r,C0+k); c.value=null; c.style={}; } };
      var writeRow=function(r,cells,styles){
        for(var k=0;k<N;k++){
          var cell=ws.getCell(r,C0+k), src=cells[k], st=cloneStyle(styles[k]||styles[Math.min(k,styles.length-1)]||{});
          var fmt=sh.colFmt[k];
          if(!st.numFmt && fmt) st.numFmt=fmt;
          if(src && src.f){
            var v=calc.value(sh,r,sh.c0+k);
            cell.value={formula:src.f, result: typeof v==="number"?v:(v==null?undefined:String(v))};
          } else if(!src || src.v==null || src.v===""){ cell.value=null; }
          else if(sh.dateCol[k] && typeof src.v==="string" && deToDate(src.v)){ cell.value=deToDate(src.v); if(!st.numFmt) st.numFmt="dd.mm.yyyy"; }
          else if(src.d && typeof src.v==="string" && deToDate(src.v)){ cell.value=deToDate(src.v); }
          else cell.value=src.v;
          cell.style=st;
        }
      };
      var total=sh.rows.length+sh.foot.length, newLast=sh.hr+1+total;
      sh.rows.forEach(function(row,i){ var r=ZF.dataRowNo(sh,i); writeRow(r,row,dataStyles[i]||fallback); });
      sh.foot.forEach(function(row,j){ var r=ZF.footRowNo(sh,j); writeRow(r,row,footStyles[j]||footStyles[footStyles.length-1]||fallback);
        var oh=heights[orig.foot[j]]; if(oh) ws.getRow(r).height=oh; });
      for(var r=newLast+1;r<=lastOrig;r++) clearRow(r);
      // gelöschte letzte Spalten leeren
      ((model.cutCols||{})[sh.origName||sh.name]||[]).forEach(function(col){ for(var r2=sh.hr+1;r2<=Math.max(newLast,lastOrig);r2++){ var c=ws.getCell(r2,col); c.value=null; c.style={}; } });
    });
    if(opts.protokoll && opts.protokoll.length){
      var old=wb.getWorksheet("Änderungsprotokoll"); if(old) wb.removeWorksheet(old.id);
      var ps=wb.addWorksheet("Änderungsprotokoll");
      ps.columns=[{header:"Zeitpunkt",width:19},{header:"Bearbeiter",width:16},{header:"Aktion",width:48},{header:"Details",width:70}];
      ps.getRow(1).font={bold:true};
      opts.protokoll.forEach(function(e){ ps.addRow([e.zeit,e.wer||"",e.aktion,(e.details||[]).join("; ")]); });
    }
    return wb.xlsx.writeBuffer();
  });
}
return {save:save};
})();
if(typeof module!=="undefined") module.exports=ZFX;
