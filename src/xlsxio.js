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
    model.sheets.forEach(function(sh){
      if(sh.hr<0) return;
      var ws=wb.getWorksheet(sh.name); if(!ws) return;
      var C0=sh.c0+1, N=sh.ncol, orig=sh.orig||{data:[],foot:[]};
      // Stile der Originalzeilen sichern, bevor überschrieben wird
      var grab=function(r){ var a=[]; for(var k=0;k<N;k++) a.push(cloneStyle(ws.getCell(r,C0+k).style)); return a; };
      var headerStyle=grab(sh.hr+1);
      var dataStyles=orig.data.map(grab), footStyles=orig.foot.map(grab);
      var lastOrig=Math.max(sh.hr+1, orig.data.length?orig.data[orig.data.length-1]:0, orig.foot.length?orig.foot[orig.foot.length-1]:0);
      var heights={}; orig.data.concat(orig.foot).forEach(function(r){ var h=ws.getRow(r).height; if(h) heights[r]=h; });
      var fallback=dataStyles.length?dataStyles[dataStyles.length-1]:headerStyle.map(function(){ return {}; });
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
