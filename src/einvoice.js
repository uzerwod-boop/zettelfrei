/* Zettelfrei – E-Rechnung.
 * Liest XRechnung (UBL) und ZUGFeRD / Factur-X (CII) exakt, ohne KI,
 * erzeugt XRechnung (UBL, Version 3.0) und überträgt Rechnungen in Tabellen.
 * Läuft im Browser und in Node (eigener kleiner XML-Leser, keine Abhängigkeiten).
 */
var ZFE = (function(){
"use strict";

/* ---------- kleiner XML-Leser ---------- */
function decode(s){ return s.replace(/&(lt|gt|amp|quot|apos|#\d+|#x[0-9a-f]+);/gi,function(m,e){
  var l=e.toLowerCase(); if(l==="lt") return "<"; if(l==="gt") return ">"; if(l==="amp") return "&"; if(l==="quot") return "\""; if(l==="apos") return "'";
  return String.fromCodePoint(l[1]==="x"?parseInt(l.slice(2),16):parseInt(l.slice(1),10)); }); }
function parseXml(text){
  text=String(text).replace(/^﻿/,"");
  var root={name:"#root",children:[],attrs:{},text:""}, stack=[root], re=/<!\[CDATA\[([\s\S]*?)\]\]>|<!--[\s\S]*?-->|<\?[\s\S]*?\?>|<!DOCTYPE[^>]*>|<\/\s*([^\s>]+)\s*>|<([^\s>\/]+)((?:\s+[^\s=]+\s*=\s*(?:"[^"]*"|'[^']*'))*)\s*(\/?)>|([^<]+)/g, m;
  while((m=re.exec(text))){
    var top=stack[stack.length-1];
    if(m[1]!=null) top.text+=m[1];
    else if(m[2]){ if(stack.length>1) stack.pop(); }
    else if(m[3]){
      var node={name:m[3].replace(/^.*:/,""),qname:m[3],attrs:{},children:[],text:""}, ar=/([^\s=]+)\s*=\s*("([^"]*)"|'([^']*)')/g, a;
      while((a=ar.exec(m[4]||""))) node.attrs[a[1].replace(/^.*:/,"")]=decode(a[3]!=null?a[3]:a[4]);
      top.children.push(node); if(!m[5]) stack.push(node);
    } else if(m[6]!=null) top.text+=decode(m[6]);
  }
  if(stack.length!==1) throw new Error("XML unvollständig");
  return root.children.find(function(n){ return n.name; })||null;
}
function kids(n,name){ return n?n.children.filter(function(c){ return c.name===name; }):[]; }
function get(n,path){ var cur=n; String(path).split("/").forEach(function(p){ cur=cur?kids(cur,p)[0]:null; }); return cur; }
function all(n,path){ var parts=String(path).split("/"), last=parts.pop(), cur=parts.length?get(n,parts.join("/")):n; return kids(cur,last); }
function txt(n,path){ var x=path?get(n,path):n; return x?String(x.text).trim():""; }
function num(n,path){ var t=txt(n,path); return t===""?null:Number(t); }
function attr(n,path,a){ var x=path?get(n,path):n; return x&&x.attrs[a]!=null?x.attrs[a]:""; }

/* ---------- Einheiten (UN/ECE Rec. 20/21) ---------- */
var UNITS={C62:"Stk",H87:"Stk",EA:"Stk",MTK:"m²",MTR:"m",MTQ:"m³",KGM:"kg",TNE:"t",LTR:"l",HUR:"Std",MIN:"Min",DAY:"Tag",SET:"Set",LS:"pauschal",
  XBG:"Sack",BG:"Sack",XSA:"Sack",XBJ:"Eimer",BJ:"Eimer",XCQ:"Kartusche",XJC:"Kanister",XRO:"Rolle",RO:"Rolle",XPK:"Pack",PK:"Pack",XPA:"Paket",XPX:"Palette",XBX:"Karton",XTU:"Tube",XCA:"Dose",KMT:"km"};
var CODES={}; Object.keys(UNITS).forEach(function(c){ var u=UNITS[c].toLowerCase(); if(!CODES[u]) CODES[u]=c; });
Object.assign(CODES,{"stk":"C62","stück":"C62","st":"C62","m2":"MTK","qm":"MTK","lfm":"MTR","std":"HUR","stunden":"HUR","stunde":"HUR","h":"HUR","säcke":"XBG","eimer":"XBJ","kartuschen":"XCQ","pauschal":"LS","psch":"LS","pausch.":"LS"});
function unitName(code){ return UNITS[code]||code||""; }
function unitCode(name){ var n=String(name||"").trim().toLowerCase(); return CODES[n]||(/^[A-Z0-9]{2,3}$/.test(String(name||""))?name:"C62"); }

/* ---------- Datum ---------- */
function isoToDe(s){ s=String(s||"").trim(); var m=s.match(/^(\d{4})-?(\d{2})-?(\d{2})/); return m?m[3]+"."+m[2]+"."+m[1]:s; }
function deToIso(s){ var m=String(s||"").trim().match(/^(\d{1,2})\.(\d{1,2})\.(\d{4})$/); return m?m[3]+"-"+m[2].padStart(2,"0")+"-"+m[1].padStart(2,"0"):String(s||""); }
function addDays(de,days){ var iso=deToIso(de), d=new Date(iso+"T12:00:00Z"); d.setUTCDate(d.getUTCDate()+days); return isoToDe(d.toISOString().slice(0,10)); }

/* ---------- Lesen ---------- */
function partyUBL(p){
  if(!p) return {};
  return { name: txt(p,"PartyLegalEntity/RegistrationName")||txt(p,"PartyName/Name"),
    street: txt(p,"PostalAddress/StreetName"), zip: txt(p,"PostalAddress/PostalZone"), city: txt(p,"PostalAddress/CityName"),
    vatId: txt(p,"PartyTaxScheme/CompanyID"), email: txt(p,"Contact/ElectronicMail")||txt(p,"EndpointID") };
}
function readUBL(root){
  var credit=root.name==="CreditNote", lineTag=credit?"CreditNoteLine":"InvoiceLine", qtyTag=credit?"CreditedQuantity":"InvoicedQuantity";
  var lines=kids(root,lineTag).map(function(l){
    var q=get(l,qtyTag);
    return { id: txt(l,"ID"), name: txt(l,"Item/Name"), articleNo: txt(l,"Item/SellersItemIdentification/ID")||txt(l,"Item/BuyersItemIdentification/ID"),
      qty: q?Number(q.text):null, unit: unitName(q?q.attrs.unitCode:""), unitCode: q?q.attrs.unitCode:"",
      price: num(l,"Price/PriceAmount"), total: num(l,"LineExtensionAmount"), rate: num(l,"Item/ClassifiedTaxCategory/Percent") };
  });
  var tax=get(root,"TaxTotal");
  return { format:"XRechnung (UBL)", type: credit?"Gutschrift":"Rechnung", number: txt(root,"ID"), date: isoToDe(txt(root,"IssueDate")),
    due: isoToDe(txt(root,"DueDate")||txt(root,"PaymentMeans/PaymentDueDate")), currency: txt(root,"DocumentCurrencyCode")||"EUR",
    buyerRef: txt(root,"BuyerReference"), orderRef: txt(root,"OrderReference/ID"), project: txt(root,"ProjectReference/ID"),
    note: kids(root,"Note").map(function(n){ return txt(n); }).join(" "),
    seller: partyUBL(get(root,"AccountingSupplierParty/Party")), buyer: partyUBL(get(root,"AccountingCustomerParty/Party")),
    lines: lines, net: num(root,"LegalMonetaryTotal/TaxExclusiveAmount"), tax: tax?num(tax,"TaxAmount"):null,
    gross: num(root,"LegalMonetaryTotal/TaxInclusiveAmount"), payable: num(root,"LegalMonetaryTotal/PayableAmount"),
    rate: lines.length?lines[0].rate:null };
}
function partyCII(p){
  if(!p) return {};
  return { name: txt(p,"Name"), street: txt(p,"PostalTradeAddress/LineOne"), zip: txt(p,"PostalTradeAddress/PostcodeCode"), city: txt(p,"PostalTradeAddress/CityName"),
    vatId: (all(p,"SpecifiedTaxRegistration").map(function(r){ return get(r,"ID"); }).find(function(i){ return i&&i.attrs.schemeID==="VA"; })||{text:""}).text.trim(),
    email: txt(p,"URIUniversalCommunication/URIID")||txt(p,"DefinedTradeContact/EmailURIUniversalCommunication/URIID") };
}
function readCII(root){
  var tx=get(root,"SupplyChainTradeTransaction"), agr=get(tx,"ApplicableHeaderTradeAgreement"), set=get(tx,"ApplicableHeaderTradeSettlement"), sum=get(set,"SpecifiedTradeSettlementHeaderMonetarySummation");
  var lines=kids(tx,"IncludedSupplyChainTradeLineItem").map(function(l){
    var q=get(l,"SpecifiedLineTradeDelivery/BilledQuantity");
    return { id: txt(l,"AssociatedDocumentLineDocument/LineID"), name: txt(l,"SpecifiedTradeProduct/Name"), articleNo: txt(l,"SpecifiedTradeProduct/SellerAssignedID")||txt(l,"SpecifiedTradeProduct/BuyerAssignedID"),
      qty: q?Number(q.text):null, unit: unitName(q?q.attrs.unitCode:""), unitCode: q?q.attrs.unitCode:"",
      price: num(l,"SpecifiedLineTradeAgreement/NetPriceProductTradePrice/ChargeAmount"),
      total: num(l,"SpecifiedLineTradeSettlement/SpecifiedTradeSettlementLineMonetarySummation/LineTotalAmount"),
      rate: num(l,"SpecifiedLineTradeSettlement/ApplicableTradeTax/RateApplicablePercent") };
  });
  var typeCode=txt(root,"ExchangedDocument/TypeCode");
  return { format:"ZUGFeRD / Factur-X (CII)", type: typeCode==="381"?"Gutschrift":"Rechnung", number: txt(root,"ExchangedDocument/ID"),
    date: isoToDe(txt(root,"ExchangedDocument/IssueDateTime/DateTimeString")), due: isoToDe(txt(set,"SpecifiedTradePaymentTerms/DueDateDateTime/DateTimeString")),
    currency: txt(set,"InvoiceCurrencyCode")||"EUR", buyerRef: txt(agr,"BuyerReference"), orderRef: txt(agr,"BuyerOrderReferencedDocument/IssuerAssignedID"),
    project: txt(agr,"SpecifiedProcuringProject/Name")||txt(agr,"SpecifiedProcuringProject/ID"),
    note: all(root,"ExchangedDocument/IncludedNote").map(function(n){ return txt(n,"Content"); }).join(" "),
    seller: partyCII(get(agr,"SellerTradeParty")), buyer: partyCII(get(agr,"BuyerTradeParty")), lines: lines,
    net: num(sum,"TaxBasisTotalAmount"), tax: num(sum,"TaxTotalAmount"), gross: num(sum,"GrandTotalAmount"), payable: num(sum,"DuePayableAmount"),
    rate: lines.length?lines[0].rate:null };
}
function read(xmlText){
  var root=parseXml(xmlText);
  if(!root) throw new Error("Keine XML-Daten");
  if(root.name==="Invoice"||root.name==="CreditNote") return readUBL(root);
  if(root.name==="CrossIndustryInvoice") return readCII(root);
  throw new Error("Kein bekanntes E-Rechnungsformat ("+root.name+")");
}
function looksLikeEInvoice(text){ return /<(\w+:)?(Invoice|CreditNote|CrossIndustryInvoice)[\s>]/.test(String(text).slice(0,4000)); }
/* ZUGFeRD: XML-Anhang aus einem PDF holen (pdf.js-Dokument) */
function fromPdf(pdfDoc){
  return pdfDoc.getAttachments().then(function(att){
    if(!att) return null;
    var keys=Object.keys(att), pick=keys.find(function(k){ return /factur-x\.xml|zugferd-invoice\.xml|xrechnung\.xml|zugferd/i.test(k); })||keys.find(function(k){ return /\.xml$/i.test(k); });
    if(!pick) return null;
    var bytes=att[pick].content, s=new TextDecoder("utf-8").decode(bytes);
    return looksLikeEInvoice(s)?{name:pick,xml:s}:null;
  });
}

/* ---------- XRechnung erzeugen (UBL, XRechnung 3.0) ---------- */
function esc(s){ return String(s==null?"":s).replace(/&/g,"&amp;").replace(/</g,"&lt;").replace(/>/g,"&gt;").replace(/"/g,"&quot;"); }
function amt(n){ return (Math.round(Number(n)*100)/100).toFixed(2); }
function r2(n){ return Math.round(Number(n)*100+(n>=0?1e-9:-1e-9))/100; }
function totals(inv){
  var groups={}, net=0;
  inv.lines.forEach(function(l){ var t=r2(l.qty*l.price); l.total=t; net=r2(net+t); var k=String(l.rate); groups[k]=groups[k]||{rate:l.rate,base:0}; groups[k].base=r2(groups[k].base+t); });
  var tax=0; Object.keys(groups).forEach(function(k){ groups[k].tax=r2(groups[k].base*groups[k].rate/100); tax=r2(tax+groups[k].tax); });
  return {net:net,tax:tax,gross:r2(net+tax),groups:Object.keys(groups).map(function(k){ return groups[k]; })};
}
function checkData(inv){
  var miss=[], s=inv.seller||{}, b=inv.buyer||{};
  if(!inv.number) miss.push("Rechnungsnummer");
  if(!s.name) miss.push("Firmenname"); if(!s.street||!s.zip||!s.city) miss.push("Firmenadresse");
  if(!s.vatId && !s.taxNo) miss.push("USt-IdNr. oder Steuernummer");
  if(!s.contactName||!s.phone||!s.email) miss.push("Ansprechpartner, Telefon und E-Mail der Firma");
  if(!s.iban) miss.push("IBAN");
  if(!b.name) miss.push("Name des Kunden"); if(!b.street||!b.zip||!b.city) miss.push("Adresse des Kunden");
  if(!b.email) miss.push("E-Mail des Kunden (elektronische Adresse)");
  if(!inv.lines||!inv.lines.length) miss.push("mindestens eine Position");
  return miss;
}
function buildXRechnung(inv){
  var s=inv.seller, b=inv.buyer, t=totals(inv), cur=inv.currency||"EUR", C='currencyID="'+cur+'"';
  var party=function(p,isSeller){
    return '<cac:Party>'+
      '<cbc:EndpointID schemeID="EM">'+esc(p.email)+'</cbc:EndpointID>'+
      (isSeller&&!p.vatId&&!p.regNo&&p.taxNo?'<cac:PartyIdentification><cbc:ID>'+esc(p.taxNo)+'</cbc:ID></cac:PartyIdentification>':'')+
      '<cac:PartyName><cbc:Name>'+esc(p.name)+'</cbc:Name></cac:PartyName>'+
      '<cac:PostalAddress><cbc:StreetName>'+esc(p.street)+'</cbc:StreetName><cbc:CityName>'+esc(p.city)+'</cbc:CityName><cbc:PostalZone>'+esc(p.zip)+'</cbc:PostalZone>'+
      '<cac:Country><cbc:IdentificationCode>'+esc(p.country||"DE")+'</cbc:IdentificationCode></cac:Country></cac:PostalAddress>'+
      (isSeller&&p.vatId?'<cac:PartyTaxScheme><cbc:CompanyID>'+esc(p.vatId)+'</cbc:CompanyID><cac:TaxScheme><cbc:ID>VAT</cbc:ID></cac:TaxScheme></cac:PartyTaxScheme>':'')+
      (isSeller&&p.taxNo?'<cac:PartyTaxScheme><cbc:CompanyID>'+esc(p.taxNo)+'</cbc:CompanyID><cac:TaxScheme><cbc:ID>FC</cbc:ID></cac:TaxScheme></cac:PartyTaxScheme>':'')+
      '<cac:PartyLegalEntity><cbc:RegistrationName>'+esc(p.name)+'</cbc:RegistrationName>'+(isSeller&&p.regNo?'<cbc:CompanyID>'+esc(p.regNo)+'</cbc:CompanyID>':'')+'</cac:PartyLegalEntity>'+
      (isSeller?'<cac:Contact><cbc:Name>'+esc(p.contactName)+'</cbc:Name><cbc:Telephone>'+esc(p.phone)+'</cbc:Telephone><cbc:ElectronicMail>'+esc(p.email)+'</cbc:ElectronicMail></cac:Contact>':'')+
      '</cac:Party>';
  };
  var x='<?xml version="1.0" encoding="UTF-8"?>\n'+
    '<ubl:Invoice xmlns:ubl="urn:oasis:names:specification:ubl:schema:xsd:Invoice-2" xmlns:cac="urn:oasis:names:specification:ubl:schema:xsd:CommonAggregateComponents-2" xmlns:cbc="urn:oasis:names:specification:ubl:schema:xsd:CommonBasicComponents-2">\n'+
    '<cbc:CustomizationID>urn:cen.eu:en16931:2017#compliant#urn:xeinkauf.de:kosit:xrechnung_3.0</cbc:CustomizationID>\n'+
    '<cbc:ProfileID>urn:fdc:peppol.eu:2017:poacc:billing:01:1.0</cbc:ProfileID>\n'+
    '<cbc:ID>'+esc(inv.number)+'</cbc:ID>\n<cbc:IssueDate>'+deToIso(inv.date)+'</cbc:IssueDate>\n'+
    (inv.due?'<cbc:DueDate>'+deToIso(inv.due)+'</cbc:DueDate>\n':'')+
    '<cbc:InvoiceTypeCode>380</cbc:InvoiceTypeCode>\n'+
    (inv.note?'<cbc:Note>'+esc(inv.note)+'</cbc:Note>\n':'')+
    '<cbc:DocumentCurrencyCode>'+cur+'</cbc:DocumentCurrencyCode>\n'+
    '<cbc:BuyerReference>'+esc(inv.buyerRef||inv.number)+'</cbc:BuyerReference>\n'+
    (inv.project?'<cac:ProjectReference><cbc:ID>'+esc(inv.project)+'</cbc:ID></cac:ProjectReference>\n':'')+
    '<cac:AccountingSupplierParty>'+party(s,true)+'</cac:AccountingSupplierParty>\n'+
    '<cac:AccountingCustomerParty>'+party(b,false)+'</cac:AccountingCustomerParty>\n'+
    (inv.deliveryDate?'<cac:Delivery><cbc:ActualDeliveryDate>'+deToIso(inv.deliveryDate)+'</cbc:ActualDeliveryDate></cac:Delivery>\n':'')+
    '<cac:PaymentMeans><cbc:PaymentMeansCode>58</cbc:PaymentMeansCode>'+
      '<cac:PayeeFinancialAccount><cbc:ID>'+esc(String(s.iban).replace(/\s/g,""))+'</cbc:ID>'+(s.bank?'<cbc:Name>'+esc(s.bank)+'</cbc:Name>':'')+
      (s.bic?'<cac:FinancialInstitutionBranch><cbc:ID>'+esc(s.bic)+'</cbc:ID></cac:FinancialInstitutionBranch>':'')+'</cac:PayeeFinancialAccount></cac:PaymentMeans>\n'+
    '<cac:PaymentTerms><cbc:Note>'+esc(inv.paymentTerms||("Zahlbar bis "+(inv.due||"")+" ohne Abzug."))+'</cbc:Note></cac:PaymentTerms>\n'+
    '<cac:TaxTotal><cbc:TaxAmount '+C+'>'+amt(t.tax)+'</cbc:TaxAmount>'+
      t.groups.map(function(g){ return '<cac:TaxSubtotal><cbc:TaxableAmount '+C+'>'+amt(g.base)+'</cbc:TaxableAmount><cbc:TaxAmount '+C+'>'+amt(g.tax)+'</cbc:TaxAmount>'+
        '<cac:TaxCategory><cbc:ID>S</cbc:ID><cbc:Percent>'+g.rate+'</cbc:Percent><cac:TaxScheme><cbc:ID>VAT</cbc:ID></cac:TaxScheme></cac:TaxCategory></cac:TaxSubtotal>'; }).join("")+
    '</cac:TaxTotal>\n'+
    '<cac:LegalMonetaryTotal><cbc:LineExtensionAmount '+C+'>'+amt(t.net)+'</cbc:LineExtensionAmount><cbc:TaxExclusiveAmount '+C+'>'+amt(t.net)+'</cbc:TaxExclusiveAmount>'+
      '<cbc:TaxInclusiveAmount '+C+'>'+amt(t.gross)+'</cbc:TaxInclusiveAmount><cbc:PayableAmount '+C+'>'+amt(t.gross)+'</cbc:PayableAmount></cac:LegalMonetaryTotal>\n'+
    inv.lines.map(function(l,i){
      return '<cac:InvoiceLine><cbc:ID>'+(i+1)+'</cbc:ID><cbc:InvoicedQuantity unitCode="'+esc(l.unitCode||unitCode(l.unit))+'">'+Number(l.qty)+'</cbc:InvoicedQuantity>'+
        '<cbc:LineExtensionAmount '+C+'>'+amt(l.total)+'</cbc:LineExtensionAmount>'+
        '<cac:Item>'+(l.description?'<cbc:Description>'+esc(l.description)+'</cbc:Description>':'')+'<cbc:Name>'+esc(l.name)+'</cbc:Name>'+
        (l.articleNo?'<cac:SellersItemIdentification><cbc:ID>'+esc(l.articleNo)+'</cbc:ID></cac:SellersItemIdentification>':'')+
        '<cac:ClassifiedTaxCategory><cbc:ID>S</cbc:ID><cbc:Percent>'+l.rate+'</cbc:Percent><cac:TaxScheme><cbc:ID>VAT</cbc:ID></cac:TaxScheme></cac:ClassifiedTaxCategory></cac:Item>'+
        '<cac:Price><cbc:PriceAmount '+C+'>'+amt(l.price)+'</cbc:PriceAmount></cac:Price></cac:InvoiceLine>\n';
    }).join("")+
    '</ubl:Invoice>\n';
  return {xml:x, totals:t};
}

/* ---------- In eine Tabelle übertragen ---------- */
function norm(s){ return String(s||"").toLowerCase().replace(/[^a-z0-9äöüß]+/g," ").trim(); }
function similar(a,b){
  var A=norm(a).split(" ").filter(Boolean), B=norm(b).split(" ").filter(Boolean);
  if(!A.length||!B.length) return 0;
  // widersprüchliche Merkmale (z. B. weiß / grau): beide Seiten haben eigene Wörter mit mind. 4 Buchstaben
  var onlyA=A.filter(function(w){ return B.indexOf(w)<0 && /^[a-zäöüß]{4,}$/.test(w); }), onlyB=B.filter(function(w){ return A.indexOf(w)<0 && /^[a-zäöüß]{4,}$/.test(w); });
  if(onlyA.length && onlyB.length) return 0;
  var hit=A.filter(function(w){ return B.indexOf(w)>=0; }).length; return hit/Math.max(A.length,B.length);
}
function colOf(ZF,sh,names){ for(var i=0;i<names.length;i++){ var k=ZF.findCol(sh,names[i]); if(k>=0) return sh.headers[k]; } return null; }
/* Liefert Änderungen im Format des Assistenten ({typ, blatt, werte, zeile}) */
function toOps(ZF, model, sheet, inv){
  var sh=sheet, calc=new ZF.Calc(model), ops=[], H=function(n){ return colOf(ZF,sh,n); };
  var hBestand=H(["Bestand","Lagerbestand","Menge im Lager"]), hArtikel=H(["Artikel","Bezeichnung","Artikelname"]), hNr=H(["Artikel-Nr","Artikelnummer","Art.-Nr"]);
  var hRechNr=H(["Rechnungs-Nr","Rechnungsnummer","Beleg-Nr","Belegnummer","Re-Nr"]), hNetto=H(["Netto €","Netto","Nettobetrag"]), hMenge=H(["Menge"]);
  if(hBestand && hArtikel){ // Lager: Bestand erhöhen oder Artikel anlegen
    var kA=ZF.findCol(sh,hArtikel), kN=hNr?ZF.findCol(sh,hNr):-1;
    inv.lines.forEach(function(l){
      var best=null, score=0;
      sh.rows.forEach(function(r,i){ var row=ZF.dataRowNo(sh,i);
        if(kN>=0 && l.articleNo && String(calc.value(sh,row,sh.c0+kN)).trim()===l.articleNo){ best=row; score=2; return; }
        var sc=similar(l.name,calc.value(sh,row,sh.c0+kA)); if(sc>score){ score=sc; best=row; } });
      if(best && score>=0.6){ var w={}; w[hBestand]={plus:l.qty}; ops.push({typ:"aendern",blatt:sh.name,zeile:best,werte:w}); }
      else { var v={}; v[hArtikel]=l.name; v[hBestand]=l.qty;
        var hE=H(["Einheit"]), hP=H(["EK-Preis €","EK-Preis","Einkaufspreis","Preis"]), hL=H(["Lieferant"]);
        if(hE) v[hE]=l.unit; if(hP) v[hP]=l.price; if(hL) v[hL]=inv.seller.name; if(hNr&&l.articleNo) v[hNr]=l.articleNo;
        ops.push({typ:"neu",blatt:sh.name,werte:v}); }
    });
    return ops;
  }
  var base=function(){ var v={};
    var set=function(names,val){ var h=H(names); if(h && val!=null && val!=="") v[h]=val; };
    set(["Rechnung am","Rechnungsdatum","Datum","Belegdatum"],inv.date); set(["Lieferant","Kreditor","Firma"],inv.seller.name);
    set(["Rechnungs-Nr","Rechnungsnummer","Beleg-Nr","Belegnummer","Re-Nr"],inv.number); set(["Fällig am","Fälligkeit","Zahlungsziel"],inv.due);
    set(["Projekt","Kommission","Baustelle"],inv.project||inv.buyerRef||inv.orderRef);
    return v; };
  if(hRechNr && hNetto && !hMenge){ // eine Zeile pro Rechnung
    var v=base(); v[hNetto]=inv.net;
    var hSatz=H(["MwSt-Satz","USt-Satz","Steuersatz"]), hTax=H(["MwSt €","USt €","MwSt","USt"]), hBr=H(["Brutto €","Brutto","Bruttobetrag"]);
    if(hSatz) v[hSatz]=inv.rate!=null?inv.rate/100:null; if(hTax && hTax!==hSatz) v[hTax]=inv.tax; if(hBr) v[hBr]=inv.gross;
    ops.push({typ:"neu",blatt:sh.name,werte:v}); return ops;
  }
  inv.lines.forEach(function(l){ // eine Zeile pro Position
    var v=base(), set=function(names,val){ var h=H(names); if(h && val!=null && val!=="") v[h]=val; };
    set(["Position","Artikel","Bezeichnung","Leistung"],l.name); set(["Menge","Anzahl"],l.qty); set(["Einheit"],l.unit);
    set(["€ / Einh.","Einzelpreis","EK-Preis €","Preis"],l.price); set(["Summe €","Summe","Gesamt","Netto €","Netto","Betrag"],l.total);
    ops.push({typ:"neu",blatt:sh.name,werte:v});
  });
  return ops;
}

return {parseXml:parseXml, read:read, fromPdf:fromPdf, looksLikeEInvoice:looksLikeEInvoice, buildXRechnung:buildXRechnung, totals:totals,
  checkData:checkData, toOps:toOps, unitCode:unitCode, unitName:unitName, isoToDe:isoToDe, deToIso:deToIso, addDays:addDays};
})();
if(typeof module!=="undefined") module.exports=ZFE;
