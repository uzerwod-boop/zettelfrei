// Erzeugt die Beispiel-E-Rechnung als XRechnung (UBL) mit dem eigenen Generator.
const fs = require("fs"), path = require("path"), ZFE = require("../src/einvoice.js");
const out = path.join(__dirname, "..", "beispiele");
const weber = {
  number: "88213", date: "15.10.2026", due: "29.10.2026", deliveryDate: "14.10.2026", buyerRef: "Küche Yilmaz", project: "Küche Yilmaz",
  seller: { name: "Fliesen-Großhandel Weber KG", street: "Industriestr. 14", zip: "04129", city: "Leipzig", vatId: "DE287654321",
    contactName: "Sabine Weber", phone: "0341 5567-0", email: "rechnung@weber-fliesen.example", iban: "DE02120300000000202051", bic: "BYLADEM1001", bank: "Deutsche Kreditbank" },
  buyer: { name: "Fliesen Krause", street: "Gneisenaustr. 22", zip: "10961", city: "Berlin", email: "buero@fliesen-krause.example" },
  lines: [
    { name: "Feinsteinzeug 60x60 anthrazit", articleNo: "FS-6060-AN", qty: 18.5, unit: "m²", price: 32.90, rate: 19 },
    { name: "Abschlussprofil Alu 2,5 m", articleNo: "AP-25-AL", qty: 8, unit: "Stk", price: 9.80, rate: 19 },
    { name: "Flexkleber S1 25 kg", articleNo: "FK-S1-25", qty: 6, unit: "Sack", price: 17.50, rate: 19 },
    { name: "Fugenmörtel weiß 5 kg", articleNo: "FM-WS-5", qty: 4, unit: "Eimer", price: 11.40, rate: 19 }
  ]
};
const r = ZFE.buildXRechnung(weber);
fs.writeFileSync(path.join(out, "Rechnung_Weber_88213_XRechnung.xml"), r.xml);
console.log("XRechnung:", r.totals);
