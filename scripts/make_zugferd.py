"""Erzeugt eine Beispiel-ZUGFeRD-Rechnung (Factur-X, Profil EN 16931): PDF mit eingebetteter CII-XML.
Benötigt: pip install factur-x reportlab
"""
import os
from io import BytesIO
from reportlab.lib.pagesizes import A4
from reportlab.pdfgen import canvas
from facturx import generate_from_binary

OUT = os.path.join(os.path.dirname(os.path.dirname(os.path.abspath(__file__))), "beispiele", "Rechnung_Kaiser_RE-2026-4502_ZUGFeRD.pdf")
LINES = [  # Artikel-Nr, Name, Menge, Einheit (Code, Anzeige), Preis
    ("A-1001", "Fliesenkleber C2TE 25 kg", 10, ("XBG", "Sack"), 14.90),
    ("A-1021", "Silikon sanitär grau 310 ml", 12, ("XCQ", "Kartusche"), 6.20),
    ("A-1040", "Tiefengrund 10 l", 2, ("XJC", "Kanister"), 29.50),
]
net = round(sum(q * p for _, _, q, _, p in LINES), 2)
tax = round(net * 0.19, 2)
gross = round(net + tax, 2)


def f(x):
    return f"{x:.2f}"


items = "".join(f"""
  <ram:IncludedSupplyChainTradeLineItem>
   <ram:AssociatedDocumentLineDocument><ram:LineID>{i}</ram:LineID></ram:AssociatedDocumentLineDocument>
   <ram:SpecifiedTradeProduct><ram:SellerAssignedID>{nr}</ram:SellerAssignedID><ram:Name>{name}</ram:Name></ram:SpecifiedTradeProduct>
   <ram:SpecifiedLineTradeAgreement><ram:NetPriceProductTradePrice><ram:ChargeAmount>{f(p)}</ram:ChargeAmount></ram:NetPriceProductTradePrice></ram:SpecifiedLineTradeAgreement>
   <ram:SpecifiedLineTradeDelivery><ram:BilledQuantity unitCode="{u[0]}">{q}</ram:BilledQuantity></ram:SpecifiedLineTradeDelivery>
   <ram:SpecifiedLineTradeSettlement>
    <ram:ApplicableTradeTax><ram:TypeCode>VAT</ram:TypeCode><ram:CategoryCode>S</ram:CategoryCode><ram:RateApplicablePercent>19</ram:RateApplicablePercent></ram:ApplicableTradeTax>
    <ram:SpecifiedTradeSettlementLineMonetarySummation><ram:LineTotalAmount>{f(q*p)}</ram:LineTotalAmount></ram:SpecifiedTradeSettlementLineMonetarySummation>
   </ram:SpecifiedLineTradeSettlement>
  </ram:IncludedSupplyChainTradeLineItem>""" for i, (nr, name, q, u, p) in enumerate(LINES, 1))

XML = f"""<?xml version="1.0" encoding="UTF-8"?>
<rsm:CrossIndustryInvoice xmlns:rsm="urn:un:unece:uncefact:data:standard:CrossIndustryInvoice:100" xmlns:ram="urn:un:unece:uncefact:data:standard:ReusableAggregateBusinessInformationEntity:100" xmlns:udt="urn:un:unece:uncefact:data:standard:UnqualifiedDataType:100">
 <rsm:ExchangedDocumentContext><ram:GuidelineSpecifiedDocumentContextParameter><ram:ID>urn:cen.eu:en16931:2017</ram:ID></ram:GuidelineSpecifiedDocumentContextParameter></rsm:ExchangedDocumentContext>
 <rsm:ExchangedDocument><ram:ID>RE-2026-4502</ram:ID><ram:TypeCode>380</ram:TypeCode>
  <ram:IssueDateTime><udt:DateTimeString format="102">20261016</udt:DateTimeString></ram:IssueDateTime></rsm:ExchangedDocument>
 <rsm:SupplyChainTradeTransaction>{items}
  <ram:ApplicableHeaderTradeAgreement>
   <ram:BuyerReference>Bad Schneider</ram:BuyerReference>
   <ram:SellerTradeParty><ram:Name>Baustoff Kaiser GmbH</ram:Name>
    <ram:PostalTradeAddress><ram:PostcodeCode>12057</ram:PostcodeCode><ram:LineOne>Sonnenallee 220</ram:LineOne><ram:CityName>Berlin</ram:CityName><ram:CountryID>DE</ram:CountryID></ram:PostalTradeAddress>
    <ram:SpecifiedTaxRegistration><ram:ID schemeID="VA">DE298765432</ram:ID></ram:SpecifiedTaxRegistration></ram:SellerTradeParty>
   <ram:BuyerTradeParty><ram:Name>Fliesen Krause</ram:Name>
    <ram:PostalTradeAddress><ram:PostcodeCode>10961</ram:PostcodeCode><ram:LineOne>Gneisenaustr. 22</ram:LineOne><ram:CityName>Berlin</ram:CityName><ram:CountryID>DE</ram:CountryID></ram:PostalTradeAddress></ram:BuyerTradeParty>
  </ram:ApplicableHeaderTradeAgreement>
  <ram:ApplicableHeaderTradeDelivery><ram:ActualDeliverySupplyChainEvent><ram:OccurrenceDateTime><udt:DateTimeString format="102">20261016</udt:DateTimeString></ram:OccurrenceDateTime></ram:ActualDeliverySupplyChainEvent></ram:ApplicableHeaderTradeDelivery>
  <ram:ApplicableHeaderTradeSettlement><ram:InvoiceCurrencyCode>EUR</ram:InvoiceCurrencyCode>
   <ram:SpecifiedTradeSettlementPaymentMeans><ram:TypeCode>58</ram:TypeCode><ram:PayeePartyCreditorFinancialAccount><ram:IBANID>DE89370400440532013000</ram:IBANID></ram:PayeePartyCreditorFinancialAccount></ram:SpecifiedTradeSettlementPaymentMeans>
   <ram:ApplicableTradeTax><ram:CalculatedAmount>{f(tax)}</ram:CalculatedAmount><ram:TypeCode>VAT</ram:TypeCode><ram:BasisAmount>{f(net)}</ram:BasisAmount><ram:CategoryCode>S</ram:CategoryCode><ram:RateApplicablePercent>19</ram:RateApplicablePercent></ram:ApplicableTradeTax>
   <ram:SpecifiedTradePaymentTerms><ram:Description>Zahlbar innerhalb 14 Tagen ohne Abzug.</ram:Description><ram:DueDateDateTime><udt:DateTimeString format="102">20261030</udt:DateTimeString></ram:DueDateDateTime></ram:SpecifiedTradePaymentTerms>
   <ram:SpecifiedTradeSettlementHeaderMonetarySummation><ram:LineTotalAmount>{f(net)}</ram:LineTotalAmount><ram:TaxBasisTotalAmount>{f(net)}</ram:TaxBasisTotalAmount>
    <ram:TaxTotalAmount currencyID="EUR">{f(tax)}</ram:TaxTotalAmount><ram:GrandTotalAmount>{f(gross)}</ram:GrandTotalAmount><ram:DuePayableAmount>{f(gross)}</ram:DuePayableAmount></ram:SpecifiedTradeSettlementHeaderMonetarySummation>
  </ram:ApplicableHeaderTradeSettlement>
 </rsm:SupplyChainTradeTransaction>
</rsm:CrossIndustryInvoice>"""


def pdf():
    buf = BytesIO()
    c = canvas.Canvas(buf, pagesize=A4)
    c.setTitle("Rechnung RE-2026-4502")
    y = 800
    c.setFont("Helvetica-Bold", 16); c.drawString(60, y, "Baustoff Kaiser GmbH"); y -= 18
    c.setFont("Helvetica", 9); c.drawString(60, y, "Sonnenallee 220 · 12057 Berlin · USt-IdNr. DE298765432"); y -= 40
    c.setFont("Helvetica", 10)
    for l in ["Fliesen Krause", "Gneisenaustr. 22", "10961 Berlin"]:
        c.drawString(60, y, l); y -= 14
    y -= 20
    c.setFont("Helvetica-Bold", 13); c.drawString(60, y, "Rechnung RE-2026-4502"); y -= 16
    c.setFont("Helvetica", 10); c.drawString(60, y, "Datum: 16.10.2026 · Kommission: Bad Schneider"); y -= 30
    c.setFont("Helvetica-Bold", 10)
    for x, t in [(60, "Art.-Nr"), (130, "Artikel"), (350, "Menge"), (420, "Einzelpreis"), (500, "Gesamt")]:
        c.drawString(x, y, t)
    y -= 16; c.setFont("Helvetica", 10)
    for nr, name, q, u, p in LINES:
        c.drawString(60, y, nr); c.drawString(130, y, name); c.drawString(350, y, f"{q} {u[1]}")
        c.drawRightString(480, y, f"{p:.2f} €".replace(".", ",")); c.drawRightString(540, y, f"{q*p:.2f} €".replace(".", ",")); y -= 15
    y -= 15
    for t, v in [("Netto", net), ("MwSt 19 %", tax), ("Brutto", gross)]:
        c.drawString(400, y, t); c.drawRightString(540, y, f"{v:.2f} €".replace(".", ",")); y -= 15
    y -= 20; c.drawString(60, y, "Zahlbar bis 30.10.2026 ohne Abzug. IBAN DE89 3704 0044 0532 0130 00")
    c.drawString(60, 60, "Diese Rechnung enthält die Rechnungsdaten als ZUGFeRD/Factur-X (EN 16931). Beispieldaten.")
    c.save()
    return buf.getvalue()


open(OUT, "wb").write(generate_from_binary(pdf(), XML.encode("utf-8"), flavor="factur-x", level="en16931", check_xsd=True))
open(OUT.replace(".pdf", ".xml"), "w").write(XML)
print("ok", OUT, net, tax, gross)
