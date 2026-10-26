# Änderungen

Alle wichtigen Änderungen an Zettelfrei. Format angelehnt an [Keep a Changelog](https://keepachangelog.com/de/1.1.0/), Versionen nach [Semantic Versioning](https://semver.org/lang/de/).

## [0.3.0] – 2026-10-26 · Neuer Arbeitsplatz

### Neu
- Neuer Aufbau wie in Excel: dunkle Leiste links (Assistent, Belege, Dokumente, Prüfen, Verlauf), Tabelle in der Mitte mit Spaltenbuchstaben, Formelleiste und Blättern unten, Seitenbereich rechts.
- Assistent als Chat: Fragen und Anweisungen bleiben als Verlauf stehen, Antworten und Vorschauen darunter; Enter sendet, Umschalt+Enter macht eine neue Zeile.
- Kennzahlen über der Tabelle: Zeilen, Spalten, Summe der wichtigsten Euro-Spalte, Hinweise zum Prüfen (ein Klick öffnet „Prüfen“).
- Vorschau zeigt Uhrzeiten und Daten im richtigen Format.
- Am Handy: Leiste oben, Tabelle und Assistent untereinander.

## [0.2.2] – 2026-10-20 · Echte Excel-Dateien

### Neu
- Überschriften über zwei Zeilen (verbundene Zellen) werden erkannt.
- Zwischensummen, Wochen-Abschnitte, Leer- und Notizzeilen bleiben an ihrem Platz; die Zeilennummern sind dieselben wie in Excel.
- Nebentabellen rechts neben der Tabelle (z. B. Stundensätze) werden nicht mehr angehängt, bleiben aber für Formeln erreichbar.
- Eigener Formel-Interpreter ohne `eval`: WENN, WENNFEHLER, SVERWEIS, SUMMEWENN, ZÄHLENWENN, TEILERGEBNIS, UND/ODER/NICHT, Text und Vergleiche – deutsch und englisch.
- Bezüge aus anderen Blättern (z. B. Übersicht) wandern beim Einfügen und Löschen mit und werden gespeichert.
- Datum und Uhrzeit auch als Text erkannt (`02.10.26`, `03/11/2026`, `7:00`); Prozent-Eingabe wie `19 %`.
- Neue Prüfungen: Zahl als Text, Text statt Zahl, Steuersatz 19 statt 19 %, Nachschlagen ohne Treffer; Arbeitszeitprüfung robuster (Pausen wie „30 min“, Namen mit Leerzeichen).
- Neue Beispiele: „Krause Betrieb Oktober 2026“ (5 Blätter) und „Stresstest Chaos Tabelle“.

### Behoben
- Summen über Bereiche lassen Text wie Excel weg.
- Doppelte Nummern nur noch in der ersten Nummernspalte (Kundennummern dürfen sich wiederholen).

## [0.2.1] – 2026-10-20 · Geprüft und gehärtet

### Sicherheit
- Formelrechner führt keinen Code mehr aus, nur Zahlen, Rechenzeichen und bekannte Funktionen.
- Werte von KI oder aus E-Rechnungen werden nie zu Formeln (Schutz vor DDE); CSV-Export entschärft `=`, `+`, `-`, `@`.
- Excel-Dateien werden mit ExcelJS gelesen; SheetJS 0.18.5 nur noch für `.xls`/`.ods`.
- PDF.js mit `isEvalSupported: false` (CVE-2024-4367); jsPDF 4.2.1.
- Server: nur JSON, nur eigene Herkunft, Prüfung des Hostnamens (CSRF, DNS-Rebinding); neue Variable `ZF_ALLOWED_HOSTS`.

### Behoben
- Gutschriften werden negativ gebucht und senken den Lagerbestand.
- CSV im deutschen Format (`;`, `1.234,50`, `12.10.2026`) wird richtig gelesen.
- Voller Browser-Speicher wird gemeldet.

### Neu
- Logo und Markenzeichen, README neu gestaltet, Screenshots hell/dunkel, Demo-GIF.
- Lizenz nach Wahl MIT oder Apache-2.0.
- Vorlagen für Fehlermeldungen und Ideen, `SECURITY.md`, `CONTRIBUTING.md`, englische Kurzbeschreibung.

## [0.2.0] – 2026-10-19 · Bürokratie, die sich selbst erledigt

### Neu
- E-Rechnung empfangen: XRechnung (UBL) und ZUGFeRD/Factur-X (CII) exakt und ohne KI lesen.
- E-Rechnung schreiben: XRechnung 3.0, geprüft gegen EN 16931 und XRechnung-Schematron.
- Arbeitszeit mit Uhrzeiten, Hinweise nach ArbZG §§ 3, 4, 5.
- Änderungsprotokoll, auch als Blatt in der Excel-Datei.
- Mahnwesen mit drei Stufen als PDF, Abnahmeprotokoll mit Unterschriften und Fotos.
- Belegliste als CSV für den Steuerberater, Wochenbericht.
- Speichern mit Formatierung der Originaldatei (ExcelJS), Spracheingabe, Bedienung am Handy.

## [0.1.0] – 2026-10-19 · Erster Prototyp

- Excel-Dateien öffnen, mit ganzen Sätzen bedienen, Formeln rechnen, prüfen, Verlauf, speichern.
- KI wahlweise über OpenRouter, eigenen Server, Claude API oder lokales Modell.
- PDFs und Fotos von Lieferantenrechnungen per KI einlesen.

[0.3.0]: https://github.com/uzerwod-boop/zettelfrei/compare/v0.2.1...main
[0.2.2]: https://github.com/uzerwod-boop/zettelfrei/commit/ef087fa
[0.2.1]: https://github.com/uzerwod-boop/zettelfrei/releases/tag/v0.2.1
[0.2.0]: https://github.com/uzerwod-boop/zettelfrei/commit/831d49f
[0.1.0]: https://github.com/uzerwod-boop/zettelfrei/commit/ce33163
