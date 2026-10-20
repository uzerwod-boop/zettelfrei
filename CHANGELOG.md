# Änderungen

Alle wichtigen Änderungen an Zettelfrei. Format angelehnt an [Keep a Changelog](https://keepachangelog.com/de/1.1.0/), Versionen nach [Semantic Versioning](https://semver.org/lang/de/).

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

[0.2.1]: https://github.com/uzerwod-boop/zettelfrei/releases/tag/v0.2.1
[0.2.0]: https://github.com/uzerwod-boop/zettelfrei/commit/831d49f
[0.1.0]: https://github.com/uzerwod-boop/zettelfrei/commit/ce33163
