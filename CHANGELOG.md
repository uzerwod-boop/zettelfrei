# Änderungen

Alle wichtigen Änderungen an Zettelfrei. Format angelehnt an [Keep a Changelog](https://keepachangelog.com/de/1.1.0/), Versionen nach [Semantic Versioning](https://semver.org/lang/de/).

## [0.7.1] – 2026-11-02 · Version immer im Blick

### Neu
- Die Versionsnummer steht unten in der linken Leiste. Ein Klick darauf (oder auf das Logo oben) öffnet „Über Zettelfrei“: Version und Datum, Programm oder Browser, KI-Modell, „Was ist neu?“, „Nach Updates suchen“, Anleitung, Lizenz und „Versionsinfo kopieren“ für Rückfragen.

## [0.7.0] – 2026-11-02 · Bedienen wie Excel

### Neu
- Bereiche markieren: mit gedrückter Maus, Umschalt + Klick/Pfeiltasten, Klick auf Zeilennummer oder Spaltenbuchstaben, ⌘/Strg + A. Oben stehen Summe, Anzahl und Mittelwert des Bereichs.
- Kopieren, Ausschneiden, Einfügen (⌘/Strg + C, X, V) – kompatibel mit Excel, Numbers und Google Sheets. Ein einzelner Wert füllt einen markierten Bereich. Berechnete Spalten werden beim Einfügen geschützt.
- Rückgängig und Wiederholen mit ⌘/Strg + Z bzw. ⌘ + Umschalt + Z / Strg + Y und mit Knöpfen oben.
- Rechtsklick auf Zeilen: darüber/darunter einfügen (auch mehrere), löschen, kopieren, einfügen, Inhalt löschen. Neue Zeilen bekommen die Formeln der Nachbarzeile, Zwischensummen und Bezüge wachsen mit.
- Rechtsklick auf Überschriften: auf- und absteigend sortieren, umbenennen (auch per Doppelklick), neue Spalte (auch mit „+“), letzte Spalte löschen.
- Blätter: neues Blatt mit „+“, umbenennen per Doppelklick oder Rechtsklick (Formeln in anderen Blättern werden angepasst), löschen mit Warnung, wenn andere Blätter darauf verweisen.
- Datei schließen mit „×“ neben dem Dateinamen – mit Rückfrage, wenn noch nicht als Excel gespeichert.
- Kurze, ruhige Animationen: Druck-Effekt auf Knöpfen, sanfter Bereichswechsel, neue Zeilen leuchten kurz auf, Menüs und Dialoge gleiten herein. Bei „Bewegung reduzieren“ im System ausgeschaltet.

### Behoben
- Nach dem Löschen einer Zeile konnten in der gespeicherten Excel-Datei Farben und Rahmen um eine Zeile verrutschen. Formatierungen bleiben jetzt an ihrer Zeile – auch beim Einfügen und Sortieren.

## [0.6.2] – 2026-11-02 · Nur Ihre eigenen Dateien

### Geändert
- Das installierte Programm zeigt keine Beispieldateien mehr – weder „Beispiel öffnen…“ noch Beispiel-Rechnungen. So verwechselt niemand Beispiele mit den eigenen Daten. Auf der Startseite gibt es dafür den Knopf „Excel-Datei öffnen“.
- In der Browser-Version bleiben die Beispiele zum Ausprobieren erhalten.
- Release-Text und Hinweis „neue Version ist da“ kommen jetzt automatisch aus diesem Änderungsprotokoll.

## [0.6.1] – 2026-11-02 · Erstes automatisches Update

### Neu
- Erste Version, die per automatischem Update kommt: Zettelfrei 0.6.0 findet sie beim Start und installiert sie nach einem Klick.
- Nach einem Update meldet Zettelfrei einmal „wurde auf 0.6.1 aktualisiert“ mit Link „Was ist neu?“.

## [0.6.0] – 2026-11-02 · Bereit für den Pilotbetrieb

### Neu
- **KI-Qualitätstest** (KI-Einstellungen): 17 Fragen an eine Testdatei mit bekannten Antworten – Auswertungen über mehrere Blätter, mehrdeutige Namen, Änderungen an vielen Zeilen, doppelte Einträge, Anweisungen in Zellen. Ergebnis je Frage als korrekt, abgefangen (Modell falsch, Zettelfrei warnt/verhindert) oder falsch, mit Zeit, Kosten und Token-Verbrauch. Frühere Läufe bleiben zum Vergleich von Modellen stehen; „Ergebnis kopieren“ für Rückmeldungen. Die geöffnete Datei bleibt unberührt.
- **Automatische Updates** im Programm: Beim Start sucht Zettelfrei nach einer neuen Version und installiert sie nach einem Klick („Jetzt aktualisieren“), danach Neustart. Updates sind signiert; installiert wird nur, was mit dem Zettelfrei-Schlüssel signiert ist. Zusätzlich „Nach Updates suchen“ in den Einstellungen.
- **Feedback** unten links: Problem, Idee oder Lob mit Version, System und KI-Modell – ohne Tabellendaten. Per E-Mail (Adresse in `web/src/config.ts`) oder über das GitHub-Formular.
- **Hilfe** unten links öffnet die neue Anleitung [Erste Schritte](docs/Erste-Schritte.md), auch als [PDF](docs/Zettelfrei-Erste-Schritte.pdf) zum Verschicken.
- Kosten und Tokens jeder KI-Anfrage werden mitgezählt (bei OpenRouter mit genauen Kosten).
- Links und E-Mails öffnen im Programm im System-Browser bzw. Mail-Programm.

## [0.5.4] – 2026-11-02 · Schutz, der nicht vom KI-Modell abhängt

Erkenntnis aus Tests mit einem schnellen, günstigen Modell: Regeln allein reichen nicht. Diese Version prüft im Programm selbst.

### Neu
- Jede Antwort zeigt, mit welchen Filtern gerechnet wurde (z. B. „Filter: Freigegeben = Ja“) – erfundene Einschränkungen fallen sofort auf.
- Zusammengezählte Einträge werden genannt, mit Teilsummen: „Sie haben ‚Kevin Müller‘ genannt, zusammengezählt wurden aber: Kevin Müller (280), Kevin Mueller (290)“.
- Selbstkorrektur: War ein Teil nicht berechenbar (falsche Spalte, falsches Blatt), bekommt die KI die Fehlermeldung samt vorhandenen Spalten und korrigiert einmal selbst („automatisch korrigiert“).
- Mehrdeutige Namen in der Vorschau: Auswahl per Knopf („Kevin Müller“ / „Kevin Mueller“), vorher nicht vorausgewählt.
- Löschen einer Zeile, deren Nummer doppelt vorkommt: Hinweis mit beiden Zeilen, nicht vorausgewählt.
- Änderungen, die aus Text in der Tabelle stammen („tu, was in den Notizen steht“) oder Text enthalten, der wie eine KI-Anweisung aussieht: Warnung, nichts vorausgewählt.
- Prüfen meldet Zellen, deren Text wie eine Anweisung an eine KI aussieht.

### Behoben
- Prüfen: Kennungen wie „L2“ oder „A-1001“ wurden als „Text statt Zahl“ gemeldet.
- Regeln: exakt genannte Namen exakt filtern, keine Filter erfinden, „offen/überfällig“ nach offenem Betrag statt Status, „ohne X“ als „≠ X“.

## [0.5.3] – 2026-11-02 · Assistent: sicher bei vielen Zeilen

### Behoben
- Assistent, Änderungen an vielen Zeilen („alle offenen Rechnungen von X auf bezahlt“): Die KI nennt jetzt eine Bedingung, die Tabelle sucht die Zeilen selbst. Vorher konnte die KI falsche Zeilennummern wählen (z. B. eine Rechnung eines anderen Kunden) oder bereits bezahlte und stornierte Rechnungen mitändern.
- Zeilennummer und Bedingung passen nicht zusammen → die Änderung wird abgelehnt statt in der falschen Zeile ausgeführt.
- Anweisungen, die in Zellen stehen (z. B. in Notizen), befolgt die KI nicht mehr; bei „tu, was in den Notizen steht“ fragt sie nach.
- Antworten bei Änderungen sind als Vorschlag formuliert („Ich schlage vor …“), nicht als „wurde geändert“.
- „Wie viel ist offen?“ summiert die Spalte mit dem offenen Betrag, statt nur Zeilen mit Status „Offen“; doppelte Einträge löscht die KI nicht selbst, sondern fragt nach.

### Neu
- Vorschau zeigt zu jeder geänderten Zeile, um welche es geht (z. B. „R-2026-0010 · Schneider GmbH“), und warnt bei großen Änderungen.
- Werte aus einer anderen Spalte derselben Zeile übernehmen, z. B. „Bezahlt EUR = Brutto EUR“ beim Bezahlt-Setzen.
- Summen sagen ehrlich, womit gerechnet wurde: Zahlen, die als Text in der Zelle stehen, nicht lesbare Werte („ca. 250“) und leere Zellen – samt dem Ergebnis, das Excel selbst anzeigen würde.
- Spalte mit Spalte vergleichen, z. B. „Bestand unter Mindestbestand“.
- Rechnen je Zeile in Auswertungen, z. B. Nachbestellkosten = (Mindestbestand − Bestand) × EK.
- Gruppieren über ein anderes Blatt, z. B. Lohn je Kunde, wenn der Kunde nur im Blatt „Projekte“ steht.
- Rechnen je Gruppe über Blätter hinweg, z. B. verbrauchter Budgetanteil je Baustelle.
- Mehrdeutige Namen werden erkannt: „Kevin“ passt zu „Kevin Müller“ und „Kevin Mueller“ → Hinweis in der Antwort; in der Vorschau ist die Zeile dann nicht vorausgewählt.
- „Bad Schneider“ erfasst auch „Bad Schneider II“ → die Antwort nennt alle erfassten Einträge.
- Konnte ein Teil einer Frage nicht berechnet werden, steht das deutlich in der Antwort statt „(?)“.
- Zwischenergebnisse einer Rechnung erscheinen nicht mehr als eigene Diagramme.

## [0.5.2] – 2026-11-02 · Excel-Dateien aus anderen Programmen

### Behoben
- Excel-Dateien, die von anderen Programmen erzeugt wurden (z. B. .NET OpenXML SDK, ERP- und Lohn-Exporte mit `x:`-Präfix, BOM und absoluten Pfaden), ließen sich nicht öffnen („Diese Datei lässt sich nicht lesen“). Sie werden jetzt angeglichen und normal gelesen; beim Speichern bleibt die Formatierung erhalten.
- Arbeitszeit-Prüfung: Genau 6 Std. ohne Pause wurde fälschlich als „Pause zu kurz“ gemeldet (Rundungsfehler). Gerechnet wird jetzt auf die Minute genau.

## [0.5.1] – 2026-10-26 · Feinschliff

### Geändert
- Auswahlfelder sehen auf macOS, Windows und im Browser gleich aus: flach, mit eigenem Pfeil, in hellem und dunklem Design.
- Schmale Scrollbalken in den Farben der Oberfläche statt der grauen Systembalken.
- „In welches Blatt?“ zeigt ohne geöffnete Datei einen Hinweis statt eines leeren Feldes.

### Behoben
- Die Bereiche rechts ließen sich seitlich verschieben (Scrollbalken unter den Beispiel-Knöpfen).

## [0.5.0] – 2026-10-26 · Jetzt als Programm

### Neu
- Zettelfrei als Programm für macOS (Apple Silicon und Intel, `.dmg`) und Windows (`.exe`, `.msi`), gebaut mit Tauri. Kein Server, kein Konto; gespeichert wird über den „Sichern unter“-Dialog des Systems.
- Neue Oberfläche in React und TypeScript (`web/`) – dieselbe im Programm, im Browser und bei `npm start`. Zellen anklicken und tippen wie in Excel (Enter, Tab, Pfeiltasten, Entf, F2, Esc).
- Excel-Datei ins Fenster ziehen öffnet sie, auch auf der leeren Startseite. Rechnungen und Fotos auf die Tabelle ziehen öffnet „Belege“.
- Chat und Belege behalten ihren Stand beim Wechseln der Bereiche.
- GitHub Actions baut die Programme automatisch in einen Release-Entwurf (`.github/workflows/desktop.yml`).
- App-Symbol aus dem Logo.
- Neuer Ende-zu-Ende-Test für die neue Oberfläche (`tests/browser_web_test.py`).

### Geändert
- Startseite (GitHub Pages und `npm start`) ist die neue Oberfläche. Die bisherige liegt unter `/klassisch/`, `/app/` leitet um.
- Texte sagen „auf diesem Rechner“ statt „im Browser“, wenn Zettelfrei als Programm läuft.

### Behoben
- Eine Zelländerung wurde im Verlauf doppelt eingetragen.
- Nach „E-Rechnung schreiben“ blieb die Bestätigung leer (auch in der klassischen Ansicht).

## [0.4.0] – 2026-10-26 · Auswertungen

### Neu
- Fragen „je/pro/nach …“: Auswertungen gruppiert nach Spalte oder Zeitraum (Tag, Woche, Monat, Jahr), sortiert, Top-N – als Balkendiagramm im Chat.
- Rechnen mit Ergebnissen: Unterschiede, Anteile, Prozent (z. B. „Anteil Neubau Weber an allen Stunden“), sicher gerechnet ohne eval.
- Blätter verknüpfen: z. B. Stunden nur für Baustellen mit offener Rechnung (Filter „in“ mit dem Ergebnis einer anderen Abfrage).
- Auswertung als neues Blatt in die Excel-Datei übernehmen, inklusive Summenformel.
- Rund 40 weitere Excel-Funktionen: XVERWEIS, WVERWEIS, SUMMEWENNS, ZÄHLENWENNS, MITTELWERTWENN(S), MAXWENNS/MINWENNS, INDEX, VERGLEICH, WENNS, ERSTERWERT, WENNNV, SUMMENPRODUKT, PRODUKT, MEDIAN, GANZZAHL, REST, OBERGRENZE, UNTERGRENZE, EDATUM, MONATSENDE, KALENDERWOCHE, NETTOARBEITSTAGE, DATEDIF, TEXT, TEIL, FINDEN, SUCHEN, WECHSELN, TEXTVERKETTEN u. a.

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

[0.7.1]: https://github.com/uzerwod-boop/zettelfrei/releases/tag/v0.7.1
[0.7.0]: https://github.com/uzerwod-boop/zettelfrei/releases/tag/v0.7.0
[0.6.2]: https://github.com/uzerwod-boop/zettelfrei/releases/tag/v0.6.2
[0.6.1]: https://github.com/uzerwod-boop/zettelfrei/releases/tag/v0.6.1
[0.6.0]: https://github.com/uzerwod-boop/zettelfrei/releases/tag/v0.6.0
[0.5.4]: https://github.com/uzerwod-boop/zettelfrei/releases/tag/v0.5.4
[0.5.3]: https://github.com/uzerwod-boop/zettelfrei/releases/tag/v0.5.3
[0.5.2]: https://github.com/uzerwod-boop/zettelfrei/releases/tag/v0.5.2
[0.5.1]: https://github.com/uzerwod-boop/zettelfrei/releases/tag/v0.5.1
[0.5.0]: https://github.com/uzerwod-boop/zettelfrei/releases/tag/v0.5.0
[0.4.0]: https://github.com/uzerwod-boop/zettelfrei/commit/472ebad
[0.3.0]: https://github.com/uzerwod-boop/zettelfrei/commit/9e32630
[0.2.2]: https://github.com/uzerwod-boop/zettelfrei/commit/ef087fa
[0.2.1]: https://github.com/uzerwod-boop/zettelfrei/releases/tag/v0.2.1
[0.2.0]: https://github.com/uzerwod-boop/zettelfrei/commit/831d49f
[0.1.0]: https://github.com/uzerwod-boop/zettelfrei/commit/ce33163
