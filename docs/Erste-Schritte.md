# Zettelfrei – Erste Schritte

**Excel bleibt. Das Tippen fällt weg.** Zettelfrei öffnet Ihre vorhandenen Excel-Dateien – Stundenzettel, Lagerliste, Auftragsliste – und lässt sie mit ganzen Sätzen bedienen. Ihre Datei bleibt auf Ihrem Rechner.

Diese Anleitung dauert etwa 10 Minuten.

---

## 1. Installieren

| System | So geht's |
|---|---|
| **Mac** | Auf der [Download-Seite](https://github.com/uzerwod-boop/zettelfrei/releases/latest) die Datei `…universal.dmg` laden, öffnen und Zettelfrei in den Ordner „Programme“ ziehen. |
| **Windows** | Auf der [Download-Seite](https://github.com/uzerwod-boop/zettelfrei/releases/latest) die Datei `…x64-setup.exe` laden und starten. Keine Administratorrechte nötig. |
| **Ohne Installation** | Im Browser öffnen: [uzerwod-boop.github.io/zettelfrei](https://uzerwod-boop.github.io/zettelfrei/) |

**Beim ersten Start fragt das System einmal nach**, weil Zettelfrei noch nicht bei Apple bzw. Microsoft registriert ist:

- **Mac:** Beim Hinweis „Fertig“ wählen. Dann **Systemeinstellungen → Datenschutz & Sicherheit → „Dennoch öffnen“**. Danach startet Zettelfrei normal aus dem Dock.
- **Windows:** Bei „Der Computer wurde durch Windows geschützt“ auf **„Weitere Informationen“ → „Trotzdem ausführen“** klicken.

Neue Versionen meldet Zettelfrei beim Start selbst („Zettelfrei 0.x ist da“) – ein Klick auf **„Jetzt aktualisieren“** genügt.

## 2. KI einrichten (einmalig, ca. 3 Minuten)

Ohne KI funktionieren schon: Datei öffnen, bearbeiten, prüfen, speichern, E-Rechnungen lesen und schreiben, Mahnungen, Abnahmeprotokoll. Für Sätze wie „Mehmet heute 7 bis 16 Uhr bei Weber“ braucht Zettelfrei eine KI.

1. Auf [openrouter.ai](https://openrouter.ai) anmelden (z. B. mit Google).
2. Unter **Credits** etwa 5–10 € aufladen. Eine Frage an die Tabelle kostet meist weniger als einen Cent.
3. Unter **Keys** einen Schlüssel anlegen. Tipp: ein Ausgabenlimit (z. B. 10 €) eintragen.
4. In Zettelfrei rechts oben im Assistenten auf den KI-Status klicken → **OpenRouter** wählen → Schlüssel einfügen.
5. Als Modell empfehlen wir `anthropic/claude-haiku-5.5` (günstig und zuverlässig). Feld leer lassen = `claude-sonnet-5.5` (genauer, etwas teurer).
6. **„Verbindung testen“** → **„Speichern“**.

Der Schlüssel bleibt nur auf Ihrem Rechner. Geben Sie ihn nie weiter.

> **Wie gut ist mein Modell?** In den KI-Einstellungen startet **„KI-Qualitätstest“** 17 Testfragen mit bekannten Antworten (1–3 Minuten, wenige Cent). So sehen Sie schwarz auf weiß, wie zuverlässig das gewählte Modell ist.

## 3. Datei öffnen

Excel-Datei einfach **ins Fenster ziehen** oder **„Excel-Datei öffnen“** klicken. Im Programm sehen Sie nur Ihre eigenen Dateien. Zum Ausprobieren ohne eigene Datei gibt es in der [Browser-Version](https://uzerwod-boop.github.io/zettelfrei/) Beispieldateien („Beispiel öffnen…“).

Die Tabelle funktioniert wie Excel:

- **Zelle anklicken und lostippen**, **Enter** speichert. Formeln mit „=“.
- **Bereich markieren** mit gedrückter Maus oder Umschalt + Pfeiltasten – oben stehen dann Summe, Anzahl und Mittelwert.
- **Kopieren und Einfügen** (⌘/Strg + C, V) – auch direkt aus Excel oder in Excel hinein.
- **Rückgängig** mit ⌘/Strg + Z, Wiederholen mit ⌘ + Umschalt + Z bzw. Strg + Y – oder mit den Pfeilen oben.
- **Rechtsklick** auf eine Zeile: darüber/darunter einfügen, löschen. Auf eine Überschrift: sortieren, umbenennen, neue Spalte. Auf ein Blatt unten: umbenennen, löschen, neues Blatt (auch mit **+**).
- **Datei schließen** mit dem **×** neben dem Dateinamen.

Formeln, Zwischensummen, Farben und mehrere Blätter bleiben beim Speichern erhalten.

> **Tipp für den Anfang:** Arbeiten Sie mit einer Kopie Ihrer Datei und speichern Sie mit **„Als Excel speichern“** unter einem neuen Namen.

## 4. Fünf Sätze zum Ausprobieren

Rechts im Assistenten eintippen (oder mit dem Mikrofon sprechen):

1. `Wie viele Stunden je Baustelle?`
2. `Mehmet heute 7 bis 16 Uhr bei Weber, 30 Minuten Pause`
3. `Welche Rechnungen sind überfällig?`
4. `Becker hat heute bezahlt`
5. `Welche Artikel liegen unter Mindestbestand?`

## 5. Sicher arbeiten – darauf achten

- **Erst Vorschau, dann „Übernehmen“.** Jede Änderung erscheint zuerst als „alt → neu“. Nichts wird ohne Ihren Klick geändert.
- **Unter jeder Antwort steht, womit gerechnet wurde** – „Berechnet aus 40 Zeilen · Filter: …“. Passt der Filter nicht zu Ihrer Frage, fragen Sie genauer.
- **Gelbe Hinweise ernst nehmen:** z. B. „‚Kevin‘ passt zu mehreren“ oder „steht als Text in der Zelle“.
- **Summen rechnet die Tabelle, nicht die KI.** Die KI sagt nur, *was* gezählt werden soll – so entstehen keine erfundenen Zahlen.
- **Verlauf:** Jeder frühere Stand lässt sich mit einem Klick zurückholen.

## 6. Weitere Bereiche

- **Belege:** E-Rechnung (XRechnung, ZUGFeRD), PDF oder Foto auf die Tabelle ziehen – wird gelesen, z. B. erhöht sich der Lagerbestand.
- **Dokumente:** E-Rechnung schreiben, Mahnung, Abnahmeprotokoll mit Unterschrift, Liste für den Steuerberater, Wochenbericht. Vorher einmal die **Firmendaten** eintragen.
- **Prüfen:** findet ohne KI leere Pflichtfelder, doppelte Nummern, Verstöße gegen das Arbeitszeitgesetz, überfällige Rechnungen und Formelfehler.

## 7. Feedback

Unten links in der Leiste: **„Feedback“**. Kurz beschreiben, was passiert ist oder was Ihnen fehlt. Mitgeschickt werden nur Version, System und KI-Modell – **nie Ihre Tabelle**.

## Datenschutz in Kürze

- Die Excel-Datei wird auf Ihrem Rechner verarbeitet und nicht hochgeladen.
- An die KI geht nur, was Sie fragen: der Tabelleninhalt als Text.
- Wer Kundendaten verarbeitet, sollte beim KI-Anbieter auf einen Auftragsverarbeitungsvertrag achten oder ein lokales Modell nutzen (siehe [README](https://github.com/uzerwod-boop/zettelfrei#ki-einrichten)).

*Die Vorlagen und Hinweise (ArbZG, § 13b UStG, Mahnungen) sind Hilfen für den Alltag, keine Rechtsberatung.*

---

Zettelfrei ist kostenlos und Open Source (MIT oder Apache-2.0) · © 2026 Bohdan Sychevskyi
