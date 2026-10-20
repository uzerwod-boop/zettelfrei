# Mitmachen

Schön, dass Sie helfen möchten! Zettelfrei ist für kleine Betriebe gedacht – **Rückmeldungen aus der Praxis sind genauso wertvoll wie Code.**

## Ohne Programmieren

- **Fehler melden:** [Issue „Fehler melden“](https://github.com/uzerwod-boop/zettelfrei/issues/new?template=fehler.yml) – bitte nur Beispieldateien mit erfundenen Daten.
- **Idee vorschlagen:** [Issue „Idee vorschlagen“](https://github.com/uzerwod-boop/zettelfrei/issues/new?template=idee.yml) – welche Aufgabe kostet Sie heute Zeit?
- **Beispieldatei beisteuern:** Eine typische Excel-Datei aus Ihrem Gewerk (anonymisiert) hilft, die Erkennung zu verbessern.

## Mit Code

```bash
git clone https://github.com/uzerwod-boop/zettelfrei.git
cd zettelfrei
npm start          # http://localhost:8787
npm test           # Engine, E-Rechnung, Sicherheit, Server
```

Für die Ende-zu-Ende-Tests: `pip install playwright openpyxl`, dann `python tests/browser_test.py` und `python tests/browser_v2_test.py`.

Grundsätze:

- **Keine Build-Schritte, keine neuen Abhängigkeiten** ohne guten Grund. Die Anwendung läuft als einfache HTML-Seite, Bibliotheken liegen in `vendor/`.
- **Die KI rechnet nicht.** Summen und Abfragen rechnet die Engine; die KI beschreibt nur, was zu tun ist.
- **Jede Änderung braucht eine Vorschau** „alt → neu“, bevor sie in die Tabelle geht.
- **Oberfläche auf Deutsch**, klar und ohne Fachjargon. Code-Kommentare gern auf Deutsch.
- Neue Funktion → neuer Test.

Pull Requests bitte klein halten und kurz beschreiben, welches Problem sie lösen.

Beiträge werden, sofern nicht anders angegeben, unter beiden Lizenzen des Projekts eingebracht (MIT oder Apache-2.0).
