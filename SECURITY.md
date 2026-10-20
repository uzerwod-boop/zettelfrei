# Sicherheit

## Lücke melden

Bitte melden Sie Sicherheitslücken **nicht** als öffentliches Issue, sondern vertraulich über
[„Report a vulnerability“](https://github.com/uzerwod-boop/zettelfrei/security/advisories/new) im Reiter *Security*.

Hilfreich sind: betroffene Version, Schritte zum Nachstellen und – falls möglich – eine Beispieldatei mit erfundenen Daten.
Ich melde mich in der Regel innerhalb von 7 Tagen und nenne Sie auf Wunsch in den Release-Notes.

## Unterstützte Versionen

| Version | Sicherheitsupdates |
|---|---|
| 0.2.x | ✅ |
| < 0.2 | ❌ |

## Wie Zettelfrei Ihre Daten schützt

- Excel-Dateien werden **im Browser** verarbeitet und nicht hochgeladen.
- An die KI geht nur, was Sie ausführen. Ohne KI funktioniert der größte Teil der Anwendung.
- Werte von KI und aus E-Rechnungen werden nie als Formel gespeichert (Schutz vor Formel-Einschleusung).
- Der optionale Server nimmt KI-Anfragen nur von der eigenen Oberfläche an und hält den API-Schlüssel vom Browser fern.

Details zu den Prüfungen in Version 0.2.1 stehen im [README](README.md#sicherheit-version-021).
