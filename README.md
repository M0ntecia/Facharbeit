# Facharbeit – Präsentation

Keynote-Style Präsentation zum Seminarfachkurs **Glück**.

## Dateien

| Datei | Beschreibung |
|-------|-------------|
| `presentations/Wie_Schule_wirkt_Keynote.pptx` | Hauptpräsentation (11 Folien, iPad/Keynote-kompatibel) |
| `presentations/SPRECHNOTIZEN.md` | Ausführliche Moderationsnotizen (~1 Min./Folie) |
| `presentations/KEYNOTE_ANLEITUNG.md` | Übergänge & Build-In-Animationen für Keynote |
| `presentations/generate_presentation.py` | Generator-Skript |

## Import auf dem iPad

1. `Wie_Schule_wirkt_Keynote.pptx` per AirDrop oder iCloud auf das iPad übertragen
2. In **Keynote** öffnen und ggf. in Keynote konvertieren
3. Sprechernotizen sind bereits in jeder Folie eingebettet
4. Für die exakten Keynote-Übergänge siehe `KEYNOTE_ANLEITUNG.md`

## Präsentation neu generieren

```bash
pip install python-pptx pillow
python presentations/generate_presentation.py
```
