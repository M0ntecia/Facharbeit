# Wie Schule wirkt – Keynote-Präsentation

Moderne, Apple-Keynote-artige Präsentation zur Facharbeit **„Wie Schule wirkt:
Strukturelle Unterschiede und ihr Effekt auf das Wohlbefinden der Schüler“**
(Seminarfachkurs Glück, Svea Timphus).

Alle Inhalte stammen **ausschließlich** aus der Facharbeit.

## Dateien

| Datei | Beschreibung |
|-------|--------------|
| `Glueck_Schule_Wohlbefinden.pptx` | Die fertige Präsentation (11 Folien, 16:9). Öffnet sich direkt in **Keynote** (iPad/Mac) und in PowerPoint. |
| `Sprechnotizen.md` | Alle Moderationsnotizen zum Nachlesen/Drucken (auch in der Datei als Präsentatornotizen hinterlegt). |
| `build_presentation.py` | Generator-Skript, das die `.pptx` erzeugt. |
| `assets/` | Pastellfarbene Illustrationen (Apple-Keynote-Stil). |

## In Keynote auf dem iPad öffnen

1. Die Datei `Glueck_Schule_Wohlbefinden.pptx` auf das iPad laden (iCloud Drive,
   Dateien-App, AirDrop oder E-Mail).
2. In der **Dateien**-App antippen und „In Keynote öffnen“ wählen.
   Keynote importiert die Datei automatisch und übernimmt Layout, Übergänge,
   Text-Animationen und Präsentatornotizen.
3. Die Präsentatornotizen erscheinen in Keynote unter *Bearbeiten ▸ Ansicht ▸
   Präsentatornotizen* bzw. im Referentenmodus.

## Design

- Farbpalette: Hellblau `#CFEFFF`, Pastellrosa `#FFD9E8`, Hellgelb `#FFF7C7`,
  Weiß, dunkelgraue Schrift.
- Große Überschriften, viel Weißraum, abgerundete Formen, dezente Schatten,
  hochwertige Illustrationen, abwechslungsreiche Layouts.
- Jede Folie hat einen eigenen Übergang sowie verzögerte Build-In-Animationen
  für die Texte.

## Neu erzeugen

```bash
pip install python-pptx Pillow
python3 build_presentation.py
```
