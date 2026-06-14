# Facharbeit – Keynote-Präsentation

Diese Präsentation fasst die Facharbeit „Wie Schule wirkt – Strukturelle
Unterschiede und ihr Effekt auf das Wohlbefinden von Schülerinnen und Schülern“
(Seminarfachkurs Glück, Svea Timphus, 2026) für einen ca. 15-minütigen Vortrag
zusammen.

## Dateien

| Datei | Inhalt |
| ----- | ------ |
| `Wie_Schule_wirkt_Praesentation.pptx` | Fertige Präsentation (11 Folien, mit Übergängen, Animationen und Sprechnotizen). Auf dem iPad einfach in **Keynote** öffnen – Keynote konvertiert die Datei automatisch. |
| `build_presentation.py` | Python-Skript, das die Präsentation aus den Inhalten der Facharbeit erzeugt (Layout, Farben, Illustrationen, Übergänge, Animationen, Sprechnotizen). |

## Foliendesign

* Pastellige Farbpalette: Hellblau `#CFEFFF`, Pastellrosa `#FFD9E8`,
  Hellgelb `#FFF7C7`, Weiß und dunkelgraue Schrift `#2B2D42`.
* Modernes, minimalistisches Layout im Apple-Keynote-Stil mit viel Weißraum,
  abgerundeten Karten, sanften Schatten und dezenten Hintergrundformen.
* Auf jeder Folie eine andere Anordnung – Karten, Mindmap, Zwei-Spalten-Layout,
  Kreise, Buttons usw.
* Statt Cliparts werden Symbole und einfache, aus Grundformen komponierte
  Illustrationen (Schule, Schülerinnen/Schüler, Waage, Pfeile) verwendet.

## Folienübergänge (je Folie ein anderer Übergang)

| # | Folie | Übergang (Keynote) | PPT-Mapping |
| - | ----- | ------------------ | ----------- |
| 1 | Titel | Fade | `fade` |
| 2 | Warum wichtig? | Zoom | `p14:zoom in` |
| 3 | Grundlagen | Move In | `cover` |
| 4 | Schulsystem | Magic Move | `p159:morph` |
| 5 | Leistungsdruck | Scale | `p14:zoom out` |
| 6 | Schulklima | Push | `push` |
| 7 | Umfrage | Dissolve | `dissolve` |
| 8 | International | Pop | `p14:flash` / `split` |
| 9 | Stärken & Schwächen | Slide | `push` |
| 10 | Reformen | Rotate | `p14:ferris` |
| 11 | Fazit | Fade | `fade` |

Außerdem wurde jeder Folie eine Build-In-Animation hinzugefügt: Texte,
Karten und Illustrationen erscheinen nacheinander mit einem leichten
Fade-In, sodass die Folien beim Vortrag schrittweise aufgebaut werden.

## Sprechnotizen

Zu jeder Folie gibt es ausführliche Moderationsnotizen (~1 Minute Sprechzeit),
sodass die gesamte Präsentation ca. 15 Minuten umfasst. Die Notizen sind in
der Keynote-Ansicht „Moderatornotizen“ sichtbar.

## Inhaltsquelle

Alle Aussagen stammen ausschließlich aus der bereitgestellten Facharbeit
(`Facharbeit_02db.pdf`). Es wurden keine zusätzlichen Fakten oder Statistiken
ergänzt.

## Neu bauen

```bash
pip install python-pptx lxml
python3 build_presentation.py
```

Das Skript erzeugt eine neue `Wie_Schule_wirkt_Praesentation.pptx` im
Projektverzeichnis.
