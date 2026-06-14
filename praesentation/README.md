# Wie Schule wirkt – Keynote-Style-Präsentation

Diese Präsentation wurde auf Basis der Facharbeit **„Wie Schule wirkt – strukturelle
Unterschiede und ihr Effekt für das Wohlbefinden der Schüler"** (Seminarfachkurs
Glück, Gymnasium Damme, Svea Timphus, 2026) erstellt.

Alle Inhalte stammen ausschließlich aus der Facharbeit – es wurden keine neuen
Fakten oder Statistiken erfunden.

## Dateien

- `Wie_Schule_wirkt.pptx` – die fertige Präsentation (für Keynote auf iPad / Mac
  und PowerPoint)
- `Wie_Schule_wirkt.pdf` – PDF-Export zur Vorschau
- `preview/` – Bilder der einzelnen Folien zur schnellen Kontrolle
- `build_presentation.py` – Python-Skript, mit dem die `.pptx`-Datei generiert
  wurde (mit `python-pptx`)

## So öffnest du die Präsentation in Keynote (iPad)

1. Lade die Datei `Wie_Schule_wirkt.pptx` z. B. nach iCloud Drive, in OneDrive,
   Dropbox oder per AirDrop auf dein iPad.
2. Öffne die Datei in der **Keynote-App**.
3. Keynote fragt, ob die PowerPoint-Datei in das Keynote-Format konvertiert
   werden soll – tippe auf **„Kopie öffnen"**. Die Datei wird automatisch in eine
   echte `.key`-Datei umgewandelt.
4. Gehe danach kurz alle Folien durch und prüfe Schriften und Übergänge
   (siehe Hinweise unten).

## Folienübersicht

| # | Titel | Übergang (vorgesehen) | XML-Effekt im PPTX |
|---|---|---|---|
| 1 | Wie Schule wirkt (Titel) | Fade | `p:fade` |
| 2 | Warum ist das Thema wichtig? | Zoom | `p14:zoom` |
| 3 | Grundlagen | Move In | `p:push` |
| 4 | Das deutsche Schulsystem | Magic Move | `p159:morph` |
| 5 | Leistungsdruck und Wohlbefinden | Scale | `p14:zoom` |
| 6 | Schulklima | Push | `p:push` |
| 7 | Kurze Umfrage | Dissolve | `p:dissolve` |
| 8 | Internationaler Vergleich | Pop | `p14:zoom` |
| 9 | Stärken und Schwächen | Slide | `p:pull` |
| 10 | Reformansätze | Rotate | `p14:newsflash` |
| 11 | Fazit | Fade | `p:fade` |

## Hinweise zu den Übergängen in Keynote

Keynote verwendet eigene Bezeichnungen und 3D-Effekte, die im PPTX-Format nicht
1:1 abgebildet werden können. Beim Import nimmt Keynote die folgenden
Annäherungen automatisch vor:

| Wunsch (Keynote) | im PPTX hinterlegt | Empfehlung in Keynote setzen |
|---|---|---|
| Fade | Fade | **Auflösen** |
| Zoom | Zoom (PowerPoint) | **Fokus** oder **Skalieren** |
| Move In | Push | **Schieben** |
| Magic Move | Morph | **Zauberhafte Bewegung** |
| Scale | Zoom | **Skalieren** |
| Push | Push | **Schieben** |
| Dissolve | Dissolve | **Auflösen** |
| Pop | Zoom (Annäherung) | **Pop** |
| Slide | Pull | **Schieben** |
| Rotate | Newsflash (Annäherung) | **Drehen** |

So passt du sie in Keynote an (sehr schnell, dauert nur wenige Sekunden pro
Folie):

1. Tippe in Keynote in der seitlichen Folienübersicht auf eine Folie.
2. Tippe auf das **Übergang-Symbol** (zwei kreisförmige Pfeile) bzw.
   `… → Übergang`.
3. Wähle den gewünschten Keynote-Übergang aus der Liste oben.
4. Wiederhole das für jede Folie gemäß der Tabelle.

## Build-In-Animationen (Texte erscheinen nacheinander)

Damit die Stichpunkte in Keynote zeitversetzt erscheinen:

1. Folie öffnen → Stichpunkt-Block antippen → **„Animieren"**.
2. **„Build-In"** → Effekt **„Erscheinen"** oder **„Auflösen"** wählen.
3. Unter **„Lieferung"** **„Nach Absatz"** auswählen, damit jeder Punkt einzeln
   eingeblendet wird.
4. Bei den Hervorhebungs-Karten ggf. den Effekt **„Pop"** verwenden, um sie
   kurz zu betonen.

Diese Build-In-Animationen lassen sich nicht zuverlässig aus PowerPoint nach
Keynote übernehmen – einmal in Keynote eingestellt, sind sie aber dauerhaft
gespeichert.

## Sprechnotizen

Alle Folien enthalten ausführliche Moderationsnotizen (ca. 1 Minute pro
Folie). Sie sind direkt in den Notizen der jeweiligen Folie gespeichert.

In Keynote findest du sie unter:
**`… → Vortragsnotizen anzeigen`** (iPad) bzw. **`Darstellung → Vortragsnotizen
einblenden`** (Mac).

## Quelle

Die Inhalte basieren ausschließlich auf:

> Timphus, Svea (2026): *Wie Schule wirkt: Strukturelle Unterschiede und ihr
> Effekt für das Wohlbefinden der Schüler.* Facharbeit im Seminarfachkurs
> Glück, Gymnasium Damme.

## Erneutes Erzeugen (optional)

```bash
pip install python-pptx
python3 build_presentation.py
```

Die Datei `Wie_Schule_wirkt.pptx` wird im selben Ordner neu erzeugt.
