# PnP2PDF

[![tests](https://github.com/Dannoster/PnP2PDF/actions/workflows/tests.yml/badge.svg)](https://github.com/Dannoster/PnP2PDF/actions/workflows/tests.yml)

[English](README.md) · [Русский](README.ru.md) · [简体中文](README.zh-CN.md) · [Español](README.es.md) · **Deutsch**

PnP2PDF setzt Print-and-Play-Karten zu einem druckfertigen PDF zusammen. Ordner
mit Kartenbildern wählen, Papier und Raster einstellen — fertig ist ein PDF mit
Schnittlinien, Sicherheitszonen und, wenn nötig, gespiegelten Rückseiten auf der
Blattrückseite.

- ein Fenster, keine Kommandozeile nötig;
- Live-Vorschau der Bögen vor dem Druck;
- Karten ohne Rückseite, eine gemeinsame Rückseite, mehrere verschiedene
  Rückseiten und Karten mit eigener Rückseite;
- Standardwerte für A4 und Karten im Format 62,5 × 88,7 mm (Poker).

## Inhalt

- [Fertige App](#fertige-app)
- [Aus dem Quellcode starten](#aus-dem-quellcode-starten)
- [Schnellstart](#schnellstart)
- [Ordnerstruktur](#ordnerstruktur)
- [Einstellungen](#einstellungen)
- [Beispiele](#beispiele)
- [Kommandozeile](#kommandozeile)
- [Eigenständige App bauen](#eigenständige-app-bauen)
- [Wenn etwas schiefgeht](#wenn-etwas-schiefgeht)
- [Entwicklung](#entwicklung)
- [Bekannte Probleme](#bekannte-probleme)
- [Lizenz](#lizenz)

## Fertige App

Lade das Archiv für dein System von
[Releases](https://github.com/Dannoster/PnP2PDF/releases) herunter, entpacke es und
starte es. Python muss nicht installiert werden.

Die macOS-Builds sind nicht mit einer Apple Developer ID signiert, deshalb muss bei
der ersten Ausführung die Quarantäne aufgehoben werden:

```bash
xattr -dr com.apple.quarantine /Applications/PnP2PDF.app
open /Applications/PnP2PDF.app
```

## Aus dem Quellcode starten

Du brauchst Python 3.11 oder neuer.

```bash
git clone https://github.com/Dannoster/PnP2PDF.git
cd PnP2PDF
pip install -r requirements.txt
python PnP2PDF.py
```

Die Installation als Paket bringt zwei Befehle mit: `pnp2pdf` öffnet das Fenster,
`pnp2pdf-cli` erzeugt das PDF ohne Fenster.

```bash
pip install .
pnp2pdf
```

## Schnellstart

1. **Cards folder** — Ordner mit den Kartenbildern wählen. Die Vorschau rechts
   erscheint, sobald ein Ordner gewählt ist.
2. **Papier und Kartengröße** — Standard ist A4 (210 × 297 mm) und eine Karte mit
   62,5 × 88,7 mm. *Smaller size (98%)* druckt die Karten etwas kleiner, sodass
   mehr Papier dazwischen bleibt.
3. **Layout** — Zeilen × Spalten, der Abstand zwischen den Karten und optional
   Sicherheitszonen (Ränder, die der Drucker nicht erreicht).
4. **Double sided cards** — anhaken, wenn der Ordner Rückseiten enthält. Siehe
   [Ordnerstruktur](#ordnerstruktur): ohne Haken werden Rückseitenbilder als
   normale Karten gedruckt.
5. **Create PDF** — der Fortschritt läuft durch, das Fenster bleibt offen, die
   Statuszeile wird grün und zeigt den Dateinamen, und **Show PDF** zeigt die Datei
   im Finder, im Explorer oder im Dateimanager.

Das PDF landet in einem Ordner mit Zeitstempel neben den Karten. Mit *Save the PDF
to a different folder* wählst du einen anderen Ort.

## Ordnerstruktur

Der *Kartenordner* ist der Ordner, den du im Fenster wählst; alles Folgende liegt
darin. Unterstützt werden `.png`, `.jpg`, `.jpeg`, `.tif`, `.tiff`.

Bilder werden auf die Kartenfläche gestreckt — verwende also Bilder mit demselben
Seitenverhältnis wie die Karten.

### Karten ohne Rückseite

Einfach die Bilder in den Ordner legen, nichts anhaken:

```
my-cards/
├── card-01.png
├── card-02.png
└── card-03.png
```

### Eine Rückseite für alle Karten

Die Rückseite ist ein Bild, dessen Name mit der Größe in Millimetern endet; die
Karten dazu liegen in einem Unterordner mit demselben Namen **ohne** den
Größenteil:

```
my-cards/
├── red-back_63,5_88,9.png     ← die Rückseite, mit Größe im Namen
└── red-back/                  ← Karten, die damit gedruckt werden
    ├── card-01.png
    └── card-02.png
```

**Double sided cards** anhaken und bauen. Auf jeden Vorderseitenbogen folgt
derselbe Bogen mit Rückseiten, gespiegelt, damit beim beidseitigen Druck alles
deckungsgleich ist.

### Mehrere Rückseiten

Dasselbe Muster wiederholen: eine Rückseitendatei, ein Ordner. Karten
verschiedener Rückseiten dürfen unterschiedlich groß sein und werden gemeinsam
gesetzt.

```
my-cards/
├── red-back_63,5_88,9.png
├── blue-back_64,5_91,3.png
├── red-back/
│   └── card-01.png …
└── blue-back/
    └── card-01.png …
```

Das Größen-Suffix ist optional: ohne es wird die Rückseite auf die Kartengröße
gestreckt; mit ihm wird eine größere Rückseite auf das Kartenfenster beschnitten,
damit Beschnittzugabe und Passermarken nicht auf der Karte landen. Als
Dezimaltrennzeichen geht Komma oder Punkt: `63,5` und `63.5` sind gleichwertig.

### Karten mit eigener Rückseite

Wenn die Rückseite zu genau einer Karte gehört, lege beide Bilder in einen Ordner
namens `Double Sided`. Die Dateien werden alphabetisch paarweise zugeordnet:
Vorderseite, Rückseite, Vorderseite, Rückseite …

```
my-cards/
└── Double Sided/
    ├── 1-front.png
    ├── 2-back.png
    └── 3-front.png …
```

Eine ungerade Anzahl von Dateien wird mit einer Erklärung abgelehnt, weil die
Paarung sonst verrutscht.

### Karten ohne Rückseite in einem gemischten Satz

Dafür gibt es den Ordner `Single Sided`:

```
my-cards/
├── Double Sided/…
└── Single Sided/
    ├── card-01.png
    └── card-02.png
```

Alle vier Strukturen lassen sich in einem Ordner mischen und in einem Lauf bauen.

> **Den Haken nicht vergessen.** Die Ordnerstruktur sagt, *welche Rückseite zu
> welcher Karte gehört*; erst die Checkbox **Double sided cards** druckt diese
> Rückseiten. Ohne Haken werden Rückseitenbilder als normale Vorderseiten
> behandelt und landen neben den Karten auf dem Bogen.

## Einstellungen

| Einstellung | Standard | Hinweis |
| --- | --- | --- |
| Papiergröße | 210 × 297 mm | A4. Jede Größe in Millimetern. |
| Kartengröße | 62,5 × 88,7 mm | Pokergröße. Bilder werden auf dieses Rechteck gestreckt. |
| Smaller size (98%) | aus | Druckt jede Karte mit 98% der angegebenen Größe. |
| Zeilen × Spalten | 3 × 3 | Karten pro Bogen. |
| Abstand | 2 mm | Platz zwischen den Karten, dort laufen die Schnittlinien. |
| Sicherheitszone oben/unten | 0 mm | Ränder, die der Drucker nicht erreicht. |
| Sicherheitszone links/rechts | 0 mm | Dasselbe waagerecht. |
| Kartenordner | — | Der oben beschriebene Ordner. |
| PDF speichern unter | neben den Karten | Optional, ein anderer Ordner ist möglich. |
| Double sided cards | aus | Rückseiten der Karten drucken. |

Einstellungen werden zwischen Läufen in `presets.json` im Datenordner der App
gespeichert (Pfade unter [Wenn etwas schiefgeht](#wenn-etwas-schiefgeht)). Beim
Start aus dem Quellcode liegt die Datei neben dem Projekt.

## Beispiele

[`examples/`](examples/) enthält fertige Kartensätze für jede Struktur — jeweils
wenige Kilobyte — und `make_examples.py` zum Neuerzeugen. Mehr dazu in
[`examples/README.md`](examples/README.md).

| Ordner | Zeigt | Seiten |
| --- | --- | --- |
| `single-sided/` | neun Karten ohne Rückseite | 1 |
| `one-back/` | sechs Karten mit einer gemeinsamen Rückseite | 2 |
| `two-backs/` | zwei Rückseiten, eine größer als die Karten | 4 |
| `double-sided/` | gepaarte und einseitige Karten in einem Lauf | 2 |

```bash
python3 -m pnp2pdf.cli --cards examples/one-back --out /tmp/ex --double-sided --quiet
```

## Kommandozeile

`pnp2pdf-cli` erzeugt dieselben PDFs ohne Fenster — praktisch für Skripte und
Stapelverarbeitung.

```bash
pnp2pdf-cli --cards ~/cards --out ~/pdf
pnp2pdf-cli --cards ~/cards --out ~/pdf --double-sided
pnp2pdf-cli --cards ~/cards --preview /tmp/first-page.png
```

| Option | Bedeutung |
| --- | --- |
| `--cards DIR` | Ordner mit den Kartenbildern (nötig). |
| `--out DIR` | Zielordner des PDFs. Standard: neben den Karten. |
| `--paper WxH` | Papiergröße in mm, Standard `210x297`. |
| `--card WxH` | Kartengröße in mm, Standard `62.5x88.7`. |
| `--grid RxC` | Zeilen × Spalten, Standard `3x3`. |
| `--gap MM` | Abstand zwischen Karten in mm, Standard `2`. |
| `--safe-ud MM` | Sicherheitszone oben/unten in mm, Standard `0`. |
| `--safe-lr MM` | Sicherheitszone links/rechts in mm, Standard `0`. |
| `--smaller-size` | Karten mit 98% der angegebenen Größe drucken. |
| `--double-sided` | Rückseiten drucken. |
| `--preview PATH` | Erste Seite als PNG schreiben und beenden. |
| `--quiet` | Keine Fortschrittsausgabe. |
| `--version` | Version anzeigen und beenden. |

Exit-Codes: `0` Erfolg, `1` unerwarteter Fehler, `2` fehlerhafte Eingabe, `130`
abgebrochen.

## Eigenständige App bauen

Ein Befehl baut die App und repariert dabei ihre Umgebung:

```bash
./build_app.sh          # macOS und Linux
build_app.bat           # Windows
```

Das Skript legt `.venv` an, wenn es fehlt, installiert oder aktualisiert die
Abhängigkeiten, ersetzt eine kaputte oder fremde virtuelle Umgebung, lässt die
Tests laufen und packt die App. Ergebnis:

- macOS: `dist/PnP2PDF.app`
- Windows: `dist/PnP2PDF.exe`
- Linux: `dist/PnP2PDF`

`--no-tests` überspringt die Tests, `PNP2PDF_BUILD_DRY=1` prüft die Umgebung und
stoppt vor dem Packen. Unter macOS wird das Bundle nach dem Schreiben der Version
neu signiert; schlägt die Prüfung der Signatur fehl, bricht der Build mit einer
klaren Meldung ab.

## Wenn etwas schiefgeht

**Die gepackte App beendet sich oder tut nichts.** Fehler, die ein Fenster-Build
sonst verschluckt, landen in `crash.log`, und der Fehlerdialog nennt den Pfad.
Bitte diese Datei bei einer Meldung mitschicken.

- macOS: `~/Library/Application Support/PnP2PDF/crash.log`
- Windows: `%APPDATA%\PnP2PDF\crash.log`
- Linux: `$XDG_DATA_HOME/pnp2pdf/crash.log` (oder `~/.local/share/pnp2pdf/crash.log`)

**„PnP2PDF is damaged and can't be opened“ (macOS).** Die heruntergeladene,
unsignierte App ist unter Quarantäne — siehe [Fertige App](#fertige-app).

**„No card images found“.** Bilder müssen auf `.png`, `.jpg`, `.jpeg`, `.tif` oder
`.tiff` enden; Karten mit Rückseite gehören in den Unterordner mit dem Namen
dieser Rückseite.

**Die Rückseiten landeten als Karten auf dem Bogen.** Die Checkbox **Double sided
cards** ist nicht gesetzt — siehe [Ordnerstruktur](#ordnerstruktur).

**Die App öffnet sich als Ordner (macOS).** Bundle oder Startdatenbank sind
veraltet: mit `./build_app.sh` neu bauen, dann
`/System/Library/Frameworks/CoreServices.framework/Frameworks/LaunchServices.framework/Support/lsregister -f dist/PnP2PDF.app`
ausführen.

## Entwicklung

```bash
pip install -r requirements-dev.txt
pytest
```

Die Tests ersetzen CustomTkinter und laufen daher ohne Bildschirm. Aufbau des
Pakets:

```
pnp2pdf/
├── app.py          Ereignisschleife und Start
├── cards.py        Ordner durchsuchen, Karten und Rückseiten paaren
├── layout.py       Karten auf Bögen verteilen
├── render.py       Seite zeichnen: Karten, Schnittlinien, Ecken, Zonen
├── pdf.py          CMYK-PDF schreiben
├── preview.py      schnelles Rendern in niedriger Auflösung für die Vorschau
├── job.py          ein Lauf: Karten → Bögen → PDF
├── cli.py          Einstiegspunkt der Kommandozeile
└── ui/             Fenster, Formularhilfen, Arbeits-Thread
```

## Bekannte Probleme

- Der Windows-Build meldet manchmal nach dem Schließen des Fensters einen Fehler;
  das PDF ist dann schon geschrieben.
- Microsoft Defender meldet PyInstaller-Builds gelegentlich als unbekannte
  Software.
- Der macOS-Build ist unsigniert, daher der Quarantäne-Schritt. Signieren und
  Notarisieren brauchen eine Apple Developer ID.

## Lizenz

[MIT](LICENSE).
