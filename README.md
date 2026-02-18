# Vimeo Transkript Generator

Liest Vimeo-Links aus einer Excel-Datei, lädt die Audio-Spuren herunter und transkribiert sie automatisch mit Whisper.

## Voraussetzungen

- **Python 3.10+** (https://www.python.org/downloads/)
- **ffmpeg** (wird von yt-dlp für Audio-Extraktion benötigt)

## Setup auf einem neuen Rechner (Schritt für Schritt)

### 1. Python installieren

Falls noch nicht vorhanden: https://www.python.org/downloads/  
Bei der Installation **"Add Python to PATH"** ankreuzen.

### 2. ffmpeg installieren

**Option A** (empfohlen, per Terminal):
```
winget install ffmpeg
```

**Option B** (manuell):
1. Gehe zu https://github.com/BtbN/FFmpeg-Builds/releases
2. Lade `ffmpeg-master-latest-win64-gpl.zip` herunter
3. Entpacke den Ordner
4. Kopiere `ffmpeg.exe` aus dem `bin/`-Ordner in dieses Projektverzeichnis (neben `transcribe.py`)

### 3. Projektordner kopieren

Kopiere den gesamten Ordner `Transkript Generator/` auf den Rechner (USB, OneDrive, etc.).

### 4. Excel-Datei bereitstellen

Kopiere `Übungen_Reihenfolge_Modul5_6_v5.xlsx` in den Projektordner (neben `transcribe.py`).  
Alternativ: `--excel "C:\Pfad\zur\Datei.xlsx"` beim Starten angeben.

### 5. Dependencies installieren

Terminal (PowerShell oder CMD) öffnen, in den Projektordner navigieren:

```
cd "C:\Pfad\zum\Transkript Generator"
pip install -r requirements.txt
```

### 6. Script starten

```
python transcribe.py
```

Das war's. Das Script:
- liest die Excel ein (143 Vimeo-Links)
- lädt die Audio-Spuren herunter (~102 einzigartige Videos)
- transkribiert alles automatisch
- speichert das Ergebnis in `output/ergebnisse.xlsx`

## Startoptionen

| Befehl | Beschreibung |
|--------|-------------|
| `python transcribe.py` | Alles ausführen (Download + Transkription + Export) |
| `python transcribe.py --skip-download` | Nur Transkription (Audio muss in `output/audio/` liegen) |
| `python transcribe.py --skip-transcribe` | Nur Audio herunterladen (keine Transkription) |
| `python transcribe.py --model medium` | Anderes Modell (tiny/small/medium/large-v3) |
| `python transcribe.py --model large-v3` | Bestes Modell (braucht GPU, ~3 GB VRAM) |
| `python transcribe.py --excel "C:\...\Datei.xlsx"` | Eigener Excel-Pfad |

## Modell-Empfehlung

| Modell | Qualität | Geschwindigkeit (GPU) | Geschwindigkeit (CPU) | VRAM |
|--------|----------|----------------------|----------------------|------|
| tiny | ausreichend | sehr schnell | schnell | ~1 GB |
| small | gut | schnell | mittel (~2-3h) | ~1 GB |
| medium | sehr gut | mittel | langsam (~4-6h) | ~2 GB |
| large-v3 | exzellent | mittel (~20-30min) | sehr langsam | ~3 GB |

**Mit NVIDIA-GPU**: `large-v3` empfohlen (beste Qualität, ~20-30 Min. für alle Videos)  
**Ohne GPU (nur CPU)**: `small` empfohlen (gute Qualität, ~2-3 Stunden)

## Resume-Fähigkeit

Das Script speichert den Fortschritt in `output/progress.json`. Bei Abbruch (Ctrl+C, Stromausfall, etc.) einfach erneut starten -- bereits heruntergeladene und transkribierte Videos werden übersprungen.

## Output

Die Ergebnisse landen in `output/ergebnisse.xlsx` und `output/ergebnisse.csv` mit folgenden Spalten:

- **Order** -- Reihenfolge aus der Excel
- **Modul** -- Modulnummer
- **Bereich** -- z.B. "Hüfte"
- **Kategorie** -- z.B. "Hüftbeugung <90 Grad"
- **Uebung** -- Übungsname
- **Link_Typ** -- "kurz" oder "lang"
- **URL** -- Original Vimeo-Link
- **Vimeo_ID** -- Numerische ID
- **Audio_Dauer_s** -- Audiodauer in Sekunden
- **Transkript** -- Volltext der Transkription
- **Status** -- ok / download_fehlgeschlagen / transkription_fehlgeschlagen

## Projektstruktur

```
Transkript Generator/
├── transcribe.py          # Haupt-Script
├── requirements.txt       # Python-Dependencies
├── README.md              # Diese Anleitung
├── Übungen_...v5.xlsx     # Input-Excel (hierhin kopieren)
└── output/
    ├── audio/             # Heruntergeladene Audio-Dateien
    ├── progress.json      # Fortschritt (für Resume)
    ├── ergebnisse.xlsx    # Ergebnis-Excel
    └── ergebnisse.csv     # Ergebnis-CSV
```
