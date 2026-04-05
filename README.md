# DVDRip

PowerShell-Skript zum automatisierten Rippen und Komprimieren von DVDs.

## Voraussetzungen

| Tool | Pfad (Standard) |
|------|----------------|
| [MakeMKV](https://www.makemkv.com/) | `C:\Program Files (x86)\MakeMKV\makemkvcon.exe` oder `C:\Program Files\MakeMKV\makemkvcon.exe` (auto-detect) |
| [HandBrakeCLI](https://handbrake.fr/) | `C:\Program Files\HandbrakeCLI\HandBrakeCLI.exe` |

## Schnellstart

```powershell
# Minimal (alles automatisch, mit Debug-Ausgabe)
.\DVDRip.ps1 -OutputDir "D:\Videos" -DebugMode

# Über den mitgelieferten Wrapper
.\DVDRip.bat
```

## Parameter

| Parameter | Typ | Standard | Beschreibung |
|-----------|-----|----------|--------------|
| `-OutputDir` | String | **Pflicht** | Zielverzeichnis; Unterordner wird automatisch angelegt |
| `-DiscTitle` | String | auto | Ordnername; wird aus Disc-Inhalt ermittelt falls leer |
| `-DriveIndex` | Int | `0` | MakeMKV-Laufwerksindex |
| `-MinLength` | Int | `1800` | Minimale Titellänge in Sekunden (MakeMKV) |
| `-HBPreset` | String | `H.265 MP4 576p25` | HandBrakeCLI-Preset (MP4 → .mp4, MKV → .mkv) |
| `-HBQuality` | Int | `22` | HandBrakeCLI RF-Qualität (kleiner = besser) |
| `-HBExtraArgs` | String | – | Zusätzliche HandBrakeCLI-Argumente |
| `-LogDir` | String | OutputDir | Verzeichnis für Log-Dateien |
| `-TempDir` | String | `%TEMP%` | Lokales Zwischenverzeichnis für Rip und Encoding |
| `-MakeMKVPath` | String | auto | Pfad zu `makemkvcon.exe` |
| `-HandBrakePath` | String | siehe oben | Pfad zu `HandBrakeCLI.exe` |
| `-TelegramToken` | String | – | Telegram Bot-Token; leer = keine Benachrichtigung |
| `-TelegramChatId` | String | – | Telegram Chat-ID des Empfängers |
| `-DebugMode` | Switch | `false` | Ausgabe auf Stdout + HandBrake-Fortschritt live |
| `-SkipRip` | Switch | `false` | MakeMKV-Schritt überspringen |
| `-SkipEncode` | Switch | `false` | HandBrakeCLI-Schritt überspringen |
| `-NoEject` | Switch | `false` | DVD nach Abschluss **nicht** auswerfen |

## Beispiele

```powershell
# Nur rippen, nicht kodieren
.\DVDRip.ps1 -OutputDir "D:\Videos" -SkipEncode -DebugMode

# Nur kodieren (bereits gerippte Dateien)
.\DVDRip.ps1 -OutputDir "D:\Videos" -DiscTitle "MeinFilm" -SkipRip -DebugMode

# Mit Preset und Qualität, ohne automatischen Auswurf
.\DVDRip.ps1 -OutputDir "D:\Videos" -HBPreset "H.265 MKV 1080p" -HBQuality 20 -NoEject

# Mit Telegram-Benachrichtigung
.\DVDRip.ps1 -OutputDir "D:\Videos" -TelegramToken "..." -TelegramChatId "..."
```

## Verzeichnisstruktur

```
OutputDir/
  <DiscTitle>/
    RAW/          <- MKV-Rohdateien von MakeMKV
    Encoded/      <- Komprimierte Dateien von HandBrakeCLI
```

## Logging

- Log-Datei: `<LogDir>/DVDRip_<Timestamp>.log`
- Ohne `-DebugMode`: keine Stdout-Ausgabe außer der finalen Erfolgsmeldung
- Mit `-DebugMode`: alle Log-Zeilen farbig auf Stdout, HandBrake-Fortschritt live

## Exit-Codes

| Code | Bedeutung |
|------|-----------|
| 0 | Erfolg |
| 1 | Voraussetzungen fehlen (MakeMKV/HandBrake nicht gefunden) |
| 2 | MakeMKV-Fehler oder keine MKV-Dateien erzeugt |
| 3 | HandBrakeCLI-Fehler |
| 99 | Unerwarteter Fehler |

## Dateien

| Datei | Beschreibung |
|-------|-------------|
| `DVDRip.ps1` | Hauptskript |
| `DVDRip.bat` | Einsatzbereiter Wrapper mit allen projektspezifischen Einstellungen |
| `DVDRip_example.bat` | Vorlage mit kommentierten Parameterbeispielen (kein Token) |
| `DVDRip.html` | Vollständige HTML-Dokumentation |
