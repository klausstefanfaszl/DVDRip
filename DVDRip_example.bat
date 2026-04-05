@echo off
setlocal

:: ============================================================
:: DVDRip_example.bat
:: Beispiel-Wrapper fuer DVDRip.ps1
:: Token-Informationen hier NICHT eintragen - nur als Vorlage!
:: ============================================================

:: --- Pflichtparameter ---
set OUTPUT_DIR=D:\Videos

:: --- Optionale Telegram-Benachrichtigung ---
:: set TELEGRAM_TOKEN=<Bot-Token hier eintragen>
:: set TELEGRAM_CHAT_ID=<Chat-ID hier eintragen>

set SCRIPT=%~dp0DVDRip.ps1

powershell.exe -NoProfile -ExecutionPolicy Bypass -File "%SCRIPT%" ^
    -OutputDir "%OUTPUT_DIR%" ^
::
:: Disc-Titel manuell vorgeben (Standard: auto aus Disc-Inhalt):
::     -DiscTitle "MeinFilm" ^
::
:: Laufwerksindex wenn mehrere optische Laufwerke vorhanden (Standard: 0):
::     -DriveIndex 0 ^
::
:: Minimale Titellänge in Sekunden fuer MakeMKV (Standard: 1800 = 30 Min):
::     -MinLength 1800 ^
::
:: HandBrake-Preset (Standard: "H.265 MKV 576p25"):
::     -HBPreset "H.265 MKV 1080p" ^
::
:: HandBrake RF-Qualitaetsfaktor, kleiner = besser (Standard: 22):
::     -HBQuality 20 ^
::
:: Zusaetzliche HandBrake-Argumente:
::     -HBExtraArgs "--crop 0:0:0:0" ^
::
:: Eigenes Log-Verzeichnis (Standard: OutputDir):
::     -LogDir "C:\Logs" ^
::
:: Pfad zu makemkvcon.exe (Standard: auto-detect):
::     -MakeMKVPath "C:\Program Files (x86)\MakeMKV\makemkvcon.exe" ^
::
:: Pfad zu HandBrakeCLI.exe (Standard: C:\Program Files\HandbrakeCLI\HandBrakeCLI.exe):
::     -HandBrakePath "C:\Program Files\HandbrakeCLI\HandBrakeCLI.exe" ^
::
:: Telegram-Benachrichtigungen:
::     -TelegramToken "%TELEGRAM_TOKEN%" ^
::     -TelegramChatId "%TELEGRAM_CHAT_ID%" ^
::
:: Nur rippen, nicht kodieren:
::     -SkipEncode ^
::
:: Nur kodieren (MKV-Dateien bereits vorhanden):
::     -SkipRip ^
::
:: DVD nach Abschluss NICHT auswerfen (Standard: auswerfen):
::     -NoEject ^
::
:: Alle Log-Zeilen farbig auf Stdout ausgeben + HandBrake-Fortschritt live:
::     -DebugMode ^
::
    %*

exit /b %ERRORLEVEL%
