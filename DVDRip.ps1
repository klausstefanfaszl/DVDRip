<#
.SYNOPSIS
    DVDRip - Automatisches Rippen und Komprimieren von DVDs

.DESCRIPTION
    Liest eine eingelegte DVD mit MakeMKV ein und komprimiert die Dateien
    anschließend mit HandBrakeCLI. Unterstützt Logging, Debug-Modus und
    flexible Parameterisierung.

.PARAMETER OutputDir
    Zielverzeichnis (Pflicht). Darunter wird ein Unterordner pro DVD angelegt.

.PARAMETER DiscTitle
    Name des DVD-Unterordners. Wird automatisch aus dem Disc-Inhalt ermittelt,
    falls nicht angegeben.

.PARAMETER DriveIndex
    MakeMKV-Laufwerksindex (Standard: 0 = erstes optisches Laufwerk).

.PARAMETER MinLength
    Minimale Titellänge in Sekunden für MakeMKV (Standard: 1800 = 30 Min).

.PARAMETER HBPreset
    HandBrakeCLI-Preset, z.B. "H.265 MKV 1080p". Leer = kein Preset.

.PARAMETER HBQuality
    HandBrakeCLI RF-Qualitätsfaktor (Standard: 22, kleiner = besser).

.PARAMETER HBExtraArgs
    Zusätzliche HandBrakeCLI-Argumente als einzelner String.

.PARAMETER LogDir
    Verzeichnis für die Log-Datei (Standard: OutputDir).

.PARAMETER DebugMode
    Schalter: Ausgabe auf Stdout zusätzlich zur Log-Datei aktivieren.

.PARAMETER SkipRip
    MakeMKV-Schritt überspringen (wenn MKV-Dateien bereits vorhanden).

.PARAMETER SkipEncode
    HandBrakeCLI-Schritt überspringen (nur rippen, nicht kodieren).

.PARAMETER MakeMKVPath
    Pfad zur makemkvcon.exe (Standard: auto-detect).

.PARAMETER TempDir
    Lokales Zwischenverzeichnis für Rip und Encoding (Standard: %TEMP%). Rip und Encode
    laufen immer lokal, danach werden die kodierten Dateien auf OutputDir verschoben und
    das Temp-Verzeichnis gelöscht. Bei erneutem Start wird geprüft ob MKV-Dateien bereits
    im TempDir vorhanden sind — falls ja, wird das Rippen übersprungen und direkt mit der
    Kodierung fortgefahren.

.PARAMETER HandBrakePath
    Pfad zur HandBrakeCLI.exe (Standard: C:\Program Files\HandbrakeCLI\HandBrakeCLI.exe).

.PARAMETER TelegramToken
    Telegram Bot-Token für Benachrichtigungen. Leer = keine Benachrichtigung.

.PARAMETER TelegramChatId
    Telegram Chat-ID des Empfängers.

.PARAMETER NoEject
    DVD nach Abschluss NICHT auswerfen. Standard: DVD wird automatisch ausgeworfen.

.PARAMETER ServiceMode
    Dauerbetrieb: Nach jeder Konvertierung wartet das Skript auf die nächste DVD.
    Erkennt DVD-Einlegen per WMI-Event; als Fallback wird alle PollInterval Minuten geprüft.

.PARAMETER PollInterval
    Polling-Intervall in Minuten im Service-Modus (Standard: 5).

.PARAMETER MinMkvSizeMB
    Minimale Dateigröße einer MKV-Datei in MB, ab der ein fehlerhafter MakeMKV-Lauf
    trotzdem als verwertbar gilt und HandBrakeCLI gestartet wird (Standard: 500).
    Dateien unterhalb dieser Größe gelten bei MakeMKV-Fehlern als unbrauchbar.

.PARAMETER Config
    Pfad zu einer JSON-Konfig-Datei. Alle dort definierten Parameter werden als
    Standardwerte verwendet. Explizit per CLI übergebene Parameter haben Vorrang.

.EXAMPLE
    .\DVDRip.ps1 -OutputDir "D:\Videos" -DebugMode
    .\DVDRip.ps1 -OutputDir "D:\Videos" -DiscTitle "MeinFilm" -HBPreset "H.265 MKV 1080p" -HBQuality 20
    .\DVDRip.ps1 -OutputDir "D:\Videos" -SkipEncode -DebugMode
#>

[CmdletBinding()]
param(
    [Parameter(Mandatory = $false)]
    [string]$Config = "",

    [Parameter(Mandatory = $false)]
    [string]$OutputDir = "",

    [Parameter(Mandatory = $false)]
    [string]$DiscTitle = "",

    [Parameter(Mandatory = $false)]
    [int]$DriveIndex = 0,

    [Parameter(Mandatory = $false)]
    [int]$MinLength = 1800,

    [Parameter(Mandatory = $false)]
    [string]$HBPreset = "HQ 576p25 Surround",

    [Parameter(Mandatory = $false)]
    [int]$HBQuality = 22,

    [Parameter(Mandatory = $false)]
    [string]$HBExtraArgs = "",

    [Parameter(Mandatory = $false)]
    [string]$LogDir = "",

    [Parameter(Mandatory = $false)]
    [switch]$DebugMode,

    [Parameter(Mandatory = $false)]
    [switch]$SkipRip,

    [Parameter(Mandatory = $false)]
    [switch]$SkipEncode,

    [Parameter(Mandatory = $false)]
    [switch]$SkipDiscScan,

    [Parameter(Mandatory = $false)]
    [string]$MakeMKVPath = "",

    [Parameter(Mandatory = $false)]
    [string]$TempDir = $env:TEMP,

    [Parameter(Mandatory = $false)]
    [string]$HandBrakePath = "C:\Program Files\HandbrakeCLI\HandBrakeCLI.exe",

    [Parameter(Mandatory = $false)]
    [string]$TelegramToken = "",

    [Parameter(Mandatory = $false)]
    [string]$TelegramChatId = "",

    [Parameter(Mandatory = $false)]
    [switch]$NoEject,

    [Parameter(Mandatory = $false)]
    [switch]$ServiceMode,

    [Parameter(Mandatory = $false)]
    [int]$PollInterval = 5,

    [Parameter(Mandatory = $false)]
    [int]$MinMkvSizeMB = 500
)

Set-StrictMode -Version Latest
$ErrorActionPreference = "Stop"

# ---------------------------------------------------------------------------
# Interne Variablen
# ---------------------------------------------------------------------------
$Script:LogFile          = $null
$Script:ExitCode         = 0
$Script:StartTime        = Get-Date
$Script:ScriptName       = [System.IO.Path]::GetFileNameWithoutExtension($MyInvocation.MyCommand.Name)
$Script:DiscSCSIError    = $false   # gesetzt wenn info-Scan SCSI-Fehler hatte
$Script:LastDiscTitle    = ""       # zuletzt verarbeiteter Disc-Titel (für Service-Modus)

# ---------------------------------------------------------------------------
# Konfig-Datei laden (falls angegeben; CLI-Parameter haben Vorrang)
# ---------------------------------------------------------------------------
if ($Config -eq "") {
    $autoConfig = Join-Path (Get-Location).Path "$Script:ScriptName.json"
    if (Test-Path $autoConfig) {
        $Config = $autoConfig
        Write-Host "Konfig-Datei automatisch geladen: $Config" -ForegroundColor Cyan
    }
}

if ($Config -ne "") {
    if (-not (Test-Path $Config)) {
        Write-Host "[ERROR] Konfig-Datei nicht gefunden: $Config" -ForegroundColor Red
        exit 1
    }
    try {
        $cfg = Get-Content -Path $Config -Raw -Encoding UTF8 | ConvertFrom-Json
    } catch {
        Write-Host "[ERROR] Konfig-Datei ungültig (kein gültiges JSON): $_" -ForegroundColor Red
        exit 1
    }
    foreach ($key in $cfg.PSObject.Properties.Name) {
        if ($key -eq "Config") { continue }
        if (-not $PSBoundParameters.ContainsKey($key)) {
            $val = $cfg.$key
            switch ($key) {
                "OutputDir"      { $script:OutputDir      = [string]$val }
                "DiscTitle"      { $script:DiscTitle      = [string]$val }
                "DriveIndex"     { $script:DriveIndex     = [int]$val }
                "MinLength"      { $script:MinLength      = [int]$val }
                "HBPreset"       { $script:HBPreset       = [string]$val }
                "HBQuality"      { $script:HBQuality      = [int]$val }
                "HBExtraArgs"    { $script:HBExtraArgs    = [string]$val }
                "LogDir"         { $script:LogDir         = [string]$val }
                "TempDir"        { $script:TempDir        = [string]$val }
                "MakeMKVPath"    { $script:MakeMKVPath    = [string]$val }
                "HandBrakePath"  { $script:HandBrakePath  = [string]$val }
                "TelegramToken"  { $script:TelegramToken  = [string]$val }
                "TelegramChatId" { $script:TelegramChatId = [string]$val }
                "PollInterval"   { $script:PollInterval   = [int]$val }
                "MinMkvSizeMB"  { $script:MinMkvSizeMB  = [int]$val }
                "DebugMode"      { if ([bool]$val) { $script:DebugMode    = [switch]::Present } }
                "SkipRip"        { if ([bool]$val) { $script:SkipRip      = [switch]::Present } }
                "SkipEncode"     { if ([bool]$val) { $script:SkipEncode   = [switch]::Present } }
                "SkipDiscScan"   { if ([bool]$val) { $script:SkipDiscScan = [switch]::Present } }
                "NoEject"        { if ([bool]$val) { $script:NoEject      = [switch]::Present } }
                "ServiceMode"    { if ([bool]$val) { $script:ServiceMode  = [switch]::Present } }
            }
        }
    }
}

if ($OutputDir -eq "") {
    Write-Host "[ERROR] -OutputDir ist Pflicht – entweder als Parameter oder in der Konfig-Datei." -ForegroundColor Red
    exit 1
}

# ---------------------------------------------------------------------------
# Logging-Funktionen
# ---------------------------------------------------------------------------
function Initialize-Log {
    param([string]$Dir)

    $timestamp = Get-Date -Format "yyyyMMdd_HHmmss"
    $logFileName = "$Script:ScriptName`_$timestamp.log"

    if (-not (Test-Path $Dir)) {
        New-Item -ItemType Directory -Path $Dir -Force | Out-Null
    }

    $Script:LogFile = Join-Path $Dir $logFileName
    Write-Log "INFO" "Log gestartet: $Script:LogFile"
    Write-Log "INFO" "Parameter: OutputDir='$OutputDir' DiscTitle='$DiscTitle' DriveIndex=$DriveIndex MinLength=$MinLength HBPreset='$HBPreset' HBQuality=$HBQuality DebugMode=$DebugMode SkipRip=$SkipRip SkipEncode=$SkipEncode"
}

function Write-Log {
    param(
        [ValidateSet("INFO","DEBUG","WARN","ERROR")]
        [string]$Level,
        [string]$Message
    )

    $ts   = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    $line = "[$ts] [$Level] $Message"

    # Immer in Log-Datei schreiben
    if ($Script:LogFile) {
        Add-Content -Path $Script:LogFile -Value $line -Encoding UTF8
    }

    # Nur im Debug-Modus auf Stdout
    if ($DebugMode) {
        switch ($Level) {
            "ERROR" { Write-Host $line -ForegroundColor Red    }
            "WARN"  { Write-Host $line -ForegroundColor Yellow }
            "DEBUG" { Write-Host $line -ForegroundColor Cyan   }
            default { Write-Host $line }
        }
    }
}

# ---------------------------------------------------------------------------
# Telegram-Benachrichtigung
# ---------------------------------------------------------------------------
function Send-TelegramMessage {
    param([string]$Text)

    if (-not $TelegramToken -or -not $TelegramChatId) { return }

    try {
        $url  = "https://api.telegram.org/bot$TelegramToken/sendMessage"
        $body = @{ chat_id = $TelegramChatId; text = $Text; parse_mode = "HTML" }
        Invoke-RestMethod -Uri $url -Method Post -Body $body -TimeoutSec 10 | Out-Null
        Write-Log "INFO" "Telegram-Nachricht gesendet: $Text"
    } catch {
        Write-Log "WARN" "Telegram-Benachrichtigung fehlgeschlagen: $($_.Exception.Message)"
    }
}

# ---------------------------------------------------------------------------
# Hilfsfunktionen
# ---------------------------------------------------------------------------
function Eject-Disc {
    param([string]$DriveLetter)
    if (-not $DriveLetter) { return }
    try {
        $shell = New-Object -ComObject Shell.Application
        $shell.NameSpace(17).ParseName($DriveLetter).InvokeVerb("Eject")
        Write-Log "INFO" "DVD ausgeworfen ($DriveLetter)."
    } catch {
        Write-Log "WARN" "DVD konnte nicht ausgeworfen werden: $($_.Exception.Message)"
    }
}

function Find-MakeMKV {
    # Bekannte Installationspfade
    $candidates = @(
        "C:\Program Files (x86)\MakeMKV\makemkvcon.exe",
        "C:\Program Files\MakeMKV\makemkvcon.exe"
    )
    foreach ($c in $candidates) {
        if (Test-Path $c) { return $c }
    }
    # PATH durchsuchen
    $found = Get-Command "makemkvcon.exe" -ErrorAction SilentlyContinue
    if ($found) { return $found.Source }
    return $null
}

function Invoke-ExternalProcess {
    param(
        [string]$Executable,
        [string[]]$Arguments,
        [string]$StepName,
        [switch]$StreamProgress   # Ausgabe live streamen (z.B. für HandBrakeCLI)
    )

    Write-Log "INFO" "Starte $StepName`: $Executable $($Arguments -join ' ')"

    if ($StreamProgress) {
        $psi = [System.Diagnostics.ProcessStartInfo]::new()
        $psi.FileName           = $Executable
        $psi.Arguments          = $Arguments -join ' '
        $psi.RedirectStandardOutput = $true
        $psi.RedirectStandardError  = $true
        $psi.UseShellExecute    = $false
        $psi.CreateNoWindow     = $true

        $proc  = [System.Diagnostics.Process]::new()
        $proc.StartInfo = $psi

        $queue = [System.Collections.Concurrent.ConcurrentQueue[string]]::new()

        $outSub = Register-ObjectEvent -InputObject $proc -EventName OutputDataReceived -MessageData $queue -Action {
            $q = $Event.MessageData
            if ($null -ne $EventArgs.Data -and $EventArgs.Data) {
                $q.Enqueue("[STDOUT] $($EventArgs.Data)")
            }
        }
        $errSub = Register-ObjectEvent -InputObject $proc -EventName ErrorDataReceived -MessageData $queue -Action {
            $q = $Event.MessageData
            if ($null -ne $EventArgs.Data) {
                $clean = $EventArgs.Data.TrimStart("`r").Trim()
                if ($clean) { $q.Enqueue("[STDERR] $clean") }
            }
        }

        $outLines = [System.Collections.Generic.List[string]]::new()
        $errLines = [System.Collections.Generic.List[string]]::new()

        $proc.Start()            | Out-Null
        $proc.BeginOutputReadLine()
        $proc.BeginErrorReadLine()

        while (-not $proc.HasExited) {
            $item = $null
            while ($queue.TryDequeue([ref]$item)) {
                Write-Log "DEBUG" "  $item"
                if ($item.StartsWith("[STDOUT]")) { $outLines.Add($item.Substring(9)) }
                else                              { $errLines.Add($item.Substring(9)) }
            }
            Start-Sleep -Milliseconds 500
        }

        # Restliche Ausgabe nach Prozessende leeren
        Start-Sleep -Milliseconds 200
        $item = $null
        while ($queue.TryDequeue([ref]$item)) {
            Write-Log "DEBUG" "  $item"
            if ($item.StartsWith("[STDOUT]")) { $outLines.Add($item.Substring(9)) }
            else                              { $errLines.Add($item.Substring(9)) }
        }

        $exitCode = $proc.ExitCode
        Unregister-Event -SourceIdentifier $outSub.Name -ErrorAction SilentlyContinue
        Unregister-Event -SourceIdentifier $errSub.Name -ErrorAction SilentlyContinue
        Remove-Job $outSub, $errSub -ErrorAction SilentlyContinue
        $proc.Dispose()

        return @{ ExitCode = $exitCode; Output = $outLines.ToArray(); Errors = $errLines.ToArray() }
    }

    # Standard: gepuffert über Temp-Dateien
    $tmpOut = [System.IO.Path]::GetTempFileName()
    $tmpErr = [System.IO.Path]::GetTempFileName()

    try {
        $proc = Start-Process `
            -FilePath $Executable `
            -ArgumentList ($Arguments -join ' ') `
            -RedirectStandardOutput $tmpOut `
            -RedirectStandardError  $tmpErr `
            -NoNewWindow `
            -PassThru `
            -Wait

        $outLines = @(Get-Content $tmpOut -Encoding UTF8 -ErrorAction SilentlyContinue)
        $errLines = @(Get-Content $tmpErr -Encoding UTF8 -ErrorAction SilentlyContinue)
        $exitCode = $proc.ExitCode
    } finally {
        Remove-Item $tmpOut, $tmpErr -ErrorAction SilentlyContinue
    }

    foreach ($l in $outLines) {
        if ($l) { Write-Log "DEBUG" "  [STDOUT] $l" }
    }
    foreach ($l in $errLines) {
        if ($l) { Write-Log "DEBUG" "  [STDERR] $l" }
    }

    return @{
        ExitCode = $exitCode
        Output   = $outLines
        Errors   = $errLines
    }
}

function Get-DiscTitle {
    param([string]$MkvCon, [int]$Index)

    Write-Log "INFO" "Ermittle Disc-Titel von Laufwerk $Index ..."

    $result = Invoke-ExternalProcess `
        -Executable $MkvCon `
        -Arguments @("--robot", "info", "disc:$Index") `
        -StepName "MakeMKV info"

    # Erste DRV-Zeile mit nicht-leerem Disc-Label gewinnt (auch wenn spätere Zeilen das Label verlieren)
    $title       = ""
    $discVisible = $false
    $tcount      = -1
    foreach ($line in $result.Output) {
        if ($line -match '^TCOUNT:(\d+)') {
            $tcount = [int]$Matches[1]
        }
        # DRV:index,visible,...,"drive_name","disc_label","device"
        # visible >= 2 = Disc eingelegt; nur erste nicht-leere Label-Zeile verwenden
        if (-not $title -and $line -match '^DRV:\d+,(\d+),\d+,\d+,"[^"]*","([^"]+)","([A-Z]:)"') {
            if ([int]$Matches[1] -ge 2) {
                $discVisible = $true
                $title = $Matches[2].Trim()
                $Script:DriveLetter = $Matches[3]
            }
        }
    }

    # Kein Disc: weder Label noch sichtbares Laufwerk, und TCOUNT=0
    if (-not $discVisible -and $tcount -eq 0) {
        Write-Log "ERROR" "Kein Disc in Laufwerk $Index erkannt (TCOUNT:0, kein Disc-Label)."
        return $null   # $null = kein Disc, Abbruch in Main
    }

    if ($tcount -eq 0) {
        $Script:DiscSCSIError = $true
        Write-Log "WARN" "Disc erkannt ('$title'), aber TCOUNT:0 - SCSI-Fehler oder Leseproblem. Rip wird dennoch versucht."
    } elseif (-not $title) {
        Write-Log "WARN" "Disc-Titel konnte nicht ermittelt werden, Fallback wird verwendet."
    } else {
        Write-Log "INFO" "Erkannter Disc-Titel: '$title'"
    }

    return $title
}

function Sanitize-FolderName {
    param([string]$Name)
    # Unerlaubte Zeichen in Windows-Ordnernamen entfernen
    $invalid = [System.IO.Path]::GetInvalidFileNameChars()
    $clean   = $Name
    foreach ($ch in $invalid) {
        $clean = $clean.Replace([string]$ch, "_")
    }
    $clean = $clean.Trim(" ._")
    return $clean
}

# ---------------------------------------------------------------------------
# Hauptlogik
# ---------------------------------------------------------------------------
function Main {

    # 1. Log initialisieren (immer lokal, am Ende auf OutputDir kopieren)
    $logBase = if ($LogDir) { $LogDir } else { $PSScriptRoot }
    Initialize-Log -Dir $logBase

    Write-Log "INFO" "=== $Script:ScriptName gestartet ==="
    Send-TelegramMessage "DVD-Import gestartet$(if ($DiscTitle) { ": <b>$DiscTitle</b>" })"

    # 2. Voraussetzungen prüfen
    Write-Log "INFO" "Prüfe Voraussetzungen ..."

    if (-not $SkipRip) {
        if ($MakeMKVPath -eq "") {
            $MakeMKVPath = Find-MakeMKV
        }
        if (-not $MakeMKVPath -or -not (Test-Path $MakeMKVPath)) {
            Write-Log "ERROR" "makemkvcon.exe nicht gefunden. Bitte -MakeMKVPath angeben oder MakeMKV installieren."
            $Script:ExitCode = 1
            return
        }
        Write-Log "INFO" "MakeMKV gefunden: $MakeMKVPath"
    }

    if (-not $SkipEncode) {
        if (-not (Test-Path $HandBrakePath)) {
            Write-Log "ERROR" "HandBrakeCLI.exe nicht gefunden unter: $HandBrakePath"
            $Script:ExitCode = 1
            return
        }
        Write-Log "INFO" "HandBrakeCLI gefunden: $HandBrakePath"
    }

    # 3. Disc-Titel bestimmen
    $folderName = $DiscTitle
    if (-not $folderName -and -not $SkipRip -and -not $SkipDiscScan) {
        $detectedTitle = Get-DiscTitle -MkvCon $MakeMKVPath -Index $DriveIndex
        if ($null -eq $detectedTitle) {
            # Get-DiscTitle hat $null zurückgegeben = kein Disc / unlesbar
            Send-TelegramMessage "DVD-Import FEHLER: Kein Disc in Laufwerk $DriveIndex erkannt."
            $Script:ExitCode = 2
            return
        }
        $folderName = $detectedTitle
    } elseif ($SkipDiscScan -and -not $folderName -and -not $SkipRip) {
        Write-Log "INFO" "DiscScan übersprungen (-SkipDiscScan), Disc-Titel wird aus MKV-Output gelesen oder Fallback verwendet."
    }

    if (-not $folderName) {
        $folderName = "DVD_$(Get-Date -Format 'yyyyMMdd_HHmmss')"
        Write-Log "WARN" "Kein Disc-Titel ermittelt, verwende Fallback-Namen: '$folderName'"
    }

    $folderName = Sanitize-FolderName -Name $folderName
    $Script:LastDiscTitle = $folderName
    Write-Log "INFO" "Zielordner-Name: '$folderName'"

    # 4. Verzeichnisse anlegen
    # Kodierte Dateien landen immer direkt in <OutputDir>/<DiscTitle>/ (kein Encoded-Unterordner).
    # Bei Netzwerkpfad (UNC \\...): Rip lokal im TempDir, danach Move auf OutputDir.
    $isNetworkPath = $OutputDir -match '^\\\\'
    $discDir = Join-Path $OutputDir $folderName
    if ($isNetworkPath) {
        $ripDir   = Join-Path $TempDir "$folderName\RAW"
        $localDir = Join-Path $TempDir $folderName   # temporäres Kodierziel
        $finalDir = $discDir
        Write-Log "INFO" "Netzwerkpfad erkannt - verwende TempDir: $TempDir"
    } else {
        $ripDir   = Join-Path $discDir "RAW"
        $localDir = $discDir
        $finalDir = $discDir
        Write-Log "INFO" "Lokales Zielverzeichnis - schreibe direkt nach: $discDir"
        if (-not $SkipEncode -and -not (Test-Path $discDir)) {
            New-Item -ItemType Directory -Path $discDir -Force | Out-Null
            Write-Log "INFO" "Zielverzeichnis erstellt: $discDir"
        }
    }

    # 5. DVD rippen mit MakeMKV
    if (-not $SkipRip) {
        $doRip     = $true
        $ripMarker = Join-Path $ripDir "_rip_complete"

        if (Test-Path $ripMarker) {
            # Vollständiger Rip vorhanden → direkt zur Kodierung
            $existingMkvs = @(Get-ChildItem -Path $ripDir -Filter "*.mkv" -Recurse -ErrorAction SilentlyContinue)
            Write-Log "INFO" "Vollständiger Rip gefunden ($($existingMkvs.Count) MKV-Datei(en)) - überspringe Rippen, starte Kodierung."
            Send-TelegramMessage "DVD <b>$folderName</b>: Rip vollständig vorhanden, starte Kodierung."
            $doRip = $false
        } elseif (Test-Path $ripDir) {
            # RAW-Verzeichnis vorhanden, aber kein Marker → unvollständiger Rip, neu starten
            $partialMkvs = @(Get-ChildItem -Path $ripDir -Filter "*.mkv" -Recurse -ErrorAction SilentlyContinue)
            Write-Log "WARN" "Unvollständiger Rip gefunden ($($partialMkvs.Count) MKV(s), kein Abschluss-Marker) - lösche und starte neu."
            Send-TelegramMessage "DVD <b>$folderName</b>: Unvollständiger Rip – wird neu gerippt."
            Remove-Item -Path $ripDir -Recurse -Force -ErrorAction SilentlyContinue
        }

        if ($doRip) {
            if (-not (Test-Path $ripDir)) {
                New-Item -ItemType Directory -Path $ripDir -Force | Out-Null
            }
            Write-Log "INFO" "Starte DVD-Rip nach: $ripDir"

            # Bei SCSI-Fehlern im Vorscan: Laufwerk Zeit geben sich zu erholen
            if ($Script:DiscSCSIError) {
                Write-Log "INFO" "SCSI-Fehler beim Scan erkannt - warte 15s fuer Laufwerks-Erholung ..."
                Start-Sleep -Seconds 15
            }

            $mkvArgs = @(
                "--robot",
                "--noscan",
                "--directio=false",
                "mkv",
                "disc:$DriveIndex",
                "all",
                "`"$ripDir`"",
                "--minlength=$MinLength"
            )

            $mkvResult = Invoke-ExternalProcess `
                -Executable $MakeMKVPath `
                -Arguments $mkvArgs `
                -StepName "MakeMKV mkv"

            # MKV-Dateien prüfen - bei Lesefehlern auf beschädigten Discs kann MakeMKV
            # non-zero zurückgeben, aber trotzdem verwertbare Dateien erzeugt haben
            $mkvFiles = @(Get-ChildItem -Path $ripDir -Filter "*.mkv" -Recurse -ErrorAction SilentlyContinue)
            if ($mkvResult.ExitCode -ne 0) {
                if ($mkvFiles -and $mkvFiles.Count -gt 0) {
                    $largestMB = [math]::Round(($mkvFiles | Measure-Object -Property Length -Maximum).Maximum / 1MB, 0)
                    if ($largestMB -ge $MinMkvSizeMB) {
                        Write-Log "WARN" "MakeMKV mit Fehlern beendet (ExitCode: $($mkvResult.ExitCode)), größte MKV-Datei ist $largestMB MB (≥ Schwellwert $MinMkvSizeMB MB) - fahre fort."
                    } else {
                        Write-Log "ERROR" "MakeMKV fehlgeschlagen (ExitCode: $($mkvResult.ExitCode)), größte MKV-Datei ist nur $largestMB MB (< Schwellwert $MinMkvSizeMB MB) - abbruch."
                        $Script:ExitCode = 2
                        return
                    }
                } else {
                    Write-Log "ERROR" "MakeMKV fehlgeschlagen (ExitCode: $($mkvResult.ExitCode)), keine MKV-Dateien erstellt."
                    $Script:ExitCode = 2
                    return
                }
            }

            if (-not $mkvFiles -or $mkvFiles.Count -eq 0) {
                Write-Log "ERROR" "Keine MKV-Dateien nach dem Rippen gefunden in: $ripDir"
                $Script:ExitCode = 2
                return
            }

            Write-Log "INFO" "DVD-Rip abgeschlossen. $($mkvFiles.Count) MKV-Datei(en) erstellt."
            Set-Content -Path $ripMarker -Value (Get-Date -Format "yyyy-MM-dd HH:mm:ss") -Encoding UTF8
            Write-Log "INFO" "Rip-Marker gesetzt: $ripMarker"
            if (-not $SkipEncode) {
                Send-TelegramMessage "Rippen abgeschlossen: <b>$folderName</b> ($($mkvFiles.Count) Titel). Kodierung laeuft..."
            }
        }
    } else {
        Write-Log "INFO" "MakeMKV-Schritt übersprungen (SkipRip)."
    }

    # 6. Kodieren mit HandBrakeCLI
    if (-not $SkipEncode) {
        $mkvFiles = @(Get-ChildItem -Path $ripDir -Filter "*.mkv" -Recurse -ErrorAction SilentlyContinue)
        if (-not $mkvFiles -or $mkvFiles.Count -eq 0) {
            Write-Log "ERROR" "Keine MKV-Dateien zum Kodieren gefunden in: $ripDir"
            $Script:ExitCode = 3
            return
        }

        Write-Log "INFO" "Starte Kodierung von $($mkvFiles.Count) Datei(en) nach: $localDir"

        if (-not (Test-Path $localDir)) {
            New-Item -ItemType Directory -Path $localDir -Force | Out-Null
        }

        # Ausgabedateiendung aus Preset ableiten (MP4 wenn Preset "MP4" enthält, sonst MKV)
        $outExt = if ($HBPreset -match '\bMKV\b') { ".mkv" } else { ".mp4" }

        # Basis-Dateiname aus Disc-Titel ableiten (Leerzeichen → Unterstrich)
        $baseName = $folderName -replace '\s+', '_'

        $encodeErrors = 0
        $fileIndex = 1
        foreach ($mkv in $mkvFiles) {
            $suffix = if ($mkvFiles.Count -gt 1) { "_$($fileIndex.ToString('D2'))" } else { "" }
            $outFile = Join-Path $localDir ($baseName + $suffix + $outExt)
            $fileIndex++

            $hbArgs = [System.Collections.Generic.List[string]]::new()
            $hbArgs.Add("-i `"$($mkv.FullName)`"")
            $hbArgs.Add("-o `"$outFile`"")
            $hbArgs.Add("-q $HBQuality")

            if ($HBPreset) {
                $hbArgs.Add("--preset `"$HBPreset`"")
            }

            if ($HBExtraArgs) {
                $hbArgs.Add($HBExtraArgs)
            }

            $hbResult = Invoke-ExternalProcess `
                -Executable $HandBrakePath `
                -Arguments $hbArgs.ToArray() `
                -StepName "HandBrakeCLI [$($mkv.Name)]" `
                -StreamProgress:$DebugMode

            if ($hbResult.ExitCode -ne 0) {
                Write-Log "ERROR" "HandBrakeCLI fehlgeschlagen für '$($mkv.Name)' (ExitCode: $($hbResult.ExitCode))"
                $encodeErrors++
            } else {
                Write-Log "INFO" "Kodierung erfolgreich: '$($mkv.Name)' -> '$outFile'"
            }
        }

        if ($encodeErrors -gt 0) {
            Write-Log "ERROR" "$encodeErrors Kodierungsfehler aufgetreten."
            Send-TelegramMessage "Fehler bei Kodierung von <b>$folderName</b>: $encodeErrors Fehler aufgetreten."
            $Script:ExitCode = 3
        } else {
            Write-Log "INFO" "Alle Dateien erfolgreich kodiert."

            # RAW-Verzeichnis löschen
            Write-Log "INFO" "Lösche RAW-Verzeichnis: $ripDir"
            Remove-Item -Path $ripDir -Recurse -Force -ErrorAction SilentlyContinue
            Write-Log "INFO" "RAW-Verzeichnis gelöscht."

            if ($isNetworkPath) {
                # Kodierte Dateien vom TempDir auf OutputDir verschieben
                Write-Log "INFO" "Verschiebe Dateien von '$localDir' nach '$finalDir' ..."
                if (-not (Test-Path $finalDir)) {
                    New-Item -ItemType Directory -Path $finalDir -Force | Out-Null
                }
                Get-ChildItem -Path $localDir -File | ForEach-Object {
                    Move-Item -Path $_.FullName -Destination $finalDir -Force
                    Write-Log "INFO" "Verschoben: $($_.Name)"
                }
                $tempFolderPath = Join-Path $TempDir $folderName
                Remove-Item -Path $tempFolderPath -Recurse -Force -ErrorAction SilentlyContinue
                Write-Log "INFO" "TempDir bereinigt."
            }
        }
    } else {
        Write-Log "INFO" "HandBrakeCLI-Schritt übersprungen (SkipEncode)."
    }
}

# ---------------------------------------------------------------------------
# Service-Hilfsfunktionen
# ---------------------------------------------------------------------------
function Test-DiscAlreadyImported {
    param([string]$Title, [string]$TargetDir)
    $dir = Join-Path $TargetDir (Sanitize-FolderName $Title)
    if (-not (Test-Path $dir)) { return $false }
    $files = @(Get-ChildItem -Path $dir -File -ErrorAction SilentlyContinue |
               Where-Object { $_.Extension -in '.mp4', '.mkv' })
    return $files.Count -gt 0
}

function Wait-ForDisc {
    # Wartet auf eine eingelegte DVD. Gibt $true zurück sobald eine Disc erkannt wurde.
    # Nutzt WMI-Event für sofortige Erkennung; fällt auf Polling zurück.
    param([int]$PollMinutes)

    Write-Log "INFO" "[Service] Warte auf DVD (WMI-Event + Polling alle $PollMinutes Min) ..."
    Send-TelegramMessage "Service wartet auf naechste DVD."

    # WMI-Event: EventType=2 = Datenträger eingelegt
    $wmiQuery = "SELECT * FROM Win32_VolumeChangeEvent WHERE EventType = 2"
    try {
        Register-WmiEvent -Query $wmiQuery -SourceIdentifier "DVDRip_DiscInserted" -ErrorAction Stop | Out-Null
        Write-Log "INFO" "[Service] WMI-Event-Listener registriert."
    } catch {
        Write-Log "WARN" "[Service] WMI-Event-Registrierung fehlgeschlagen, nutze reines Polling: $($_.Exception.Message)"
    }

    try {
        while ($true) {
            $event = Wait-Event -SourceIdentifier "DVDRip_DiscInserted" -Timeout ($PollMinutes * 60) -ErrorAction SilentlyContinue
            if ($event) {
                Remove-Event -SourceIdentifier "DVDRip_DiscInserted" -ErrorAction SilentlyContinue
                Write-Log "INFO" "[Service] WMI-Event: Datenträger eingelegt."
            } else {
                Write-Log "INFO" "[Service] Polling-Timeout - prüfe Laufwerk ..."
            }

            # Disc-Status über MakeMKV prüfen
            $checkResult = Invoke-ExternalProcess `
                -Executable $MakeMKVPath `
                -Arguments @("--robot", "info", "disc:$DriveIndex") `
                -StepName "MakeMKV disc-check"

            $discReady = $checkResult.Output | Where-Object { $_ -match "^DRV:$DriveIndex,([2-9]|\d{2})" }
            if ($discReady) {
                Write-Log "INFO" "[Service] Disc erkannt."
                return $true
            }
            Write-Log "INFO" "[Service] Kein Disc im Laufwerk - weiter warten ..."
        }
    } finally {
        Unregister-Event -SourceIdentifier "DVDRip_DiscInserted" -ErrorAction SilentlyContinue
    }
}

function Invoke-SingleImport {
    # Führt einen kompletten Import-Durchlauf durch. Gibt $true bei Erfolg zurück.
    $Script:ExitCode  = 0
    $Script:StartTime = Get-Date
    $Script:DriveLetter = $null
    $Script:DiscSCSIError = $false

    try {
        Main
    } catch {
        $errMsg = $_.Exception.Message
        if ($Script:LogFile) {
            Write-Log "ERROR" "Unerwarteter Fehler: $errMsg"
            Write-Log "ERROR" "Stack Trace: $($_.ScriptStackTrace)"
        } else {
            Write-Host "[ERROR] Unerwarteter Fehler: $errMsg" -ForegroundColor Red
        }
        $Script:ExitCode = 99
    }

    $elapsed    = (Get-Date) - $Script:StartTime
    $elapsedStr = "{0:hh\:mm\:ss}" -f $elapsed

    if ($Script:ExitCode -eq 0) {
        $finalMsg = "ERFOLG - Import beendet nach $elapsedStr. Log: $Script:LogFile"
        Write-Log "INFO" $finalMsg
        Write-Host $finalMsg -ForegroundColor Green
        if (-not $NoEject) { Eject-Disc -DriveLetter $Script:DriveLetter }
        Send-TelegramMessage "Fertig: <b>$Script:LastDiscTitle</b> ($elapsedStr)&#10;Bitte naechste DVD einlegen."
    } else {
        $finalMsg = "FEHLER (Code $Script:ExitCode) - Import beendet nach $elapsedStr. Log: $Script:LogFile"
        Write-Log "ERROR" $finalMsg
        Write-Host $finalMsg -ForegroundColor Red
        if (-not $NoEject) { Eject-Disc -DriveLetter $Script:DriveLetter }
        Send-TelegramMessage "DVD-Import FEHLER (Code $Script:ExitCode) nach $elapsedStr."
    }

    # Log auf OutputDir kopieren
    if (-not $LogDir -and $Script:LogFile -and (Test-Path $Script:LogFile)) {
        try {
            if (-not (Test-Path $OutputDir)) { New-Item -ItemType Directory -Path $OutputDir -Force | Out-Null }
            Copy-Item -Path $Script:LogFile -Destination $OutputDir -Force
        } catch { }
    }

    return ($Script:ExitCode -eq 0)
}

function Start-ServiceLoop {
    Write-Host "=== DVDRip Service-Modus gestartet ===" -ForegroundColor Cyan
    Write-Log "INFO" "[Service] Service-Modus aktiv. PollInterval: $PollInterval Min."
    Send-TelegramMessage "DVDRip Service gestartet."

    while ($true) {
        # Auf Disc warten
        Wait-ForDisc -PollMinutes $PollInterval | Out-Null

        # Disc-Titel vorab lesen um zu prüfen ob bereits importiert
        Write-Log "INFO" "[Service] Lese Disc-Titel ..."
        $checkTitle = $null
        try {
            $checkTitle = Get-DiscTitle -MkvCon $MakeMKVPath -Index $DriveIndex
        } catch { }

        if ($checkTitle -and (Test-DiscAlreadyImported -Title $checkTitle -TargetDir $OutputDir)) {
            Write-Log "INFO" "[Service] '$checkTitle' bereits importiert - werfe DVD aus."
            Send-TelegramMessage "DVD <b>$checkTitle</b> bereits importiert - ausgeworfen."
            Eject-Disc -DriveLetter $Script:DriveLetter
        } else {
            Write-Log "INFO" "[Service] Starte Import ..."
            $startTitle = if ($checkTitle) { $checkTitle } else { "Unbekannt" }
            Send-TelegramMessage "DVD <b>$startTitle</b> eingelegt – Import gestartet."
            Invoke-SingleImport | Out-Null
        }
    }
}

# ---------------------------------------------------------------------------
# Skript-Einstiegspunkt
# ---------------------------------------------------------------------------
# MakeMKV-Pfad einmalig auflösen (wird auch im Service-Loop benötigt)
if ($MakeMKVPath -eq "") { $MakeMKVPath = Find-MakeMKV }

if ($ServiceMode) {
    if (-not $MakeMKVPath -or -not (Test-Path $MakeMKVPath)) {
        Write-Host "[ERROR] makemkvcon.exe nicht gefunden. Service-Modus nicht möglich." -ForegroundColor Red
        exit 1
    }
    # Log-Basisverzeichnis für Service-Modus initialisieren
    $logBase = if ($LogDir) { $LogDir } else { $PSScriptRoot }
    Initialize-Log -Dir $logBase
    Start-ServiceLoop
    exit 0
}

try {
    Main
} catch {
    $errMsg = $_.Exception.Message
    if ($Script:LogFile) {
        Write-Log "ERROR" "Unerwarteter Fehler: $errMsg"
        Write-Log "ERROR" "Stack Trace: $($_.ScriptStackTrace)"
    } else {
        Write-Host "[ERROR] Unerwarteter Fehler: $errMsg" -ForegroundColor Red
    }
    $Script:ExitCode = 99
}

# ---------------------------------------------------------------------------
# Abschlussmeldung (Einzel-Modus)
# ---------------------------------------------------------------------------
$elapsed = (Get-Date) - $Script:StartTime
$elapsedStr = "{0:hh\:mm\:ss}" -f $elapsed

if ($Script:ExitCode -eq 0) {
    $finalMsg = "ERFOLG - Skript beendet nach $elapsedStr. Log: $Script:LogFile"
    Write-Log "INFO" $finalMsg
    Write-Host $finalMsg -ForegroundColor Green
    if (-not $NoEject) { Eject-Disc -DriveLetter $Script:DriveLetter }
    Send-TelegramMessage "Fertig: <b>$Script:LastDiscTitle</b> ($elapsedStr)&#10;Bitte naechste DVD einlegen."
} else {
    $finalMsg = "FEHLER (Code $Script:ExitCode) - Skript beendet nach $elapsedStr. Log: $Script:LogFile"
    Write-Log "ERROR" $finalMsg
    Write-Host $finalMsg -ForegroundColor Red
    Send-TelegramMessage "DVD-Import FEHLER (Code $Script:ExitCode) nach $elapsedStr."
}

# Log auf OutputDir kopieren (nur wenn kein explizites -LogDir angegeben)
if (-not $LogDir -and $Script:LogFile -and (Test-Path $Script:LogFile)) {
    try {
        if (-not (Test-Path $OutputDir)) {
            New-Item -ItemType Directory -Path $OutputDir -Force | Out-Null
        }
        Copy-Item -Path $Script:LogFile -Destination $OutputDir -Force
        Write-Host "Log kopiert nach: $OutputDir" -ForegroundColor Cyan
    } catch {
        Write-Host "Log konnte nicht auf OutputDir kopiert werden: $($_.Exception.Message)" -ForegroundColor Yellow
    }
}

exit $Script:ExitCode
