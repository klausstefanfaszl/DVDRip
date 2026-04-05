<#
.SYNOPSIS
    Installiert DVDRip als Windows-Dienst via NSSM.

.DESCRIPTION
    Lädt NSSM (Non-Sucking Service Manager) herunter falls nicht vorhanden,
    und registriert DVDRip.ps1 im Service-Modus als automatisch startenden
    Windows-Dienst.

.PARAMETER OutputDir
    Zielverzeichnis für importierte DVDs (Pflicht).

.PARAMETER TelegramToken
    Telegram Bot-Token für Benachrichtigungen.

.PARAMETER TelegramChatId
    Telegram Chat-ID des Empfängers.

.PARAMETER PollInterval
    Polling-Intervall in Minuten (Standard: 5).

.PARAMETER ServiceName
    Name des Windows-Dienstes (Standard: DVDRip).

.PARAMETER NssmPath
    Pfad zu nssm.exe. Wird automatisch heruntergeladen falls nicht angegeben.

.PARAMETER HBPreset
    HandBrake-Preset (Standard: "H.265 MP4 576p25").

.PARAMETER HBQuality
    HandBrake RF-Qualität (Standard: 22).

.PARAMETER Uninstall
    Deinstalliert den Dienst anstatt ihn zu installieren.

.EXAMPLE
    .\Install-DVDRipService.ps1 -OutputDir "D:\Videos" -TelegramToken "..." -TelegramChatId "..."
    .\Install-DVDRipService.ps1 -Uninstall
#>

#Requires -RunAsAdministrator

[CmdletBinding()]
param(
    [Parameter(Mandatory = $false)]
    [string]$OutputDir = "",

    [Parameter(Mandatory = $false)]
    [string]$TelegramToken = "",

    [Parameter(Mandatory = $false)]
    [string]$TelegramChatId = "",

    [Parameter(Mandatory = $false)]
    [int]$PollInterval = 5,

    [Parameter(Mandatory = $false)]
    [string]$ServiceName = "DVDRip",

    [Parameter(Mandatory = $false)]
    [string]$NssmPath = "",

    [Parameter(Mandatory = $false)]
    [string]$HBPreset = "H.265 MP4 576p25",

    [Parameter(Mandatory = $false)]
    [int]$HBQuality = 22,

    [Parameter(Mandatory = $false)]
    [switch]$Uninstall
)

$ErrorActionPreference = "Stop"
$ScriptDir = $PSScriptRoot
$DVDRipScript = Join-Path $ScriptDir "DVDRip.ps1"

function Write-Step {
    param([string]$Msg)
    Write-Host "  $Msg" -ForegroundColor Cyan
}
function Write-Ok   { param([string]$Msg); Write-Host "  [OK] $Msg" -ForegroundColor Green }
function Write-Fail { param([string]$Msg); Write-Host "  [FEHLER] $Msg" -ForegroundColor Red }

# ---------------------------------------------------------------------------
# NSSM finden oder herunterladen
# ---------------------------------------------------------------------------
function Get-Nssm {
    param([string]$Hint)

    if ($Hint -and (Test-Path $Hint)) { return $Hint }

    # In bekannten Pfaden suchen
    $candidates = @(
        (Join-Path $ScriptDir "nssm.exe"),
        "C:\nssm\nssm.exe",
        "C:\tools\nssm\nssm.exe"
    )
    $found = Get-Command "nssm.exe" -ErrorAction SilentlyContinue
    if ($found) { $candidates = @($found.Source) + $candidates }

    foreach ($c in $candidates) {
        if (Test-Path $c) { return $c }
    }

    # Herunterladen
    Write-Step "nssm.exe nicht gefunden - lade herunter ..."
    $nssmDir  = Join-Path $ScriptDir "nssm"
    $nssmZip  = Join-Path $env:TEMP "nssm.zip"
    $nssmExe  = Join-Path $nssmDir "nssm.exe"

    try {
        $nssmUrl = "https://nssm.cc/release/nssm-2.24.zip"
        Invoke-WebRequest -Uri $nssmUrl -OutFile $nssmZip -UseBasicParsing
        Expand-Archive -Path $nssmZip -DestinationPath (Join-Path $env:TEMP "nssm_extract") -Force
        Remove-Item $nssmZip -ErrorAction SilentlyContinue

        # Passende nssm.exe (64-bit) heraussuchen
        $extracted = Get-ChildItem (Join-Path $env:TEMP "nssm_extract") -Recurse -Filter "nssm.exe" |
                     Where-Object { $_.FullName -match "win64" } |
                     Select-Object -First 1
        if (-not $extracted) {
            $extracted = Get-ChildItem (Join-Path $env:TEMP "nssm_extract") -Recurse -Filter "nssm.exe" |
                         Select-Object -First 1
        }
        if (-not $extracted) { throw "nssm.exe nicht in ZIP gefunden." }

        if (-not (Test-Path $nssmDir)) { New-Item -ItemType Directory -Path $nssmDir -Force | Out-Null }
        Copy-Item $extracted.FullName -Destination $nssmExe -Force
        Remove-Item (Join-Path $env:TEMP "nssm_extract") -Recurse -Force -ErrorAction SilentlyContinue
        Write-Ok "nssm.exe heruntergeladen: $nssmExe"
        return $nssmExe
    } catch {
        throw "NSSM konnte nicht heruntergeladen werden: $($_.Exception.Message)"
    }
}

# ---------------------------------------------------------------------------
# Deinstallation
# ---------------------------------------------------------------------------
if ($Uninstall) {
    Write-Host "`nDeinstalliere Dienst '$ServiceName' ..." -ForegroundColor Yellow

    $nssm = Get-Nssm -Hint $NssmPath

    $svc = Get-Service -Name $ServiceName -ErrorAction SilentlyContinue
    if (-not $svc) {
        Write-Host "  Dienst '$ServiceName' nicht gefunden." -ForegroundColor Yellow
        exit 0
    }

    Write-Step "Stoppe Dienst ..."
    & $nssm stop $ServiceName confirm 2>&1 | Out-Null

    Write-Step "Entferne Dienst ..."
    & $nssm remove $ServiceName confirm 2>&1 | Out-Null

    Write-Ok "Dienst '$ServiceName' deinstalliert."
    exit 0
}

# ---------------------------------------------------------------------------
# Installation
# ---------------------------------------------------------------------------
Write-Host "`n=== DVDRip Service-Installation ===" -ForegroundColor Cyan

if (-not $OutputDir) {
    Write-Fail "-OutputDir ist erforderlich."
    exit 1
}
if (-not (Test-Path $DVDRipScript)) {
    Write-Fail "DVDRip.ps1 nicht gefunden: $DVDRipScript"
    exit 1
}

$nssm = Get-Nssm -Hint $NssmPath
Write-Ok "NSSM: $nssm"

# Prüfen ob Dienst bereits existiert
$existingSvc = Get-Service -Name $ServiceName -ErrorAction SilentlyContinue
if ($existingSvc) {
    Write-Step "Dienst '$ServiceName' existiert bereits - stoppe und entferne ihn zuerst ..."
    & $nssm stop   $ServiceName confirm 2>&1 | Out-Null
    & $nssm remove $ServiceName confirm 2>&1 | Out-Null
    Start-Sleep -Seconds 2
}

# PowerShell-Argumente für DVDRip zusammenstellen
$psArgs  = "-NoProfile -ExecutionPolicy Bypass -File `"$DVDRipScript`""
$psArgs += " -OutputDir `"$OutputDir`""
$psArgs += " -ServiceMode"
$psArgs += " -PollInterval $PollInterval"
$psArgs += " -HBPreset `"$HBPreset`""
$psArgs += " -HBQuality $HBQuality"

if ($TelegramToken)  { $psArgs += " -TelegramToken `"$TelegramToken`"" }
if ($TelegramChatId) { $psArgs += " -TelegramChatId `"$TelegramChatId`"" }

$powershellExe = (Get-Command powershell.exe).Source

Write-Step "Registriere Dienst '$ServiceName' ..."
& $nssm install $ServiceName $powershellExe $psArgs 2>&1 | Out-Null

# Dienst konfigurieren
Write-Step "Konfiguriere Dienst ..."
& $nssm set $ServiceName DisplayName  "DVDRip Service"                         2>&1 | Out-Null
& $nssm set $ServiceName Description  "Automatisches DVD-Rippen und Komprimieren" 2>&1 | Out-Null
& $nssm set $ServiceName Start        SERVICE_AUTO_START                           2>&1 | Out-Null
& $nssm set $ServiceName AppStdout    (Join-Path $OutputDir "DVDRip_service_stdout.log") 2>&1 | Out-Null
& $nssm set $ServiceName AppStderr    (Join-Path $OutputDir "DVDRip_service_stderr.log") 2>&1 | Out-Null
& $nssm set $ServiceName AppRotateFiles    1   2>&1 | Out-Null
& $nssm set $ServiceName AppRotateBytes    10485760 2>&1 | Out-Null  # 10 MB
& $nssm set $ServiceName AppRestartDelay   5000 2>&1 | Out-Null      # 5s Pause vor Neustart

Write-Step "Starte Dienst ..."
& $nssm start $ServiceName 2>&1 | Out-Null
Start-Sleep -Seconds 2

$svc = Get-Service -Name $ServiceName -ErrorAction SilentlyContinue
if ($svc -and $svc.Status -eq "Running") {
    Write-Ok "Dienst '$ServiceName' läuft."
} else {
    Write-Host "  [WARN] Dienst gestartet, Status: $($svc.Status)" -ForegroundColor Yellow
    Write-Host "         Prüfe: Get-Service $ServiceName" -ForegroundColor Yellow
}

Write-Host "`nInstallation abgeschlossen." -ForegroundColor Green
Write-Host "  Dienst verwalten:" -ForegroundColor Gray
Write-Host "    Start:        Start-Service $ServiceName" -ForegroundColor Gray
Write-Host "    Stop:         Stop-Service  $ServiceName" -ForegroundColor Gray
Write-Host "    Deinstall:    .\Install-DVDRipService.ps1 -Uninstall" -ForegroundColor Gray
Write-Host "    NSSM-Editor:  $nssm edit $ServiceName" -ForegroundColor Gray
