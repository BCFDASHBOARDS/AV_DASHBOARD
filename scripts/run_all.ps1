# ============================================================
# run_all.ps1  —  Master-Skript Dashboard-Update
# Ruft alle Teilskripte auf, loggt Ergebnisse, pusht zu GitHub
# ============================================================
# Portabel: läuft für jeden Benutzer via %USERNAME% / $env:USERNAME
# Arbeitsordner: ...\050 Claude planned Tasks\051 Dashboard\
# ============================================================

$SCRIPTS = Split-Path -Parent $MyInvocation.MyCommand.Path   # .../051 Dashboard/scripts/
$BASE    = Split-Path -Parent $SCRIPTS                        # .../051 Dashboard/
$LOG_DIR = Join-Path $BASE "logs"
$DATE    = Get-Date -Format "yyyy-MM-dd"
$TIME    = Get-Date -Format "HH:mm:ss"

if (-not (Test-Path $LOG_DIR)) { New-Item -ItemType Directory -Path $LOG_DIR | Out-Null }
$LOG = Join-Path $LOG_DIR "run_$DATE.log"

function Log($msg, $color = "White") {
    $entry = "$(Get-Date -Format 'HH:mm:ss') | $msg"
    Write-Host $entry -ForegroundColor $color
    Add-Content -Path $LOG -Value $entry -Encoding UTF8
}

Log "===== Dashboard-Update gestartet ($DATE $TIME) =====" "Cyan"

$steps = @(
    @{ Name = "Netzwerkdateien kopieren";      Script = "01_copy_network.ps1"       }
    @{ Name = "Draht-Übersicht Query-Refresh"; Script = "02_refresh_ubersicht.ps1"  }
    @{ Name = "SharePoint-Dateien laden";      Script = "03_sharepoint_download.ps1"}
    @{ Name = "Excel → JSON konvertieren";     Script = "04_excel_to_json.ps1"      }
)

$failed = @()

foreach ($step in $steps) {
    Log "--- $($step.Name) ---" "Yellow"
    $scriptPath = Join-Path $SCRIPTS $step.Script

    if (-not (Test-Path $scriptPath)) {
        Log "[SKIP] Skript nicht gefunden: $($step.Script)" "DarkGray"
        continue
    }

    try {
        & $scriptPath 2>&1 | ForEach-Object { Log "  $_" }
        if ($LASTEXITCODE -ne 0) { throw "Exit-Code $LASTEXITCODE" }
        Log "[OK] $($step.Name)" "Green"
    } catch {
        Log "[ERR] $($step.Name): $_" "Red"
        $failed += $step.Name
    }
}

# GitHub Push (nur wenn kein kritischer Fehler in JSON-Export)
if ($failed -notcontains "Excel → JSON konvertieren") {
    Log "--- GitHub Push ---" "Yellow"
    $dashDir = Join-Path $BASE "dashboard"

    try {
        # _data nach dashboard/_data kopieren damit GitHub Pages die JSON-Dateien hostet
        $dataSrc = Join-Path $BASE "_data"
        $dataDst = Join-Path $dashDir "_data"
        if (-not (Test-Path $dataDst)) { New-Item -ItemType Directory -Path $dataDst | Out-Null }
        Copy-Item -Path "$dataSrc\*" -Destination $dataDst -Recurse -Force

        Push-Location $BASE
        git add -A                                   2>&1 | ForEach-Object { Log "  git: $_" }
        git commit -m "Auto-Update $DATE $TIME"      2>&1 | ForEach-Object { Log "  git: $_" }
        git push                                     2>&1 | ForEach-Object { Log "  git: $_" }
        Pop-Location

        Log "[OK] GitHub Push erfolgreich" "Green"
    } catch {
        Log "[ERR] GitHub Push fehlgeschlagen: $_" "Red"
        $failed += "GitHub Push"
        Pop-Location -ErrorAction SilentlyContinue
    }
} else {
    Log "[SKIP] GitHub Push übersprungen (JSON-Export fehlerhaft)" "DarkGray"
}

# Abschlussstatus
Log "====================================================" "Cyan"
if ($failed.Count -eq 0) {
    Log "FERTIG — Alle Schritte erfolgreich." "Cyan"
    exit 0
} else {
    Log "FERTIG MIT FEHLERN — Fehlgeschlagen: $($failed -join ', ')" "Red"
    exit 1
}
