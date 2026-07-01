# ============================================================
# run_all.ps1  --  Master-Skript Dashboard-Update
# Ruft alle Teilskripte auf, loggt Ergebnisse, pusht zu GitHub
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

$failed = @()

# -- Schritt 1: Netzwerkdateien kopieren --------------------
Log "--- Netzwerkdateien kopieren ---" "Yellow"
$s1 = Join-Path $SCRIPTS "01_copy_network.ps1"
if (Test-Path $s1) {
    try {
        & $s1 2>&1 | ForEach-Object { Log "  $_" }
        if ($LASTEXITCODE -ne 0) { throw "Exit-Code $LASTEXITCODE" }
        Log "[OK] Netzwerkdateien kopieren" "Green"
    } catch {
        Log "[ERR] Netzwerkdateien kopieren: $_" "Red"
        $failed += "Netzwerkdateien kopieren"
    }
} else { Log "[SKIP] 01_copy_network.ps1 nicht gefunden" "DarkGray" }

# -- Schritt 2: Draht-Uebersicht Query-Refresh --------------
Log "--- Draht-Uebersicht Query-Refresh ---" "Yellow"
$s2 = Join-Path $SCRIPTS "02_refresh_ubersicht.ps1"
if (Test-Path $s2) {
    try {
        & $s2 2>&1 | ForEach-Object { Log "  $_" }
        if ($LASTEXITCODE -ne 0) { throw "Exit-Code $LASTEXITCODE" }
        Log "[OK] Draht-Uebersicht Query-Refresh" "Green"
    } catch {
        Log "[ERR] Draht-Uebersicht Query-Refresh: $_" "Red"
        $failed += "Draht-Uebersicht Query-Refresh"
    }
} else { Log "[SKIP] 02_refresh_ubersicht.ps1 nicht gefunden" "DarkGray" }

# -- Schritt 3: SharePoint-Dateien laden --------------------
# KEIN Pipe-Redirect: Device-Code-Prompt muss im Terminal sichtbar sein
Log "--- SharePoint-Dateien laden ---" "Yellow"
$s3 = Join-Path $SCRIPTS "03_sharepoint_download.ps1"
if (Test-Path $s3) {
    try {
        & $s3
        if ($LASTEXITCODE -ne 0) { throw "Exit-Code $LASTEXITCODE" }
        Log "[OK] SharePoint-Dateien laden" "Green"
    } catch {
        Log "[ERR] SharePoint-Dateien laden: $_" "Red"
        $failed += "SharePoint-Dateien laden"
    }
} else { Log "[SKIP] 03_sharepoint_download.ps1 nicht gefunden" "DarkGray" }

# -- Schritt 4: Excel -> JSON (deaktiviert bis Datenstruktur definiert) --
Log "--- Excel -> JSON konvertieren: UEBERSPRUNGEN (in Arbeit) ---" "DarkGray"

# -- GitHub Push --------------------------------------------
if ($true) {
    Log "--- GitHub Push ---" "Yellow"
    $dashDir = Join-Path $BASE "docs"

    try {
        $dataSrc = Join-Path $BASE "_data"
        $dataDst = Join-Path $dashDir "_data"
        if (-not (Test-Path $dataDst)) { New-Item -ItemType Directory -Path $dataDst | Out-Null }
        Copy-Item -Path "$dataSrc\*" -Destination $dataDst -Recurse -Force

        Push-Location $BASE
        git add -A                                  2>&1 | ForEach-Object { Log "  git: $_" }
        git commit -m "Auto-Update $DATE $TIME"     2>&1 | ForEach-Object { Log "  git: $_" }
        git push                                    2>&1 | ForEach-Object { Log "  git: $_" }
        Pop-Location

        Log "[OK] GitHub Push erfolgreich" "Green"
    } catch {
        Log "[ERR] GitHub Push fehlgeschlagen: $_" "Red"
        $failed += "GitHub Push"
        Pop-Location -ErrorAction SilentlyContinue
    }
} else {
    Log "[SKIP] GitHub Push uebersprungen (JSON-Export fehlerhaft)" "DarkGray"
}

# -- Abschluss ----------------------------------------------
Log "====================================================" "Cyan"
if ($failed.Count -eq 0) {
    Log "FERTIG -- Alle Schritte erfolgreich." "Cyan"
    exit 0
} else {
    Log "FERTIG MIT FEHLERN -- Fehlgeschlagen: $($failed -join ', ')" "Red"
    exit 1
}
