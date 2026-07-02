# ============================================================
# run_all.ps1  --  Master-Skript Dashboard-Update
# Ruft alle Teilskripte auf, loggt Ergebnisse, pusht zu GitHub
#
# OneDrive wird zu Beginn pausiert (Stop-Process) und am Ende
# in einem finally-Block garantiert neu gestartet -- auch bei
# Fehler oder vorzeitigem Abbruch.
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

# ============================================================
# ONEDRIVE PAUSIEREN
# Stoppt den OneDrive-Prozess damit der Sync keine .git-Lock-
# Dateien blockiert. Im finally-Block wird OneDrive garantiert
# neu gestartet -- auch wenn das Script mit Fehler abbricht.
# ============================================================
$odWasRunning = $false
$odExe = "C:\Program Files\Microsoft OneDrive\OneDrive.exe"

$odProc = Get-Process -Name OneDrive -ErrorAction SilentlyContinue
if ($odProc) {
    $odWasRunning = $true
    Log "OneDrive: Sync pausieren ..." "Yellow"
    Stop-Process -Name OneDrive -Force -ErrorAction SilentlyContinue
    # Warten bis Prozess wirklich weg (max. 15 Sek)
    $waited = 0
    while ((Get-Process -Name OneDrive -ErrorAction SilentlyContinue) -and ($waited -lt 15)) {
        Start-Sleep -Seconds 1
        $waited++
    }
    Log "[OK] OneDrive gestoppt (nach ${waited}s)" "Green"
} else {
    Log "[INFO] OneDrive lief nicht -- kein Stop noetig" "DarkGray"
}

# ============================================================
# HAUPTPIPELINE  (try/finally garantiert OneDrive-Neustart)
# ============================================================
$failed = @()

try {

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

    # -- Schritt 5: Auftragsbestand-Extraktion --------------------
    Log "--- Auftragsbestand Extraktion ---" "Yellow"
    $s5 = Join-Path $SCRIPTS "05_auftragsbestand.ps1"

    if (Test-Path $s5) {
        try {
            & $s5 2>&1 | Out-String -Stream | ForEach-Object { Log "  $_" }
            if ($LASTEXITCODE -ne 0) { throw "Exit-Code $LASTEXITCODE" }
            Log "[OK] Auftragsbestand Extraktion" "Green"
        } catch {
            Log "[ERR] Auftragsbestand Extraktion: $_" "Red"
            $failed += "Auftragsbestand Extraktion"
        }
    } else { Log "[SKIP] 05_auftragsbestand.ps1 nicht gefunden" "DarkGray" }

    # -- Schritt 6: Material-Extraktion ---------------------------
    Log "--- Material Extraktion ---" "Yellow"
    $s6 = Join-Path $SCRIPTS "06_material.ps1"
    if (Test-Path $s6) {
        try {
            & $s6 2>&1 | Out-String -Stream | ForEach-Object { Log "  $_" }
            if ($LASTEXITCODE -ne 0) { throw "Exit-Code $LASTEXITCODE" }
            Log "[OK] Material Extraktion" "Green"
        } catch {
            Log "[ERR] Material Extraktion: $_" "Red"
            $failed += "Material Extraktion"
        }
    } else { Log "[SKIP] 06_material.ps1 nicht gefunden" "DarkGray" }

    # -- GitHub Push via Plumbing ----------------------------------
    # Plumbing-Methode: umgeht lokale Branch-Divergenz, kein git add/commit
    # Pushed immer auf Basis des aktuellen Remote-HEAD
    Log "--- GitHub Push ---" "Yellow"
    $dashDir = Join-Path $BASE "docs"
    $tmpIdx  = [System.IO.Path]::GetTempFileName()
    try {
        # JSON-Daten nach docs/_data kopieren
        $dataSrc = Join-Path $BASE "_data"
        $dataDst = Join-Path $dashDir "_data"
        if (-not (Test-Path $dataDst)) { New-Item -ItemType Directory -Path $dataDst | Out-Null }
        Copy-Item -Path "$dataSrc\*" -Destination $dataDst -Recurse -Force
        Log "  _data nach docs/_data kopiert" "DarkGray"

        $env:GIT_INDEX_FILE = $tmpIdx
        Push-Location $BASE

        # Remote HEAD ermitteln
        $lsOut     = (git ls-remote origin HEAD 2>&1)
        $remoteHead = ($lsOut | ForEach-Object {
            if ($_ -match '^([0-9a-f]{40})\s+HEAD') { $Matches[1] }
        } | Select-Object -First 1)
        if (-not $remoteHead) { throw "Remote HEAD nicht ermittelbar: $lsOut" }
        Log "  Remote HEAD: $remoteHead" "DarkGray"

        # Remote-Tree in Temp-Index laden
        git read-tree $remoteHead 2>&1 | Out-Null

        # Dateien hashen und zum Index hinzufuegen
        # Nur JSON-Daten -- HTML-Seiten (exec/, team/) werden separat via
        # git-Plumbing gepusht wenn sich das Dashboard aendert.
        # Root-HTML-Dateien (auftragsbestand.html etc.) sind veraltet;
        # kanonische Seiten liegen unter docs/exec/ und docs/team/.
        $pushFiles = @(
            "docs/_data/auftragsbestand.json",
            "docs/_data/material.json"
        )
        foreach ($rel in $pushFiles) {
            $abs = Join-Path $BASE ($rel -replace "/", "\")
            if (-not (Test-Path $abs)) {
                Log "  [SKIP] nicht gefunden: $rel" "DarkGray"
                continue
            }
            $blob = (git hash-object -w $abs 2>$null)
            if ($blob -notmatch '^[0-9a-f]{40}$') { throw "hash-object fehlge