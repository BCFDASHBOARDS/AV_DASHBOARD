# ============================================================
# 00_setup.ps1  --  Einmaliges Setup fuer das BCF AV Dashboard
# Ausfuehren als Administrator:
# powershell -ExecutionPolicy Bypass -File "...\00_setup.ps1"
# ============================================================

$ErrorActionPreference = "Continue"
$BASE = Split-Path -Parent $MyInvocation.MyCommand.Path

function Banner($text) {
    Write-Host ""
    Write-Host ("-" * 55) -ForegroundColor DarkGray
    Write-Host "  $text" -ForegroundColor Cyan
    Write-Host ("-" * 55) -ForegroundColor DarkGray
}
function OK($msg)   { Write-Host "  [OK]  $msg" -ForegroundColor Green  }
function WARN($msg) { Write-Host "  [!]   $msg" -ForegroundColor Yellow }
function ERR($msg)  { Write-Host "  [ERR] $msg" -ForegroundColor Red    }
function INF($msg)  { Write-Host "  >>    $msg" -ForegroundColor Gray   }

Clear-Host
Write-Host ""
Write-Host "  BCF AV Dashboard - Einrichtung" -ForegroundColor White
Write-Host "  $(Get-Date -Format 'dd.MM.yyyy HH:mm')" -ForegroundColor DarkGray

# -- 1. PowerShell-ExecutionPolicy --------------------------
Banner "1/6  ExecutionPolicy"
$pol = Get-ExecutionPolicy -Scope CurrentUser
if ($pol -eq "Restricted" -or $pol -eq "Undefined") {
    Set-ExecutionPolicy -Scope CurrentUser -ExecutionPolicy RemoteSigned -Force
    OK "ExecutionPolicy auf RemoteSigned gesetzt"
} else {
    OK "ExecutionPolicy bereits OK ($pol)"
}

# -- 2. PS-Module -------------------------------------------
Banner "2/6  PowerShell-Module"

$modules = @(
    @{ Name = "ImportExcel";                      Min = "7.0" }
    @{ Name = "Microsoft.Graph.Authentication";   Min = "2.0" }
    @{ Name = "Microsoft.Graph.Files";            Min = "2.0" }
)

foreach ($m in $modules) {
    $installed = Get-Module -ListAvailable -Name $m.Name | Sort-Object Version -Descending | Select-Object -First 1
    if ($installed) {
        OK "$($m.Name) v$($installed.Version) bereits installiert"
    } else {
        INF "Installiere $($m.Name) ..."
        try {
            Install-Module $m.Name -Scope CurrentUser -Force -AllowClobber -ErrorAction Stop
            OK "$($m.Name) installiert"
        } catch {
            ERR "$($m.Name) konnte nicht installiert werden: $_"
        }
    }
}

# -- 3. Git -------------------------------------------------
Banner "3/6  Git"
$gitPath = Get-Command git -ErrorAction SilentlyContinue
if ($gitPath) {
    $gitVer = (git --version 2>&1)
    OK "Git gefunden: $gitVer"
} else {
    WARN "Git ist nicht installiert."
    INF "Download-Seite wird geoeffnet: https://git-scm.com/download/win"
    Start-Process "https://git-scm.com/download/win"
    Write-Host ""
    Write-Host "  Bitte Git installieren und danach dieses Skript" -ForegroundColor Yellow
    Write-Host "  erneut ausfuehren." -ForegroundColor Yellow
    Read-Host "  [Enter] zum Beenden"
    exit 1
}

# Git-Konfiguration pruefen
$gitUser  = git config --global user.name  2>&1
$gitEmail = git config --global user.email 2>&1
if (-not $gitUser -or $gitUser -like "*exit*") {
    Write-Host ""
    $name = Read-Host "  Git-Name (z.B. WiBa Baussmann)"
    git config --global user.name  $name
}
if (-not $gitEmail -or $gitEmail -like "*exit*") {
    $email = Read-Host "  Git-E-Mail (deine GitHub-E-Mail)"
    git config --global user.email $email
}
OK "Git-Benutzer: $(git config --global user.name) <$(git config --global user.email)>"

# -- 4. Ordnerstruktur --------------------------------------
Banner "4/6  Ordnerstruktur"
$dirs = @("_source","_data","scripts","docs","docs\_data","logs")
foreach ($d in $dirs) {
    $p = Join-Path $BASE $d
    if (-not (Test-Path $p)) {
        New-Item -ItemType Directory -Path $p | Out-Null
        OK "Erstellt: $d"
    } else {
        OK "Vorhanden: $d"
    }
}

# .gitignore pruefen
$gitignore = Join-Path $BASE ".gitignore"
if (Test-Path $gitignore) {
    OK ".gitignore vorhanden"
} else {
    WARN ".gitignore fehlt -- bitte manuell anlegen"
}

# -- 5. Git-Repository --------------------------------------
Banner "5/6  Git-Repository"
$gitDir = Join-Path $BASE ".git"
if (-not (Test-Path $gitDir)) {
    Push-Location $BASE
    git init --initial-branch=main 2>&1 | Out-Null
    OK "Git-Repository initialisiert (Branch: main)"
    Pop-Location
} else {
    OK "Git-Repository bereits vorhanden"
}

# GitHub-Remote einrichten (Token fuer passwortlosen Push)
Write-Host ""

# Token aus Token.txt lesen
$tokenFile = Join-Path $BASE "Token.txt"
$ghToken = ""
if (Test-Path $tokenFile) {
    $ghToken = (Get-Content $tokenFile -Raw).Trim()
    OK "GitHub-Token aus Token.txt geladen"
} else {
    WARN "Token.txt nicht gefunden -- Push wird ohne Token konfiguriert"
}

$repoSlug = Read-Host "  GitHub-Repo (FORMAT: USERNAME/REPONAME, leer = spaeter)"

if ($repoSlug.Trim() -ne "") {
    if ($ghToken -ne "") {
        $repoUrl        = "https://$ghToken@github.com/$($repoSlug.Trim()).git"
        $repoUrlDisplay = "https://***TOKEN***@github.com/$($repoSlug.Trim()).git"
    } else {
        $repoUrl        = "https://github.com/$($repoSlug.Trim()).git"
        $repoUrlDisplay = $repoUrl
    }

    Push-Location $BASE
    $remoteExists = git remote get-url origin 2>&1
    if ($LASTEXITCODE -eq 0) {
        git remote set-url origin $repoUrl
        OK "Remote 'origin' aktualisiert: $repoUrlDisplay"
    } else {
        git remote add origin $repoUrl
        OK "Remote 'origin' gesetzt: $repoUrlDisplay"
    }

    # Erster Commit + Push
    git add -A 2>&1 | Out-Null
    $commitOut = git commit -m "Initial setup -- BCF AV Dashboard" 2>&1
    if ($LASTEXITCODE -eq 0) {
        OK "Erster Commit erstellt"
        INF "Push zu GitHub ..."
        git push -u origin main 2>&1
        if ($LASTEXITCODE -eq 0) {
            $webUrl = "https://github.com/$($repoSlug.Trim())"
            OK "Push erfolgreich!"
            Write-Host ""
            Write-Host "  Naechster Schritt: GitHub Pages aktivieren" -ForegroundColor Cyan
            Write-Host "  >> $webUrl/settings/pages" -ForegroundColor Gray
            Write-Host "     Source: Deploy from branch" -ForegroundColor Gray
            Write-Host "     Branch: main  |  Folder: /dashboard" -ForegroundColor Gray
            Start-Process "$webUrl/settings/pages"
        } else {
            WARN "Push fehlgeschlagen -- bitte manuell: cd '$BASE' && git push -u origin main"
        }
    } else {
        INF "Nichts zu committen (bereits aktuell)"
    }
    Pop-Location
} else {
    WARN "Kein Remote eingerichtet. Spaeter nachholen mit:"
    INF "  git remote add origin https://TOKEN@github.com/USERNAME/REPO.git"
    INF "  git push -u origin main"
}

# -- 6. Task Scheduler --------------------------------------
Banner "6/6  Windows Task Scheduler"
Write-Host ""
$zeitStr = Read-Host "  Uhrzeit fuer taeglichen Update-Lauf (HH:MM, z.B. 06:00)"
if ($zeitStr -match "^\d{1,2}:\d{2}$") {
    $zeit = [datetime]::ParseExact($zeitStr, "H:mm", $null)

    $scriptPath = Join-Path $BASE "scripts\run_all.ps1"
    $action  = New-ScheduledTaskAction `
        -Execute "powershell.exe" `
        -Argument "-NonInteractive -ExecutionPolicy Bypass -WindowStyle Hidden -File `"$scriptPath`""

    $trigger  = New-ScheduledTaskTrigger -Daily -At $zeit
    $settings = New-ScheduledTaskSettingsSet `
        -ExecutionTimeLimit  (New-TimeSpan -Minutes 30) `
        -RestartCount        2 `
        -RestartInterval     (New-TimeSpan -Minutes 5) `
        -StartWhenAvailable  $true

    try {
        Register-ScheduledTask `
            -TaskName   "BCF_Dashboard_Update" `
            -Action     $action `
            -Trigger    $trigger `
            -Settings   $settings `
            -RunLevel   Highest `
            -Force `
            -ErrorAction Stop | Out-Null
        OK "Task 'BCF_Dashboard_Update' taeglich um $zeitStr eingerichtet"
    } catch {
        ERR "Task konnte nicht angelegt werden: $_"
        WARN "Bitte PowerShell als Administrator ausfuehren und Schritt 6 wiederholen."
    }
} else {
    WARN "Ungueltige Zeitangabe -- Task Scheduler uebersprungen."
    INF "Manuell nachholen: Taskplaner >> BCF_Dashboard_Update"
}

# -- Abschluss ----------------------------------------------
Write-Host ""
Write-Host ("=" * 55) -ForegroundColor DarkGray
Write-Host "  Setup abgeschlossen!" -ForegroundColor Green
Write-Host ""
Write-Host "  Naechste Schritte:" -ForegroundColor White
Write-Host "  1. GitHub Pages aktivieren (falls noch nicht)" -ForegroundColor Gray
Write-Host "     >> Repository Settings >> Pages >> Branch: main / /dashboard" -ForegroundColor DarkGray
Write-Host "  2. SharePoint-Login (einmalig beim ersten Lauf):" -ForegroundColor Gray
Write-Host "     >> scripts\03_sharepoint_download.ps1 manuell starten" -ForegroundColor DarkGray
Write-Host "  3. Ersten Test-Lauf starten:" -ForegroundColor Gray
Write-Host "     >> scripts\run_all.ps1" -ForegroundColor DarkGray
Write-Host ("=" * 55) -ForegroundColor DarkGray
Write-Host ""
Read-Host "  [Enter] zum Beenden"
