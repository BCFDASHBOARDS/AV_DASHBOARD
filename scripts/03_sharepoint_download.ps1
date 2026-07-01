# ============================================================
# 03_sharepoint_download.ps1
# Laedt SharePoint-Dateien via MSAL.PS (Refresh-Token auf Disk)
# Erster Lauf: Device Code Login. Danach: lautlos.
# ============================================================

$SCRIPTS = $PSScriptRoot
$BASE    = Split-Path -Parent $SCRIPTS
$SOURCE  = Join-Path $BASE "_source"

$SP_FILES = @{
    "Offene_Bestellungen_Draht.xlsx" = "https://baussmannfi-my.sharepoint.com/personal/j_bischopink_baussmann_de/Documents/01_Einkauf/Offene%20Bestellungen_Draht.xlsx"
    "Planung_Pressen_NEU.xlsx"       = "https://baussmannfi-my.sharepoint.com/personal/j_rubner_baussmann_de/Documents/AV_NEU/Produktionsplanung%20und%20BDE/Pressen/Planung_Pressen_NEU.xlsx"
}

$clientId = "14d82eec-204b-4c2f-b7e8-296a70dab67e"
$tenantId = "baussmannfi.onmicrosoft.com"
$scopes   = @("https://graph.microsoft.com/Files.Read.All","https://graph.microsoft.com/Sites.Read.All")

# --- MSAL.PS installieren falls noetig ----------------------
if (-not (Get-Module -ListAvailable -Name MSAL.PS)) {
    Write-Host "Installiere MSAL.PS ..." -ForegroundColor Yellow
    Install-Module MSAL.PS -Scope CurrentUser -Force -AllowClobber -ErrorAction Stop
}
Import-Module MSAL.PS -ErrorAction Stop
Write-Host "[OK] MSAL.PS geladen." -ForegroundColor Green

# --- Token holen (silent aus Cache, sonst Device Code) ------
Write-Host "Hole Graph-Token ..."
$tok = $null

try {
    $tok = Get-MsalToken -ClientId $clientId -TenantId $tenantId -Scopes $scopes -Silent -ErrorAction Stop
    Write-Host "[OK] Token aus Cache." -ForegroundColor Green
} catch {
    Write-Host "Erster Login erforderlich (einmalig):" -ForegroundColor Yellow
    try {
        $tok = Get-MsalToken -ClientId $clientId -TenantId $tenantId -Scopes $scopes -DeviceCode -ErrorAction Stop
        Write-Host "[OK] Login erfolgreich." -ForegroundColor Green
    } catch {
        Write-Host "[ERR] Login fehlgeschlagen: $_" -ForegroundColor Red
        exit 1
    }
}

if (-not $tok -or -not $tok.AccessToken) {
    Write-Host "[ERR] Kein Access Token." -ForegroundColor Red
    exit 1
}

$headers = @{ Authorization = "Bearer $($tok.AccessToken)" }

# --- Dateien laden ------------------------------------------
$errors = 0

foreach ($entry in $SP_FILES.GetEnumerator()) {
    $fileName = $entry.Key
    $webUrl   = $entry.Value
    $dstFile  = [string](Join-Path $SOURCE $fileName)

    Write-Host "Lade: $fileName ..."

    try {
        $bytes    = [System.Text.Encoding]::UTF8.GetBytes($webUrl)
        $b64      = [Convert]::ToBase64String($bytes).TrimEnd('=').Replace('+','-').Replace('/','_')
        $shareId  = "u!$b64"
        $graphUri = "https://graph.microsoft.com/v1.0/shares/$shareId/driveItem/content"

        Invoke-WebRequest -Uri $graphUri -Headers $headers -OutFile $dstFile -ErrorAction Stop
        Write-Host "[OK]  $fileName" -ForegroundColor Green
    } catch {
        Write-Host "[ERR] $fileName : $_" -ForegroundColor Red
        $errors++
    }
}

if ($errors -gt 0) {
    Write-Warning "03_sharepoint_download: $errors Datei(en) fehlgeschlagen."
    exit 1
} else {
    Write-Host "03_sharepoint_download: Alle Dateien geladen." -ForegroundColor Cyan
    exit 0
}
