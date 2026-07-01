# ============================================================
# 03_sharepoint_download.ps1
# Laedt SharePoint-Dateien via Microsoft Graph REST API.
# Kein MSAL.PS noetig — Refresh-Token wird DPAPI-verschluesselt
# auf Disk gespeichert. Einmalig per Device Code einloggen,
# danach ca. 90 Tage lautlos (rolling refresh).
# ============================================================

$SCRIPTS = $PSScriptRoot
$BASE    = Split-Path -Parent $SCRIPTS
$SOURCE  = Join-Path $BASE "_source"

$SP_FILES = @{
    "Offene_Bestellungen_Draht.xlsx" = "https://baussmannfi-my.sharepoint.com/personal/j_bischopink_baussmann_de/Documents/01_Einkauf/Offene%20Bestellungen_Draht.xlsx"
    "Planung_Pressen_NEU.xlsx"       = "https://baussmannfi-my.sharepoint.com/personal/j_rubner_baussmann_de/Documents/AV_NEU/Produktionsplanung%20und%20BDE/Pressen/Planung_Pressen_NEU.xlsx"
}

$clientId  = "14d82eec-204b-4c2f-b7e8-296a70dab67e"
$tenantId  = "baussmannfi.onmicrosoft.com"
$scopeStr  = "https://graph.microsoft.com/Files.Read.All https://graph.microsoft.com/Sites.Read.All offline_access"
$TOKEN_URL = "https://login.microsoftonline.com/$tenantId/oauth2/v2.0/token"
$DC_URL    = "https://login.microsoftonline.com/$tenantId/oauth2/v2.0/devicecode"
$CACHE     = Join-Path $env:LOCALAPPDATA "BCF_Dashboard\rt.bin"

# DPAPI verfuegbar machen
Add-Type -AssemblyName System.Security

function Save-RT([string]$rt) {
    $dir = Split-Path -Parent $CACHE
    if (-not (Test-Path $dir)) { New-Item -ItemType Directory -Path $dir | Out-Null }
    $bytes = [System.Text.Encoding]::UTF8.GetBytes($rt)
    $enc   = [System.Security.Cryptography.ProtectedData]::Protect(
                 $bytes, $null,
                 [System.Security.Cryptography.DataProtectionScope]::CurrentUser)
    [System.IO.File]::WriteAllBytes($CACHE, $enc)
}

function Load-RT {
    if (-not (Test-Path $CACHE)) { return $null }
    try {
        $enc   = [System.IO.File]::ReadAllBytes($CACHE)
        $bytes = [System.Security.Cryptography.ProtectedData]::Unprotect(
                     $enc, $null,
                     [System.Security.Cryptography.DataProtectionScope]::CurrentUser)
        return [System.Text.Encoding]::UTF8.GetString($bytes)
    } catch {
        Write-Host "[WARN] Token-Cache unlesbar, neu einloggen: $_" -ForegroundColor Yellow
        return $null
    }
}

# ── Token holen ──────────────────────────────────────────────
$accessToken = $null

# Versuch 1: Refresh Token aus Cache
$rt = Load-RT
if ($rt) {
    Write-Host "Hole Graph-Token per Refresh ..." -ForegroundColor DarkGray
    try {
        $resp = Invoke-RestMethod -Method Post -Uri $TOKEN_URL `
            -ContentType "application/x-www-form-urlencoded" `
            -Body "client_id=$clientId&scope=$([uri]::EscapeDataString($scopeStr))&refresh_token=$([uri]::EscapeDataString($rt))&grant_type=refresh_token" `
            -ErrorAction Stop
        $accessToken = $resp.access_token
        if ($resp.refresh_token) { Save-RT $resp.refresh_token }   # rolling refresh
        Write-Host "[OK] Token erneuert (kein Login noetig)." -ForegroundColor Green
    } catch {
        Write