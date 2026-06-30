# ============================================================
# 03_sharepoint_download.ps1
# Lädt freigegebene SharePoint/OneDrive-Dateien via
# Microsoft Graph API herunter (MSAL Device-Code-Flow,
# Token wird gecacht für spätere Läufe).
# ============================================================
# Voraussetzung (einmalig):
#   Install-Module Microsoft.Graph.Authentication -Scope CurrentUser
#   Install-Module Microsoft.Graph.Files         -Scope CurrentUser
# ============================================================

$BASE   = Split-Path -Parent (Split-Path -Parent $MyInvocation.MyCommand.Path)
$SOURCE = Join-Path $BASE "_source"
$TOKEN_CACHE = Join-Path $BASE "scripts\.graph_token_cache.json"

# --- SharePoint-Dateien (DriveItem-Download-URLs) -----------
# Format: @{ Dateiname = "SharePoint-Web-URL" }
$SP_FILES = @{
    "Offene_Bestellungen_Draht.xlsx" = "https://baussmannfi-my.sharepoint.com/personal/j_bischopink_baussmann_de/Documents/01_Einkauf/Offene%20Bestellungen_Draht.xlsx"
    "Planung_Pressen_NEU.xlsx"       = "https://baussmannfi-my.sharepoint.com/personal/j_rubner_baussmann_de/Documents/AV_NEU/Produktionsplanung%20und%20BDE/Pressen/Planung_Pressen_NEU.xlsx"
}

# --- Graph-Verbindung herstellen ----------------------------
Write-Host "Verbinde mit Microsoft Graph ..."

try {
    Import-Module Microsoft.Graph.Authentication -ErrorAction Stop
    Import-Module Microsoft.Graph.Files         -ErrorAction Stop
} catch {
    Write-Host "[ERR] Microsoft.Graph Module nicht installiert." -ForegroundColor Red
    Write-Host "Bitte einmalig ausführen:" -ForegroundColor Yellow
    Write-Host "  Install-Module Microsoft.Graph.Authentication -Scope CurrentUser" -ForegroundColor Yellow
    Write-Host "  Install-Module Microsoft.Graph.Files         -Scope CurrentUser" -ForegroundColor Yellow
    exit 1
}

# Interaktiver Login (Tenant: baussmannfi)
$tenantId = "baussmannfi.onmicrosoft.com"
Connect-MgGraph -TenantId $tenantId -Scopes "Files.Read.All", "Sites.Read.All" -NoWelcome

$errors = 0

foreach ($entry in $SP_FILES.GetEnumerator()) {
    $fileName = $entry.Key
    $webUrl   = $entry.Value
    $dstFile  = Join-Path $SOURCE $fileName

    Write-Host "Lade: $fileName ..."

    try {
        # SharePoint-URL → DriveItem auflösen
        # Kodierung bereinigen für Graph-Aufruf
        $encodedUrl = [System.Web.HttpUtility]::UrlEncode($webUrl)
        $graphUrl   = "https://graph.microsoft.com/v1.0/shares/u!$([Convert]::ToBase64String([System.Text.Encoding]::UTF8.GetBytes($webUrl)).Replace('=','').Replace('+','-').Replace('/','_'))/driveItem/content"

        # Download via Graph
        $headers = @{ Authorization = "Bearer $((Get-MgContext).AccessToken)" }
        Invoke-WebRequest -Uri $graphUrl -Headers $headers -OutFile $dstFile -ErrorAction Stop

        Write-Host "[OK]  $fileName → $dstFile" -ForegroundColor Green

    } catch {
        # Fallback: direkt als authenticated download versuchen
        try {
            $token = (Get-MgContext).AccessToken
            if (-not $token) { throw "Kein Access Token verfügbar" }

            # SharePoint REST API Fallback
            $shareUrl   = $webUrl -replace "%20", " "
            $siteBase   = ($shareUrl -split "/personal/")[0] + "/personal/" + ($shareUrl -split "/personal/")[1].Split("/")[0]
            $docPath    = "/" + (($shareUrl -split "/personal/")[1] -split "/",2)[1]

            $restUrl    = "$siteBase/_api/web/GetFileByServerRelativeUrl('$docPath')/`$value"
            $authHeader = @{ Authorization = "Bearer $token"; Accept = "application/json;odata=verbose" }

            Invoke-WebRequest -Uri $restUrl -Headers $authHeader -OutFile $dstFile -ErrorAction Stop
            Write-Host "[OK]  $fileName (REST-Fallback) → $dstFile" -ForegroundColor Green

        } catch {
            Write-Host "[ERR] $fileName : $_" -ForegroundColor Red
            $errors++
        }
    }
}

Disconnect-MgGraph -ErrorAction SilentlyContinue | Out-Null

if ($errors -gt 0) {
    Write-Warning "03_sharepoint_download: $errors Datei(en) fehlgeschlagen."
    exit 1
} else {
    Write-Host "03_sharepoint_download: Alle SharePoint-Dateien geladen." -ForegroundColor Cyan
    exit 0
}
