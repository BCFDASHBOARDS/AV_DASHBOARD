# ============================================================
# 04_excel_to_json.ps1
# Liest _source\*.xlsx und schreibt JSON nach _data\
# ============================================================

$SCRIPTS = [string]$PSScriptRoot
$BASE    = [string](Split-Path -Parent $SCRIPTS)
$SOURCE  = [string](Join-Path $BASE "_source")
$DATA    = [string](Join-Path $BASE "_data")

Write-Host "BASE:   $BASE"
Write-Host "SOURCE: $SOURCE"
Write-Host "DATA:   $DATA"

if (-not (Test-Path $DATA)) { New-Item -ItemType Directory -Path $DATA | Out-Null }

try {
    Import-Module ImportExcel -ErrorAction Stop
} catch {
    Write-Host "[ERR] ImportExcel nicht installiert: Install-Module ImportExcel -Scope CurrentUser" -ForegroundColor Red
    exit 1
}

$FILE_MAP = @(
    @{ File = "1100_Datenimport_Pressen.xlsx";    Key = "pressen";            Sheet = "" }
    @{ File = "1180_Datenimport_Walzen.xlsx";     Key = "walzen";             Sheet = "" }
    @{ File = "Auftragsbestand.xlsx";             Key = "auftragsbestand";    Sheet = "" }
    @{ File = "Istmengen.xlsx";                   Key = "istmengen";          Sheet = "" }
    @{ File = "Lagerbestand.xlsx";                Key = "lagerbestand";       Sheet = "" }
    @{ File = "ubersicht_Draht.xlsx";             Key = "draht_ubersicht";    Sheet = "" }
    @{ File = "Offene_Bestellungen_Draht.xlsx";   Key = "draht_bestellungen"; Sheet = "" }
    @{ File = "Planung_Pressen_NEU.xlsx";         Key = "planung_pressen";    Sheet = "" }
)

$errors  = 0
$updated = Get-Date -Format "yyyy-MM-ddTHH:mm:ss"

foreach ($entry in $FILE_MAP) {
    $filePath = [string](Join-Path $SOURCE $entry.File)

    if (-not (Test-Path $filePath)) {
        Write-Warning "[$($entry.Key)] nicht gefunden: $($entry.File) -- uebersprungen."
        continue
    }

    Write-Host "Lese: $($entry.File) ..."

    try {
        $params = @{ Path = $filePath; ErrorAction = "Stop" }
        if ($entry.Sheet -ne "") { $params["WorksheetName"] = $entry.Sheet }

        $data = $null
        try {
            $data = Import-Excel @params
        } catch {
            if ("$_" -like "*Duplicate column*") {
                Write-Warning "  Duplikat-Spalten in $($entry.File) -- lese ohne Header."
                $params.Remove("ErrorAction")
                $params["NoHeader"] = $true
                $params["ErrorAction"] = "Stop"
                $data = Import-Excel @params
            } else {
                throw
            }
        }

        # Leere Zeilen entfernen
        $rows = @($data | Where-Object {
            ($_.PSObject.Properties.Value |
             Where-Object { $_ -ne $null -and "$_".Trim() -ne "" }).Count -gt 0
        })

        # JSON in Variable sammeln -- kein Objekt-Durchlauf durch Pipeline
        $json = $rows | ConvertTo-Json -Depth 5 -Compress
        if ($null -eq $json -or $json -eq "") { $json = "[]" }

        $outPath = [string](Join-Path $DATA "$($entry.Key).json")
        [System.IO.File]::WriteAllText($outPath, $json, [System.Text.Encoding]::UTF8)

        Write-Host "[OK]  $($entry.File) -> $($rows.Count) Zeilen -> $($entry.Key).json" -ForegroundColor Green

    } catch {
        Write-Host "[ERR] $($entry.File): $_" -ForegroundColor Red
        $errors++
    }
}

# Meta
$metaJson = "{""updated"":""$updated"",""source"":""PowerShell-Export""}"
$metaPath = [string](Join-Path $DATA "_meta.json")
[System.IO.File]::WriteAllText($metaPath, $metaJson, [System.Text.Encoding]::UTF8)

if ($errors -gt 0) {
    Write-Warning "04_excel_to_json: $errors Fehler."
    exit 1
} else {
    Write-Host "04_excel_to_json: Export abgeschlossen." -ForegroundColor Cyan
    exit 0
}
