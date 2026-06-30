# ============================================================
# 04_excel_to_json.ps1
# Liest alle _source\*.xlsx via ImportExcel-Modul und
# exportiert jedes erste Tabellenblatt als JSON nach _data\
# ============================================================
# Voraussetzung (einmalig):
#   Install-Module ImportExcel -Scope CurrentUser
# ============================================================

$BASE   = Split-Path -Parent (Split-Path -Parent $MyInvocation.MyCommand.Path)
$SOURCE = Join-Path $BASE "_source"
$DATA   = Join-Path $BASE "_data"

try {
    Import-Module ImportExcel -ErrorAction Stop
} catch {
    Write-Host "[ERR] ImportExcel-Modul nicht installiert." -ForegroundColor Red
    Write-Host "Bitte einmalig ausführen: Install-Module ImportExcel -Scope CurrentUser" -ForegroundColor Yellow
    exit 1
}

# Mapping: Dateiname → JSON-Key + Blattname (leer = erstes Blatt)
$FILE_MAP = @(
    @{ File = "1100_Datenimport_Pressen.xlsx";    Key = "pressen";           Sheet = "" }
    @{ File = "1180_Datenimport_Walzen.xlsx";     Key = "walzen";            Sheet = "" }
    @{ File = "Auftragsbestand.xlsx";             Key = "auftragsbestand";   Sheet = "" }
    @{ File = "Istmengen.xlsx";                   Key = "istmengen";         Sheet = "" }
    @{ File = "Lagerbestand.xlsx";                Key = "lagerbestand";      Sheet = "" }
    @{ File = "ubersicht_Draht.xlsx";             Key = "draht_ubersicht";   Sheet = "" }
    @{ File = "Offene_Bestellungen_Draht.xlsx";   Key = "draht_bestellungen";Sheet = "" }
    @{ File = "Planung_Pressen_NEU.xlsx";         Key = "planung_pressen";   Sheet = "" }
)

$result  = @{}
$errors  = 0
$updated = Get-Date -Format "yyyy-MM-ddTHH:mm:ss"

foreach ($entry in $FILE_MAP) {
    $filePath = Join-Path $SOURCE $entry.File

    if (-not (Test-Path $filePath)) {
        Write-Warning "[$($entry.Key)] Datei nicht gefunden: $($entry.File) — übersprungen."
        continue
    }

    try {
        $params = @{ Path = $filePath; ErrorAction = "Stop" }
        if ($entry.Sheet -ne "") { $params["WorksheetName"] = $entry.Sheet }

        $data = Import-Excel @params

        # Leere Zeilen entfernen (alle Felder null/leer)
        $data = $data | Where-Object {
            ($_.PSObject.Properties.Value | Where-Object { $_ -ne $null -and "$_".Trim() -ne "" }).Count -gt 0
        }

        $result[$entry.Key] = $data
        Write-Host "[OK]  $($entry.File) → $($data.Count) Zeilen" -ForegroundColor Green

    } catch {
        Write-Host "[ERR] $($entry.File): $_" -ForegroundColor Red
        $errors++
    }
}

# Meta-Informationen hinzufügen
$result["_meta"] = @{
    updated  = $updated
    source   = "PowerShell-Export"
    files    = ($FILE_MAP | ForEach-Object { $_.Key })
}

# Ausgabe als JSON
$jsonPath = Join-Path $DATA "dashboard_data.json"
$result | ConvertTo-Json -Depth 10 -Compress | Out-File -FilePath $jsonPath -Encoding UTF8 -Force

Write-Host "JSON gespeichert: $jsonPath" -ForegroundColor Cyan

# Einzelne JSON-Dateien pro Datensatz (für späteres lazy loading)
foreach ($key in $result.Keys) {
    if ($key -eq "_meta") { continue }
    $singlePath = Join-Path $DATA "$key.json"
    $result[$key] | ConvertTo-Json -Depth 10 -Compress | Out-File -FilePath $singlePath -Encoding UTF8 -Force
}

if ($errors -gt 0) {
    Write-Warning "04_excel_to_json: $errors Fehler aufgetreten."
    exit 1
} else {
    Write-Host "04_excel_to_json: Export abgeschlossen." -ForegroundColor Cyan
    exit 0
}
