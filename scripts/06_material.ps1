# ============================================================
# 06_material.ps1
# Extraktion Materialdaten aus ubersicht_Draht.xlsx
# Sheet "Verbrauch N YTD" -> docs/_data/material.json
#
# Gruppen-Logik: Artikel-Zeilen werden der Maschinengruppe der
# naechsten Subtotal-Zeile (B="Nagelpressen"/"Doppeldruck")
# zugeordnet (pending-flush-Verfahren).
# Materialkategorie: 01/04=Stiftdraht, 03/06=Kaltstauchdraht,
#                    08/09=Edelstahl
# ============================================================

$SCRIPTS  = $PSScriptRoot
$BASE     = Split-Path -Parent $SCRIPTS
$SRC_XLS  = Join-Path $BASE "_source\ubersicht_Draht.xlsx"
$JSON_OUT = Join-Path $BASE "_data\material.json"

# ---- Hilfsfunktionen ----------------------------------------

function Dbl($v) {
    if ($v -eq $null) { return 0.0 }
    try { return [double]$v } catch { return 0.0 }
}

function Get-Material([string]$artnr) {
    $pfx = if ($artnr.Length -ge 2) { $artnr.Substring(0,2) } else { "" }
    switch ($pfx) {
        "01" { return "Stiftdraht" }
        "04" { return "Stiftdraht" }
        "03" { return "Kaltstauchdraht" }
        "06" { return "Kaltstauchdraht" }
        "08" { return "Edelstahl" }
        "09" { return "Edelstahl" }
        default { return "Sonstiges" }
    }
}

# ---- Excel oeffnen ------------------------------------------

try { Import-Module ImportExcel -ErrorAction Stop }
catch { Write-Host "[ERR] ImportExcel nicht installiert." -ForegroundColor Red; exit 1 }

if (-not (Test-Path $SRC_XLS)) {
    Write-Host "[ERR] Quelldatei nicht gefunden: $SRC_XLS" -ForegroundColor Red; exit 1
}

Write-Host "Oeffne $SRC_XLS ..."
$pkg = Open-ExcelPackage -Path $SRC_XLS
$ws  = $pkg.Workbook.Worksheets["Verbrauch N YTD"]
if (-not $ws) {
    Write-Host "[ERR] Sheet 'Verbrauch N YTD' nicht gefunden." -ForegroundColor Red
    Close-ExcelPackage $pkg -NoSave; exit 1
}

$maxRow = $ws.Dimension.Rows
$maxCol = $ws.Dimension.Columns
Write-Host "Sheet geladen: $maxRow Zeilen, $maxCol Spalten."

# ---- Spalten-Indices (1-basiert) ----------------------------
# B=2  C=3  D=4  E=5  AG=33(2026 YTD)  AH=34(Jan26)..AM=39(Jun26)

$COL_ARTNR   = 2
$COL_KTEXT   = 3
$COL_BESTAND = 5
$COL_YTD     = -1   # "YYYY YTD" - dynamisch suchen
$COL_VM      = -1   # "Mmm YY" Vormonat - dynamisch suchen

# Vormonat-Label ermitteln
$monatKuerzel = @("Jan","Feb","Mrz","Apr","Mai","Jun","Jul","Aug","Sep","Okt","Nov","Dez")
$prevDate     = (Get-Date).AddMonths(-1)
$prevLabel    = $monatKuerzel[$prevDate.Month - 1] + " " + $prevDate.ToString("yy")
$yearLabel    = (Get-Date).Year.ToString() + " YTD"

# Spalten durch Scan von Zeile 1 finden
for ($c = 1; $c -le $maxCol; $c++) {
    $hdr = [string]($ws.Cells[1, $c].Value)
    if ($hdr -eq $yearLabel)  { $COL_YTD = $c }
    if ($hdr -eq $prevLabel)  { $COL_VM  = $c }
}

if ($COL_YTD -lt 0) {
    Write-Host "[WARN] Spalte '$yearLabel' nicht gefunden, Fallback auf AG (33)." -ForegroundColor Yellow
    $COL_YTD = 33
}
if ($COL_VM -lt 0) {
    Write-Host "[WARN] Spalte '$prevLabel' nicht gefunden, Vormonat-Werte = 0." -ForegroundColor Yellow
}

Write-Host "Spalten: Bestand=$COL_BESTAND  YTD=$COL_YTD ($yearLabel)  VM=$COL_VM ($prevLabel)"

# ---- Zeilen parsen ------------------------------------------

$ART_RE   = '^[0-9]{2}-'
$pending  = [System.Collections.Generic.List[object]]::new()
$allArts  = [System.Collections.Generic.List[object]]::new()

for ($r = 2; $r -le $maxRow; $r++) {
    $artnr = "$($ws.Cells[$r, $COL_ARTNR].Value)".Trim()

    if ($artnr -match $ART_RE) {
        # Artikel-Zeile
        $vm = if ($COL_VM -gt 0) { Dbl ($ws.Cells[$r, $COL_VM].Value) } else { 0.0 }
        $art = [PSCustomObject]@{
            artnr    = $artnr
            text     = "$($ws.Cells[$r, $COL_KTEXT].Value)".Trim()
            bestand  = Dbl ($ws.Cells[$r, $COL_BESTAND].Value)
            ytd      = Dbl ($ws.Cells[$r, $COL_YTD].Value)
            vormonat = $vm
            machine  = $null
            material = Get-Material $artnr
        }
        $pending.Add($art)
    } else {
        # Subtotal / Gruppen-Zeile -- Maschine den wartenden Artikeln zuweisen
        $machine = $artnr.Trim()
        if ($machine -eq "Nagelpressen" -or $machine -eq "Doppeldruck") {
            foreach ($pa in $pending) {
                if ($pa.machine -eq $null) {
                    $pa.machine = $machine
                    $allArts.Add($pa)
                }
            }
            $pending.RemoveAll({ param($pa) $pa.machine -ne $null }) | Out-Null
        }
    }
}

Write-Host "Artikel gefunden: $($allArts.Count)"

# ---- JSON-Struktur aufbauen ---------------------------------

# Maschinenreihenfolge
$machineOrder    = @("Nagelpressen","Doppeldruck")
$materialOrder   = @("Stiftdraht","Kaltstauchdraht","Edelstahl","Sonstiges")

$maschinen = [System.Collections.Generic.List[object]]::new()

foreach ($mach in $machineOrder) {
    $mArts = $allArts | Where-Object { $_.machine -eq $mach }
    if (-not $mArts) { continue }

    $mBestand  = ($mArts | Measure-Object -Property bestand  -Sum).Sum
    $mYtd      = ($mArts | Measure-Object -Property ytd      -Sum).Sum
    $mVormonat = ($mArts | Measure-Object -Property vormonat -Sum).Sum

    $materialgruppen = [System.Collections.Generic.List[object]]::new()

    foreach ($mat in $materialOrder) {
        $matArts = $mArts | Where-Object { $_.material -eq $mat }
        if (-not $matArts) { continue }

        $matB  = ($matArts | Measure-Object -Property bestand  -Sum).Sum
        $matY  = ($matArts | Measure-Object -Property ytd      -Sum).Sum
        $matVM = ($matArts | Measure-Object -Property vormonat -Sum).Sum

        $artikelList = $matArts | ForEach-Object {
            [PSCustomObject]@{
                artnr    = $_.artnr
                text     = $_.text
                bestand  = [Math]::Round($_.bestand,  2)
                ytd      = [Math]::Round($_.ytd,      2)
                vormonat = [Math]::Round($_.vormonat, 2)
            }
        }

        $materialgruppen.Add([PSCustomObject]@{
            key      = $mat
            label    = $mat
            bestand  = [Math]::Round($matB,  2)
            ytd      = [Math]::Round($matY,  2)
            vormonat = [Math]::Round($matVM, 2)
            artikel  = @($artikelList)
        })
    }

    $maschinen.Add([PSCustomObject]@{
        key             = $mach
        label           = $mach
        bestand         = [Math]::Round($mBestand,  2)
        ytd             = [Math]::Round($mYtd,      2)
        vormonat        = [Math]::Round($mVormonat, 2)
        materialgruppen = @($materialgruppen)
    })
}

# Gesamtsummen
$gesamtBestand  = [Math]::Round(($allArts | Measure-Object -Property bestand  -Sum).Sum, 2)
$gesamtYtd      = [Math]::Round(($allArts | Measure-Object -Property ytd      -Sum).Sum, 2)
$gesamtVormonat = [Math]::Round(($allArts | Measure-Object -Property vormonat -Sum).Sum, 2)

# Globale Summen je Materialkategorie (maschinenuebergreifend)
$sumByMat = @{}
foreach ($mat in $materialOrder) {
    $mats = $allArts | Where-Object { $_.material -eq $mat }
    if ($mats) {
        $sumByMat[$mat] = [Math]::Round(($mats | Measure-Object -Property vormonat -Sum).Sum, 2)
    }
}

$result = [PSCustomObject]@{
    timestamp             = (Get-Date -Format "yyyy-MM-ddTHH:mm:ss")
    vormonat_label        = $prevLabel
    gesamt_bestand        = $gesamtBestand
    gesamt_ytd            = $gesamtYtd
    gesamt_vormonat       = $gesamtVormonat
    vormonat_stiftdraht   = if ($sumByMat["Stiftdraht"])        { $sumByMat["Stiftdraht"] }        else { 0.0 }
    vormonat_kaltstauch   = if ($sumByMat["Kaltstauchdraht"])   { $sumByMat["Kaltstauchdraht"] }   else { 0.0 }
    vormonat_edelstahl    = if ($sumByMat["Edelstahl"])         { $sumByMat["Edelstahl"] }         else { 0.0 }
    maschinen             = @($maschinen)
}

# ---- Ausgabe ------------------------------------------------

$dataDir = Join-Path $BASE "_data"
if (-not (Test-Path $dataDir)) { New-Item -ItemType Directory -Path $dataDir | Out-Null }

$json = $result | ConvertTo-Json -Depth 6 -Compress:$false
[System.IO.File]::WriteAllText($JSON_OUT, $json, [System.Text.Encoding]::UTF8)

Write-Host "[OK] material.json geschrieben: $JSON_OUT" -ForegroundColor Green
Write-Host "     Bestand: $gesamtBestand kg  |  YTD: $gesamtYtd kg  |  Vormonat ($prevLabel): $gesamtVormonat kg"

Close-ExcelPackage $pkg -NoSave
exit 0
          