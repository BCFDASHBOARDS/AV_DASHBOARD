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
#
# Neue Zeilen in der Excel koennen einfach eingefuegt werden --
# das Script liest alle Zeilen dynamisch, kein Hardcoding.
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
# B=2  C=3  D=4(offene Prod)  E=5(Bestand)
# AG=33(2026 YTD)  AH=34(Jan26)..

$COL_ARTNR   = 2
$COL_KTEXT   = 3
$COL_OFFENE  = 4   # D: Offene Produktion
$COL_BESTAND = 5   # E: Lagerbestand
$COL_YTD     = -1  # "YYYY YTD" - dynamisch
$COL_AKTUELL = -1  # Neuester Monat mit Daten - dynamisch

# Spalten durch Scan von Zeile 1 finden
$monatKuerzel = @("Jan","Feb","Mrz","Apr","Mai","Jun","Jul","Aug","Sep","Okt","Nov","Dez")
$curYear      = (Get-Date).Year
$yearLabel    = "$curYear YTD"
$aktuellLabel = ""

$monthColMap = @{}  # month number -> col index

for ($c = 1; $c -le $maxCol; $c++) {
    $hdr = [string]($ws.Cells[1, $c].Value)
    if ($hdr -eq $yearLabel) { $COL_YTD = $c }
    # Monatsspalten des aktuellen Jahres: "Mmm YY"
    $yr2 = ($curYear.ToString()).Substring(2,2)
    for ($m = 1; $m -le 12; $m++) {
        if ($hdr -eq ($monatKuerzel[$m-1] + " " + $yr2)) {
            $monthColMap[$m] = $c
        }
    }
}

if ($COL_YTD -lt 0) {
    Write-Host "[WARN] Spalte '$yearLabel' nicht gefunden, Fallback auf AG (33)." -ForegroundColor Yellow
    $COL_YTD = 33
}

# Neuesten Monat mit Daten finden (von aktuellem Monat rueckwaerts)
$curMonth = (Get-Date).Month
for ($m = $curMonth; $m -ge 1; $m--) {
    if (-not $monthColMap.ContainsKey($m)) { continue }
    $colM = $monthColMap[$m]
    $hasData = $false
    for ($r = 2; $r -le $maxRow; $r++) {
        $v = $ws.Cells[$r, $colM].Value
        if ($v -ne $null -and (Dbl $v) -ne 0) { $hasData = $true; break }
    }
    if ($hasData) {
        $COL_AKTUELL  = $colM
        $aktuellLabel = $monatKuerzel[$m-1] + " " + ($curYear.ToString()).Substring(2,2)
        break
    }
}

if ($COL_AKTUELL -lt 0) {
    Write-Host "[WARN] Kein Monat mit Daten gefunden." -ForegroundColor Yellow
    $COL_AKTUELL = $COL_YTD
    $aktuellLabel = "n/a"
}

# Vormonat: eine Spalte vor Aktuell (gleiche Jahresreihe, aufeinanderfolgende Spalten)
$COL_VORMONAT  = -1
$vormonatLabel = ""
if ($COL_AKTUELL -gt 34) {
    $COL_VORMONAT  = $COL_AKTUELL - 1
    $vormonatLabel = [string]($ws.Cells[1, $COL_VORMONAT].Value)
}

Write-Host "Spalten: Bestand=$COL_BESTAND  OffeneProd=$COL_OFFENE  YTD=$COL_YTD ($yearLabel)  Aktuell=$COL_AKTUELL ($aktuellLabel)  VM=$COL_VORMONAT ($vormonatLabel)"

# ---- Zeilen parsen ------------------------------------------
# Neue Zeilen in der Excel koennen einfach hinzugefuegt werden,
# solange sie vor der naechsten Subtotal-Zeile der Maschine stehen.

$ART_RE  = '^[0-9]{2}-'
$pending = [System.Collections.Generic.List[object]]::new()
$allArts = [System.Collections.Generic.List[object]]::new()

for ($r = 2; $r -le $maxRow; $r++) {
    $artnr = "$($ws.Cells[$r, $COL_ARTNR].Value)".Trim()

    if ($artnr -match $ART_RE) {
        # Artikel-Zeile
        $vm = if ($COL_VORMONAT -gt 0) { Dbl ($ws.Cells[$r, $COL_VORMONAT].Value) } else { 0.0 }
        $art = [PSCustomObject]@{
            artnr       = $artnr
            text        = "$($ws.Cells[$r, $COL_KTEXT].Value)".Trim()
            offene_prod = Dbl ($ws.Cells[$r, $COL_OFFENE].Value)
            bestand     = Dbl ($ws.Cells[$r, $COL_BESTAND].Value)
            ytd         = Dbl ($ws.Cells[$r, $COL_YTD].Value)
            vormonat    = $vm
            aktuell     = Dbl ($ws.Cells[$r, $COL_AKTUELL].Value)
            machine     = $null
            material    = Get-Material $artnr
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
            $pending.RemoveAll([Predicate[object]]{ param($pa) $pa.machine -ne $null }) | Out-Null
        }
    }
}

Write-Host "Artikel gefunden: $($allArts.Count)"

# ---- JSON-Struktur aufbauen ---------------------------------

$machineOrder  = @("Nagelpressen","Doppeldruck")
$materialOrder = @("Stiftdraht","Kaltstauchdraht","Edelstahl","Sonstiges")

$maschinen = [System.Collections.Generic.List[object]]::new()

foreach ($mach in $machineOrder) {
    $mArts = $allArts | Where-Object { $_.machine -eq $mach }
    if (-not $mArts) { continue }

    $mBestand  = ($mArts | Measure-Object -Property bestand    -Sum).Sum
    $mYtd      = ($mArts | Measure-Object -Property ytd        -Sum).Sum
    $mVormonat = ($mArts | Measure-Object -Property vormonat   -Sum).Sum
    $mAktuell  = ($mArts | Measure-Object -Property aktuell    -Sum).Sum
    $mOffene   = ($mArts | Measure-Object -Property offene_prod -Sum).Sum

    $materialgruppen = [System.Collections.Generic.List[object]]::new()

    foreach ($mat in $materialOrder) {
        $matArts = $mArts | Where-Object { $_.material -eq $mat }
        if (-not $matArts) { continue }

        $matB   = ($matArts | Measure-Object -Property bestand    -Sum).Sum
        $matY   = ($matArts | Measure-Object -Property ytd        -Sum).Sum
        $matVM  = ($matArts | Measure-Object -Property vormonat   -Sum).Sum
        $matA   = ($matArts | Measure-Object -Property aktuell    -Sum).Sum
        $matOP  = ($matArts | Measure-Object -Property offene_prod -Sum).Sum

        $artikelList = $matArts | ForEach-Object {
            [PSCustomObject]@{
                artnr       = $_.artnr
                text        = $_.text
                offene_prod = [Math]::Round($_.offene_prod, 2)
                bestand     = [Math]::Round($_.bestand,    2)
                ytd         = [Math]::Round($_.ytd,        2)
                vormonat    = [Math]::Round($_.vormonat,   2)
                aktuell     = [Math]::Round($_.aktuell,    2)
            }
        }

        $materialgruppen.Add([PSCustomObject]@{
            key         = $mat
            label       = $mat
            offene_prod = [Math]::Round($matOP, 2)
            bestand     = [Math]::Round($matB,  2)
            ytd         = [Math]::Round($matY,  2)
            vormonat    = [Math]::Round($matVM, 2)
            aktuell     = [Math]::Round($matA,  2)
            artikel     = @($artikelList)
        })
    }

    $maschinen.Add(