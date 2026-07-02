# ============================================================
# 07_magazinierung.ps1
# Extrakt MAG-Produktionsmengen aus Istmengen.xlsx
# Gruppen: CDF_UTC (50-/51-), Papertape (45-/46-/47-),
#          Drahtcoil (30-), Streifennaegel_KS (40-/41-/42-),
#          Ankernaegel (40-40/40-60/42-40/50-40/47-400, ohne 62-),
#          KRL (60-), BSN (55-)
# Menge netto = Spalte C minus Spalte D
# ============================================================

$SCRIPTS   = [string]$PSScriptRoot
$BASE      = [string](Split-Path -Parent $SCRIPTS)
$SRC_DIR   = [string](Join-Path $BASE "_source")
$OUT_DIR   = [string](Join-Path $BASE "_data")
$SRC_XLS   = [string](Join-Path $SRC_DIR "Istmengen.xlsx")
$JSON_OUT  = [string](Join-Path $OUT_DIR "magazinierung.json")
$ae        = [char]0x00E4
$SRC_STAMM = [string](Join-Path $SRC_DIR "fertigungsauftr${ae}ge_stammdaten_gesamt.xlsx")

if (-not (Test-Path $OUT_DIR)) { New-Item -ItemType Directory -Path $OUT_DIR | Out-Null }

Write-Host "=== 07_magazinierung.ps1 ===" -ForegroundColor Cyan

try {
    Import-Module ImportExcel -ErrorAction Stop
} catch {
    Write-Host "[ERR] ImportExcel nicht installiert." -ForegroundColor Red; exit 1
}

# ---- Hilfsfunktionen -----------------------------------------

function Get-MAGGruppe([string]$artnr) {
    if ([string]::IsNullOrWhiteSpace($artnr)) { return $null }
    $p3 = if ($artnr.Length -ge 3) { $artnr.Substring(0,3) } else { $artnr }
    switch ($p3) {
        "30-" { return "Drahtcoil" }
        "40-" { return "Streifennaegel_KS" }
        "41-" { return "Streifennaegel_KS" }
        "42-" { return "Streifennaegel_KS" }
        "45-" { return "Papertape" }
        "46-" { return "Papertape" }
        "47-" { return "Papertape" }
        "50-" { return "CDF_UTC" }
        "51-" { return "CDF_UTC" }
        "55-" { return "BSN" }
        "60-" { return "KRL" }
        # 62- wird bewusst nicht erfasst (Lose, nicht MAG-relevant)
        default { return $null }
    }
}

function Test-IsAnkernagel([string]$artnr) {
    # Ankernagel bleibt in Hauptgruppe, wird aber zusaetzlich als davon_ankernaegel gezaehlt
    if ($artnr.Length -ge 5) {
        $p5 = $artnr.Substring(0,5)
        if ($p5 -in @("40-40","40-60","42-40","50-40")) { return $true }
    }
    if ($artnr.Length -ge 6 -and $artnr.Substring(0,6) -eq "47-400") { return $true }
    return $false
}

function Test-ShouldSkip([string]$artnr, [string]$gruppe) {
    # Regel 1: BSN (55-) -- nur Artikel mit Kundennummer-Suffix (-DDDDD) behalten
    if ($gruppe -eq "BSN") {
        if ($artnr -notmatch '-\d{5}$') { return $true }
    }
    # Regel 2: CDF_UTC -- Artikel mit *815-11913, *905-11913, *605-11913 ignorieren
    if ($gruppe -eq "CDF_UTC") {
        if ($artnr -match '(815|905|605)-11913$') { return $true }
    }
    return $false
}

# ---- Kurztexte aus Stammdaten_gesamt laden -------------------
$ktextMap = @{}
Write-Host "Lese Stammdaten (MAG und HKL) ..." -ForegroundColor Yellow
if (Test-Path $SRC_STAMM) {
    try {
        $stammdaten = Import-Excel -Path $SRC_STAMM -WorksheetName "MAG und HKL" `
            -HeaderName @("Auftrag","Artikelnummer","Zustand","Kurztext","Menge_Auftrag","ME",
                          "StkPal","GewPro1000","StkKarton","Kundenauftrag","Kunde","Zusatz","Typ") `
            -StartRow 2 -ErrorAction Stop
        foreach ($row in $stammdaten) {
            $an = [string]$row.Artikelnummer
            $kt = [string]$row.Kurztext
            if ($an -and $kt -and -not $ktextMap.ContainsKey($an)) {
                $ktextMap[$an] = $kt.Trim()
            }
        }
        Write-Host "  Kurztexte geladen: $($ktextMap.Count)" -ForegroundColor DarkGray
    } catch {
        Write-Host "  [WARN] Kurztexte nicht ladbar: $_" -ForegroundColor Yellow
    }
} else {
    Write-Host "  [WARN] Stammdaten nicht gefunden: $SRC_STAMM" -ForegroundColor Yellow
}

# ---- Excel einlesen ------------------------------------------
Write-Host "Lese Istmengen.xlsx ..." -ForegroundColor Yellow
if (-not (Test-Path $SRC_XLS)) {
    Write-Host "[ERR] Quelldatei nicht gefunden: $SRC_XLS" -ForegroundColor Red; exit 1
}

$xlRows = Import-Excel -Path $SRC_XLS -WorksheetName "Tabelle1" `
          -HeaderName @("Auftrag","Ressource","Menge_C","Menge_D") `
          -StartRow 2 -ErrorAction Stop

Write-Host "  $($xlRows.Count) Zeilen gelesen." -ForegroundColor DarkGray

# ---- Auswerten -----------------------------------------------
$gruppen    = @{ CDF_UTC=0.0; Papertape=0.0; Drahtcoil=0.0;
                 Streifennaegel_KS=0.0; KRL=0.0; BSN=0.0 }
$artMengen  = @{}   # artnr -> netto menge (konsolidiert)
$artGruppe  = @{}   # artnr -> gruppe
$ankerMenge = 0.0   # davon_ankernaegel (subset, bleibt in Hauptgruppe)
$ankerArts  = @{}   # artnr -> menge (nur Ankernagel-Artikel)

foreach ($row in $xlRows) {
    $artnr = [string]($row.Ressource)
    $c     = [double]($row.Menge_C -replace ',','.' )
    $d     = [double]($row.Menge_D -replace ',','.' )
    $netto = $c - $d
    if ($netto -le 0) { continue }

    $g = Get-MAGGruppe $artnr
    if ($null -eq $g) { continue }
    if (Test-ShouldSkip $artnr $g) { continue }

    if (-not $artMengen.ContainsKey($artnr)) {
        $artMengen[$artnr] = 0.0
        $artGruppe[$artnr] = $g
    }
  