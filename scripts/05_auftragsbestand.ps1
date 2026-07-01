# ============================================================
# 05_auftragsbestand.ps1
# Extraktion Auftragsbestand.xlsx:
#   - Gesamtbestand, Max-Auftragsnr, Tageseingang
#   - Wocheneingang (letzte KW), Monatseingang (lfd. Monat)
#   - Top-5 Kunden (Wuerth/ITW/Kyocera konsolidiert)
#   - Artikelgruppen + Sondergruppe Ankernagel
#   - Plausibilitaets-Check
#   - Tageslog als XLSX + JSON fuer Dashboard
# ============================================================

$SCRIPTS = [string]$PSScriptRoot
$BASE    = [string](Split-Path -Parent $SCRIPTS)
$SRC_DIR = [string](Join-Path $BASE "_source")
$OUT_DIR = [string](Join-Path $BASE "_data")
$SRC_XLS = [string](Join-Path $SRC_DIR "Auftragsbestand.xlsx")
$LOG_XLS = [string](Join-Path $OUT_DIR "auftragsbestand_log.xlsx")
$JSON_OUT = [string](Join-Path $OUT_DIR "auftragsbestand.json")

if (-not (Test-Path $OUT_DIR)) { New-Item -ItemType Directory -Path $OUT_DIR | Out-Null }

Write-Host "=== 05_auftragsbestand.ps1 ===" -ForegroundColor Cyan

# --- Module ----------------------------------------------------
try {
    Import-Module ImportExcel -ErrorAction Stop
} catch {
    Write-Host "[ERR] ImportExcel nicht installiert." -ForegroundColor Red; exit 1
}

# --- Hilfsfunktionen -------------------------------------------

function Get-Gruppe([string]$artnr) {
    if ([string]::IsNullOrWhiteSpace($artnr)) { return "Sonstiges" }
    $p = if ($artnr.Length -ge 3) { $artnr.Substring(0,3) } else { $artnr }
    switch ($p) {
        "01-" { return "Maschinenstifte" }
        "03-" { return "Kernschrauben" }
        "08-" { return "Lose_Edelstahl" }
        "09-" { return "Lose_Edelstahl" }
        "20-" { return "Industrieklammern" }
        "21-" { return "Industrieklammern" }
        "27-" { return "Umpack_Dachpapp" }
        "30-" { return "Drahtcoil" }
        "36-" { return "Umpack_Dachpapp" }
        "40-" { return "Streifennaegel_KS" }
        "41-" { return "Streifennaegel_KS" }
        "42-" { return "Streifennaegel_KS" }
        "45-" { return "Papertape" }
        "46-" { return "Papertape" }
        "47-" { return "Papertape" }
        "50-" { return "CDF_UTC" }
        "62-" { return "Ankernaegel_Lose" }
        default { return "Sonstiges" }
    }
}

function Test-IsAnkernagel([string]$artnr) {
    if ([string]::IsNullOrWhiteSpace($artnr)) { return $false }
    $p5 = if ($artnr.Length -ge 5) { $artnr.Substring(0,5) } else { $artnr }
    $p6 = if ($artnr.Length -ge 6) { $artnr.Substring(0,6) } else { $artnr }
    if ($p5 -in @("40-40","40-60","42-40","50-40","62-40","62-60")) { return $true }
    if ($p6 -eq "47-400") { return $true }  # Ankernagel Papertape
    return $false
}

function Get-KonsolidierterKunde([string]$kunde) {
    if ([string]::IsNullOrWhiteSpace($kunde)) { return $kunde }
    $kl = $kunde.ToLower()
    if ($kl -match "w.rth|wuerth")   { return "Wuerth" }
    if ($kl -match "\bitw\b")         { return "ITW" }
    if ($kl -match "kyocera")         { return "Kyocera" }
    if ($kl -match "hilti")           { return "Hilti" }
    if ($kl -match "tti|milwaukee")   { return "Milwaukee" }
    return $kunde
}

# --- Excel einlesen --------------------------------------------
Write-Host "Lese Auftragsbestand.xlsx ..."
if (-not (Test-Path $SRC_XLS)) {
    Write-Host "[ERR] Quelldatei nicht gefunden: $SRC_XLS" -ForegroundColor Red; exit 1
}

# Duplikate in Kopfzeile abfangen: Header manuell lesen + eindeutig machen
$pkg = Open-ExcelPackage -Path $SRC_XLS
$ws  = $pkg.Workbook.Worksheets["Tabelle1"]
$colCount = $ws.Dimension.Columns
$headerNames = @()
$seen = @{}
for ($c = 1; $c -le $colCount; $c++) {
    $h = [string]$ws.Cells[1, $c].Value
    if ([string]::IsNullOrWhiteSpace($h)) { $h = "Spalte$c" }
    if ($seen.ContainsKey($h)) {
        $seen[$h]++
        $h = "${h}_$($seen[$h])"
    } else {
        $seen[$h] = 0
    }
    $headerNames += $h
}
Close-ExcelPackage $pkg -NoSave
Write-Host "  Spalten ($colCount): $($headerNames -join ', ')"

$xlRows = Import-Excel -Path $SRC_XLS -WorksheetName "Tabelle1" `
          -HeaderName $headerNames -StartRow 2 -ErrorAction Stop

# --- Auswerten -------------------------------------------------
$gruppen = @{
    "Maschinenstifte"   = 0.0
    "Kernschrauben"     = 0.0
    "Lose_Edelstahl"    = 0.0
    "Industrieklammern" = 0.0
    "Drahtcoil"         = 0.0
    "Streifennaegel_KS" = 0.0
    "Papertape"         = 0.0
    "CDF_UTC"           = 0.0
    "Ankernaegel_Lose"  = 0.0
    "Umpack_Dachpapp"   = 0.0
    "Sonstiges"         = 0.0
}
$kundenMap    = @{}
$kundenDetail = @{}   # Artikel je Kunde (konsolidiert): $kundenDetail[$kName][$artnr] = @{Wert=...; Text=...}
$gruppenDetail = @{}  # Top-Artikel + Top-Kunden (roh) je Gruppe
$gesamt       = 0.0
$maxAuftrag   = 0
$davonAnker   = 0.0

foreach ($xlRow in $xlRows) {
    $wert    = 0.0
    $auftrNr = 0
    try { $wert    = [double]$xlRow.'Wert'   } catch {}
    try { $auftrNr = [int]$xlRow.'Auftrag'   } catch {}
    $artnr  = [string]$xlRow.'Artikelnummer'
    $ktext  = [string]$xlRow.'Kurztext'
    $kRaw   = [string]$xlRow.'Kunde'
    $kName  = Get-KonsolidierterKunde $kRaw

    $gesamt += $wert
    if ($auftrNr -gt $maxAuftrag) { $maxAuftrag = $auftrNr }

    $gr = Get-Gruppe $artnr
    $gruppen[$gr] += $wert

    if (Test-IsAnkernagel $artnr) { $davonAnker += $wert }

    if ($kundenMap.ContainsKey($kName)) { $kundenMap[$kName] += $wert }
    else { $kundenMap[$kName] = $wert }

    # Kunden-Detail: Artikel aggregieren (je konsolidiertem Kunden)
    if (-not $kundenDetail.ContainsKey($kName)) { $kundenDetail[$kName] = @{} }
    if ($kundenDetail[$kName].ContainsKey($artnr)) {
        $kundenDetail[$kName][$artnr].Wert += $wert
    } else {
        $kundenDetail[$kName][$artnr] = @{ Wert = $wert; Text = $ktext }
    }

    # Gruppen-Detail: Top-Artikel + Top-Kunden (roh, nicht konsolidiert)
    if (-not $gruppenDetail.ContainsKey($gr)) {
        $gruppenDetail[$gr] = @{ Artikel = @{}; Kunden = @{} }
    }
    if ($gruppenDetail[$gr].Artikel.ContainsKey($artnr)) {
        $gruppenDetail[$gr].Artikel[$artnr].Wert += $wert
    } else {
        $gruppenDetail[$gr].Artikel[$artnr] = @{ Wert = $wert; Text = $ktext }
    }
    if ($gruppenDetail[$gr].Kunden.ContainsKey($kRaw)) {
        $gruppenDetail[$gr].Kunden[$kRaw] += $wert
    } else {
        $gruppenDetail[$gr].Kunden[$kRaw] = $wert
    }
}

# Kunden-Detail: Hashtable -> sortiertes Array je Kunde
$kundenDetailOut = @{}
foreach ($kd in $kundenDetail.GetEnumerator()) {
    $kundenDetailOut[$kd.Key] = @(
        $kd.Value.GetEnumerator() |
        Sort-Object { [double]$_.Value.Wert } -Descending |
        ForEach-Object {
            [PSCustomObject]@{
                artnr = [string]$_.Key
                text  = [string]$_.Value.Text
                wert  = [Math]::Round([double]$_.Value.Wert, 2)
            }
        }
    )
}

# Gruppen-Detail: Top-5 Artikel + Top-5 Kunden je Gruppe
$gruppenDetailOut = @{}
foreach ($gd in $gruppenDetail.GetEnumerator()) {
    $topArtikel = @(
        $gd.Value.Artikel.GetEnumerator() |
        Sort-Object { [double]$_.Value.Wert } -Descending |
        Select-Object -First 5 |
        ForEach-Object {
            [PSCustomObject]@{
                artnr = [string]$_.Key
                text  = [string]$_.Value.Text
                wert  = [Math]::Round([double]$_.Value.Wert, 2)
            }
        }
    )
    $topKunden = @(
        $gd.Value.Kunden.GetEnumerator() |
        Sort-Object { [double]$_.Value } -Descending |
        Select-Object -First 5 |
        ForEach-Object {
            [PSCustomObject]@{
                kunde = [string]$_.Key
                wert  = [Math]::Round([double]$_.Value, 2)
            }
        }
    )
    $gruppenDetailOut[$gd.Key] = [PSCustomObject]@{
        top_artikel = $topArtikel
        top_kunden  = $topKunden
    }
}

# Alle Kunden (vollstaendige Liste, absteigend sortiert)
$alleKunden = $kundenMap.GetEnumerator() |
    Sort-Object Value -Descending |
    ForEach-Object { [PSCustomObject]@{ Kunde = $_.Key; Wert = [Math]::Round($_.Value, 2) } }

# Kunden-Summe fuer Plausi
$kundenGesamt = 0.0
foreach ($v in $kundenMap.Values) { $kundenGesamt += $v }
$kundenCheckOk = [Math]::Abs($gesamt - $kundenGesamt) -lt 0.02

# Plausibilitaets-Check (Gruppen-Summe)
$gruppenGesamt = 0.0
foreach ($v in $gruppen.Values) { $gruppenGesamt += $v }
$checkDiff = [Math]::Abs($gesamt - $gruppenGesamt)
$checkOk   = $checkDiff -lt 0.02

Write-Host "  Gesamtbestand:  $('{0:N2}' -f $gesamt) EUR"
Write-Host "  Max. Auftragsnr: $maxAuftrag"
Write-Host "  Check OK: $checkOk (Differenz: $('{0:N4}' -f $checkDiff))"

# --- Log einlesen fuer Tageseingang ----------------------------
$heute         = (Get-Date).Date
$tageseingang  = 0.0
$gesternMaxNr  = 0
$logDaten      = @()

if (Test-Path $LOG_XLS) {
    try {
        $logDaten = @(Import-Excel -Path $LOG_XLS -WorksheetName "Log" -ErrorAction Stop)
        # Letzten Werktag-Eintrag suchen (nicht heute)
        $letzterEintrag = $logDaten |
            Where-Object { $_.Datum -ne $null } |
            Sort-Object Datum -Descending |
            Where-Object { ([datetime]$_.Datum).Date -lt $heute } |
            Select-Object -First 1

        if ($letzterEintrag) {
            $gesternMaxNr = [int]$letzterEintrag.Max_Auftragsnr
            Write-Host "  Vortages-Max-Auftragsnr: $gesternMaxNr"
        }
    } catch {
        Write-Warning "Log-Datei konnte nicht gelesen werden: $_"
    }
}

# Tageseingang = Summe aller Auftraege mit Nr > gesternMaxNr
if ($gesternMaxNr -gt 0) {
    foreach ($xlRow in $xlRows) {
        $auftrNr = 0
        $wert    = 0.0
        try { $auftrNr = [int]$xlRow.'Auftrag'  } catch {}
        try { $wert    = [double]$xlRow.'Wert'  } catch {}
        if ($auftrNr -gt $gesternMaxNr) { $tageseingang += $wert }
    }
    Write-Host "  Tageseingang: $('{0:N2}' -f $tageseingang) EUR"
} else {
    Write-Host "  Tageseingang: kein Vortages-Log -- erster Lauf, Tageseingang = 0" -ForegroundColor Yellow
}

# --- Wochen- und Monatseingang aus Log -------------------------
$wocheneingang = 0.0
$monatseingang = 0.0

$montag = $heute.AddDays(-(([int]$heute.DayOfWeek + 6) % 7))
$ersterDesMonats = [datetime]::new($heute.Year, $heute.Month, 1)

# Letzte Woche (Mo-Fr der Vorwoche)
$letzteWocheMo = $montag.AddDays(-7)
$letzteWocheFr = $montag.AddDays(-3)
$letzteWocheEingang = 0.0

foreach ($zeile in $logDaten) {
    if ($zeile.Datum -eq $null) { continue }
    try {
        $zDat = ([datetime]$zeile.Datum).Date
        $zEin = 0.0
        try { $zEin = [double]$zeile.Tageseingang_EUR } catch {}

        # Lfd. Monat
        if ($zDat -ge $ersterDesMonats -and $zDat -lt $heute) {
            $monatseingang += $zEin
        }
        # Lfd. Woche (Mo bis gestern)
        if ($zDat -ge $montag -and $zDat -lt $heute) {
            $wocheneingang += $zEin
        }
        # Letzte Woche
        if ($zDat -ge $letzteWocheMo -and $zDat -le $letzteWocheFr) {
            $letzteWocheEingang += $zEin
        }
    } catch {}
}

# Heutigen Tageseingang dazurechnen
$monatseingang += $tageseingang
$wocheneingang += $tageseingang

Write-Host "  Wocheneingang (lfd.):  $('{0:N2}' -f $wocheneingang) EUR"
Write-Host "  Monatseingang (lfd.):  $('{0:N2}' -f $monatseingang) EUR"
Write-Host "  Letzte Woche gesamt:   $('{0:N2}' -f $letzteWocheEingang) EUR"

# --- Log-Zeile schreiben ---------------------------------------
$logZeile = [PSCustomObject]@{
    Datum                 = $heute.ToString("yyyy-MM-dd")
    Gesamtbestand_EUR     = [Math]::Round($gesamt, 2)
    Max_Auftragsnr        = $maxAuftrag
    Tageseingang_EUR      = [Math]::Round($tageseingang, 2)
    Maschinenstifte_EUR   = [Math]::Round($gruppen["Maschinenstifte"], 2)
    Kernschrauben_EUR     = [Math]::Round($gruppen["Kernschrauben"], 2)
    Lose_Edelstahl_EUR    = [Math]::Round($gruppen["Lose_Edelstahl"], 2)
    Industrieklammern_EUR = [Math]::Round($gruppen["Industrieklammern"], 2)
    Drahtcoil_EUR         = [Math]::Round($gruppen["Drahtcoil"], 2)
    Streifennaegel_KS_EUR = [Math]::Round($gruppen["Streifennaegel_KS"], 2)
    Papertape_EUR         = [Math]::Round($gruppen["Papertape"], 2)
    CDF_UTC_EUR           = [Math]::Round($gruppen["CDF_UTC"], 2)
    Ankernaegel_Lose_EUR  = [Math]::Round($gruppen["Ankernaegel_Lose"], 2)
    Umpack_Dachpapp_EUR   = [Math]::Round($gruppen["Umpack_Dachpapp"], 2)
    Sonstiges_EUR         = [Math]::Round($gruppen["Sonstiges"], 2)