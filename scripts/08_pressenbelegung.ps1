# ============================================================
# 08_pressenbelegung.ps1
# Extrakt Pressenbelegung aus Planung_Pressen_NEU.xlsx
#
# Gruppen:
#   DKopf_Offset  -> N51-2, N90-D2, N90-D3, N90-D4, NQ03-013
#   N31           -> N31-1, N31-2, N31-3
#   N41           -> N41-1, N41-2, N41-3, N41-5, N41-6, N41-7
#   N51           -> N51-1, N51-3
#   N61           -> N61-1, N61-2, N61-3
#   N90           -> N90-2,5  N90-3,1  N90-D1
#   Enkotec       -> Enkotec 2,5 NA03  Enkotec 2,8 NI01  Enkotec 2,8 NU03
#   Doppeldruck   -> Hilgeland COH/COLH/HD6-40/HD6-60/HD7/HC5-60/C2AZ
#                    FWB 20C-1/20C-2  ChunZu CH12LL  Klose MTH350
#
# Auslastung Tage:
#   N-Maschinen : Stunden / 16  (2 Schichten x 8h)
#   DD-Maschinen: Stunden /  8  (1 Schicht  x 8h)
# ============================================================

$SCRIPTS = [string]$PSScriptRoot
$BASE    = [string](Split-Path -Parent $SCRIPTS)
$SRC_DIR = [string](Join-Path $BASE "_source")
$OUT_DIR = [string](Join-Path $BASE "_data")
$JSON_OUT= [string](Join-Path $OUT_DIR "pressenbelegung.json")
$SRC_XLS = [string](Join-Path $SRC_DIR "Planung_Pressen_NEU.xlsx")

if (-not (Test-Path $OUT_DIR)) { New-Item -ItemType Directory -Path $OUT_DIR | Out-Null }

Write-Host "=== 08_pressenbelegung.ps1 ===" -ForegroundColor Cyan

# Python bevorzugen (08_pressenbelegung.py enthaelt vollstaendige Logik inkl.
# Extrapolation unverteilter FA, Gewichte-Lookup, NQ-Sonderregel)
$pyScript = Join-Path $SCRIPTS "08_pressenbelegung.py"
$py = $null
foreach ($cmd in @("python", "python3", "py")) {
    try {
        $v = & $cmd --version 2>&1
        if ($v -match "Python") { $py = $cmd; break }
    } catch {}
}
if ($py -and (Test-Path $pyScript)) {
    Write-Host "  Python gefunden ($py) -- fuehre 08_pressenbelegung.py aus" -ForegroundColor DarkGray
    & $py $pyScript 2>&1 | ForEach-Object { Write-Host "  $_" }
    if ($LASTEXITCODE -eq 0) {
        Write-Host "[OK] Pressenbelegung via Python extrahiert" -ForegroundColor Green
        exit 0
    }
    Write-Host "[WARN] Python-Skript fehlgeschlagen (ExitCode=$LASTEXITCODE) -- Fallback auf PS" -ForegroundColor Yellow
}

# --- Fallback: einfache PS-Extraktion (ohne Extrapolation) ---
Write-Host "  Fallback: PS-Extraktion (ohne Prognose-Daten)" -ForegroundColor Yellow

try {
    Import-Module ImportExcel -ErrorAction Stop
} catch {
    Write-Host "[ERR] ImportExcel nicht installiert." -ForegroundColor Red; exit 1
}

if (-not (Test-Path $SRC_XLS)) {
    Write-Host "[ERR] Quelldatei nicht gefunden: $SRC_XLS" -ForegroundColor Red; exit 1
}

# ---- Gruppen-Mapping ------------------------------------------
# Key: Maschinenname (exakt wie in Spalte A des Tabs)
# Value: Gruppenkey
$MASCHINE_GRUPPE = @{
    # D-Kopf + Offset
    "N51-2"            = "DKopf_Offset"
    "N90-D2"           = "DKopf_Offset"
    "N90-D3"           = "DKopf_Offset"
    "N90-D4"           = "DKopf_Offset"
    "NQ03-013"         = "DKopf_Offset"
    # N31
    "N31-1"            = "N31"
    "N31-2"            = "N31"
    "N31-3"            = "N31"
    # N41
    "N41-7"            = "N41"
    "N41-6"            = "N41"
    "N41-2"            = "N41"
    "N41-5"            = "N41"
    "N41-1"            = "N41"
    "N41-3"            = "N41"
    # N51 (ohne N51-2)
    "N51-3"            = "N51"
    "N51-1"            = "N51"
    # N61
    "N61-3"            = "N61"
    "N61-2"            = "N61"
    "N61-1"            = "N61"
    # N90 (ohne D2/D3/D4)
    "N90-2,5"          = "N90"
    "N90-3,1"          = "N90"
    "N90-D1"           = "N90"
    # Enkotec
    "Enkotec 2,5 NA03" = "Enkotec"
    "Enkotec 2,8 NI01" = "Enkotec"
    "Enkotec 2,8 NU03" = "Enkotec"
    # Doppeldruck
    "Hilgeland COH"    = "Doppeldruck"
    "Hilgeland COLH"   = "Doppeldruck"
    "FWB 20C-1 wei`u00df"  = "Doppeldruck"   # wird auch normalisiert abgeglichen
    "FWB 20C-2 gr`u00fcn"  = "Doppeldruck"
    "Hilgeland HD7"    = "Doppeldruck"
    "Hilgeland HD6-40" = "Doppeldruck"
    "Hilgeland HD6-60" = "Doppeldruck"
    "Hilgeland HC5-60" = "Doppeldruck"
    "Hilgeland C2AZ"   = "Doppeldruck"
    "ChunZu CH12LL"    = "Doppeldruck"
    "Klose MTH350"     = "Doppeldruck"
}

# Unicode-Sonderzeichen fuer FWB-Namen
$MASCHINE_GRUPPE["FWB 20C-1 wei$([char]0x00DF)"] = "Doppeldruck"
$MASCHINE_GRUPPE["FWB 20C-2 gr$([char]0x00FC)n"] = "Doppeldruck"

# Schichten pro Tag je Gruppe
$SCHICHTEN = @{
    "DKopf_Offset" = 2
    "N31"          = 2
    "N41"          = 2
    "N51"          = 2
    "N61"          = 2
    "N90"          = 2
    "Enkotec"      = 2
    "Doppeldruck"  = 1
}

# Gruppen-Definitionen (Reihenfolge + Label)
$GRUPPEN_DEF = [ordered]@{
    "DKopf_Offset" = "D-Kopf + Offset"
    "N31"          = "N31"
    "N41"          = "N41"
    "N51"          = "N51"
    "N61"          = "N61"
    "N90"          = "N90"
    "Enkotec"      = "Enkotec"
    "Doppeldruck"  = "Doppeldruck"
}

# ---- Tabs die geparst werden ----------------------------------
$TABS = @("Doppeldruck", "N31+41", "N51+61", "N90 + ENK", "NQ03-013")

# ---- Hilfsfunktion: Zustand normalisieren ---------------------
function Get-Zustand([string]$raw) {
    $r = $raw.Trim().ToLower()
    if ($r -match 'freigegeb') { return "freigegeben" }
    if ($r -match 'terminiert') { return "terminiert" }
    if ($r -match 'unterbrochen') { return "unterbrochen" }
    return "unbekannt"
}

# ---- Maschinenname normalisieren (Leerzeichen/nbsp trimmen) --
function Normalize-Name([string]$n) {
    return ($n -replace '\s+', ' ').Trim().Trim([char]0x00A0)
}

# ---- Alle Maschinen sammeln ----------------------------------
$allMaschinen = [ordered]@{}   # name -> @{gruppe, permille, kg, stunden, auftraege[]}

foreach ($tab in $TABS) {
    Write-Host "  Lese Tab: $tab" -ForegroundColor Yellow
    $ws = Import-Excel -Path $SRC_XLS -WorksheetName $tab -NoHeader -ErrorAction Stop
    $aktMaschine = $null

    foreach ($row in $ws) {
        $cols = @($row.PSObject.Properties.Value)
        # Sicherstellen dass genug Spalten vorhanden
        while ($cols.Count -lt 12) { $cols += $null }

        $c0 = if ($cols[0] -ne $null) { [string]$cols[0] } else { "" }
        $c1 = $cols[1]
        $c4 = $cols[4]
        $c5 = if ($cols[5] -ne $null) { [string]$cols[5] } else { "" }
        $c6 = $cols[6]

        # --- Maschinenheader erkennen ---
        # Standard-Tabs: col[5] == "‰"
        # NQ03-013:      col[1] == "NQ03-013" (col[0] ist null)
        $isMachineHeader = $false

        if ($c5 -eq [char]0x2030 -and $c4 -ne $null -and $c4 -is [double]) {
            $isMachineHeader = $true
            $maschinenRaw = Normalize-Name $c0
        } elseif ($c1 -ne $null -and [string]$c1 -eq "NQ03-013") {
            $isMachineHeader = $true
            $maschinenRaw = "NQ03-013"
        }

        if ($isMachineHeader) {
            $aktMaschine = $maschinenRaw
            if (-not $allMaschinen.Contains($aktMaschine)) {
                $allMaschinen[$aktMaschine] = @{
                    name      = $aktMaschine
                    sum_pm    = 0
                    sum_kg    = 0.0
                    sum_h     = 0.0
                    auftraege = [System.Collections.Generic.List[hashtable]]::new()
                }
            }
            continue
        }

        # --- Spaltenheader-Zeile ueberspringen ---
        if ($c0 -eq "S" -and [string]$c1 -eq "FA") { continue }

        # --- Auftragszeile ---
        if ($aktMaschine -and $c1 -ne $null -and $c1 -is [double] -and [double]$c1 -gt 0) {
            # Zustand-Filter: nur echte Status
            $zustandRaw = $c0
            if ($zustandRaw -match 'nicht vorhanden|#N/A|#WERT' -or
                [string]$zustandRaw -eq "") { continue }

            $artnr   = if ($cols[2] -ne $null) { [string]$cols[2] } else { "" }
            if ($artnr -match 'nicht vorhanden|#N/A') { continue }

            $kurztext = if ($cols[3] -ne $null) { [string]$cols[3] } else { "" }
            $mengeVal = $cols[4]
            $istVal   = $cols[5]
            $kgVal    = $cols[6]
            $terminVal= $cols[7]

            # Spaltenstruktur je Tab
            if ($tab -eq "Doppeldruck") {
                $folgeVal  = $cols[8]    # INFO
                $bemerkVal = $null
                $stdVal    = $cols[9]
            } elseif ($aktMaschine -eq "NQ03-013") {
                $folgeVal  = $cols[8]
                $bemerkVal = if ($cols.Count -gt 10) { $cols[10] } else { $null }
                $stdVal    = if ($cols.Count -gt 11) { $cols[11] } else { $null }
            } else {
                $folgeVal  = $cols[8]
                $bemerkVal = $cols[9]
                $stdVal    = if ($cols.Count -gt 10) { $cols[10] } else { $null }
            }

            $menge   = if ($mengeVal -ne $null -and $mengeVal -is [double]) { [Math]::Round([double]$mengeVal, 0) } else { $null }
            $ist     = if ($istVal -ne $null -and $istVal -is [double]) { [Math]::Round([double]$istVal, 0) } else { $null }
            $kg      = if ($kgVal -ne $null -and $kgVal -is [double]) { [Math]::Round([double]$kgVal, 1) } else { $null }
            $stunden = if ($stdVal -ne $null -and $stdVal -is [double]) { [Math]::Round([double]$stdVal, 1) } else { $null }

            # kg-Plausibilitaetsfilter: max 500 kg/‰ (Naegel)
            if ($kg -ne $null -and $menge -ne $null -and $menge -gt 0) {
                if (($kg / $menge) -gt 500) {
                    Write-Host ("  [FILTER] {0} FA {1} {2}: {3:N0}kg/{4}‰={5:N0}" -f `
                        $aktMaschine, [int][double]$c1, $artnr, $kg, $menge, [int]($kg/$menge)) -ForegroundColor Yellow
                    $kg = $null
                }
            }

            $termin  = if ($terminVal -ne $null -and $terminVal -is [datetime]) {
                           $terminVal.ToString("yyyy-MM-dd") } else { $null }
            $folge   = if ($folgeVal -ne $null) { ([string]$folgeVal).Trim() } else { "" }
            $bemerk  = if ($bemerkVal -ne $null) { ([string]$bemerkVal).Trim() } else { "" }

            $allMaschinen[$aktMaschine].auftraege.Add(@{
                zustand  = Get-Zustand $zustandRaw
                fa       = [int][double]$c1
                artnr    = $artnr.Trim()
                kurztext = $kurztext.Trim()
                menge    = $menge
                ist      = $ist
                kg       = $kg
                termin   = $termin
                folge_ag = $folge
                bemerkung= $bemerk
                stunden  = $stunden
            })
            $allMaschinen[$aktMaschine].sum_pm += if ($menge)   { $menge }   else { 0 }
            $allMaschinen[$aktMaschine].sum_kg += if ($kg)      { $kg }      else { 0.0 }
            $allMaschinen[$aktMaschine].sum_h  += if ($stunden) { $stunden } else { 0.0 }
        }
    }
}

Write-Host "  Maschinen gefunden: $($allMaschinen.Count)" -ForegroundColor DarkGray

# ---- In Gruppen einteilen ------------------------------------
$gruppenOut = [ordered]@{}
foreach ($gKey in $GRUPPEN_DEF.Keys) {
    $gruppenOut[$gKey] = @{
        key      = $gKey
        name     = $GRUPPEN_DEF[$gKey]
        permille = 0.0
        kg       = 0.0
        stunden  = 0.0
        maschinen= [System.Collections.Generic.List[object]]::new()
    }
}

foreach ($mName in $allMaschinen.Keys) {
    $m = $allMaschinen[$mName]
    # Gruppe bestimmen (normalisierter Namensabgleich)
    $gKey = $null
    foreach ($k in $MASCHINE_GRUPPE.Keys) {
        if ((Normalize-Name $k) -eq (Normalize-Name $mName)) {
            $gKey = $MASCHINE_GRUPPE[$k]
            break
        }
    }
    if ($null -eq $gKey) {
        # Fuzzy-Fallback: starts-with Abgleich
        $nNorm = (Normalize-Name $mName).ToLower()
        if ($nNorm -match '^n31') { $gKey = "N31" }
        elseif ($nNorm -match '^n41') { $gKey = "N41" }
        elseif ($nNorm -match '^n51') { $gKey = "N51" }
        elseif ($nNorm -match '^n61') { $gKey = "N61" }
        elseif ($nNorm -match '^n90') { $gKey = "N90" }
        elseif ($nNorm -match '^enkotec') { $gKey = "Enkotec" }
        elseif ($nNorm -match '^hilgeland|^fwb|^chunzu|^klose') { $gKey = "Doppeldruck" }
        else { Write-Host "  [WARN] Keine Gruppe fuer: $mName" -ForegroundColor Yellow; continue }
    }

    $schichten    = $SCHICHTEN[$gKey]
    $auslastTage  = if ($m.sum_h -gt 0 -and $schichten -gt 0) {
                        [Math]::Round($m.sum_h / ($schichten * 8.0), 2)
                    } else { 0.0 }

    $maschinenObj = [ordered]@{
        name           = $m.name
        permille       = [Math]::Round($m.sum_pm, 0)
        kg             = [Math]::Round($m.sum_kg, 1)
        stunden        = [Math]::Round($m.sum_h, 1)
        auslastung_tage= $auslastTage
        schichten      = $schichten
        auftraege      = @($m.auftraege)
    }

    $g = $gruppenOut[$gKey]
    $g.maschinen.Add($maschinenObj)
    $g.permille += $m.sum_pm
    $g.kg       += $m.sum_kg
    $g.stunden  += $m.sum_h
}

# ---- KPI gesamt ----------------------------------------------
$kpi_permille = 0; $kpi_kg = 0.0; $kpi_stunden = 0.0
foreach ($g in $gruppenOut.Values) {
    $kpi_permille += $g.permille
    $kpi_kg       += $g.kg
    $kpi_stunden  += $g.stunden
    # Runden fuer JSON
    $g.permille = [Math]::Round($g.permille, 0)
    $g.kg       = [Math]::Round($g.kg, 1)
    $g.stunden  = [Math]::Round($g.stunden, 1)
}

# ---- Ausgabe -------------------------------------------------
$result = [ordered]@{
    kpi = [ordered]@{
        gesamt_permille = [Math]::Round($kpi_permille, 0)
        gesamt_kg       = [Math]::Round($kpi_kg, 1)
        gesamt_stunden  = [Math]::Round($kpi_stunden, 1)
    }
    gruppen   = @($gruppenOut.Values)
    timestamp = (Get-Date -Format "o")
}

$json = $result | ConvertTo-Json -Depth 8 -Compress:$false
[System.IO.File]::WriteAllText($JSON_OUT, $json, [System.Text.Encoding]::UTF8)

Write-Host "[OK] pressenbelegung.json geschrieben" -ForegroundColor Green
foreach ($g in $gruppenOut.Values) {
    Write-Host ("  {0,-18} {1,3} Maschinen  {2,8:N0} ‰  {3,8:N1} h" -f `
        $g.name, $g.maschinen.Count, $g.permille, $g.stunden) -ForegroundColor DarkGray
}
exit 0
