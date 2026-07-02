# ============================================================
# 07_magazinierung.ps1
# Extrakt MAG-Produktionsmengen aus Istmengen.xlsx
# Gruppen: CDF_UTC (50-/51-), Papertape (45-/46-/47-),
#          Drahtcoil (30-), Streifennaegel_KS (40-/41-/42-),
#          Ankernaegel (40-40/40-60/42-40/50-40/47-400, ohne 62-),
#          KRL (60-), BSN (55-)
# Menge netto = Spalte C minus Spalte D
# ============================================================

$SCRIPTS = [string]$PSScriptRoot
$BASE    = [string](Split-Path -Parent $SCRIPTS)
$SRC_DIR = [string](Join-Path $BASE "_source")
$OUT_DIR = [string](Join-Path $BASE "_data")
$SRC_XLS = [string](Join-Path $SRC_DIR "Istmengen.xlsx")
$JSON_OUT = [string](Join-Path $OUT_DIR "magazinierung.json")

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
    $p5 = if ($artnr.Length -ge 5) { $artnr.Substring(0,5) } else { $artnr }
    $p6 = if ($artnr.Length -ge 6) { $artnr.Substring(0,6) } else { $artnr }

    # Ankernagel-Sondergruppe zuerst pruefen (vor regulaerer Gruppen-Zuweisung)
    if ($p5 -in @("40-40","40-60","42-40","50-40")) { return "Ankernaegel" }
    if ($p6 -eq "47-400") { return "Ankernaegel" }

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
                 Streifennaegel_KS=0.0; Ankernaegel=0.0; KRL=0.0; BSN=0.0 }
$artMengen  = @{}   # artnr -> netto menge (konsolidiert)
$artGruppe  = @{}   # artnr -> gruppe

foreach ($row in $xlRows) {
    $artnr = [string]($row.Ressource)
    $c     = [double]($row.Menge_C -replace ',','.' )
    $d     = [double]($row.Menge_D -replace ',','.' )
    $netto = $c - $d
    if ($netto -le 0) { continue }

    $g = Get-MAGGruppe $artnr
    if ($null -eq $g) { continue }

    if (-not $artMengen.ContainsKey($artnr)) {
        $artMengen[$artnr] = 0.0
        $artGruppe[$artnr] = $g
    }
    $artMengen[$artnr] += $netto
    $gruppen[$g] += $netto
}

$gesamt = ($gruppen.Values | Measure-Object -Sum).Sum

# ---- Gruppen-Detail: Top-5 Artikel je Gruppe -----------------
$gruppenDetail = @{}
foreach ($g in $gruppen.Keys) {
    $topArts = $artMengen.Keys |
        Where-Object { $artGruppe[$_] -eq $g } |
        Sort-Object { -$artMengen[$_] } |
        Select-Object -First 5
    $gruppenDetail[$g] = @{
        top_artikel = @($topArts | ForEach-Object {
            @{ artnr = $_; menge = [Math]::Round($artMengen[$_], 0) }
        })
    }
}

# ---- Ausgabe -------------------------------------------------
$result = [ordered]@{
    gesamt         = [Math]::Round($gesamt, 0)
    gruppen        = [ordered]@{
        CDF_UTC           = [Math]::Round($gruppen["CDF_UTC"], 0)
        Papertape         = [Math]::Round($gruppen["Papertape"], 0)
        Drahtcoil         = [Math]::Round($gruppen["Drahtcoil"], 0)
        Streifennaegel_KS = [Math]::Round($gruppen["Streifennaegel_KS"], 0)
        Ankernaegel       = [Math]::Round($gruppen["Ankernaegel"], 0)
        KRL               = [Math]::Round($gruppen["KRL"], 0)
        BSN               = [Math]::Round($gruppen["BSN"], 0)
    }
    gruppen_detail = $gruppenDetail
    timestamp      = (Get-Date -Format "o")
}

$json = $result | ConvertTo-Json -Depth 6 -Compress:$false
[System.IO.File]::WriteAllText($JSON_OUT, $json, [System.Text.Encoding]::UTF8)

Write-Host "[OK] magazinierung.json geschrieben ($([Math]::Round($gesamt,0).ToString('N0')) Stk. gesamt)" -ForegroundColor Green

# Log
$gruppen.GetEnumerator() | Sort-Object Name | ForEach-Object {
    $pct = if ($gesamt -gt 0) { [Math]::Round($_.Value/$gesamt*100,1) } else { 0 }
    Write-Host ("  {0,-22} {1,10:N0} Stk.  ({2,5:N1}%)" -f $_.Key, $_.Value, $pct) -ForegroundColor DarkGray
}
exit 0
