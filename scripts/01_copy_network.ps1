# ============================================================
# 01_copy_network.ps1
# Kopiert Quelldateien vom Netzlaufwerk nach _source\
# ============================================================

$BASE   = Split-Path -Parent (Split-Path -Parent $MyInvocation.MyCommand.Path)
$SOURCE = Join-Path $BASE "_source"
$NET    = "\\srv12020\Allgemein\AV\Produktionsplanung und BDE\Quelldateien\Erforderliche Datenimports"

$FILES = @(
    "1100_Datenimport_Pressen.xlsx",
    "1180_Datenimport_Walzen.xlsx",
    "Auftragsbestand.xlsx",
    "Istmengen.xlsx",
    "Lagerbestand.xlsx"
)

$errors = 0

foreach ($file in $FILES) {
    $src = Join-Path $NET $file
    $dst = Join-Path $SOURCE $file
    try {
        Copy-Item -Path $src -Destination $dst -Force -ErrorAction Stop
        Write-Host "[OK]  $file" -ForegroundColor Green
    } catch {
        Write-Host "[ERR] $file : $_" -ForegroundColor Red
        $errors++
    }
}

if ($errors -gt 0) {
    Write-Warning "01_copy_network: $errors Datei(en) konnten nicht kopiert werden."
    exit 1
} else {
    Write-Host "01_copy_network: Alle Dateien erfolgreich kopiert." -ForegroundColor Cyan
    exit 0
}
