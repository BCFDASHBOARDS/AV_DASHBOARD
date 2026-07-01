# ============================================================
# 02_refresh_ubersicht.ps1
# Nur Lagerbestand, Restmenge, Verbrauch 2026 refreshen
# ============================================================

$SCRIPTS  = $PSScriptRoot
$BASE     = Split-Path -Parent $SCRIPTS
$SOURCE   = Join-Path $BASE "_source"
$NET_FILE = "\\srv12020\Allgemein\AV\Draht\$([char]0xFC)bersicht.xlsx"
$DST_FILE = Join-Path $SOURCE "ubersicht_Draht.xlsx"

$REFRESH_FILTER = @("Lagerbestand", "Restmenge", "Verbrauch 2026")

Write-Host "Starte Excel-COM fuer uebersicht.xlsx ..."

$excel = $null
$wb    = $null

try {
    $excel = New-Object -ComObject Excel.Application
    $excel.Visible          = $false
    $excel.DisplayAlerts    = $false
    $excel.AskToUpdateLinks = $false

    $wb = $excel.Workbooks.Open($NET_FILE, 0, $false)
    Write-Host "Datei geoeffnet."

    foreach ($conn in $wb.Connections) {
        $match = $false
        foreach ($f in $REFRESH_FILTER) {
            if ($conn.Name -like "*$f*") { $match = $true; break }
        }
        if (-not $match) {
            Write-Host "  [SKIP] $($conn.Name)"
            continue
        }
        try {
            # BackgroundQuery abschalten wenn moeglich (OLEDBConnection)
            try { $conn.OLEDBConnection.BackgroundQuery = $false } catch {}
            $conn.Refresh()
            Write-Host "  [OK]   $($conn.Name)" -ForegroundColor Green
        } catch {
            Write-Warning "  [WARN] $($conn.Name): $_"
        }
    }

    # Warten bis Excel fertig (CalculateUntilAsyncQueriesDone blockiert)
    Write-Host "Warte auf Query-Abschluss ..."
    try {
        $excel.CalculateUntilAsyncQueriesDone()
    } catch {
        # Nicht alle Excel-Versionen unterstuetzen diese Methode -- Fallback
        $elapsed = 0
        while (-not $excel.Ready -and $elapsed -lt 90) {
            Start-Sleep -Seconds 3
            $elapsed += 3
        }
    }

    # Speichern mit Retry
    $saved = $false
    for ($i = 1; $i -le 6; $i++) {
        try {
            $wb.Save()
            $saved = $true
            Write-Host "Gespeichert." -ForegroundColor Green
            break
        } catch {
            Write-Host "  Speichern Versuch $i fehlgeschlagen, warte 5 Sek. ..."
            Start-Sleep -Seconds 5
        }
    }
    if (-not $saved) { throw "Speichern fehlgeschlagen nach 6 Versuchen" }

    $wb.Close($false)
    $wb = $null

    Copy-Item -Path $NET_FILE -Destination $DST_FILE -Force
    Write-Host "[OK] ubersicht_Draht.xlsx -> $DST_FILE" -ForegroundColor Green
    exit 0

} catch {
    Write-Host "[ERR] $($_.Exception.Message)" -ForegroundColor Red
    exit 1

} finally {
    if ($wb)    { try { $wb.Close($false)    } catch {} }
    if ($excel) { try { $excel.Quit()        } catch {} }
    if ($excel) {
        [System.Runtime.InteropServices.Marshal]::ReleaseComObject($excel) | Out-Null
    }
    [GC]::Collect()
    [GC]::WaitForPendingFinalizers()
}
