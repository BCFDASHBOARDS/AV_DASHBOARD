# ============================================================
# 02_refresh_ubersicht.ps1
# Öffnet übersicht.xlsx via Excel COM, aktualisiert alle
# Power Query-Verbindungen, speichert und kopiert nach _source\
# ============================================================

$BASE   = Split-Path -Parent (Split-Path -Parent $MyInvocation.MyCommand.Path)
$SOURCE = Join-Path $BASE "_source"
$NET_FILE = "\\srv12020\Allgemein\AV\Draht\übersicht.xlsx"
$DST_FILE = Join-Path $SOURCE "ubersicht_Draht.xlsx"

Write-Host "Starte Excel-COM für übersicht.xlsx ..."

try {
    $excel = New-Object -ComObject Excel.Application
    $excel.Visible        = $false
    $excel.DisplayAlerts  = $false
    $excel.AskToUpdateLinks = $false

    # Datei direkt vom Netzwerkpfad öffnen
    $wb = $excel.Workbooks.Open($NET_FILE, 0, $false)

    Write-Host "Datei geöffnet. Aktualisiere alle Verbindungen ..."

    # Alle Verbindungen (Power Query / externe Datenquellen) refreshen
    foreach ($conn in $wb.Connections) {
        try {
            $conn.Refresh()
            Write-Host "  [OK] Verbindung: $($conn.Name)"
        } catch {
            Write-Warning "  [WARN] Verbindung '$($conn.Name)' konnte nicht aktualisiert werden: $_"
        }
    }

    # Warte kurz damit async Queries fertig werden (max. 120 Sek.)
    $timeout = 120
    $elapsed = 0
    while ($wb.Queries.Count -gt 0 -and $elapsed -lt $timeout) {
        $running = $false
        foreach ($q in $wb.Queries) {
            # BackgroundQuery-Status prüfen (wenn verfügbar)
        }
        Start-Sleep -Seconds 2
        $elapsed += 2
    }

    # Speichern und schließen
    $wb.Save()
    $wb.Close($false)
    Write-Host "Datei gespeichert und geschlossen."

    # Lokale Kopie nach _source
    Copy-Item -Path $NET_FILE -Destination $DST_FILE -Force
    Write-Host "[OK] übersicht.xlsx → $DST_FILE" -ForegroundColor Green

    exit 0

} catch {
    Write-Host "[ERR] Fehler bei übersicht.xlsx: $_" -ForegroundColor Red
    exit 1

} finally {
    if ($wb)    { try { $wb.Close($false)    } catch {} }
    if ($excel) { try { $excel.Quit()        } catch {} }
    if ($excel) { [System.Runtime.InteropServices.Marshal]::ReleaseComObject($excel) | Out-Null }
    [GC]::Collect()
}
