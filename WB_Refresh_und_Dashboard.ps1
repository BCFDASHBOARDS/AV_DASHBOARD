# ============================================================
#  WB_Refresh_und_Dashboard.ps1
#  1. Excel öffnen, Power-Query-Abfragen aktualisieren, speichern
#  2. Python-Skript aufrufen → dashboard.html neu generieren
# ============================================================

$ErrorActionPreference = "Stop"
$scriptDir = Split-Path -Parent $MyInvocation.MyCommand.Path
$xlsxPath  = "C:\Users\wibau\OneDrive - Baussmann\2026\WB_Übersicht.xlsx"
$pyScript  = Join-Path $scriptDir "generate_dashboard.py"
$logFile   = Join-Path $scriptDir "refresh.log"

function Log($msg) {
    $ts = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    "$ts  $msg" | Tee-Object -FilePath $logFile -Append
}

Log "=== Start ==="

# ---- Schritt 1: Excel-Refresh ----
try {
    Log "Excel öffnen: $xlsxPath"
    $excel = New-Object -ComObject Excel.Application
    $excel.Visible        = $false
    $excel.DisplayAlerts  = $false

    $wb = $excel.Workbooks.Open($xlsxPath)
    Log "Workbook geöffnet. Abfragen aktualisieren …"

    $excel.CalculateUntilAsyncQueriesDone()
    $wb.RefreshAll()
    $excel.CalculateUntilAsyncQueriesDone()

    Log "Warte 30 s auf Hintergrundabfragen …"
    Start-Sleep -Seconds 30

    $wb.Save()
    $wb.Close($false)
    $excel.Quit()
    [System.Runtime.InteropServices.Marshal]::ReleaseComObject($excel) | Out-Null
    Log "Excel gespeichert und geschlossen."
} catch {
    Log "FEHLER (Excel): $_"
    exit 1
}

# ---- Schritt 2: Dashboard generieren ----
try {
    Log "Python-Skript starten: $pyScript"
    $result = & python $pyScript 2>&1
    Log "Python-Output: $result"
    Log "Dashboard erfolgreich generiert."
} catch {
    Log "FEHLER (Python): $_"
    exit 2
}

Log "=== Fertig ==="
