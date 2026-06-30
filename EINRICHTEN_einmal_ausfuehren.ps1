# ============================================================
#  EINRICHTEN_einmal_ausfuehren.ps1
#  Dieses Skript NUR EINMAL als Administrator ausführen.
#  Es registriert einen täglichen Windows-Task um 08:00 Uhr.
# ============================================================

$taskName  = "WB_Dashboard_Refresh"
$scriptDir = Split-Path -Parent $MyInvocation.MyCommand.Path
$ps1       = Join-Path $scriptDir "WB_Refresh_und_Dashboard.ps1"

# Argument: Execution-Policy umgehen, Fenster versteckt
$argument  = "-NonInteractive -WindowStyle Hidden -ExecutionPolicy Bypass -File `"$ps1`""

$action    = New-ScheduledTaskAction -Execute "powershell.exe" -Argument $argument
$trigger   = New-ScheduledTaskTrigger -Daily -At "08:00"
$settings  = New-ScheduledTaskSettingsSet `
                -RunOnlyIfNetworkAvailable `
                -StartWhenAvailable `
                -ExecutionTimeLimit (New-TimeSpan -Minutes 10)

# Task unter dem aktuellen Benutzer registrieren (kein Passwort nötig)
Register-ScheduledTask `
    -TaskName  $taskName `
    -Action    $action `
    -Trigger   $trigger `
    -Settings  $settings `
    -RunLevel  Highest `
    -Force

Write-Host ""
Write-Host "✓ Task '$taskName' erfolgreich registriert." -ForegroundColor Green
Write-Host "  Läuft täglich um 08:00 Uhr."
Write-Host "  Skript: $ps1"
Write-Host ""
Write-Host "Zum sofortigen Testen:"
Write-Host "  Start-ScheduledTask -TaskName '$taskName'"
