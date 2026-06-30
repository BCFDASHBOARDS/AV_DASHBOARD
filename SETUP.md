# BCF AV Dashboard — Setup-Anleitung

## Ordnerstruktur

```
051 Dashboard\
├── _source\          ← Lokale Kopien der Quelldateien (temporär)
├── _data\            ← JSON-Exporte (werden zu GitHub gepusht)
├── dashboard\
│   ├── index.html    ← Das Dashboard (online via GitHub Pages)
│   └── _data\        ← JSON-Kopie für GitHub Pages
├── scripts\
│   ├── run_all.ps1              ← Master-Skript (täglich ausführen)
│   ├── 01_copy_network.ps1
│   ├── 02_refresh_ubersicht.ps1
│   ├── 03_sharepoint_download.ps1
│   └── 04_excel_to_json.ps1
├── logs\             ← Tageslog-Dateien
└── SETUP.md          ← Diese Datei
```

---

## Schritt 1 — PowerShell-Module installieren (einmalig)

PowerShell als Admin öffnen und ausführen:

```powershell
Install-Module ImportExcel             -Scope CurrentUser -Force
Install-Module Microsoft.Graph.Authentication -Scope CurrentUser -Force
Install-Module Microsoft.Graph.Files   -Scope CurrentUser -Force
```

---

## Schritt 2 — GitHub Repository anlegen

1. Auf https://github.com einloggen
2. **New repository** → Name z.B. `bcf-av-dashboard` → **Private** (empfohlen)
3. Repository lokal klonen in den Dashboard-Ordner:

```powershell
cd "C:\Users\%USERNAME%\OneDrive - Baussmann\2026\050 Claude planned Tasks\051 Dashboard"
git init
git remote add origin https://github.com/DEIN-USERNAME/bcf-av-dashboard.git
git add .
git commit -m "Initial commit"
git push -u origin main
```

4. Auf GitHub: **Settings → Pages → Source: Deploy from branch → Branch: main → Folder: /dashboard**
5. Nach 1–2 Minuten ist das Dashboard erreichbar unter:  
   `https://DEIN-USERNAME.github.io/bcf-av-dashboard/`

---

## Schritt 3 — SharePoint-Login (einmalig)

Beim ersten Lauf von `03_sharepoint_download.ps1` erscheint ein Gerätecode-Login.  
Den Code auf https://microsoft.com/devicelogin eingeben und mit dem Baussmann-Account einloggen.  
Danach läuft es automatisch.

---

## Schritt 4 — Windows Task Scheduler einrichten

PowerShell als Admin:

```powershell
$action  = New-ScheduledTaskAction `
    -Execute "powershell.exe" `
    -Argument "-NonInteractive -ExecutionPolicy Bypass -File `"C:\Users\$env:USERNAME\OneDrive - Baussmann\2026\050 Claude planned Tasks\051 Dashboard\scripts\run_all.ps1`""

$trigger = New-ScheduledTaskTrigger -Daily -At "06:00"   # Uhrzeit anpassen!

$settings = New-ScheduledTaskSettingsSet `
    -ExecutionTimeLimit (New-TimeSpan -Minutes 30) `
    -RestartCount 2 `
    -RestartInterval (New-TimeSpan -Minutes 5)

Register-ScheduledTask `
    -TaskName   "BCF_Dashboard_Update" `
    -Action     $action `
    -Trigger    $trigger `
    -Settings   $settings `
    -RunLevel   Highest `
    -Force
```

---

## Datenquellen im Überblick

| Skript | Quelle | Zieldatei in _source |
|--------|--------|----------------------|
| 01 | `\\srv12020\...\1100_Datenimport_Pressen.xlsx` | 1100_Datenimport_Pressen.xlsx |
| 01 | `\\srv12020\...\1180_Datenimport_Walzen.xlsx` | 1180_Datenimport_Walzen.xlsx |
| 01 | `\\srv12020\...\Auftragsbestand.xlsx` | Auftragsbestand.xlsx |
| 01 | `\\srv12020\...\Istmengen.xlsx` | Istmengen.xlsx |
| 01 | `\\srv12020\...\Lagerbestand.xlsx` | Lagerbestand.xlsx |
| 02 | `\\srv12020\Allgemein\AV\Draht\übersicht.xlsx` (Query-Refresh) | ubersicht_Draht.xlsx |
| 03 | SharePoint: Offene Bestellungen_Draht.xlsx | Offene_Bestellungen_Draht.xlsx |
| 03 | SharePoint: Planung_Pressen_NEU.xlsx | Planung_Pressen_NEU.xlsx |

---

## Nächste Schritte (nach Grundsetup)

- [ ] Spaltenfilter und Beschriftungen in `index.html` → `COL_LABELS` anpassen
- [ ] Weitere Dateien in `scripts\04_excel_to_json.ps1` → `$FILE_MAP` ergänzen
- [ ] Tabellenblatt-Namen für jede Datei in `$FILE_MAP` eintragen
- [ ] Dashboard-URL intern bekanntgeben
