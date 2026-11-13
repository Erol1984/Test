<#
  Liest die jeweils neueste Forecast-Mail fuer MBB und MBW von "Model Group BOEP"
  aus dem Posteingang (inkl. Unterordner), oeffnet den Excel-Anhang (SAP-BO-Export
  "Forecast Fertigungsauftraege") und baut daraus den Datensatz fuer das
  Forecast-Dashboard.

  Verwendete Blaetter:
    - "Details nach Tag":        eine Zeile pro Fertigungsauftrag
    - "Gesamt KW Auftragsart":   Summen pro KW x Warengruppe x Auftragsart

  Die Lesefunktionen liegen in forecast-common.ps1.

  Usage:
    powershell -File extract-forecast.ps1 [-OutFile path\to\data.json]
#>
param(
    [string]$OutFile = "$PSScriptRoot\..\forecast-data.json"
)

$ErrorActionPreference = "Stop"
$logFile = Join-Path $PSScriptRoot "extract-forecast.log"
. (Join-Path $PSScriptRoot "forecast-common.ps1")

function Write-Log($msg) {
    $line = "$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')  $msg"
    Add-Content -Path $logFile -Value $line
    Write-Output $line
}

# Laeuft taeglich als geplante Aufgabe - Fehler landen im Log statt nur in der Konsole
trap {
    Write-Log "FEHLER: $_"
    exit 1
}

Write-Log "Verbinde mit Outlook..."
$outlook = New-Object -ComObject Outlook.Application
$inbox = $outlook.GetNamespace("MAPI").GetDefaultFolder(6)

# Neueste Mail je Werk
$latest = @{}
foreach ($f in (Find-ForecastMails $inbox)) {
    if (-not $latest[$f.plant] -or $f.received -gt $latest[$f.plant].received) { $latest[$f.plant] = $f }
}

# Eigener Ordner pro Lauf, damit gleichzeitige Laeufe sich nicht die Datei wegnehmen
$tmp = Join-Path $env:TEMP "forecast-extract-$PID"
New-Item -ItemType Directory -Force $tmp | Out-Null

$excel = New-Object -ComObject Excel.Application
$excel.DisplayAlerts = $false
$plants = [ordered]@{}
$wgNames = @{}
try {
    foreach ($plant in @("MBB", "MBW")) {
        if (-not $latest[$plant]) { Write-Log "  ${plant}: keine Mail gefunden"; continue }
        $mail = $latest[$plant].mail
        Write-Log "  ${plant}: $($mail.Subject) ($($mail.ReceivedTime))"
        $det = Read-ForecastMail $excel $mail $tmp -WithKwArt

        foreach ($k in $det.wgNames.Keys) { $wgNames[$k] = $det.wgNames[$k] }
        $plants[$plant] = [ordered]@{
            mail     = [string]$mail.Subject
            received = $latest[$plant].received.ToString("yyyy-MM-ddTHH:mm")
            rows     = $det.rows.ToArray()
            kwArt    = $det.kwArt
        }
        Write-Log "    $($det.rows.Count) Auftragszeilen, $($det.kwArt.Count) KW/Auftragsart-Werte"
    }
} finally {
    $excel.Quit()
    [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($excel)
    Remove-Item $tmp -Recurse -Force -ErrorAction SilentlyContinue
}

$out = [ordered]@{
    generated = (Get-Date).ToString("yyyy-MM-ddTHH:mm")
    wgNames   = $wgNames
    plants    = $plants
}
$json = $out | ConvertTo-Json -Depth 6 -Compress
[System.IO.File]::WriteAllText([System.IO.Path]::GetFullPath($OutFile), $json, (New-Object System.Text.UTF8Encoding $false))
Write-Log "Daten -> $OutFile"

# Eigenstaendige Dashboard-Datei mit eingebetteten Daten erzeugen
$template = "$PSScriptRoot\..\forecast-dashboard.html"
if (Test-Path $template) {
    $standalone = "$PSScriptRoot\..\forecast-dashboard-standalone.html"
    $html = [System.IO.File]::ReadAllText([System.IO.Path]::GetFullPath($template), [System.Text.Encoding]::UTF8)
    $html = $html.Replace("/*__FORECAST_DATA__*/null", $json.Replace("</", "<\/"))
    [System.IO.File]::WriteAllText([System.IO.Path]::GetFullPath($standalone), $html, (New-Object System.Text.UTF8Encoding $false))
    Write-Log "Dashboard -> $standalone"
}
