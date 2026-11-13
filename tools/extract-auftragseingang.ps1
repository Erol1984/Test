<#
  Auftragseingang aus den taeglichen Forecast-Mails von "Model Group BOEP".

  Jede Forecast-Mail (MBB und MBW) enthaelt alle offenen Fertigungsauftraege.
  Dieses Skript vergleicht jeden Tag mit dem vorherigen Tag desselben Werks:
    - neu:       Auftrag ist heute im Export, gestern nicht   -> Auftragseingang
    - entfallen: Auftrag war gestern im Export, heute nicht   -> produziert/geliefert oder storniert
    - geaendert: Menge oder Liefertermin hat sich geaendert

  Jede Mail wird nur einmal gelesen und als Snapshot in auftragseingang-cache\
  abgelegt (gitignored, echte Kundendaten). Folgelaeufe lesen nur neue Mails.

  Usage:
    powershell -File extract-auftragseingang.ps1 [-Days 60] [-KeepDays 120] [-OutFile path\to\data.json]
#>
param(
    [int]$Days = 60,
    [int]$KeepDays = 120,
    [string]$OutFile = "$PSScriptRoot\..\auftragseingang-data.json"
)

$ErrorActionPreference = "Stop"
$logFile = Join-Path $PSScriptRoot "extract-auftragseingang.log"
$cacheDir = [System.IO.Path]::GetFullPath("$PSScriptRoot\..\auftragseingang-cache")
. (Join-Path $PSScriptRoot "forecast-common.ps1")

function Write-Log($msg) {
    $line = "$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')  $msg"
    Add-Content -Path $logFile -Value $line
    Write-Output $line
}
trap {
    Write-Log "FEHLER: $_"
    exit 1
}

$utf8 = New-Object System.Text.UTF8Encoding $false

function Slim($r) {
    # Nur was das Dashboard braucht (ohne Maschinen)
    return [ordered]@{ auf = $r.auf; d = $r.d; kunde = $r.kunde; kgrp = $r.kgrp; wg = $r.wg; q = $r.q; st = $r.st; m2 = $r.m2 }
}
New-Item -ItemType Directory -Force $cacheDir | Out-Null

Write-Log "Verbinde mit Outlook..."
$outlook = New-Object -ComObject Outlook.Application
$inbox = $outlook.GetNamespace("MAPI").GetDefaultFolder(6)
# Einen Tag mehr holen, damit auch der erste Tag im Zeitraum einen Vortag zum Vergleich hat
$since = (Get-Date).Date.AddDays(-$Days - 1)
$mails = Find-ForecastMails $inbox $since

# Pro Werk und Tag nur die neueste Mail (falls ein Report mal doppelt kommt)
$perDay = @{}
foreach ($f in $mails) {
    $key = "$($f.plant)_$($f.received.ToString('yyyy-MM-dd'))"
    if (-not $perDay[$key] -or $f.received -gt $perDay[$key].received) { $perDay[$key] = $f }
}
Write-Log "$($mails.Count) Forecast-Mails seit $($since.ToString('yyyy-MM-dd')), $($perDay.Count) Werk-Tage"

# Fehlende Snapshots aus den Excel-Anhaengen erzeugen
$todo = @($perDay.GetEnumerator() | Where-Object { -not (Test-Path (Join-Path $cacheDir "$($_.Key).json")) })
if ($todo.Count) {
    Write-Log "$($todo.Count) neue Snapshots werden gelesen..."
    $tmp = Join-Path $env:TEMP "auftragseingang-$PID"
    New-Item -ItemType Directory -Force $tmp | Out-Null
    $excel = New-Object -ComObject Excel.Application
    $excel.DisplayAlerts = $false
    try {
        foreach ($e in ($todo | Sort-Object Key)) {
            $f = $e.Value
            $det = Read-ForecastMail $excel $f.mail $tmp
            $snap = [ordered]@{
                plant    = $f.plant
                received = $f.received.ToString("yyyy-MM-ddTHH:mm")
                wgNames  = $det.wgNames
                rows     = @($det.rows | ForEach-Object { Slim $_ })
            }
            $json = $snap | ConvertTo-Json -Depth 5 -Compress
            [System.IO.File]::WriteAllText((Join-Path $cacheDir "$($e.Key).json"), $json, $utf8)
            Write-Log "  $($e.Key): $($det.rows.Count) Auftraege"
        }
    } finally {
        $excel.Quit()
        [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($excel)
        Remove-Item $tmp -Recurse -Force -ErrorAction SilentlyContinue
    }
}

# Cache begrenzen: Snapshots aelter als $KeepDays loeschen (ca. 0,3 MB pro Werk und Tag)
$keepCut = (Get-Date).Date.AddDays(-$KeepDays).ToString("yyyy-MM-dd")
Get-ChildItem $cacheDir -Filter "*.json" | Where-Object { $_.BaseName -match '_(\d{4}-\d{2}-\d{2})$' -and $Matches[1] -lt $keepCut } | Remove-Item

# Snapshots laden (auch aeltere aus dem Cache, falls die Mails inzwischen geloescht wurden)
$cutoff = $since.ToString("yyyy-MM-dd")
$snaps = @{ MBB = @(); MBW = @() }
foreach ($file in (Get-ChildItem $cacheDir -Filter "*.json" | Sort-Object Name)) {
    if ($file.BaseName -notmatch '^(MBB|MBW)_(\d{4}-\d{2}-\d{2})$') { continue }
    if ($Matches[2] -lt $cutoff) { continue }
    $snaps[$Matches[1]] += , ([System.IO.File]::ReadAllText($file.FullName, $utf8) | ConvertFrom-Json)
}


$wgNames = @{}
$plants = [ordered]@{}
foreach ($plant in @("MBB", "MBW")) {
    $list = $snaps[$plant]
    $dayList = New-Object System.Collections.Generic.List[object]
    $prev = $null
    foreach ($s in $list) {
        foreach ($p in $s.wgNames.PSObject.Properties) { $wgNames[$p.Name] = $p.Value }
        $cur = @{}
        foreach ($r in $s.rows) { $cur[$r.auf] = $r }
        $tot = [ordered]@{ n = $cur.Count; st = 0.0; m2 = 0.0 }
        foreach ($r in $s.rows) { $tot.st += $r.st; $tot.m2 += $r.m2 }
        $tot.st = [math]::Round($tot.st); $tot.m2 = [math]::Round($tot.m2)

        $day = [ordered]@{ date = $s.received.Substring(0, 10); received = $s.received; prev = $null; bestand = $tot }
        if ($prev) {
            $neu = New-Object System.Collections.Generic.List[object]
            $weg = New-Object System.Collections.Generic.List[object]
            $chg = New-Object System.Collections.Generic.List[object]
            foreach ($k in $cur.Keys) {
                $r = $cur[$k]
                if (-not $prev.map.ContainsKey($k)) { $neu.Add((Slim $r)); continue }
                $o = $prev.map[$k]
                if ($o.st -ne $r.st -or $o.d -ne $r.d) {
                    $c = Slim $r
                    $c.st0 = $o.st; $c.m20 = $o.m2; $c.d0 = $o.d
                    $chg.Add($c)
                }
            }
            foreach ($k in $prev.map.Keys) { if (-not $cur.ContainsKey($k)) { $weg.Add((Slim $prev.map[$k])) } }
            $day.prev = $prev.date
            $day.neu = $neu.ToArray()
            $day.weg = $weg.ToArray()
            $day.chg = $chg.ToArray()
        }
        if ($day.date -ge (Get-Date).Date.AddDays(-$Days).ToString("yyyy-MM-dd") -or $day.prev) { $dayList.Add($day) }
        $prev = @{ date = $day.date; map = $cur }
    }
    $plants[$plant] = [ordered]@{ days = $dayList.ToArray() }
    $withDiff = @($dayList | Where-Object { $_.prev })
    Write-Log ("  {0}: {1} Tage, {2} neue / {3} entfallene / {4} geaenderte Auftraege" -f $plant, $dayList.Count,
        ($withDiff | ForEach-Object { $_.neu.Count } | Measure-Object -Sum).Sum,
        ($withDiff | ForEach-Object { $_.weg.Count } | Measure-Object -Sum).Sum,
        ($withDiff | ForEach-Object { $_.chg.Count } | Measure-Object -Sum).Sum)
}

$out = [ordered]@{
    generated = (Get-Date).ToString("yyyy-MM-ddTHH:mm")
    days      = $Days
    wgNames   = $wgNames
    plants    = $plants
}
$json = $out | ConvertTo-Json -Depth 6 -Compress
[System.IO.File]::WriteAllText([System.IO.Path]::GetFullPath($OutFile), $json, $utf8)
Write-Log "Daten -> $OutFile"

$template = "$PSScriptRoot\..\auftragseingang-dashboard.html"
if (Test-Path $template) {
    $standalone = "$PSScriptRoot\..\auftragseingang-dashboard-standalone.html"
    $html = [System.IO.File]::ReadAllText([System.IO.Path]::GetFullPath($template), [System.Text.Encoding]::UTF8)
    $html = $html.Replace("/*__AE_DATA__*/null", $json.Replace("</", "<\/"))
    [System.IO.File]::WriteAllText([System.IO.Path]::GetFullPath($standalone), $html, $utf8)
    Write-Log "Dashboard -> $standalone"
}
