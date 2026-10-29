<#
  Liest die jeweils neueste Forecast-Mail fuer MBB und MBW von "Model Group BOEP"
  aus dem Posteingang (inkl. Unterordner), oeffnet den Excel-Anhang (SAP-BO-Export
  "Forecast Fertigungsauftraege") und baut daraus den Datensatz fuer das
  Forecast-Dashboard.

  Verwendete Blaetter:
    - "Details nach Tag":        eine Zeile pro Fertigungsauftrag
    - "Gesamt KW Auftragsart":   Summen pro KW x Warengruppe x Auftragsart

  Wie im U/Ue-Skript wird jedes Blatt in EINEM COM-Aufruf (UsedRange.Value2)
  gelesen.

  Usage:
    powershell -File extract-forecast.ps1 [-OutFile path\to\data.json]
#>
param(
    [string]$OutFile = "$PSScriptRoot\..\forecast-data.json"
)

$ErrorActionPreference = "Stop"
$logFile = Join-Path $PSScriptRoot "extract-forecast.log"

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

function Get-Grid($ws) {
    $ur = $ws.UsedRange
    return @{ v = $ur.Value2; rows = $ur.Rows.Count; cols = $ur.Columns.Count }
}

function Cell($g, $r, $c) {
    if ($r -lt 1 -or $c -lt 1 -or $r -gt $g.rows -or $c -gt $g.cols) { return $null }
    return $g.v[$r, $c]
}

function Txt($v) {
    if ($null -eq $v) { return "" }
    return ([string]$v).Trim()
}

function Find-Row($g, $col, $text) {
    for ($r = 1; $r -le $g.rows; $r++) { if ((Txt (Cell $g $r $col)) -eq $text) { return $r } }
    throw "Zeile '$text' nicht gefunden"
}

function Convert-Date($v) {
    if ($v -is [double]) { return [datetime]::FromOADate($v).ToString("yyyy-MM-dd") }
    $dt = [datetime]::MinValue
    if ([datetime]::TryParseExact((Txt $v), "dd.MM.yyyy", $null, 0, [ref]$dt)) { return $dt.ToString("yyyy-MM-dd") }
    return $null
}

function Convert-Kw($v) {
    # SAP liefert "44.2026"; Excel kann das als Zahl interpretieren
    $s = if ($v -is [double]) { $v.ToString("00.0000", [cultureinfo]::InvariantCulture) } else { Txt $v }
    if ($s -match '^(\d{1,2})\.(\d{4})$') { return "{0}-W{1:00}" -f $Matches[2], [int]$Matches[1] }
    return $null
}

function Convert-Machine($v) {
    $s = Txt $v
    if ($s -eq "" -or $s -eq "0" -or $s -eq "Nicht zugeordnet") { return $null }
    if ($v -is [double]) { return ([long]$v).ToString() }
    return $s
}

function Read-Details($ws) {
    $g = Get-Grid $ws
    $hdr = Find-Row $g 1 "Liefert. Plan"
    # Wertspalten: Warengruppe steht 3 Zeilen ueber der Einheitenzeile
    $valCols = @()
    for ($c = 11; $c -le $g.cols; $c++) {
        $wg = Txt (Cell $g ($hdr - 2) $c)
        $unit = Txt (Cell $g $hdr $c)
        if ($wg -match '^\d{4}$' -and ($unit -eq "ST" -or $unit -eq "M2")) {
            $valCols += @{ c = $c; wg = $wg; unit = $unit; name = Txt (Cell $g ($hdr - 1) $c) }
        }
    }
    $wgNames = @{}
    foreach ($vc in $valCols) { if ($vc.name) { $wgNames[$vc.wg] = $vc.name } }

    $rows = New-Object System.Collections.Generic.List[object]
    for ($r = $hdr + 1; $r -le $g.rows; $r++) {
        $date = Convert-Date (Cell $g $r 1)
        $qual = Txt (Cell $g $r 2)
        if (-not $date -or $qual -eq "Ergebnis" -or $qual -eq "") { continue }
        $byWg = @{}
        foreach ($vc in $valCols) {
            $v = Cell $g $r $vc.c
            if ($v -isnot [double]) { continue }
            if (-not $byWg.ContainsKey($vc.wg)) { $byWg[$vc.wg] = @{ st = 0.0; m2 = 0.0 } }
            if ($vc.unit -eq "ST") { $byWg[$vc.wg].st = $v } else { $byWg[$vc.wg].m2 = $v }
        }
        $machines = @(2..5 | ForEach-Object { Convert-Machine (Cell $g $r ($_ + 5)) } | Where-Object { $_ })
        foreach ($wg in $byWg.Keys) {
            $rows.Add([ordered]@{
                d     = $date
                q     = $qual
                kunde = Txt (Cell $g $r 3)
                kdnr  = Txt (Cell $g $r 4)
                kgrp  = Txt (Cell $g $r 5)
                auf   = Txt (Cell $g $r 6)
                m     = $machines
                wg    = $wg
                st    = [math]::Round($byWg[$wg].st)
                m2    = [math]::Round($byWg[$wg].m2, 1)
            })
        }
    }
    return @{ rows = $rows; wgNames = $wgNames }
}

function Read-KwArt($ws) {
    $g = Get-Grid $ws
    $artRow = Find-Row $g 1 "Auftragsart"
    $hdr = Find-Row $g 1 "Liefert. Plan"
    $wgRow = Find-Row $g 1 "Warengruppe"
    $cols = @()
    $wg = ""
    for ($c = 2; $c -le $g.cols; $c++) {
        $w = Txt (Cell $g $wgRow $c)
        if ($w) { $wg = $w }
        $art = Txt (Cell $g $artRow $c)
        $unit = Txt (Cell $g $hdr $c)
        if ($wg -match '^\d{4}$' -and $art -and $art -ne "Ergebnis" -and ($unit -eq "ST" -or $unit -eq "M2")) {
            $cols += @{ c = $c; wg = $wg; art = $art; unit = $unit }
        }
    }
    $acc = [ordered]@{}
    for ($r = $hdr + 1; $r -le $g.rows; $r++) {
        $kw = Convert-Kw (Cell $g $r 1)
        if (-not $kw) { continue }
        foreach ($col in $cols) {
            $v = Cell $g $r $col.c
            if ($v -isnot [double]) { continue }
            $key = "$kw|$($col.wg)|$($col.art)"
            if (-not $acc.Contains($key)) { $acc[$key] = [ordered]@{ kw = $kw; wg = $col.wg; art = $col.art; st = 0; m2 = 0 } }
            if ($col.unit -eq "ST") { $acc[$key].st = [math]::Round($v) } else { $acc[$key].m2 = [math]::Round($v, 1) }
        }
    }
    return @($acc.Values)
}

Write-Log "Verbinde mit Outlook..."
$outlook = New-Object -ComObject Outlook.Application
$inbox = $outlook.GetNamespace("MAPI").GetDefaultFolder(6)

# Neueste Mail je Werk suchen (Mails werden per Regel in Unterordner verschoben)
$latest = @{}
function Search-Folder($folder) {
    if ($folder.DefaultItemType -eq 0) {
        $items = $folder.Items.Restrict("@SQL=""urn:schemas:httpmail:fromname"" LIKE '%BOEP%' AND ""urn:schemas:httpmail:subject"" LIKE '%Forecast%'")
        foreach ($m in $items) {
            $plant = if ($m.Subject -like "MBB*") { "MBB" } elseif ($m.Subject -like "MBW*") { "MBW" } else { $null }
            if ($plant -and (-not $latest[$plant] -or $m.ReceivedTime -gt $latest[$plant].ReceivedTime)) { $latest[$plant] = $m }
        }
    }
    foreach ($sub in $folder.Folders) { Search-Folder $sub }
}
Search-Folder $inbox

$tmp = Join-Path $env:TEMP "forecast-extract"
New-Item -ItemType Directory -Force $tmp | Out-Null

$excel = New-Object -ComObject Excel.Application
$excel.DisplayAlerts = $false
$plants = [ordered]@{}
$wgNames = @{}
try {
    foreach ($plant in @("MBB", "MBW")) {
        $mail = $latest[$plant]
        if (-not $mail) { Write-Log "  ${plant}: keine Mail gefunden"; continue }
        $att = $mail.Attachments | Where-Object { $_.FileName -like "*.xlsx" } | Select-Object -First 1
        $path = Join-Path $tmp $att.FileName
        $att.SaveAsFile($path)
        Write-Log "  ${plant}: $($mail.Subject) ($($mail.ReceivedTime))"

        $wb = $excel.Workbooks.Open($path, 0, $true)
        try {
            $det = Read-Details ($wb.Worksheets.Item("Details nach Tag"))
            $kwArt = Read-KwArt ($wb.Worksheets | Where-Object { $_.Name.Trim() -eq "Gesamt KW Auftragsart" } | Select-Object -First 1)
        } finally { $wb.Close($false) }
        Remove-Item $path

        foreach ($k in $det.wgNames.Keys) { $wgNames[$k] = $det.wgNames[$k] }
        $plants[$plant] = [ordered]@{
            mail     = [string]$mail.Subject
            received = ([datetime]$mail.ReceivedTime).ToString("yyyy-MM-ddTHH:mm")
            rows     = $det.rows.ToArray()
            kwArt    = $kwArt
        }
        Write-Log "    $($det.rows.Count) Auftragszeilen, $($kwArt.Count) KW/Auftragsart-Werte"
    }
} finally {
    $excel.Quit()
    [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($excel)
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
