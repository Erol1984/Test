<#
  Gemeinsame Lesefunktionen fuer die SAP-BO-Exporte "Forecast Fertigungsauftraege"
  (Mails von Model Group BOEP). Wird von extract-forecast.ps1 und
  extract-auftragseingang.ps1 per Dot-Sourcing geladen.

  Jedes Blatt wird in EINEM COM-Aufruf (UsedRange.Value2) gelesen.
#>

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

# Alle Forecast-Mails von BOEP im Posteingang inkl. Unterordnern (Mails werden per Regel verschoben).
# Liefert Objekte { plant; mail } - optional nur Mails ab $Since.
function Find-ForecastMails($inbox, [datetime]$Since = [datetime]::MinValue) {
    $filter = "@SQL=""urn:schemas:httpmail:fromname"" LIKE '%BOEP%' AND ""urn:schemas:httpmail:subject"" LIKE '%Forecast%'"
    if ($Since -gt [datetime]::MinValue) {
        $filter += " AND ""urn:schemas:httpmail:datereceived"" >= '$($Since.ToUniversalTime().ToString("yyyy-MM-dd HH:mm"))'"
    }
    $found = New-Object System.Collections.Generic.List[object]
    $walk = {
        param($folder)
        if ($folder.DefaultItemType -eq 0) {
            foreach ($m in $folder.Items.Restrict($filter)) {
                # Sonntags kommt zusaetzlich "... nach Kalenderwoche" (nur Blatt "Gesamt nach KW") - nicht verwendbar
                if ($m.Subject -like "*Kalenderwoche*") { continue }
                $plant = if ($m.Subject -like "MBB*") { "MBB" } elseif ($m.Subject -like "MBW*") { "MBW" } else { $null }
                if ($plant) { $found.Add([pscustomobject]@{ plant = $plant; mail = $m; received = [datetime]$m.ReceivedTime }) }
            }
        }
        foreach ($sub in $folder.Folders) { & $walk $sub }
    }
    & $walk $inbox
    return $found
}

# Excel-Anhang einer Mail speichern und "Details nach Tag" (+ optional "Gesamt KW Auftragsart") lesen.
# $tmp sollte pro Lauf eindeutig sein, damit parallele Laeufe sich nicht in die Quere kommen.
function Read-ForecastMail($excel, $mail, $tmp, [switch]$WithKwArt) {
    $att = $mail.Attachments | Where-Object { $_.FileName -like "*.xlsx" } | Select-Object -First 1
    if (-not $att) { throw "Kein Excel-Anhang in '$($mail.Subject)'" }
    $path = Join-Path $tmp $att.FileName
    $att.SaveAsFile($path)
    $wb = $excel.Workbooks.Open($path, 0, $true)
    try {
        $det = Read-Details ($wb.Worksheets.Item("Details nach Tag"))
        $kwArt = if ($WithKwArt) { Read-KwArt ($wb.Worksheets | Where-Object { $_.Name.Trim() -eq "Gesamt KW Auftragsart" } | Select-Object -First 1) } else { $null }
    } finally { $wb.Close($false) }
    Remove-Item $path
    return @{ rows = $det.rows; wgNames = $det.wgNames; kwArt = $kwArt }
}
