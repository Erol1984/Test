<#
  Liest alle Mails aus dem Outlook-Ordner "U/Ü Lieferungen" (Posteingang/SCM),
  extrahiert die Excel-Anhaenge (taegliche Snapshot-Reports) und baut daraus
  einen deduplizierten Datensatz fuer das U/Ue-Lieferungen-Dashboard.

  Liest jeden Sheet-Inhalt in EINEM COM-Aufruf (UsedRange.Value2) statt
  zellweise - das ist deutlich schneller und vermeidet eine Encoding-
  Beschaedigung, die bei vielen einzelnen Cells.Item(...).Text-Aufrufen
  ueber eine gemeinsam genutzte Excel-Instanz auftritt.

  Usage:
    powershell -File extract-ue-lieferungen.ps1 [-MaxMails 10] [-OutFile path\to\data.json]
#>
param(
    [int]$MaxMails = 0,             # 0 = alle Mails
    [string]$OutFile = "$PSScriptRoot\..\ue-lieferungen-data.json"
)

$ErrorActionPreference = "Stop"

function Find-Folder($root, $name) {
    foreach ($f in $root.Folders) {
        if ($f.Name -eq $name) { return $f }
        $sub = Find-Folder $f $name
        if ($sub) { return $sub }
    }
    return $null
}

function Convert-CellDate($v) {
    if ($null -eq $v) { return $null }
    if ($v -is [double] -or $v -is [int]) {
        try { return [datetime]::FromOADate([double]$v).ToString("yyyy-MM-dd") } catch { return $null }
    }
    $dt = [datetime]::MinValue
    if ([datetime]::TryParse([string]$v, [ref]$dt)) { return $dt.ToString("yyyy-MM-dd") }
    return $null
}

function Is-ExcelError($v) {
    # Zellen mit Formelfehlern (#N/A, #VALUE!, #REF! ...) kommen aus Value2 als
    # Int32-Fehlercode zurueck (z.B. -2146826246 = #N/A), nicht als Text/Zahl.
    return ($v -is [int]) -and ($v -lt -1000000)
}

function Convert-CellNum($v) {
    if ($null -eq $v -or (Is-ExcelError $v)) { return $null }
    if ($v -is [double] -or $v -is [int]) { return [double]$v }
    $parsed = 0.0
    if ([double]::TryParse([string]$v, [ref]$parsed)) { return $parsed }
    return $null
}

function Convert-CellText($v) {
    if ($null -eq $v -or (Is-ExcelError $v)) { return "" }
    return ([string]$v).Trim()
}

Write-Output "Verbinde mit Outlook..."
$outlook = New-Object -ComObject Outlook.Application
$ns = $outlook.GetNamespace("MAPI")
$target = $null
foreach ($store in $ns.Folders) {
    $target = Find-Folder $store "U/Ü Lieferungen"
    if ($target) { break }
}
if (-not $target) { throw "Ordner 'U/Ü Lieferungen' nicht gefunden." }

$items = $target.Items
$items.Sort("[ReceivedTime]", $true)
$total = $items.Count
if ($MaxMails -gt 0 -and $MaxMails -lt $total) { $total = $MaxMails }
Write-Output "Verarbeite $total von $($items.Count) Mails..."

$tempDir = Join-Path $env:TEMP "ue_lieferungen_extract"
New-Item -ItemType Directory -Force -Path $tempDir | Out-Null

function New-ExcelApp {
    $e = New-Object -ComObject Excel.Application
    $e.Visible = $false
    $e.DisplayAlerts = $false
    $e.AutomationSecurity = 3   # msoAutomationSecurityForceDisable
    return $e
}

function Stop-ExcelApp($e) {
    try { $e.Quit() } catch { }
    try { [System.Runtime.Interopservices.Marshal]::ReleaseComObject($e) | Out-Null } catch { }
    [GC]::Collect()
    [GC]::WaitForPendingFinalizers()
}

Write-Output "Starte Excel (Makros deaktiviert)..."
$excel = New-ExcelApp

# Excel-COM-Automation wird ueber viele sequenzielle Workbook-Oeffnungen instabil
# (kann abstuerzen). Deshalb: alle N Dateien vorsorglich neu starten, und bei
# einem Fehler die betroffene Datei nach Neustart einmal erneut versuchen.
$RESTART_EVERY = 25

$allRows = New-Object System.Collections.Generic.List[object]
$processed = 0
$errors = 0

for ($i = 1; $i -le $total; $i++) {
    if ($processed -gt 0 -and ($processed % $RESTART_EVERY -eq 0)) {
        Stop-ExcelApp $excel
        $excel = New-ExcelApp
    }

    $item = $items.Item($i)
    $meldeDatum = $item.ReceivedTime.ToString("yyyy-MM-dd")
    $tmpFile = Join-Path $tempDir "mail_$i.xlsm"

    $attempt = 0
    $done = $false
    while (-not $done -and $attempt -lt 2) {
        $attempt++
        try {
            if ($item.Attachments.Count -lt 1) { $done = $true; continue }
            $att = $item.Attachments.Item(1)
            if ($att.FileName -notmatch '\.xls[mx]$') { $done = $true; continue }
            $att.SaveAsFile($tmpFile)

            $wb = $excel.Workbooks.Open($tmpFile, $false, $true)
            $ws = $wb.Worksheets.Item(1)
            $arr = $ws.UsedRange.Value2   # ein einziger COM-Aufruf -> 2D-Array, 1-indiziert

            $rowCount = $arr.GetLength(0)
            for ($r = 3; $r -le $rowCount; $r++) {
                $kundenAuftrag = Convert-CellText $arr[$r, 1]
                if ([string]::IsNullOrWhiteSpace($kundenAuftrag)) { continue }

                $row = [ordered]@{
                    kundenAuftrag  = $kundenAuftrag
                    pos            = Convert-CellText $arr[$r, 2]
                    art            = Convert-CellText $arr[$r, 3]
                    fauf           = Convert-CellText $arr[$r, 4]
                    material       = Convert-CellText $arr[$r, 5]
                    liefertermin   = Convert-CellDate  $arr[$r, 6]
                    kunde          = Convert-CellText $arr[$r, 7]
                    soll           = Convert-CellNum  $arr[$r, 8]
                    ist            = Convert-CellNum  $arr[$r, 9]
                    minMenge       = Convert-CellNum  $arr[$r, 10]
                    maxMenge       = Convert-CellNum  $arr[$r, 11]
                    erfasser       = Convert-CellText $arr[$r, 15]
                    maschNr        = Convert-CellText $arr[$r, 16]
                    maschine       = Convert-CellText $arr[$r, 17]
                    nachproduktion = Convert-CellNum  $arr[$r, 19]
                    meldedatum     = $meldeDatum
                }
                $allRows.Add((New-Object PSObject -Property $row))
            }

            $wb.Close($false)
            [System.Runtime.Interopservices.Marshal]::ReleaseComObject($ws) | Out-Null
            [System.Runtime.Interopservices.Marshal]::ReleaseComObject($wb) | Out-Null
            Remove-Item $tmpFile -Force -ErrorAction SilentlyContinue
            $processed++
            $done = $true
            if ($processed % 20 -eq 0) { Write-Output "  ... $processed / $total verarbeitet" }
        } catch {
            if ($attempt -ge 2) {
                $errors++
                Write-Warning "Fehler bei Mail $i ($meldeDatum): $($_.Exception.Message)"
            } else {
                Write-Warning "Excel-Fehler bei Mail $i, starte Excel neu und versuche erneut..."
                Stop-ExcelApp $excel
                $excel = New-ExcelApp
            }
        }
    }
}

Stop-ExcelApp $excel

Write-Output "Fertig. $processed Mails verarbeitet, $errors Fehler, $($allRows.Count) Rohzeilen gesammelt."

# Dedup: neuester Meldedatum-Stand pro eindeutigem Schluessel gewinnt
$byKey = @{}
foreach ($row in $allRows) {
    $key = "$($row.kundenAuftrag)|$($row.pos)|$($row.fauf)"
    if (-not $byKey.ContainsKey($key) -or $row.meldedatum -gt $byKey[$key].meldedatum) {
        $byKey[$key] = $row
    }
}
$final = $byKey.Values | ForEach-Object {
    $abweichung = $null
    $abweichungPct = $null
    $typ = "OK"
    if ($null -ne $_.soll -and $null -ne $_.ist -and $_.soll -ne 0) {
        $abweichung = [math]::Round($_.ist - $_.soll, 2)
        $abweichungPct = [math]::Round((($_.ist - $_.soll) / $_.soll) * 100, 2)
        if ($abweichung -lt 0) { $typ = "Unterlieferung" }
        elseif ($abweichung -gt 0) { $typ = "Ueberlieferung" }
    }
    [ordered]@{
        kundenAuftrag = $_.kundenAuftrag
        pos           = $_.pos
        art           = $_.art
        fauf          = $_.fauf
        material      = $_.material
        datum         = $_.liefertermin
        kunde         = $_.kunde
        soll          = $_.soll
        ist           = $_.ist
        abweichung    = $abweichung
        abweichungPct = $abweichungPct
        typ           = $typ
        maschNr       = $_.maschNr
        maschine      = $_.maschine
        erfasser      = $_.erfasser
        meldedatum    = $_.meldedatum
    }
}

$final = $final | Sort-Object { $_.datum }
Write-Output "Eindeutige Datensaetze nach Dedup: $($final.Count)"

$outDir = Split-Path $OutFile -Parent
if ($outDir -and -not (Test-Path $outDir)) { New-Item -ItemType Directory -Force -Path $outDir | Out-Null }
$final | ConvertTo-Json -Depth 5 | Out-File -FilePath $OutFile -Encoding utf8
Write-Output "Geschrieben nach: $OutFile"
