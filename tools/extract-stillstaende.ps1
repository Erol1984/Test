<#
  Sammelt die MBB-Stillstandsmeldungen eines Zeitraums aus dem ganzen Postfach
  (alle Ordner, auch Gesendete/Geloeschte) fuer das Dashboard "MBB Stillstaende".

  Erfasst werden Mails an den Verteiler "MBB Stoerung (DL)" (An oder Cc, jeder
  Betreff) sowie Mails, deren Betreff "ausfall", "stillst", "stoerung"/"Stoerung",
  "defekt", "steht" oder "Schichtprotokoll" enthaelt, jeweils nur wenn sie MBB betreffen:
  Ordner MBB\..., Betreff mit MBB/Bentheim, Absender mit GAL-City "Bad Bentheim"
  oder MBB/Bad Bentheim im Text.

  Ausgabe: JSON-Liste, eine Zeile je Mail, gruppiert nach Konversation, mit
  gekuerztem Text (Zitat der Vorgaenger-Mail und Signatur abgeschnitten).
  Die Zuordnung zu Stillstaenden (Beginn, Ende, Dauer, Ursache) macht die geplante
  Claude-Aufgabe "MBB Stillstaende aktualisieren", nicht dieses Skript.

  Usage:
    powershell -File extract-stillstaende.ps1 [-Days 3] [-OutFile path\to\neu.json]
#>
param(
    [int]$Days = 3,
    [string]$OutFile = "$PSScriptRoot\..\stillstaende-neu.json"
)

$ErrorActionPreference = "Stop"
$logFile = Join-Path $PSScriptRoot "extract-stillstaende.log"
function Write-Log($msg) {
    $line = "$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')  $msg"
    Add-Content -Path $logFile -Value $line
    Write-Output $line
}
trap { Write-Log "FEHLER: $_"; exit 1 }

$since = (Get-Date).Date.AddDays(-$Days)
$ol = New-Object -ComObject Outlook.Application
$ns = $ol.GetNamespace("MAPI")
$root = $ns.DefaultStore.GetRootFolder()

# Outlook filtert nur nach Datum (DASL, UTC, ISO - die Jet-Syntax [ReceivedTime] haengt vom
# Gebietsschema ab). Betreff und Empfaenger prueft das Skript selbst: Restrict mit LIKE auf
# Empfaengerfelder liefert bei Exchange keine Treffer, und viele Meldungen an den Verteiler
# "MBB Stoerung (DL)" tragen nur den Maschinennamen im Betreff ("MWU", "Emba1", "MARTIN").
$sinceUtc = $since.ToUniversalTime().ToString("yyyy-MM-dd HH:mm")
$dateFilter = "@SQL=""urn:schemas:httpmail:datereceived"" >= '$sinceUtc'"
# "." statt Umlaut, damit die Datei reines ASCII bleibt (PowerShell 5.1 liest sie sonst als ANSI)
$subjectPattern = 'ausfall|stillst|st(oe|.)rung|defekt|\ssteht|Schichtprotokoll'
$listPattern = 'MBB St(oe|.)rung'

$cityCache = @{}
function Get-City($it) {
    $k = $it.SenderEmailAddress
    if ($cityCache.ContainsKey($k)) { return $cityCache[$k] }
    $c = $null
    try { if ($it.SenderEmailType -eq "EX") { $c = $it.Sender.GetExchangeUser().City } } catch {}
    $cityCache[$k] = $c
    return $c
}

function Get-ShortBody([string]$b) {
    $b = $b -replace "`r", ""
    # Zitierte Vorgaenger-Mail abschneiden
    $m = [regex]::Match($b, "\n(_{5,}|Von: |From: |-----Original)")
    if ($m.Success) { $b = $b.Substring(0, $m.Index) }
    $keep = foreach ($l in $b -split "`n") {
        $s = $l.Trim()
        if (-not $s) { continue }
        if ($s -match '^(Stufe ?\d: Zusammen|Planung:|Bei Bedarf Liste|WK?V-Leitung|Model GmbH|Stockholmer|48455|T \+49|M \+49|modelgroup\.com|\[www|Amtsgericht|Gesch.ftsf|Production Packaging|Zum Formular|Formularweb|[a-z.0-9]+@modelgroup)') { continue }
        $s
    }
    $t = ($keep -join " / ")
    if ($t.Length -gt 2500) { $t = $t.Substring(0, 2500) + " ..." }
    return $t
}

$rows = New-Object System.Collections.ArrayList
function Walk($f, $path) {
    if ($f.DefaultItemType -eq 0 -and $path -notlike '*Synchronisierungsprobleme*') {
        try {
            $items = $f.Items.Restrict($dateFilter)
            foreach ($it in $items) {
                if ($it.Class -ne 43) { continue }
                $toList = [bool](($it.To + ";" + $it.CC) -match $listPattern)
                if (-not $toList -and $it.Subject -notmatch $subjectPattern) { continue }
                $isMbb = ($path -like '*\MBB*') -or ($it.Subject -match 'MBB|Bentheim') -or ((Get-City $it) -eq 'Bad Bentheim')
                if (-not $isMbb -and $it.Body -notmatch 'MBB|Bad Bentheim') { continue }
                $null = $rows.Add([pscustomobject]@{
                    empfangen    = $it.ReceivedTime.ToString("yyyy-MM-dd HH:mm")
                    von          = $it.SenderName
                    betreff      = $it.Subject
                    an_verteiler = $toList
                    konversation = $it.ConversationTopic
                    ordner       = $path
                    text         = Get-ShortBody $it.Body
                })
            }
        } catch { Write-Log "WARNUNG Ordner $path : $_" }
    }
    foreach ($s in $f.Folders) { Walk $s "$path\$($s.Name)" }
}
Walk $root ""

# Doppelte (gleiche Mail in mehreren Ordnern) entfernen
$uniq = $rows | Sort-Object empfangen, von, betreff -Unique | Sort-Object konversation, empfangen
$json = ConvertTo-Json -InputObject @($uniq) -Depth 3
[IO.File]::WriteAllText((Resolve-Path -LiteralPath (Split-Path $OutFile) | Join-Path -ChildPath (Split-Path $OutFile -Leaf)), $json, (New-Object System.Text.UTF8Encoding $false))
Write-Log "OK: $($uniq.Count) Mails seit $($since.ToString('yyyy-MM-dd')) -> $OutFile"
