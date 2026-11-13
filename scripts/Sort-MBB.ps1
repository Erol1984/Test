# Sortiert Mails des Werks Model Bad Bentheim (MBB) aus dem Posteingang in
# Posteingang\MBB\Informationen bzw. Posteingang\MBB\Reports MBB.
#
# Erkennung:
#  1) Absender laut Exchange-Adressbuch (GAL) mit City = "Bad Bentheim"  -> MBB
#  2) Automatische Berichte, deren Betreff "MBB"/"Bad Bentheim" enthaelt und die von
#     den unten genannten Systemabsendern kommen                        -> MBB / Reports MBB
# Reports MBB: nur Systemabsender-Berichte (Schichtprotokoll, P-Print Verbrauch,
#   Fertigungsauftragseingang-Forecast, Leimkueche-Verbrauch). Alles andere -> Informationen.
# Gekennzeichnete (Flag) Mails werden NICHT verschoben (gleiche Regel wie Cleanup-Inbox.ps1).
# Ohne -Execute ist es ein Probelauf (nichts wird verschoben).

#
# -Historie: scannt zusaetzlich alle Unterordner des Posteingangs (ausser MBB selbst), nicht
#   nur den Posteingang. Im Probelauf wird je Ordner die Anzahl ausgegeben.

#   Mit -Ordner "A","B" werden bei -Historie nur diese Posteingang-Unterordner gescannt
#   (dann ohne den Posteingang selbst).

param([switch]$Execute, [switch]$Historie, [string[]]$Ordner)

$ErrorActionPreference = "Stop"
$logFile = Join-Path $PSScriptRoot "Sort-MBB.log"
function Write-Log($msg) {
    $line = "$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')  $msg"
    Add-Content -Path $logFile -Value $line
    Write-Output $line
}

$mbbCity = "Bad Bentheim"

# Systemabsender (Anzeigename bzw. SMTP), die MBB-Berichte versenden -> Reports MBB,
# aber nur wenn der Betreff MBB / Bad Bentheim enthaelt.
$reportSenders = @("Formularweb MGM", "Model Group BOEP", "fernwartung@es-ha-es.de")
$mbbSubjectPattern = "MBB|Bad Bentheim"

$ol = New-Object -ComObject Outlook.Application
$ns = $ol.GetNamespace("MAPI")
$inbox = $ns.GetDefaultFolder(6)
$mbb = $inbox.Folders.Item("MBB")
$fInfo = $mbb.Folders.Item("Informationen")
$fRep  = $mbb.Folders.Item("Reports MBB")

$cityCache = @{}
function Get-SenderCity($it) {
    $key = "$($it.SenderEmailType)|$($it.SenderEmailAddress)|$($it.SenderName)"
    if ($cityCache.ContainsKey($key)) { return $cityCache[$key] }
    $city = $null
    try {
        if ($it.SenderEmailType -eq "EX") {
            $u = $it.Sender.GetExchangeUser()
            if ($u) { $city = $u.City }
        }
    } catch {}
    $cityCache[$key] = $city
    return $city
}

function Get-AllSubfolders($folder, $prefix) {
    foreach ($f in $folder.Folders) {
        $name = "$prefix$($f.Name)"
        if ($prefix -eq "" -and $f.Name -eq "MBB") { continue }
        [PSCustomObject]@{ Folder = $f; Name = $name }
        Get-AllSubfolders $f "$name\"
    }
}

$scanTargets = @([PSCustomObject]@{ Folder = $inbox; Name = "Posteingang" })
if ($Historie) {
    $all = @(Get-AllSubfolders $inbox "")
    if ($Ordner) {
        $scanTargets = @($all | Where-Object { $Ordner -contains $_.Name })
    } else {
        $scanTargets += $all
    }
}

$plan = @()
$skippedFlag = @()
foreach ($st in $scanTargets) {
foreach ($it in @($st.Folder.Items)) {
    try {
        if ($it.Class -ne 43) { continue }
        $subj = [string]$it.Subject
        $isSystem = ($reportSenders -contains $it.SenderName) -and ($subj -match $mbbSubjectPattern)
        $isMbbPerson = ((Get-SenderCity $it) -eq $mbbCity)
        if (-not ($isSystem -or $isMbbPerson)) { continue }
        $target = if ($isSystem) { "Reports MBB" } else { "Informationen" }
        $row = [PSCustomObject]@{ Item = $it; Target = $target; Subject = $subj; From = $it.SenderName; Date = $it.ReceivedTime.ToString("yyyy-MM-dd HH:mm"); Unread = $it.UnRead; Quelle = $st.Name }
        if ($it.FlagStatus -ne 0) { $skippedFlag += $row; continue }
        $plan += $row
    } catch {}
}
}

Write-Log ("Modus: " + $(if ($Execute) { "AUSFUEHREN" } else { "PROBELAUF" }))
Write-Log "Zu verschieben: $($plan.Count)  (davon ungelesen: $(@($plan | Where-Object Unread).Count)); wegen Kennzeichnung uebersprungen: $($skippedFlag.Count)"
foreach ($g in ($plan | Group-Object Target)) { Write-Log "  -> $($g.Name): $($g.Count)" }
foreach ($g in ($plan | Group-Object Quelle | Sort-Object Count -Descending)) { Write-Log "  Quelle $($g.Name): $($g.Count)" }
foreach ($s in $skippedFlag) { Write-Log "  uebersprungen (gekennzeichnet): $($s.Quelle) | $($s.Date) | $($s.From) | $($s.Subject)" }

if (-not $Execute) {
    $plan | Sort-Object Quelle, Target, From | ForEach-Object { "{0,-22} | {1,-12} | {2} | {3} | {4}" -f $_.Quelle, $_.Target, $_.Date, $_.From, $_.Subject }
    exit 0
}

$ok = 0
foreach ($e in $plan) {
    try {
        $dest = if ($e.Target -eq "Reports MBB") { $fRep } else { $fInfo }
        $e.Item.Move($dest) | Out-Null
        $ok++
        Add-Content -Path $logFile -Value "  verschoben ($($e.Quelle)) -> $($e.Target) | $($e.Date) | $($e.From) | $($e.Subject)"
    } catch {
        Write-Log "FEHLER beim Verschieben '$($e.Subject)' -> $($e.Target): $($_.Exception.Message)"
    }
}
Write-Log "Fertig: $ok von $($plan.Count) verschoben."
