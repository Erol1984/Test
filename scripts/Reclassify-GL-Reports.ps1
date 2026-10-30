# Einmaliger Nachlauf: verschiebt bereits vorhandene Mails aus den Themenordnern
# (nicht nur aus dem Posteingang) nach GL bzw. Reports, gemaess derselben Absender-/
# Betreff-Regeln wie im taeglichen Cleanup-Inbox.ps1. Gekennzeichnete (markierte)
# Mails werden dabei ausgelassen. Die 24h-Regel gilt hier nicht, da es sich um
# bereits abgelegte, historische Mails handelt.

$ErrorActionPreference = "Stop"
$logFile = Join-Path $PSScriptRoot "Reclassify-GL-Reports.log"

function Write-Log($msg) {
    $line = "$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')  $msg"
    Add-Content -Path $logFile -Value $line
    Write-Output $line
}

# Anzeige-Name (wie er im Exchange-Adressbuch/als SenderName in Mails erscheint) -> Zielordner
$nameMap = @{
    "Model Sarah"            = "GL"
    "Model David"            = "GL"
    "Uebersax Michael"       = "GL"
    "Hugentobler Matthias"   = "GL"
    "Hess Matthias"          = "GL"
    "Wilms Michael"          = "GL"
    "Höller Günter"          = "GL"
    "Nicolaus Franz Joseph"  = "GL"
    "Huber Fabian"           = "Reports"
    "Berset Loïck"           = "Reports"
}

$reportsSubjectKeywords = @(
    "Monatsabschluss",
    "Quartalabschluss",
    "Quartalsabschluss",
    "Jahresabschluss",
    "Produktionsrapport",
    "Budget"
)

$foldersToScan = @("SM","Projekte","privat","MEM","Kennzahlen","SCM","HR","IT-Tickets",
    "System-Meldungen","Termin-Rueckmeldungen","Newsletter","Verein FCWB","Unzugeordnet",
    "Machine supplier","Close Contacts")

try {
    if (-not (Get-Process -Name "OUTLOOK" -ErrorAction SilentlyContinue)) {
        Write-Log "Outlook laeuft nicht, starte es..."
        $exe = Get-ChildItem "C:\Program Files\Microsoft Office","C:\Program Files (x86)\Microsoft Office" -Recurse -Filter "OUTLOOK.EXE" -ErrorAction SilentlyContinue | Select-Object -First 1
        if ($exe) { Start-Process $exe.FullName; Start-Sleep -Seconds 20 }
    }
    $outlook = New-Object -ComObject Outlook.Application
    $namespace = $outlook.GetNamespace("MAPI")
    $inbox = $namespace.GetDefaultFolder(6)
} catch {
    Write-Log "FEHLER: Outlook COM nicht erreichbar - $($_.Exception.Message)"
    exit 1
}

$scanTargets = @([PSCustomObject]@{ Folder = $inbox; Name = "Posteingang" })
foreach ($fname in $foldersToScan) {
    try {
        $f = $inbox.Folders.Item($fname)
        if ($f) { $scanTargets += [PSCustomObject]@{ Folder = $f; Name = $fname } }
    } catch { Write-Log "Ordner '$fname' nicht gefunden, uebersprungen." }
}

$seen = @{}
$toMove = @()

foreach ($target in $scanTargets) {
    $folder = $target.Folder

    # 1) Absender-basiert (GL + Reports-Personen), exakter SenderName-Abgleich
    foreach ($name in $nameMap.Keys) {
        try {
            $escaped = $name.Replace("'", "''")
            $items = $folder.Items.Restrict("[SenderName] = '$escaped'")
            foreach ($it in @($items)) {
                try {
                    if ($it.Class -ne 43) { continue }
                    if ($it.FlagStatus -ne 0) { continue }
                    if ($seen.ContainsKey($it.EntryID)) { continue }
                    $seen[$it.EntryID] = $true
                    $toMove += [PSCustomObject]@{ Item = $it; Target = $nameMap[$name]; Subject = $it.Subject; From = $target.Name }
                } catch {}
            }
        } catch { Write-Log "Restrict-Fehler (Sender '$name') in '$($target.Name)': $($_.Exception.Message)" }
    }

    # 2) Betreff-Stichworte -> Reports
    foreach ($kw in $reportsSubjectKeywords) {
        try {
            $filter = "@SQL=" + '"urn:schemas:httpmail:subject" LIKE ''%' + $kw + '%'''
            $items = $folder.Items.Restrict($filter)
            foreach ($it in @($items)) {
                try {
                    if ($it.Class -ne 43) { continue }
                    if ($it.FlagStatus -ne 0) { continue }
                    if ($seen.ContainsKey($it.EntryID)) { continue }
                    $seen[$it.EntryID] = $true
                    $toMove += [PSCustomObject]@{ Item = $it; Target = "Reports"; Subject = $it.Subject; From = $target.Name }
                } catch {}
            }
        } catch { Write-Log "Restrict-Fehler (Betreff '$kw') in '$($target.Name)': $($_.Exception.Message)" }
    }
}

Write-Log "Gefunden: $($toMove.Count) Mail(s) zum Verschieben nach GL/Reports."

$counts = @{}
foreach ($entry in $toMove) {
    try {
        $targetFolder = $inbox.Folders.Item($entry.Target)
        $entry.Item.Move($targetFolder) | Out-Null
        $key = "$($entry.From) -> $($entry.Target)"
        if ($counts.ContainsKey($key)) { $counts[$key]++ } else { $counts[$key] = 1 }
    } catch {
        Write-Log "FEHLER beim Verschieben '$($entry.Subject)' ($($entry.From) -> $($entry.Target)): $($_.Exception.Message)"
    }
}

foreach ($k in ($counts.Keys | Sort-Object)) {
    Write-Log "  $k : $($counts[$k]) Mail(s)"
}
Write-Log "Nachlauf beendet."
