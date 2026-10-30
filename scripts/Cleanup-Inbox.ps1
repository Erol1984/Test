# Verschiebt gelesene, nicht gekennzeichnete Mails aus dem Posteingang in thematische
# Zielordner, basierend auf der Absender-Domain (Zuordnung abgeleitet aus der bereits
# bestehenden Einsortierung in diesen Ordnern).

$ErrorActionPreference = "Stop"
$logFile = Join-Path $PSScriptRoot "Cleanup-Inbox.log"

function Write-Log($msg) {
    $line = "$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')  $msg"
    Add-Content -Path $logFile -Value $line
    Write-Output $line
}

# Domain -> Zielordner (Unterordner von Posteingang). Nur eindeutige Domains,
# die in genau einem Zielordner vorkommen.
$domainMap = @{
    "vogel-communications.ch" = "SM"
    "aitraining.institute"    = "SM"
    "ayes.cz"                 = "SM"
    "timgelhausen.de"         = "SM"
    "netzwerk-experten.com"   = "SM"
    "linkedin.com"            = "SM"
    "mail.xing.com"           = "SM"

    "eu.temuemail.com"        = "privat"
    "backslash.ch"            = "privat"
    "messaging.microsoft.com" = "privat"
    "tkb.ch"                  = "privat"
    "volume-trader.de"        = "privat"
    "testorush.eu"            = "privat"
    "mail.maraoster.com"      = "privat"
    "bluewin.ch"              = "privat"
    "email.apple.com"         = "privat"
    "news.vicampo.de"         = "privat"
    "revolut.com"             = "privat"
    "update.just-eat.ch"      = "privat"
    "news.twint.ch"           = "privat"
    "crm.topcc.ch"            = "privat"
    "e.sixt.ch"               = "privat"
    "eu-shop.temuemail.com"   = "privat"
    "viseca.ch"               = "privat"
    "threema.ch"              = "privat"
    "e.trabzonspor.com.tr"    = "privat"
    "uber.com"                = "privat"
    "spotify.com"             = "privat"
    "fussballtraining.com"    = "privat"
    "the-naturalcoach.com"    = "privat"
    "stayfriends.ch"          = "privat"
    "simple-sixpack.com"      = "privat"
    "baspo.admin.ch"          = "privat"

    "nl.chmedia.ch"              = "Newsletter"
    "mailings.nzz.ch"            = "Newsletter"
    "news.simplify.de"           = "Newsletter"
    "packaging-journal.de"       = "Newsletter"
    "replyto.neue-verpackung.de" = "Newsletter"
    "news.impulse.de"            = "Newsletter"
    "info.kununu.com"            = "Newsletter"
    "ptspaper.de"                = "Newsletter"
    "it-production.com"         = "Newsletter"
    "display.de"                 = "Newsletter"
    "mc.technischefachmedien.de" = "Newsletter"
    "account.chmedia.ch"         = "Newsletter"
    "em.linkedin.com"            = "Newsletter"

    "fc-wb.ch"    = "Verein FCWB"
    "veo.co"      = "Verein FCWB"
    "football.ch" = "Verein FCWB"

    "oversys.es"       = "Machine supplier"
    "bobst.com"        = "Machine supplier"
    "minda.com"        = "Machine supplier"
    "bhs-world.com"    = "Machine supplier"
    "bahmueller.de"    = "Machine supplier"
    "koenig-bauer.com" = "Machine supplier"
    "mybobst.com"      = "Machine supplier"
}

# Domains, die in mehreren Zielordnern vorkommen -> bewusst NICHT automatisch verschieben
$ambiguousDomains = @()

# Absender-E-Mail -> Zielordner (GL = Geschaeftsleitung, Reports = Absender fuer Berichte)
$senderMap = @{
    "sarah.model@modelgroup.com"        = "GL"
    "david.model@modelgroup.com"        = "GL"
    "michael.uebersax@modelgroup.com"   = "GL"
    "matthias.hugentobler@modelgroup.com" = "GL"
    "matthias.hess@modelgroup.com"      = "GL"
    "michael.wilms@modelgroup.com"      = "GL"
    "guenter.hoeller@modelgroup.com"    = "GL"
    "franzjoseph.nicolaus@modelgroup.com" = "GL"

    "fabian.huber@modelgroup.com" = "Reports"
    "loick.berset@modelgroup.com" = "Reports"
}

# Betreff-Stichworte -> "Reports"-Ordner (case-insensitive Teilstring-Suche)
$reportsSubjectKeywords = @(
    "Monatsabschluss",
    "Quartalabschluss",
    "Quartalsabschluss",
    "Jahresabschluss",
    "Produktionsrapport",
    "Budget"
)

# MBB (Model Bad Bentheim): Absender mit City = "Bad Bentheim" im Exchange-Adressbuch ->
# MBB\Informationen; Systemabsender mit "MBB"/"Bad Bentheim" im Betreff -> MBB\Reports MBB.
# Abweichend von den Regeln oben gilt fuer MBB: auch ungelesene Mails, keine 24h-Wartezeit
# (gekennzeichnete Mails bleiben weiterhin liegen). Hat Vorrang vor Domain-/Absender-/Betreff-Regeln.
$mbbCity = "Bad Bentheim"
$mbbReportSenders = @("Formularweb MGM", "Model Group BOEP", "fernwartung@es-ha-es.de")
$mbbSubjectPattern = "MBB|Bad Bentheim"
$mbbCityCache = @{}

function Get-SenderCity($mailItem) {
    $key = "$($mailItem.SenderEmailType)|$($mailItem.SenderEmailAddress)|$($mailItem.SenderName)"
    if ($mbbCityCache.ContainsKey($key)) { return $mbbCityCache[$key] }
    $city = $null
    try {
        if ($mailItem.SenderEmailType -eq "EX") {
            $u = $mailItem.Sender.GetExchangeUser()
            if ($u) { $city = $u.City }
        }
    } catch {}
    $mbbCityCache[$key] = $city
    return $city
}

function Get-TargetFolder($inboxFolder, $path) {
    $f = $inboxFolder
    foreach ($part in $path.Split("\")) { $f = $f.Folders.Item($part) }
    return $f
}

function Get-SmtpAddress($mailItem) {
    try {
        if ($mailItem.SenderEmailType -eq "EX") {
            $exUser = $mailItem.Sender.GetExchangeUser()
            if ($exUser) { return $exUser.PrimarySmtpAddress }
            return $null
        } else {
            return $mailItem.SenderEmailAddress
        }
    } catch {
        return $null
    }
}

try {
    if (-not (Get-Process -Name "OUTLOOK" -ErrorAction SilentlyContinue)) {
        Write-Log "Outlook laeuft nicht, starte es..."
        $exe = Get-ChildItem "C:\Program Files\Microsoft Office","C:\Program Files (x86)\Microsoft Office" -Recurse -Filter "OUTLOOK.EXE" -ErrorAction SilentlyContinue | Select-Object -First 1
        if ($exe) {
            Start-Process $exe.FullName
            Start-Sleep -Seconds 20
        }
    }
    $outlook = New-Object -ComObject Outlook.Application
    $namespace = $outlook.GetNamespace("MAPI")
    $inbox = $namespace.GetDefaultFolder(6)
} catch {
    Write-Log "FEHLER: Outlook COM nicht erreichbar - $($_.Exception.Message)"
    exit 1
}

$toMove = @()
$mbbIds = @{}

# MBB-Durchlauf ueber alle Mails im Posteingang (gelesen und ungelesen)
foreach ($it in @($inbox.Items)) {
    try {
        if ($it.Class -ne 43) { continue }
        if ($it.FlagStatus -ne 0) { continue }
        $subj = [string]$it.Subject
        $isSystem = ($mbbReportSenders -contains $it.SenderName) -and ($subj -match $mbbSubjectPattern)
        $isMbbPerson = ((Get-SenderCity $it) -eq $mbbCity)
        if (-not ($isSystem -or $isMbbPerson)) { continue }
        $t = if ($isSystem) { "MBB\Reports MBB" } else { "MBB\Informationen" }
        $mbbIds[$it.EntryID] = $true
        $toMove += [PSCustomObject]@{ Item = $it; Target = $t; Subject = $subj }
    } catch {}
}

$candidates = $inbox.Items.Restrict("[Unread] = False")

foreach ($it in @($candidates)) {
    try {
        if ($it.Class -ne 43) { continue }          # nur Mail-Items (olMail)
        if ($mbbIds.ContainsKey($it.EntryID)) { continue }   # bereits MBB zugeordnet
        if ($it.FlagStatus -ne 0) { continue }        # 0 = olNoFlag = nicht gekennzeichnet
        # Outlook speichert keinen expliziten "gelesen am"-Zeitstempel; LastModificationTime
        # aendert sich beim Setzen auf gelesen und dient hier als Naeherung dafuer.
        if (((Get-Date) - $it.LastModificationTime).TotalHours -lt 24) { continue }
        $addr = Get-SmtpAddress $it
        if (-not $addr -or $addr -notmatch "@") { continue }
        $addrLower = $addr.ToLower()
        $domain = $addrLower.Split("@")[1]

        $target = $null
        if (-not ($ambiguousDomains -contains $domain) -and $domainMap.ContainsKey($domain)) {
            $target = $domainMap[$domain]
        } elseif ($senderMap.ContainsKey($addrLower)) {
            $target = $senderMap[$addrLower]
        } else {
            foreach ($kw in $reportsSubjectKeywords) {
                if ($it.Subject -and $it.Subject -like "*$kw*") { $target = "Reports"; break }
            }
        }

        if ($target) {
            $toMove += [PSCustomObject]@{ Item = $it; Target = $target; Subject = $it.Subject }
        }
    } catch {}
}

Write-Log "Gefunden: $($toMove.Count) gelesene, nicht gekennzeichnete Mail(s) zum Verschieben."

$counts = @{}
foreach ($entry in $toMove) {
    try {
        $targetFolder = Get-TargetFolder $inbox $entry.Target
        $entry.Item.Move($targetFolder) | Out-Null
        if ($counts.ContainsKey($entry.Target)) { $counts[$entry.Target]++ } else { $counts[$entry.Target] = 1 }
    } catch {
        Write-Log "FEHLER beim Verschieben '$($entry.Subject)' -> $($entry.Target): $($_.Exception.Message)"
    }
}

foreach ($k in $counts.Keys) {
    Write-Log "  -> $k : $($counts[$k]) Mail(s) verschoben"
}
Write-Log "Lauf beendet."
