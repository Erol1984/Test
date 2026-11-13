<#
  Liest Metadaten aller Mails der letzten N Tage aus dem Posteingang (inkl.
  aller Unterordner) und aus "Gesendete Elemente" und schreibt sie als JSON
  fuer das Mail-Statistik-Dashboard.

  Es werden KEINE Betreffs oder Inhalte exportiert: der Gespraechsthema-Text
  wird nur als kurzer Hash gespeichert, damit Antworten ihrer Ursprungsmail
  zugeordnet werden koennen.

  Nutzt die Outlook-Table-API (folder.GetTable) statt einzelner MailItems -
  das ist um Groessenordnungen schneller.

  Usage:
    powershell -File extract-mail-stats.ps1 [-Days 90] [-OutFile path\to\data.json]
#>
param(
    [int]$Days = 90,
    [string]$OutFile = "$PSScriptRoot\..\mail-stats-data.json"
)

$ErrorActionPreference = "Stop"

$P_SMTP  = "http://schemas.microsoft.com/mapi/proptag/0x5D01001F"  # PR_SENDER_SMTP_ADDRESS
$P_TOPIC = "http://schemas.microsoft.com/mapi/proptag/0x0070001F"  # PR_CONVERSATION_TOPIC
$P_FLAG  = "http://schemas.microsoft.com/mapi/proptag/0x10900003"  # PR_FLAG_STATUS (2 = gekennzeichnet)
$P_ATT   = "http://schemas.microsoft.com/mapi/proptag/0x0E1B000B"  # PR_HASATTACH

$sha = [System.Security.Cryptography.SHA1]::Create()
function Get-TopicHash($topic) {
    $t = ([string]$topic).Trim().ToLowerInvariant()
    if (-not $t) { return "" }
    $bytes = $sha.ComputeHash([System.Text.Encoding]::UTF8.GetBytes($t))
    return ([BitConverter]::ToString($bytes, 0, 6) -replace "-", "").ToLowerInvariant()
}

function Get-Domain($smtp) {
    $s = [string]$smtp
    $i = $s.LastIndexOf("@")
    if ($i -lt 0) { return "" }
    return $s.Substring($i + 1).ToLowerInvariant()
}

$since = (Get-Date).Date.AddDays(-$Days)
$sinceUtc = $since.ToUniversalTime().ToString("yyyy-MM-dd HH:mm")

Write-Output "Verbinde mit Outlook..."
$outlook = New-Object -ComObject Outlook.Application
$ns = $outlook.GetNamespace("MAPI")
$inbox = $ns.GetDefaultFolder(6)
$sentFolder = $ns.GetDefaultFolder(5)
$me = ([string]$ns.CurrentUser.AddressEntry.GetExchangeUser().PrimarySmtpAddress).ToLowerInvariant()

function Read-Folder($folder, $dateField, $dateCol, $path) {
    $filter = "@SQL=""urn:schemas:httpmail:$dateField"" >= '$sinceUtc' AND ""http://schemas.microsoft.com/mapi/proptag/0x001A001F"" LIKE 'IPM.Note%'"
    $table = $folder.GetTable($filter)
    $table.Columns.RemoveAll()
    foreach ($c in @($dateCol, "SenderName", $P_SMTP, $P_TOPIC, "UnRead", $P_FLAG, "Importance", $P_ATT)) {
        [void]$table.Columns.Add($c)
    }
    $rows = New-Object System.Collections.Generic.List[object]
    while (-not $table.EndOfTable) {
        $r = $table.GetNextRow()
        $v = $r.GetValues()
        $smtp = ([string]$v[2]).ToLowerInvariant()
        $rows.Add([ordered]@{
            t      = ([datetime]$v[0]).ToString("yyyy-MM-ddTHH:mm")
            name   = [string]$v[1]
            domain = Get-Domain $smtp
            fromMe = ($smtp -eq $me)
            conv   = Get-TopicHash $v[3]
            unread = [bool]$v[4]
            flag   = ([int]$v[5] -eq 2)
            imp    = [int]$v[6]
            att    = [bool]$v[7]
            folder = $path
        })
    }
    return $rows
}

function Read-Tree($folder, $path) {
    $all = New-Object System.Collections.Generic.List[object]
    if ($folder.DefaultItemType -eq 0) {
        Write-Output "  $path" | Out-Host
        foreach ($row in (Read-Folder $folder "datereceived" "ReceivedTime" $path)) { $all.Add($row) }
    }
    foreach ($sub in $folder.Folders) {
        foreach ($row in (Read-Tree $sub "$path/$($sub.Name)")) { $all.Add($row) }
    }
    return $all
}

Write-Output "Lese Posteingang (seit $($since.ToString('yyyy-MM-dd')))..."
$received = Read-Tree $inbox $inbox.Name

Write-Output "Lese Gesendete Elemente..."
$sent = foreach ($row in (Read-Folder $sentFolder "date" "SentOn" $sentFolder.Name)) {
    [ordered]@{ t = $row.t; conv = $row.conv }
}

$out = [ordered]@{
    generated = (Get-Date).ToString("yyyy-MM-ddTHH:mm")
    since     = $since.ToString("yyyy-MM-dd")
    days      = $Days
    myDomain  = Get-Domain $me
    received  = @($received)
    sent      = @($sent)
}

$json = $out | ConvertTo-Json -Depth 4 -Compress
[System.IO.File]::WriteAllText([System.IO.Path]::GetFullPath($OutFile), $json, (New-Object System.Text.UTF8Encoding $false))
Write-Output "Fertig: $($received.Count) empfangen, $(@($sent).Count) gesendet -> $OutFile"

# Eigenstaendige Dashboard-Datei mit eingebetteten Daten erzeugen
$template = "$PSScriptRoot\..\mail-stats-dashboard.html"
if (Test-Path $template) {
    $standalone = "$PSScriptRoot\..\mail-stats-dashboard-standalone.html"
    $html = [System.IO.File]::ReadAllText([System.IO.Path]::GetFullPath($template), [System.Text.Encoding]::UTF8)
    $html = $html.Replace("/*__MAIL_DATA__*/null", $json.Replace("</", "<\/"))
    [System.IO.File]::WriteAllText([System.IO.Path]::GetFullPath($standalone), $html, (New-Object System.Text.UTF8Encoding $false))
    Write-Output "Dashboard -> $standalone"
}
