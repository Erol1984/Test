<#
  Startet die Dashboard-Extraktoren fuer die geplante Aufgabe
  "Forecast Dashboard Aktualisieren" nacheinander und wiederholt einen
  fehlgeschlagenen Lauf.

  Hintergrund: Wird ein verpasster Lauf direkt nach dem Aufwachen aus dem
  Standby nachgeholt, scheitern Prozesse teils schon beim Start
  (0xC0000142, DLL-Initialisierung) oder brechen beim Start von Excel ab.
  Jeder Extraktor laeuft deshalb als eigener Prozess; bei Exitcode <> 0
  wird nach einer Pause erneut versucht.

  Usage:
    powershell -File run-dashboards.ps1 [-Attempts 3] [-WaitSeconds 120]
#>
param(
    [int]$Attempts = 3,
    [int]$WaitSeconds = 120
)

$logFile = Join-Path $PSScriptRoot "run-dashboards.log"
function Write-Log($msg) {
    $line = "$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')  $msg"
    Add-Content -Path $logFile -Value $line
    Write-Output $line
}

$scripts = @("extract-forecast.ps1", "extract-auftragseingang.ps1")
$failed = 0
foreach ($s in $scripts) {
    $path = Join-Path $PSScriptRoot $s
    for ($i = 1; $i -le $Attempts; $i++) {
        $p = Start-Process powershell.exe -ArgumentList "-NoProfile -WindowStyle Hidden -ExecutionPolicy Bypass -File `"$path`"" `
            -WindowStyle Hidden -Wait -PassThru
        if ($p.ExitCode -eq 0) { Write-Log "$s ok (Versuch $i)"; break }
        Write-Log ("$s fehlgeschlagen, Exitcode 0x{0:X8} (Versuch $i von $Attempts)" -f $p.ExitCode)
        if ($i -lt $Attempts) { Start-Sleep -Seconds $WaitSeconds } else { $failed++ }
    }
}
exit $failed
