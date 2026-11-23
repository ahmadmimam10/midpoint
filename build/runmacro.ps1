# Entwicklungshilfe: Makro in der geoeffneten Mappe in eigenem Prozess mit Zeitlimit ausfuehren.
param([string]$WorkbookName, [string]$Macro, [string[]]$MacroArgs = @(), [int]$TimeoutSec = 60)
$argList = ($MacroArgs | ForEach-Object { "'" + ($_ -replace "'", "''") + "'" }) -join ','
$call = if ($argList) { "`$r = `$xl.Run(`$ref, $argList)" } else { "`$r = `$xl.Run(`$ref)" }
$script = @"
[Console]::OutputEncoding = [Text.Encoding]::UTF8
`$ErrorActionPreference = 'Stop'
try {
`$xl = [Runtime.InteropServices.Marshal]::GetActiveObject('Excel.Application')
`$ref = "'$WorkbookName'!$Macro"
$call
'OK: ' + [string]`$r
} catch { 'FEHLER: ' + `$_.Exception.Message }
"@
$f = Join-Path $env:TEMP 'dbk_run_helper.ps1'
Set-Content -Path $f -Value $script -Encoding UTF8
$out = Join-Path $env:TEMP 'dbk_run_out.txt'
if (Test-Path $out) { Remove-Item $out -Force }
$p = Start-Process powershell -ArgumentList '-NoProfile','-ExecutionPolicy','Bypass','-File',$f -PassThru -WindowStyle Hidden -RedirectStandardOutput $out
if (-not $p.WaitForExit($TimeoutSec * 1000)) {
    Stop-Process -Id $p.Id -Force
    'TIMEOUT (vermutlich Dialog oder Laufzeitfehler sichtbar)'
} else {
    Get-Content $out -Encoding UTF8
}
