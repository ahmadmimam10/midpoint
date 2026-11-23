# Entwicklungshilfe: Lasttestdatei bauen, Kopie oeffnen (Startzeit messen) und mSelfTest.Measure ausfuehren.
# Ergebnis: pruefung\leistung\Leistung_Ergebnis.txt
param([switch]$NoBuild)
$root = Split-Path $PSScriptRoot -Parent
Get-Process EXCEL -ErrorAction SilentlyContinue | Stop-Process -Force
Start-Sleep 2
if (-not $NoBuild) {
    $sw = [Diagnostics.Stopwatch]::StartNew()
    & powershell -ExecutionPolicy Bypass -File (Join-Path $PSScriptRoot 'build.ps1') -Target stress | Select-Object -Last 1
    "Aufbau der Lasttestdatei: {0:0} s" -f $sw.Elapsed.TotalSeconds
}
$dir = Join-Path $root 'pruefung\leistung'
New-Item -ItemType Directory -Force $dir | Out-Null
$copy = Join-Path $dir 'Lasttest.xlsm'
Copy-Item (Join-Path $root 'dist\MidPoint_Lasttest.xlsm') $copy -Force
$size = (Get-Item $copy).Length / 1MB
$sw = [Diagnostics.Stopwatch]::StartNew()
& powershell -ExecutionPolicy Bypass -File (Join-Path $PSScriptRoot 'devopen.ps1') -Path $copy | Out-Null
# warten, bis die Startseite gerendert ist (Excel antwortet wieder)
$xl = [Runtime.InteropServices.Marshal]::GetActiveObject('Excel.Application')
for ($i = 0; $i -lt 120; $i++) { try { if ($xl.Ready) { break } } catch {} ; Start-Sleep -Milliseconds 500 }
$open = $sw.Elapsed.TotalSeconds
$o = & powershell -ExecutionPolicy Bypass -File (Join-Path $PSScriptRoot 'runmacro.ps1') -WorkbookName 'Lasttest.xlsm' -Macro 'mSelfTest.Measure' -TimeoutSec 900
$res = Join-Path $dir 'Leistung_Ergebnis.txt'
$lines = @("Dateigröße: {0:0.0} MB" -f $size, ("Datei öffnen bis bedienbar (inkl. Excel-Start): {0:0.0} s" -f $open), "")
if (Test-Path $res) { $lines += Get-Content $res -Encoding Default } else { $lines += $o }
$lines | Set-Content $res -Encoding UTF8
$lines
