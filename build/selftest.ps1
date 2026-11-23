# Entwicklungshilfe: Demodatei bauen, Kopie oeffnen und automatisierte Pruefungen ausfuehren.
# Ergebnis: pruefung\selbsttest\Selbsttest_Ergebnis.txt (plus T17/T19)
param([switch]$NoBuild)
$root = Split-Path $PSScriptRoot -Parent
function Run-M([string]$m, [string[]]$a = @()) {
    if ($a.Count -gt 0) { $o = & powershell -ExecutionPolicy Bypass -File (Join-Path $PSScriptRoot 'runmacro.ps1') -WorkbookName 'Selbsttest.xlsm' -Macro $m -MacroArgs $a -TimeoutSec 240 }
    else { $o = & powershell -ExecutionPolicy Bypass -File (Join-Path $PSScriptRoot 'runmacro.ps1') -WorkbookName 'Selbsttest.xlsm' -Macro $m -TimeoutSec 240 }
    return ($o -join "`n")
}
Get-Process EXCEL -ErrorAction SilentlyContinue | Stop-Process -Force
Start-Sleep 2
if (-not $NoBuild) { & powershell -ExecutionPolicy Bypass -File (Join-Path $PSScriptRoot 'build.ps1') -Target demo | Select-Object -Last 1 }
$dir = Join-Path $root 'pruefung\selbsttest'
if (Test-Path $dir) { Remove-Item $dir -Recurse -Force }
New-Item -ItemType Directory -Force $dir | Out-Null
$copy = Join-Path $dir 'Selbsttest.xlsm'
Copy-Item (Join-Path $root 'dist\MidPoint_Demo.xlsm') $copy -Force
& powershell -ExecutionPolicy Bypass -File (Join-Path $PSScriptRoot 'devopen.ps1') -Path $copy | Out-Null
Start-Sleep 5
Run-M 'mSelfTest.RunSelfTest' | Out-Null
$res = Join-Path $dir 'Selbsttest_Ergebnis.txt'
$lines = @()
if (Test-Path $res) { $lines += Get-Content $res -Encoding Default } else { $lines += 'Kein Ergebnis (Dialog oder Laufzeitfehler sichtbar?)' }

# T19 Sicherung und Wiederherstellung
$lines += ''
$lines += (Run-M 'mSelfTest.BackupRestoreTest') -replace '^OK: ', ''

# T17 Speichern, Excel vollstaendig schliessen, erneut oeffnen
$before = (Run-M 'mSelfTest.SnapshotToFile' @('snap_vor.txt'))
$save = Run-M 'mSelfTest.SaveForTest'
$xl = [Runtime.InteropServices.Marshal]::GetActiveObject('Excel.Application')
$xl.DisplayAlerts = $false
foreach ($w in @($xl.Workbooks)) { $w.Close($false) }
$xl.Quit()
[void][Runtime.InteropServices.Marshal]::ReleaseComObject($xl)
$xl = $null; [GC]::Collect(); [GC]::WaitForPendingFinalizers()
# Excel Zeit zum regulaeren Beenden geben (hartes Beenden wuerde beim naechsten Oeffnen
# einen "schwerwiegenden Fehler" melden und den Test blockieren)
for ($i = 0; $i -lt 40; $i++) { if (-not (Get-Process EXCEL -ErrorAction SilentlyContinue)) { break }; Start-Sleep -Milliseconds 500 }
Get-Process EXCEL -ErrorAction SilentlyContinue | Stop-Process -Force
Start-Sleep 2
& powershell -ExecutionPolicy Bypass -File (Join-Path $PSScriptRoot 'devopen.ps1') -Path $copy | Out-Null
Start-Sleep 6
$after = (Run-M 'mSelfTest.SnapshotToFile' @('snap_nach.txt'))
$a = Get-Content (Join-Path $dir 'snap_vor.txt') -Encoding Default
$b = Get-Content (Join-Path $dir 'snap_nach.txt') -Encoding Default
$lines += ''
$lines += ("T17`t" + $(if ($a -eq $b -and $save -match 'OK: GESPEICHERT') { 'BESTANDEN' } else { 'NICHT BESTANDEN' }) + "`tSpeichern ($save), Excel beendet, erneut geoeffnet; Fingerabdruck identisch: " + ($a -eq $b))
$lines += "vorher:  $a"
$lines += "nachher: $b"

# T01 Masterdatei kopieren, umbenennen, unabhaengig oeffnen
Get-Process EXCEL -ErrorAction SilentlyContinue | Stop-Process -Force
Start-Sleep 2
$m1 = Join-Path $dir 'Meine_Materialverwaltung.xlsm'
Copy-Item (Join-Path $root 'dist\MidPoint_Master.xlsm') $m1 -Force
& powershell -ExecutionPolicy Bypass -File (Join-Path $PSScriptRoot 'devopen.ps1') -Path $m1 | Out-Null
Start-Sleep 6
$pr = & powershell -ExecutionPolicy Bypass -File (Join-Path $PSScriptRoot 'runmacro.ps1') -WorkbookName 'Meine_Materialverwaltung.xlsm' -Macro 'mSelfTest.Probe' -TimeoutSec 60
$lines += ''
$lines += ("T01`t" + $(if ($pr -match 'gestartet=Wahr' -and $pr -match 'schreiben erlaubt=Wahr') { 'BESTANDEN' } else { 'NICHT BESTANDEN' }) + "`tKopie der Masterdatei unter neuem Namen geoeffnet: $pr")

# T25 Leere Datei: Tabelle anlegen, einzige Tabelle loeschen, erneut anlegen (dreimal)
$cd = & powershell -ExecutionPolicy Bypass -File (Join-Path $PSScriptRoot 'runmacro.ps1') -WorkbookName 'Meine_Materialverwaltung.xlsm' -Macro 'mSelfTest.CreateDeleteCreateTest' -TimeoutSec 120
$lines += ("T25`t" + $(if ($cd -match 'OK: OK') { 'BESTANDEN' } else { 'NICHT BESTANDEN' }) + "`tLeere Masterdatei: Tabelle angelegt, als einzige Tabelle geloescht, erneut angelegt: $cd")

# T18 Nur-Lesen: zweite Instanz oeffnet dieselbe Datei schreibgeschuetzt
$xl2 = New-Object -ComObject Excel.Application
$xl2.Visible = $true
$xl2.AutomationSecurity = 1
$wb2 = $xl2.Workbooks.Open($copy, 0, $true)
Start-Sleep 6
try { $ro = $xl2.Run("'Selbsttest.xlsm'!mSelfTest.Probe") } catch { $ro = 'FEHLER: ' + $_.Exception.Message }
$lines += ("T18`t" + $(if ($ro -match 'lesemodus=Wahr' -and $ro -match 'schreiben erlaubt=Falsch') { 'BESTANDEN' } else { 'NICHT BESTANDEN' }) + "`tDatei schreibgeschuetzt geoeffnet: $ro")
$xl2.DisplayAlerts = $false
$wb2.Close($false)
$xl2.Quit()
[void][Runtime.InteropServices.Marshal]::ReleaseComObject($xl2)
Start-Sleep 2

# T24 Makros deaktiviert
Get-Process EXCEL -ErrorAction SilentlyContinue | Stop-Process -Force
Start-Sleep 2
$xl3 = New-Object -ComObject Excel.Application
$xl3.Visible = $true
$xl3.AutomationSecurity = 3
$wb3 = $xl3.Workbooks.Open($m1)
Start-Sleep 3
$vis = @(); foreach ($s in $wb3.Worksheets) { if ($s.Visible -eq -1) { $vis += $s.Name } }
$hint = [string]$wb3.Worksheets.Item('Hinweis').Range('B4').Value2
$lines += ("T24`t" + $(if ($vis.Count -eq 1 -and $vis[0] -eq 'Hinweis') { 'BESTANDEN' } else { 'NICHT BESTANDEN' }) + "`tMakros deaktiviert: sichtbar nur [" + ($vis -join ', ') + "]: " + $hint)
$xl3.DisplayAlerts = $false
$wb3.Close($false)
$xl3.Quit()
[void][Runtime.InteropServices.Marshal]::ReleaseComObject($xl3)
Get-Process EXCEL -ErrorAction SilentlyContinue | Stop-Process -Force
$lines | Set-Content $res -Encoding UTF8
$lines
