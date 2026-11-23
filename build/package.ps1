# Entwicklungshilfe: Lieferpaket paket\MidPoint_<Version> (Ordner und ZIP) zusammenstellen.
# Gliederung seit 1.8.1: Masterdatei oben; Vorstellung (Demodatei, Video 1.1); Lernen (Lernportal, Lernvideos, Mitbauen aus video\out\lern);
# Dokumentation; IT und Datenschutz (Video fuer IT und Datenschutz, Technische Dokumentation, Pruefung, Quellcode).
param([string]$Version = "1.8.2")
$root = Split-Path $PSScriptRoot -Parent
$name = "MidPoint_$Version"
$out = Join-Path $root "paket\$name"
if (Test-Path $out) { Remove-Item $out -Recurse -Force }
$dV = Join-Path $out 'Vorstellung'; $dL = Join-Path $out 'Lernen'; $dD = Join-Path $out 'Dokumentation'; $dI = Join-Path $out 'IT und Datenschutz'
$d3 = Join-Path $dI 'Pruefung'; $d4 = Join-Path $dI 'Quellcode'
foreach ($d in @($dV, $dL, $dD, $dI, $d3, $d4)) { New-Item -ItemType Directory -Force $d | Out-Null }
Copy-Item (Join-Path $root 'doku\Liesmich.md') $out
Copy-Item (Join-Path $root 'LICENSE') (Join-Path $out 'LICENSE.txt')
Copy-Item (Join-Path $root 'dist\MidPoint_Master.xlsm') $out
Copy-Item (Join-Path $root 'dist\MidPoint_Demo.xlsm') $dV
# Lernen: Lernportal mit den 17 Lernvideos (und Untertiteln) und der Ordner Mitbauen
$lern = Join-Path $root 'video\out\lern'
$vids = @(Get-ChildItem (Join-Path $lern 'Lernvideos') -Filter *.mp4)
if ($vids.Count -ne 17) { throw "Paket gestoppt: 17 Lernvideos erwartet, gefunden $($vids.Count)" }
Copy-Item (Join-Path $lern 'Lernportal.html') $dL
Copy-Item (Join-Path $lern 'Lernvideos\1.1_Warum_MidPoint.*') $dV
New-Item -ItemType Directory -Force (Join-Path $dL 'Lernvideos') | Out-Null
Copy-Item (Join-Path $lern 'Lernvideos\*') (Join-Path $dL 'Lernvideos') -Include *.mp4, *.vtt
Copy-Item (Join-Path $lern 'Mitbauen') $dL -Recurse
foreach ($f in @('Hilfe.md', 'Kurzanleitung_In_zehn_Schritten.md', 'Versionshinweise.md')) { Copy-Item (Join-Path $root "doku\$f") $dD }
# IT und Datenschutz: das Video dazu, Technische Dokumentation, Pruefberichte, Quellcode mit Build-Anleitung
Copy-Item (Join-Path $lern 'Lernvideos\Extra_Fuer_IT_und_Datenschutz.*') $dI
Copy-Item (Join-Path $root 'doku\Technische_Dokumentation.md') $dI
foreach ($f in @('Testmatrix.md', 'Leistungsbericht.md')) { Copy-Item (Join-Path $root "doku\$f") $d3 }
Copy-Item (Join-Path $root 'pruefung\selbsttest\Selbsttest_Ergebnis.txt') $d3
Copy-Item (Join-Path $root 'pruefung\leistung\Leistung_Ergebnis.txt') $d3
Copy-Item (Join-Path $root 'pruefung\sicherheit\Sicherheitspruefung.md') $d3
Copy-Item (Join-Path $root 'pruefung\listimport\Listimport_Ergebnis.txt') $d3
Copy-Item (Join-Path $root 'pruefung\listimport\Listimport_Tabellen.txt') $d3
Copy-Item (Join-Path $root 'dist\MidPoint_Lasttest.xlsm') $d3
# Testlisten fuer "Liste uebernehmen" (13 Verwaltungslisten, 2ask-Nachbau und echte 2ask-Exporte)
$tl = Join-Path $d3 'Testlisten'
New-Item -ItemType Directory -Force $tl | Out-Null
Copy-Item (Join-Path $root 'testdaten\importlisten\*') $tl -Recurse
# Bildschirmaufnahmen: aktuelle Version und unveraenderte Ansichten der Vorversionen
$shots = Join-Path $d3 'screenshots'
New-Item -ItemType Directory -Force $shots | Out-Null
Copy-Item (Join-Path $root 'pruefung\screenshots_180\*.jpg') $shots
$o171 = Join-Path $shots 'aus_Version_1.7.1'
New-Item -ItemType Directory -Force $o171 | Out-Null
Copy-Item (Join-Path $root 'pruefung\screenshots_171\*.jpg') $o171
$o170 = Join-Path $shots 'aus_Version_1.7.0'
New-Item -ItemType Directory -Force $o170 | Out-Null
Copy-Item (Join-Path $root 'pruefung\screenshots_170\*.jpg') $o170
$o161 = Join-Path $shots 'aus_Version_1.6.1'
New-Item -ItemType Directory -Force $o161 | Out-Null
Copy-Item (Join-Path $root 'pruefung\screenshots_161\*.jpg') $o161
$o160 = Join-Path $shots 'aus_Version_1.6.0'
New-Item -ItemType Directory -Force $o160 | Out-Null
Copy-Item (Join-Path $root 'pruefung\screenshots_160\*.jpg') $o160
$o150 = Join-Path $shots 'aus_Version_1.5.0'
New-Item -ItemType Directory -Force $o150 | Out-Null
Copy-Item (Join-Path $root 'pruefung\screenshots_150\*.jpg') $o150
$o140 = Join-Path $shots 'aus_Version_1.4.0'
New-Item -ItemType Directory -Force $o140 | Out-Null
Copy-Item (Join-Path $root 'pruefung\screenshots_140\*.jpg') $o140
$o130 = Join-Path $shots 'aus_Version_1.3.0'
New-Item -ItemType Directory -Force $o130 | Out-Null
Copy-Item (Join-Path $root 'pruefung\screenshots_130\*.jpg') $o130
$o120 = Join-Path $shots 'aus_Version_1.2.0'
New-Item -ItemType Directory -Force $o120 | Out-Null
Copy-Item (Join-Path $root 'pruefung\screenshots_120\*.jpg') $o120
$old = Join-Path $root 'paket\Excel_Datenbankbaukasten_1.1.0\03_Pruefung\screenshots'
if (Test-Path $old) {
    $o = Join-Path $shots 'aus_Version_1.1.0'
    New-Item -ItemType Directory -Force $o | Out-Null
    foreach ($f in @('03_Board.jpg', '04_Kalender.jpg', '05_Projektleiste_Gantt.jpg', '07_Dashboard_Details_Drilldown.jpg', '14_Filterdialog.jpg',
                     '16_Hilfe.jpg', '17_Technische_Dokumentation.jpg', '22_Ruecklauf_Vorschau.jpg', '24_Board_nach_Verschieben.jpg', '25_Import_Fehlerbericht.jpg')) {
        if (Test-Path (Join-Path $old $f)) { Copy-Item (Join-Path $old $f) $o }
    }
}
Copy-Item (Join-Path $root 'src') (Join-Path $d4 'src') -Recurse
New-Item -ItemType Directory -Force (Join-Path $d4 'build') | Out-Null
Copy-Item (Join-Path $root 'build\*.ps1') (Join-Path $d4 'build')
Copy-Item (Join-Path $root 'doku\Build_Anleitung.md') $d4
# Spuren des Entwicklungsrechners: Excel-Dateien bereinigen und pruefen, Textdateien pruefen; ein Rest stoppt das Paket
& powershell -ExecutionPolicy Bypass -File (Join-Path $PSScriptRoot 'entspuren.ps1') -Path $out
if ($LASTEXITCODE -ne 0) { throw 'Paket gestoppt: Spuren des Entwicklungsrechners in Excel-Dateien' }
$pat = '(?i)(?<![a-z])' + [regex]::Escape($env:USERNAME) + '(?![a-z])|[a-z]:\\users\\'
$spur = Get-ChildItem $out -Recurse -File |
    Where-Object { $_.Extension -in '.md', '.txt', '.csv', '.ps1', '.bas', '.cls', '.vba', '.frm', '.html', '.json' } |
    Select-String -Pattern $pat
if ($spur) {
    $spur | ForEach-Object { '  SPUR: ' + $_.Filename + ': ' + $_.Line.Trim() }
    throw 'Paket gestoppt: Spuren des Entwicklungsrechners in Textdateien'
}
'Textdateien ohne Pfade und Benutzernamen des Entwicklungsrechners'
$zip = Join-Path $root "paket\$name.zip"
if (Test-Path $zip) { Remove-Item $zip -Force }
Compress-Archive -Path $out -DestinationPath $zip
"Paket: $out"
"ZIP: $zip ({0:0.0} MB)" -f ((Get-Item $zip).Length / 1MB)
