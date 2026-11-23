# Entwicklungshilfe: "Liste uebernehmen" mit allen Testlisten pruefen.
# Kopiert die Masterdatei, oeffnet sie in einer eigenen sichtbaren Excel-Instanz und ruft je Testliste
# mListImport.LiTestCase auf (Tabellen, Rueckfragen, Felder, Pruefung und Uebernahme in eine neue Tabelle).
# Danach werden alle uebernommenen Tabellen mit Feldern, Optionen und ersten Eintraegen ausgegeben.
# Ergebnis: pruefung\listimport\Listimport_Ergebnis.txt und Listimport_Tabellen.txt (UTF-8)
# Bricht ab, wenn bereits Excel laeuft (fremde Sitzungen werden nie beendet).
# -Set neu: zweiter Satz mit den "gewachsenen" Listen 14 bis 16, vier Listen von Behoerden (testdaten\online,
# nicht im Lieferpaket) und den Exporten der zweiten 2ask-Befragung; Ergebnis dann in pruefung\listimport_neu
param([switch]$NoImport, [ValidateSet('standard', 'neu')][string]$Set = 'standard')
$root = Split-Path $PSScriptRoot -Parent
if (Get-Process EXCEL -ErrorAction SilentlyContinue) { 'ABBRUCH: Excel ist bereits geoeffnet. Bitte zuerst alle Excel-Fenster schliessen.'; return }
$dir = Join-Path $root $(if ($Set -eq 'neu') { 'pruefung\listimport_neu' } else { 'pruefung\listimport' })
if (Test-Path $dir) { Remove-Item $dir -Recurse -Force }
New-Item -ItemType Directory -Force $dir | Out-Null
# Grosse Liste fuer die Laufzeitmessung: 5.000 Zeilen, 12 Spalten (eigene, unsichtbare Excel-Instanz)
$big = Join-Path $dir 'Grosse_Liste_5000.xlsx'
$EN = [Globalization.CultureInfo]'en-US'
if ($Set -eq 'standard') {
$xg = New-Object -ComObject Excel.Application
$xg.Visible = $false; $xg.DisplayAlerts = $false
try {
    $wbg = $xg.Workbooks.Add()
    $wsg = $wbg.Worksheets.Item(1)
    $n = 5000
    $data = New-Object 'object[,]' ($n + 1), 12
    $hdr = @('Nr.', 'Name', 'Vorname', 'Dienststelle', 'Eingang', 'Betrag', 'Status', 'Teilnahme', 'E-Mail', 'PLZ', 'Ort', 'Bemerkung')
    for ($c = 0; $c -lt 12; $c++) { $data[0, $c] = $hdr[$c] }
    $amt = @('Jugendamt', 'Sozialamt', 'Amt fuer Schulen', 'Gesundheitsamt', 'Kita Sonnenblume', 'Kita Regenbogen')
    $st = @('offen', 'in Arbeit', 'erledigt', 'zurueckgestellt')
    $orte = @('Musterstadt', 'Nordheim', 'Suedstadt')
    $rnd = New-Object System.Random 7
    $base = (Get-Date '2025-01-01').ToOADate()
    for ($r = 1; $r -le $n; $r++) {
        $data[$r, 0] = $r
        $data[$r, 1] = 'Name' + $r
        $data[$r, 2] = 'Vorname' + ($r % 300)
        $data[$r, 3] = $amt[$rnd.Next($amt.Count)]
        $data[$r, 4] = $base + $rnd.Next(365)
        $data[$r, 5] = [math]::Round($rnd.NextDouble() * 5000, 2)
        $data[$r, 6] = $st[$rnd.Next($st.Count)]
        if ($rnd.Next(2) -eq 1) { $data[$r, 7] = 'x' }
        $data[$r, 8] = 'person' + $r + '@stadt.example'
        $data[$r, 9] = 55000 + $rnd.Next(999)
        $data[$r, 10] = $orte[$rnd.Next($orte.Count)]
        if ($rnd.Next(10) -eq 0) { $data[$r, 11] = 'Rueckfrage zu Vorgang ' + $r }
    }
    $wsg.Range($wsg.Cells.Item(1, 1), $wsg.Cells.Item($n + 1, 12)).Value2 = $data
    [void][System.__ComObject].InvokeMember('NumberFormat', [Reflection.BindingFlags]::SetProperty, $null, $wsg.Range('E2:E' + ($n + 1)), @('dd.mm.yyyy'), $null, $EN, $null)
    $wbg.SaveAs($big, 51)
    $wbg.Close($false)
} finally {
    $xg.Quit()
    # alle Verweise freigeben, sonst bleibt die unsichtbare Instanz am Leben und nimmt spaeter die Makroaufrufe an
    foreach ($o in @($wsg, $wbg, $xg)) { if ($null -ne $o) { [void][Runtime.InteropServices.Marshal]::ReleaseComObject($o) } }
    $wsg = $null; $wbg = $null; $xg = $null
    [GC]::Collect(); [GC]::WaitForPendingFinalizers()
}
for ($i = 0; $i -lt 30; $i++) { if (-not (Get-Process EXCEL -ErrorAction SilentlyContinue)) { break }; Start-Sleep 1 }
if (Get-Process EXCEL -ErrorAction SilentlyContinue) { 'ABBRUCH: Die Excel-Instanz fuer die grosse Testliste hat sich nicht beendet.'; return }
}
# eindeutiger Name je Lauf (Excel merkt sich abgebrochene Dateien unter ihrem Pfad und fragt sonst beim Oeffnen nach)
$wbName = 'Listimport_Test_' + (Get-Date -Format 'HHmmss') + '.xlsm'
$copy = Join-Path $dir $wbName
Copy-Item (Join-Path $root 'dist\MidPoint_Master.xlsm') $copy -Force
& powershell -ExecutionPolicy Bypass -File (Join-Path $PSScriptRoot 'devopen.ps1') -Path $copy | Out-Null
# warten, bis die Testkopie in der laufenden Instanz offen und bereit ist
$ready = $false
for ($i = 0; $i -lt 60 -and -not $ready; $i++) {
    try {
        $xa = [Runtime.InteropServices.Marshal]::GetActiveObject('Excel.Application')
        foreach ($w in @($xa.Workbooks)) { if ($w.Name -eq $wbName -and $xa.Ready) { $ready = $true } }
        [void][Runtime.InteropServices.Marshal]::ReleaseComObject($xa)
    } catch { }
    if (-not $ready) { Start-Sleep 1 }
}
if (-not $ready) { 'ABBRUCH: Die Testkopie ist nach 60 s nicht geoeffnet (Dialog sichtbar?).'; return }
Start-Sleep 2
$td = Join-Path $root 'testdaten\importlisten'
# Datei, Tabelle (0 = Vorschlag)
$cases = @(
    @('01_Teilnehmerliste_Fortbildung.xlsx', 0),
    @('02_Fristenliste_Wiedervorlage.xlsx', 0),
    @('03_Haushalt_Mittelabfluss.xlsx', 0),
    @('04_Belegung_Kitas_Monate.xlsx', 0),
    @('05_Adressliste_Ansprechpersonen.xlsx', 0),
    @('06_Inventar_IT.xlsx', 0),
    @('07_Massnahmenplan.xlsx', 0),
    @('08_Bedarfsmeldung_je_Kita.xlsx', 0),
    @('09_Projekte_und_Aufgaben.xlsx', 1),
    @('09_Projekte_und_Aufgaben.xlsx', 2),
    @('10_Beschlusskontrolle.xlsx', 0),
    @('11_Fachverfahren_Export.csv', 0),
    @('12_Antrag_Formular.xlsx', 1),
    @('12_Antrag_Formular.xlsx', 2),
    @('13_Raumbelegung_Wochenplan.xlsx', 0),
    @('2ask_Export_kodiert.xlsx', 0),
    @('2ask_echt\Original_XLSX_kodiert_mit_Codeplan.xlsx', 0),
    @('2ask_echt\Original_XLSX_kodiert_Codeplan_Umkodierung_99_98_97.xlsx', 0),
    @('2ask_Export_unkodiert.xlsx', 0),
    @($big, 0)
)
if ($Set -eq 'neu') {
    $on = Join-Path $root 'testdaten\online'
    $k = Join-Path $root 'testdaten\importlisten\2ask_komplex'
    $cases = @(
        @('14_Telefonliste_Amt_FINAL_neu.xlsx', 0),
        @('15_Veranstaltungsplanung_2026_Kopie.xlsx', 0),
        @('16_Fortbildungsbedarf_Rueckmeldungen.xlsx', 0),
        @((Join-Path $on '01_Destatis_Gemeindeverzeichnis\Destatis_Gemeindeverzeichnis_30092026.xlsx'), 0),
        @((Join-Path $on '02_Statistik_RLP_Schulverzeichnis\Schulverzeichnis_RLP_2025_26.xlsx'), 0),
        @((Join-Path $on '03_Berlin_EFRE_Liste_der_Vorhaben\EFRE_Berlin_Liste_der_Vorhaben_2023-06-30.xlsx'), 0),
        @((Join-Path $on '04_Auswaertiges_Amt_Staatennamen\Liste_der_Staatennamen_AA.xlsx'), 0),
        @((Join-Path $k 'Buergerbefragung_kodiert_Codeplan.xlsx'), 0),
        @((Join-Path $k 'Buergerbefragung_kodiert_Codeplan_Umkodierung_99_98_97.xlsx'), 0),
        @((Join-Path $k 'Buergerbefragung_unkodiert_Codeplan.xlsx'), 0)
    )
}
$imp = if ($NoImport) { '0' } else { '1' }
$lines = @()
$n = 0
foreach ($c in $cases) {
    $n++
    $p = if ([IO.Path]::IsPathRooted($c[0])) { $c[0] } else { Join-Path $td $c[0] }
    $part = Join-Path $dir ('fall_' + $n + '.txt')
    $spec = $p + '|' + [string]$c[1] + '|' + $imp + '|' + $part
    $t0 = Get-Date
    $o = & powershell -ExecutionPolicy Bypass -File (Join-Path $PSScriptRoot 'runmacro.ps1') -WorkbookName $wbName -Macro 'mListImport.LiTestCase' -MacroArgs $spec -TimeoutSec 240
    $sec = [math]::Round(((Get-Date) - $t0).TotalSeconds, 1)
    $lines += ('=== ' + (Split-Path $c[0] -Leaf) + ' (Tabelle ' + $c[1] + ', ' + $sec + ' s, ' + (($o -join ' ') -replace '^OK: ', '') + ')')
    if (Test-Path $part) { $lines += Get-Content $part -Encoding UTF8; Remove-Item $part -Force }
    $lines += ''
}
$res = Join-Path $dir 'Listimport_Ergebnis.txt'
$lines | Set-Content $res -Encoding UTF8
if (-not $NoImport) {
    $dump = Join-Path $dir 'Listimport_Tabellen.txt'
    [void](& powershell -ExecutionPolicy Bypass -File (Join-Path $PSScriptRoot 'runmacro.ps1') -WorkbookName $wbName -Macro 'mListImport.LiDumpTests' -MacroArgs $dump -TimeoutSec 240)
}
# eigene Testinstanz schliessen (ohne Speichern, ohne Rueckfrage der Anwendung)
try {
    $xl = [Runtime.InteropServices.Marshal]::GetActiveObject('Excel.Application')
    # geplante Aufrufe der Anwendung abmelden (das Schliessen ohne Ereignisse ueberspringt das sonst;
    # Excel wuerde die Datei fuer den naechsten Zoom-Takt wieder oeffnen)
    foreach ($m in @('mNav.StopWatch', 'mBackup.CancelAutoSave')) { try { [void]$xl.Run("'" + $wbName + "'!" + $m) } catch { } }
    $xl.EnableEvents = $false
    $xl.DisplayAlerts = $false
    foreach ($w in @($xl.Workbooks)) { $w.Close($false) }
    $xl.Quit()
    [void][Runtime.InteropServices.Marshal]::ReleaseComObject($xl)
} catch { }
"Ergebnis: $res"
