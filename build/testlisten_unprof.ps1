# Drei "gewachsene" Verwaltungslisten, wie sie in Aemtern tatsaechlich entstehen: verbundene Zellen,
# Farbe statt Status, Legende und Notizen unter der Tabelle, wilde Datums- und Betragsangaben,
# mehrere Werte in einer Zelle, Blaetter mit abweichender Spaltenreihenfolge. Alle Namen und Daten sind erfunden.
# Aufruf:   powershell -ExecutionPolicy Bypass -File build\testlisten_unprof.ps1
# Ergebnis: testdaten\importlisten\14_*.xlsx, 15_*.xlsx, 16_*.xlsx (Beschreibung in LIESMICH.md dort)
# Die Datei muss als UTF-8 mit BOM gespeichert sein (Umlaute in den Beispieldaten).
param([string]$OutDir = "")
$ErrorActionPreference = 'Stop'
$root = Split-Path $PSScriptRoot -Parent
if ($OutDir -eq "") { $OutDir = Join-Path $root 'testdaten\importlisten' }
New-Item -ItemType Directory -Force $OutDir | Out-Null

function RGB([int]$r, [int]$g, [int]$b) { return $r + 256 * $g + 65536 * $b }
function D([int]$y, [int]$m, [int]$d) { return (New-Object DateTime $y, $m, $d).ToOADate() }
$GRAU = RGB 217 217 217; $GELB = RGB 255 242 0; $ORANGE = RGB 255 192 0; $HELLBLAU = RGB 189 215 238
$EN = [Globalization.CultureInfo]'en-US'
function Format($ws, [string]$addr, [string]$fmt) { [void][System.__ComObject].InvokeMember('NumberFormat', [Reflection.BindingFlags]::SetProperty, $null, $ws.Range($addr), @($fmt), $null, $EN, $null) }
function Put($ws, [int]$r0, [int]$c0, [object[]]$rows) {
    $n = $rows.Count; $m = 1
    foreach ($row in $rows) { if (@($row).Count -gt $m) { $m = @($row).Count } }
    $a = New-Object 'object[,]' $n, $m
    for ($i = 0; $i -lt $n; $i++) { $row = @($rows[$i]); for ($j = 0; $j -lt $row.Count; $j++) { $a[$i, $j] = $row[$j] } }
    $rng = $ws.Range($ws.Cells.Item($r0, $c0), $ws.Cells.Item($r0 + $n - 1, $c0 + $m - 1))
    $rng.Value2 = $a
    return $rng
}
function Txt($ws, [int]$r, [int]$c, [string]$s) { $cell = $ws.Cells.Item($r, $c); $cell.NumberFormat = '@'; $cell.Value2 = $s }
function Breiten($ws, [double[]]$w) { for ($i = 0; $i -lt $w.Count; $i++) { $ws.Columns.Item($i + 1).ColumnWidth = $w[$i] } }
function Linien($rng) { foreach ($i in 7, 8, 9, 10, 11, 12) { try { $b = $rng.Borders.Item($i); $b.LineStyle = 1; $b.Weight = 2 } catch { } } }

$xl = New-Object -ComObject Excel.Application
$xl.Visible = $false
$xl.DisplayAlerts = $false
function NeueMappe([string[]]$blaetter) {
    $wb = $xl.Workbooks.Add()
    while ($wb.Worksheets.Count -lt $blaetter.Count) { [void]$wb.Worksheets.Add([Type]::Missing, $wb.Worksheets.Item($wb.Worksheets.Count)) }
    while ($wb.Worksheets.Count -gt $blaetter.Count) { $wb.Worksheets.Item($wb.Worksheets.Count).Delete() }
    for ($i = 1; $i -le $blaetter.Count; $i++) { $s = $wb.Worksheets.Item($i); $s.Name = $blaetter[$i - 1]; $s.Cells.Font.Name = 'Arial'; $s.Cells.Font.Size = 10 }
    return $wb
}
function Speichern($wb, [string]$name) {
    foreach ($s in @($wb.Worksheets)) { if ($s.Visible -eq -1) { try { $s.Activate(); $s.Range('A1').Select() | Out-Null } catch { } } }
    $wb.Worksheets.Item(1).Activate()
    $p = Join-Path $OutDir $name
    if (Test-Path $p) { Remove-Item $p -Force }
    $wb.SaveAs($p, 51)
    $wb.Close($false)
    Write-Output "erstellt: $name"
}

try {
# ---------------------------------------------------------------------------------------------
# 14 Telefonliste: Abteilung als senkrecht verbundene Zelle, Durchwahlen in allen Schreibweisen,
# zwei Personen in einer Zelle, Farbe und Durchstreichen statt Status, Legende darunter,
# ausgeblendete Spalte, Kommentar, Raum "2.01" von Excel zum Datum gemacht
# ---------------------------------------------------------------------------------------------
$wb = NeueMappe @('Tabelle1')
$ws = $wb.Worksheets.Item(1)
$ws.Range('A1').Value2 = 'Telefonliste Amt für Bürgerdienste'
$ws.Range('A1:G1').Merge(); $ws.Range('A1').Font.Size = 14; $ws.Range('A1').Font.Bold = $true
$ws.Range('A2').Value2 = 'Stand: siehe unten!!'; $ws.Range('A2').Font.Italic = $true; $ws.Range('A2').Font.Color = (RGB 192 0 0)
[void](Put $ws 4 1 @(, @('Abt.', 'Name, Vorname', 'Funktion', 'Tel.', 'Zi.', 'E-Mail', 'Bemerkung / Vertretung', 'alte Durchwahl')))
$ws.Range('A4:H4').Font.Bold = $true; $ws.Range('A4:H4').Interior.Color = $GRAU
$rows = @(
    @('Abt. 1 Bürgerbüro', 'Kaya, Selin', 'Leitung', '0123 45-3401', '1.12', 'selin.kaya@stadt.example', 'Vertr.: Hr. Brandt', '2201'),
    @('', 'Brandt, Jonas (Teilzeit)', 'Sachbearbeitung', '12-3402', '1.14', 'jonas.brandt@stadt.example', 'nur Mo-Mi', '2202'),
    @('', 'Okafor, Grace', 'SB', '-3403', '1.14', '', 'Elternzeit bis 31.08.', '2203'),
    @('', 'Schulz, Peter', 'SB Meldewesen', 'intern 3404', '1.15', 'peter.schulz@stadt.example', '', '2204'),
    @('', 'Neu: Lindqvist, Maja', 'Auszubildende', '3405', '1.15', 'maja.lindqvist@stadt.example', 'ab 01.09.', ''),
    @('', '', '', '', '', '', '', ''),
    @('Abt. 2 Kfz-Zulassung', 'Yilmaz, Deniz', 'Leitung', '0123 45-3410', '', 'deniz.yilmaz@stadt.example', '', '2301'),
    @('', 'Hartmann, Felix', 'Sachbearbeitung', '12-3411 / -3412', '2.03', 'felix.hartmann@stadt.example', 'Vertretung für Fr. Yilmaz', '2302'),
    @('', 'Berger, Anna; Romano, Luca', 'Schalter', '3413', '2.04', 'zulassung@stadt.example', 'Funktionspostfach', ''),
    @('', 'Nowak, Sven', 'Sachbearbeitung', '3414', '2.05', 'sven.nowak@stadt.example', 'Mo, Di, Do', '2304'),
    @('Abt. 3 Fundbüro', 'Wagner, Lena', 'SB', '3420', 'EG', 'fundbuero@stadt.example', '', ''),
    @('', 'Fischer, Tom', '', '3421', 'EG', '', '?', '')
)
for ($i = 0; $i -lt $rows.Count; $i++) { for ($j = 0; $j -lt 8; $j++) { Txt $ws (5 + $i) (1 + $j) ([string]$rows[$i][$j]) } }
# Raum 2.01 hat Excel beim Tippen in den 2. Januar verwandelt
Format $ws 'E11' 'General'; $ws.Cells.Item(11, 5).Value2 = (D 2026 1 2); Format $ws 'E11' 'dd. mmm'
$ws.Range('A5:A9').Merge(); $ws.Range('A11:A14').Merge(); $ws.Range('A15:A16').Merge()
$ws.Range('A5:A16').VerticalAlignment = -4108; $ws.Range('A5:A16').WrapText = $true; $ws.Range('A5:A16').Font.Bold = $true
$ws.Range('A7:G7').Interior.Color = $GELB
$ws.Range('A8:G8').Font.Strikethrough = $true
Linien $ws.Range('A4:H16')
[void]$ws.Cells.Item(7, 5).AddComment('Büro wird geteilt')
$ws.Range('A18').Value2 = 'gelb = nicht im Dienst, durchgestrichen = ausgeschieden'
$ws.Range('A19').Value2 = 'Stand: 03/2026, Änderungen bitte an Zimmer 1.12 melden'
$ws.Range('A18:A19').Font.Italic = $true
Breiten $ws @(16, 26, 16, 16, 7, 28, 26, 10)
$ws.Columns.Item(8).Hidden = $true   # nach den Breiten (eine Breite blendet wieder ein)
Speichern $wb '14_Telefonliste_Amt_FINAL_neu.xlsx'

# ---------------------------------------------------------------------------------------------
# 15 Veranstaltungsplanung: Kopf mit Zeilenumbruechen, Datum in allen Varianten, Kosten mit Text,
# Ja/Nein-Varianten, mehrere Zielgruppen mit / und +, Summenzeile mitten drin, Ideen-Liste darunter,
# leere Blaetter Tabelle2 und Tabelle3 (eines mit Notizen)
# ---------------------------------------------------------------------------------------------
$wb = NeueMappe @('Planung 2026', 'Tabelle2', 'Tabelle3')
$ws = $wb.Worksheets.Item(1)
$ws.Range('A1').Value2 = 'Veranstaltungen Stadtteilbüro 2026 (Entwurf!!)'; $ws.Range('A1').Font.Bold = $true; $ws.Range('A1').Font.Size = 12
[void](Put $ws 3 1 @(, @('Was?', "Datum`n(geplant)", 'Uhrzeit', 'Ort', 'Zielgruppe', 'Kosten', "Anmeldung`nnötig?", 'Wer kümmert sich?', 'Stand')))
$ws.Range('A3:I3').Font.Bold = $true; $ws.Range('A3:I3').Interior.Color = $HELLBLAU; $ws.Range('A3:I3').WrapText = $true
$daten = @(
    @('Frühlingsfest', (D 2026 4 25), '14-18 Uhr', 'Bürgerhaus', 'alle', 'ca. 1.200 €', 'nein', 'Fr. Kaya', 'läuft'),
    @('Repair-Café', 'jeden 1. Samstag', '10:00-13:00', 'Stadtteilbüro', 'Erwachsene/Senioren', 'kostenlos', 'nein', 'Hr. Brandt + Ehrenamt', 'ok'),
    @('Ferienspiele', 'KW 31-33', '9-15 Uhr', 'Grundschule Am Park', 'Kinder (6-12)', 3500, 'ja!', 'Jugendpflege', 'Antrag gestellt'),
    @('Seniorennachmittag', 'Mai 2026', 'ab 15 Uhr', 'Café Mitte', 'Senioren', '500,-', 'x', 'Fr. Okafor', '?'),
    @('Laternenumzug', '11.11.', '17:30', 'Treffpunkt Kirche', 'Familien + Kinder', 's.o.', '(ja)', 'Fr. Kaya / Hr. Nowak', 'Genehmigung fehlt'),
    @('', '', '', '', '', '', '', '', ''),
    @('Summe Q2', '', '', '', '', 5200, '', '', ''),
    @('Bürgerversammlung', (D 2026 6 15), '19 Uhr', 'Bürgerhaus, gr. Saal', 'alle Bürger', '', 'nein', 'Amtsleitung', 'fix'),
    @('Kinderflohmarkt', '15./16.06.2026', '', 'Schulhof', 'Familien', 0, 'j', 'Elternbeirat', 'angefragt'),
    @('Lesung', 'tbd', '', 'Bibliothek', 'Erwachsene', 'ca. 300 €', 'Ja', 'Hr. Romano', '')
)
for ($i = 0; $i -lt $daten.Count; $i++) {
    for ($j = 0; $j -lt 9; $j++) {
        $v = $daten[$i][$j]
        if ($v -is [string]) { Txt $ws (4 + $i) (1 + $j) $v } else { $ws.Cells.Item(4 + $i, 1 + $j).Value2 = $v }
    }
}
Format $ws 'B4' 'dd.mm.yyyy'; Format $ws 'B11' 'dd.mm.yyyy'
$ws.Range('A10:I10').Font.Bold = $true
$ws.Range('A7:I7').Interior.Color = $ORANGE
Linien $ws.Range('A3:I13')
$ws.Range('A16').Value2 = 'Ideen (noch offen):'; $ws.Range('A16').Font.Bold = $true
$ws.Range('A17').Value2 = 'Kino im Park'
$ws.Range('A18').Value2 = 'Stadtteilrallye (Herbst?)'
$ws.Range('A19').Value2 = 'Tauschbörse, evtl. mit Kita'
Breiten $ws @(20, 16, 12, 22, 20, 12, 11, 22, 18)
$ws.Rows.Item(3).RowHeight = 28
$w3 = $wb.Worksheets.Item(3)
$w3.Range('A1').Value2 = 'Notizen Besprechung'
$w3.Range('A2').Value2 = 'Budget insgesamt ca. 8.000 €, Rest über Spenden'
Speichern $wb '15_Veranstaltungsplanung_2026_Kopie.xlsx'

# ---------------------------------------------------------------------------------------------
# 16 Fortbildungsbedarf: je Amt ein Blatt mit abweichender Spaltenfolge und -benennung,
# mehrere Wuensche per Zeilenumbruch in einer Zelle, Kaestchen-Zeichen, Dauer mit Einheiten,
# "dito" und "siehe oben", Auswertungsblatt mit Formeln, ausgeblendetes Altblatt
# ---------------------------------------------------------------------------------------------
$wb = NeueMappe @('Amt 10', 'Amt 20', 'Amt 51 (spät)', 'Auswertung', 'alt')
$hinweis = 'Fortbildungsbedarf 2027: bitte bis 15.05. zurück an die Personalentwicklung'
$ja = [string][char]0x2611; $nein = [string][char]0x2610
# Amt 10
$ws = $wb.Worksheets.Item(1)
$ws.Range('A1').Value2 = $hinweis; $ws.Range('A1').Font.Italic = $true
[void](Put $ws 2 1 @(, @('Name', 'Fortbildungswunsch', 'Dauer', 'dringend?', 'Kosten ca.', 'Bemerkung')))
$ws.Range('A2:F2').Font.Bold = $true; $ws.Range('A2:F2').Interior.Color = $GRAU
$a10 = @(
    @('Kaya, Selin', "Excel Aufbau`nProjektmanagement", '2 Tage', $ja, 450, ''),
    @('Brandt, Jonas', 'Konfliktmanagement', '1', $nein, 'ca. 600 €', 'gern online'),
    @('Okafor, Grace', "Datenschutz (Pflicht)`nE-Akte", '0,5 Tag', $ja, 'kostenlos (intern)', 'nach Elternzeit'),
    @('Schulz, Peter', 'siehe oben', '?', '', '', '')
)
for ($i = 0; $i -lt $a10.Count; $i++) { for ($j = 0; $j -lt 6; $j++) { $v = $a10[$i][$j]; if ($v -is [string]) { Txt $ws (3 + $i) (1 + $j) $v } else { $ws.Cells.Item(3 + $i, 1 + $j).Value2 = $v } } }
$ws.Range('B3:B6').WrapText = $true
Linien $ws.Range('A2:F6'); Breiten $ws @(18, 30, 10, 10, 18, 22)
# Amt 20: andere Reihenfolge, zusaetzliche Spalte
$ws = $wb.Worksheets.Item(2)
$ws.Range('A1').Value2 = $hinweis; $ws.Range('A1').Font.Italic = $true
[void](Put $ws 2 1 @(, @('Name', 'Wunsch', 'dringend?', 'Dauer', 'Kosten ca.', 'Genehmigt AL', 'Bemerkung')))
$ws.Range('A2:G2').Font.Bold = $true; $ws.Range('A2:G2').Interior.Color = $GRAU
$a20 = @(
    @('Yilmaz, Deniz', "Führung in Teilzeit`nGesundheitsmanagement", 'x', '3 x 2 Std.', '250,-', 'ja', ''),
    @('Hartmann, Felix', 'Excel Aufbau', '', '2 Tage', 450, 'ja', 'wie Amt 10'),
    @('Berger, Anna', 'dito', 'ja', '2 Tage', 450, '', ''),
    @('Nowak, Sven', 'Gebärdensprache Grundkurs', 'eher nein', '5 Tage', 'ca. 900 €', 'nein', 'Budget prüfen')
)
for ($i = 0; $i -lt $a20.Count; $i++) { for ($j = 0; $j -lt 7; $j++) { $v = $a20[$i][$j]; if ($v -is [string]) { Txt $ws (3 + $i) (1 + $j) $v } else { $ws.Cells.Item(3 + $i, 1 + $j).Value2 = $v } } }
$ws.Range('B3:B6').WrapText = $true
Linien $ws.Range('A2:G6'); Breiten $ws @(18, 30, 10, 12, 12, 13, 18)
# Amt 51: Kopf anders benannt, Leerzeile, Nachzuegler
$ws = $wb.Worksheets.Item(3)
$ws.Range('A1').Value2 = $hinweis + ' (kam erst am 20.05.)'; $ws.Range('A1').Font.Italic = $true
[void](Put $ws 2 1 @(, @('Mitarbeiter/in', 'Fortbildungswunsch', 'Dauer', 'dringend?', 'Kosten', 'Bemerkung')))
$ws.Range('A2:F2').Font.Bold = $true; $ws.Range('A2:F2').Interior.Color = $GRAU
$a51 = @(
    @('Wagner, Lena', "Kinderschutz (§ 8a)`nGesprächsführung`nDokumentation", '3 Tage', $ja, 'ca. 750 €', 'dringend!!'),
    @('', '', '', '', '', ''),
    @('Fischer, Tom', 'Excel Grundlagen', '1 Tag', $nein, 0, ''),
    @('Romano, Luca', 'keiner', '', '', '', '')
)
for ($i = 0; $i -lt $a51.Count; $i++) { for ($j = 0; $j -lt 6; $j++) { $v = $a51[$i][$j]; if ($v -is [string]) { Txt $ws (3 + $i) (1 + $j) $v } else { $ws.Cells.Item(3 + $i, 1 + $j).Value2 = $v } } }
$ws.Range('B3:B6').WrapText = $true
Linien $ws.Range('A2:F6'); Breiten $ws @(18, 30, 10, 10, 12, 16)
# Auswertung mit Formeln
$ws = $wb.Worksheets.Item(4)
[void](Put $ws 1 1 @(@('Auswertung (nicht ausfüllen)', ''), @('Amt', 'Rückmeldungen'), @('Amt 10', ''), @('Amt 20', ''), @('Amt 51', ''), @('Summe', '')))
$ws.Range('B3').Formula = "=COUNTA('Amt 10'!A3:A20)"
$ws.Range('B4').Formula = "=COUNTA('Amt 20'!A3:A20)"
$ws.Range('B5').Formula = "=COUNTA('Amt 51 (spät)'!A3:A20)"
$ws.Range('B6').Formula = '=SUM(B3:B5)'
$ws.Range('A1').Font.Bold = $true; $ws.Range('A2:B2').Font.Bold = $true; $ws.Range('A6:B6').Font.Bold = $true
Breiten $ws @(28, 14)
# ausgeblendetes Altblatt vom Vorjahr
$ws = $wb.Worksheets.Item(5)
[void](Put $ws 1 1 @(@('Name', 'Thema 2026', 'erledigt'), @('Kaya, Selin', 'Zeitmanagement', 'ja'), @('Brandt, Jonas', 'Excel Grundlagen', 'nein')))
$ws.Visible = 0
Speichern $wb '16_Fortbildungsbedarf_Rueckmeldungen.xlsx'
} finally {
    $xl.Quit()
    [void][Runtime.InteropServices.Marshal]::ReleaseComObject($xl)
    [GC]::Collect(); [GC]::WaitForPendingFinalizers()
}
