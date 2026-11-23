# Testsatz fuer den Listen-Import: typische Verwaltungslisten mit unterschiedlichem Aufbau
# und nachgebaute 2ask-Exporte. Alle Namen und Daten sind erfunden.
# Aufruf:   powershell -ExecutionPolicy Bypass -File build\testlisten.ps1
# Ergebnis: testdaten\importlisten\*.xlsx und *.csv; Beschreibung und Soll-Erkennung in LIESMICH.md dort
# Die Datei muss als UTF-8 mit BOM gespeichert sein (Umlaute in den Beispieldaten).
param([string]$OutDir = "")
$ErrorActionPreference = 'Stop'
$root = Split-Path $PSScriptRoot -Parent
if ($OutDir -eq "") { $OutDir = Join-Path $root 'testdaten\importlisten' }
New-Item -ItemType Directory -Force $OutDir | Out-Null

function RGB([int]$r, [int]$g, [int]$b) { return $r + 256 * $g + 65536 * $b }
function D([int]$y, [int]$m, [int]$d, [int]$h = 0, [int]$mi = 0) { return (New-Object DateTime $y, $m, $d, $h, $mi, 0).ToOADate() }
$GRAU = RGB 217 217 217; $BLAU = RGB 221 235 247; $HELLGRAU = RGB 242 242 242
$ROT = RGB 255 199 206; $GELB = RGB 255 235 156; $GRUEN = RGB 198 239 206
$LILA = RGB 228 223 236; $ORANGE = RGB 252 228 214; $TUERKIS = RGB 218 238 243

# Zeilen (Array von Arrays) ab Zelle r0/c0 als Block schreiben; liefert den Bereich
function Put($ws, [int]$r0, [int]$c0, [object[]]$rows) {
    $n = $rows.Count; $m = 1
    foreach ($row in $rows) { if (@($row).Count -gt $m) { $m = @($row).Count } }
    $a = New-Object 'object[,]' $n, $m
    for ($i = 0; $i -lt $n; $i++) {
        $row = @($rows[$i])
        for ($j = 0; $j -lt $row.Count; $j++) { $a[$i, $j] = $row[$j] }
    }
    $rng = $ws.Range($ws.Cells.Item($r0, $c0), $ws.Cells.Item($r0 + $n - 1, $c0 + $m - 1))
    $rng.Value2 = $a
    return $rng
}
# Zelle ausdruecklich als Text (sonst wandelt Excel "12.3.25" oder "1.12" in Datum bzw. Zahl um)
function Txt($ws, [int]$r, [int]$c, [string]$s) { $cell = $ws.Cells.Item($r, $c); $cell.NumberFormat = '@'; $cell.Value2 = $s }
function TextSpalte($ws, [string]$addr) { $ws.Range($addr).NumberFormat = '@' }
# Zahlenformate mit englischen Codes setzen. PowerShell ruft Excel mit der Benutzersprache auf: ein deutsches Excel
# liest "dd.mm.yyyy" sonst als "dd.01.yyyy" und "#,##0.00" als "#,##000". InvokeMember mit en-US behebt das.
$EN = [Globalization.CultureInfo]'en-US'
function Format($ws, [string]$addr, [string]$fmt) { [void][System.__ComObject].InvokeMember('NumberFormat', [Reflection.BindingFlags]::SetProperty, $null, $ws.Range($addr), @($fmt), $null, $EN, $null) }
function Kopf($rng, $fill) { $rng.Font.Bold = $true; if ($null -ne $fill) { $rng.Interior.Color = $fill }; $rng.VerticalAlignment = -4108 }
function Linien($rng) {
    foreach ($i in 7, 8, 9, 10, 11, 12) {
        try { $b = $rng.Borders.Item($i); $b.LineStyle = 1; $b.Weight = 2; $b.Color = (RGB 166 166 166) } catch { }
    }
}
function Breiten($ws, [double[]]$w) { for ($i = 0; $i -lt $w.Count; $i++) { $ws.Columns.Item($i + 1).ColumnWidth = $w[$i] } }
function Titel($ws, [string]$addr, [string]$text, [double]$size = 13) { $c = $ws.Range($addr); $c.Value2 = $text; $c.Font.Bold = $true; $c.Font.Size = $size }
function Fixieren($wb, $ws, [int]$rows, [int]$cols = 0) {
    try {
        $ws.Activate()
        $win = $wb.Windows.Item(1)
        $win.FreezePanes = $false
        $win.SplitColumn = $cols; $win.SplitRow = $rows
        $win.FreezePanes = $true
    } catch { }
}

$xl = New-Object -ComObject Excel.Application
$xl.Visible = $false
$xl.DisplayAlerts = $false
function NeueMappe([string[]]$blaetter, [string]$schrift = 'Calibri', [double]$groesse = 11) {
    $wb = $xl.Workbooks.Add()
    while ($wb.Worksheets.Count -lt $blaetter.Count) { [void]$wb.Worksheets.Add([Type]::Missing, $wb.Worksheets.Item($wb.Worksheets.Count)) }
    while ($wb.Worksheets.Count -gt $blaetter.Count) { $wb.Worksheets.Item($wb.Worksheets.Count).Delete() }
    for ($i = 1; $i -le $blaetter.Count; $i++) {
        $s = $wb.Worksheets.Item($i)
        $s.Name = $blaetter[$i - 1]
        $s.Cells.Font.Name = $schrift
        $s.Cells.Font.Size = $groesse
    }
    return $wb
}
function Speichern($wb, [string]$name) {
    foreach ($s in @($wb.Worksheets)) { try { $s.Activate(); $s.Range('A1').Select() | Out-Null } catch { } }
    $wb.Worksheets.Item(1).Activate()
    $p = Join-Path $OutDir $name
    if (Test-Path $p) { Remove-Item $p -Force }
    $wb.SaveAs($p, 51)
    $wb.Close($false)
    Write-Output "erstellt: $name"
}

try {
#==============================================================================
# 01 Teilnehmerliste: Titelzeilen, zweizeiliger Kopf, Kreuze ueber mehrere Spalten, Summenzeile
#==============================================================================
$wb = NeueMappe @('Anmeldungen')
$ws = $wb.Worksheets.Item(1)
Titel $ws 'A1' 'Fortbildungsreihe Kinderschutz 2025' 14
$ws.Range('A2').Value2 = 'Anmeldungen, Stand 15.04.2025'
$ws.Range('A2').Font.Italic = $true
[void](Put $ws 4 1 @(
    @('Nr.', 'Name', 'Vorname', 'Dienststelle', 'E-Mail', 'Teilnahme', $null, $null, 'Verpflegung', 'Bemerkung'),
    @($null, $null, $null, $null, $null, 'Modul 1 (12.05.)', 'Modul 2 (16.06.)', 'Modul 3 (08.09.)', $null, $null)
))
$daten = @(
    @(1, 'Becker', 'Anna', 'Jugendamt', 'anna.becker@stadt.example', 'x', 'x', 'x', 'vegetarisch', $null),
    @(2, 'Schulz', 'Thomas', 'Jugendamt ', 'thomas.schulz@stadt.example', 'x', $null, 'x', $null, $null),
    @(3, 'Yilmaz', 'Elif', 'Kita Sonnenblume', 'e.yilmaz@kita-sonnenblume.example', 'x', 'x', $null, 'vegan', 'kommt evtl. später'),
    @(4, 'Hoffmann', 'Jan', 'Sozialamt', 'jan.hoffmann@stadt.example', 'X', 'X', 'X', $null, $null),
    @(5, 'Wagner', 'Laura', 'Amt für Schulen', 'laura.wagner@stadt.example', 'x', 'entschuldigt', 'x', 'keine Angabe', 'Modul 2 Nachholtermin?'),
    @(6, 'Demir', 'Mehmet', 'Jugendamt', 'mehmet.demir@stadt.example', $null, 'x', 'x', 'vegetarisch', $null),
    @(7, 'Schäfer', 'Julia', 'Kita Regenbogen', 'j.schaefer@kita-regenbogen.example', 'x', 'x', $null, 'Veg.', $null),
    @(8, 'Klein', 'Peter', 'Gesundheitsamt', 'peter.klein@stadt.example', 'x', $null, $null, $null, 'nur Modul 1'),
    @(9, 'Richter', 'Sophie', 'Jugendamt', 'sophie.richter@stadt.example', 'x', 'x', 'x', 'vegetarisch', $null),
    @(10, 'Wolf', 'Daniel', 'Kita Sonnenblume', 'd.wolf@kita-sonnenblume.example', 'x', 'x', 'x', $null, $null),
    @(11, 'Neumann', 'Sarah', 'Sozialamt', 'sarah.neumann@stadt.example', $null, 'x', 'x', 'vegan', $null),
    @(12, 'Schwarz', 'Michael', 'Amt für Schulen', 'michael.schwarz@stadt.example', 'x', 'x', 'x', $null, $null),
    @(13, 'Zimmermann', 'Lea', 'Jugendamt', 'lea.zimmermann@stadt.example', 'x', $null, 'x', 'vegetarisch', $null),
    @(14, 'Braun', 'Tim', 'Kita Pusteblume', 't.braun@kita-pusteblume.example', 'x', 'x', $null, $null, $null),
    @(15, 'Hartmann', 'Nina', 'Gesundheitsamt', 'nina.hartmann@stadt.example', 'x', 'x', 'x', 'keine Angabe', $null),
    @(16, 'Lange', 'Felix', 'Jugendamt', 'felix.lange@stadt.example', 'x', 'x', 'x', $null, $null),
    @(17, 'Werner', 'Mia', 'Kita Regenbogen', 'm.werner@kita-regenbogen.example', $null, $null, 'x', 'vegetarisch', 'Nachmeldung'),
    @(18, 'Krause', 'Paul', 'Sozialamt', 'paul.krause@stadt.example', 'x', 'x', 'x', $null, $null),
    @(19, 'Lehmann', 'Emma', 'Jugendamt', 'emma.lehmann@stadt.example', 'x', 'x', $null, 'vegan', $null),
    @(20, 'Köhler', 'Ben', 'Amt für Schulen', 'ben.koehler@stadt.example', 'x', $null, 'x', $null, $null)
)
[void](Put $ws 6 1 $daten)
[void](Put $ws 26 1 @(@('Summe', $null, $null, $null, $null, '=COUNTA(F6:F25)', '=COUNTA(G6:G25)', '=COUNTA(H6:H25)', $null, $null), @($null)))
foreach ($a in 'A4:A5', 'B4:B5', 'C4:C5', 'D4:D5', 'E4:E5', 'F4:H4', 'I4:I5', 'J4:J5') { $ws.Range($a).Merge() }
Kopf $ws.Range('A4:J5') $GRAU
$ws.Range('A4:J5').HorizontalAlignment = -4108
$ws.Range('F6:H26').HorizontalAlignment = -4108
$ws.Range('A26:J26').Font.Bold = $true
Linien $ws.Range('A4:J26')
Breiten $ws @(5, 13, 11, 18, 36, 11, 11, 11, 13, 24)
$ws.Range('F5:H5').WrapText = $true
Fixieren $wb $ws 5
Speichern $wb '01_Teilnehmerliste_Fortbildung.xlsx'

#==============================================================================
# 02 Fristenliste: Kopf in Zeile 1, Filter, Datum teils als Text, Kuerzel, Status-Schreibweisen,
#    Ampel nur als Zellfarbe (ohne Legende)
#==============================================================================
$wb = NeueMappe @('WV 2025') 'Arial' 10
$ws = $wb.Worksheets.Item(1)
TextSpalte $ws 'A2:A23'
$daten = @(
    @('Az.', 'Betreff', 'Eingang', 'Frist', 'zust.', 'Status', 'erledigt am', 'Bemerkung'),
    @('51.2-0101/25', 'Zuschussantrag Jugendverein Ferienfreizeit', (D 2025 1 8), (D 2025 2 15), 'Mü', 'erledigt', (D 2025 2 10), $null),
    @('51.2-0107/25', 'Beschwerde Spielplatz Lindenstraße', (D 2025 1 14), (D 2025 1 31), 'Sch', 'Erledigt', (D 2025 1 29), 'Antwort per Mail'),
    @('51.1-0112/25', 'Anfrage Fraktion zu Kitaplätzen', (D 2025 1 20), (D 2025 2 3), 'Kl', 'erl.', (D 2025 2 3), $null),
    @('51.2-0120/25', 'Verwendungsnachweis Förderverein prüfen', $null, (D 2025 4 30), 'Ba', 'in Arbeit', $null, 'Belege fehlen'),
    @('51.3-0125/25', 'Stellungnahme Bebauungsplan Nord', (D 2025 3 18), (D 2025 4 11), 'Mü', 'erledigt', (D 2025 4 9), $null),
    @('51.2-0131/25', 'Kooperationsvertrag Schulsozialarbeit', 'KW 14', (D 2025 6 30), 'Sch', 'offen', $null, $null),
    @('51.1-0138/25', 'Bericht Jugendhilfeausschuss Juni', (D 2025 4 7), (D 2025 5 23), 'Kl', 'in Bearb.', $null, 'Entwurf liegt vor'),
    @('51.2-0142/25', 'Antrag Mittagsverpflegung Hort', (D 2025 4 15), (D 2025 5 15), 'MÜ', 'erledigt', (D 2025 5 12), $null),
    @('51.3-0150/25', 'Elternbeschwerde Öffnungszeiten', (D 2025 4 22), (D 2025 5 6), 'Ba', 'zurückgestellt', $null, 'wartet auf Träger'),
    @('51.2-0155/25', 'Fördermittel Bundesprogramm abrufen', 'Anfang Mai', (D 2025 6 13), 'Mü', 'in Arbeit', $null, $null),
    @('51.1-0161/25', 'Datenschutzanfrage Bürger', (D 2025 5 9), (D 2025 6 9), 'Kl', 'erledigt', $null, $null),
    @('51.2-0166/25', 'Raumanfrage Familienzentrum', (D 2025 5 14), (D 2025 5 28), 'Sch', 'offen', $null, $null),
    @('51.3-0170/25', 'Statistikmeldung Land', (D 2025 5 20), (D 2025 7 15), 'Ba', 'offen', $null, 'Zahlen von Amt 40 fehlen'),
    @('51.2-0174/25', 'Zuschuss Stadtteilfest', (D 2025 5 26), (D 2025 6 20), 'Mü', 'Erledigt', (D 2025 6 18), $null),
    @('51.1-0181/25', 'Personalbedarfsplanung Kita', (D 2025 6 2), (D 2025 7 31), 'Kl', 'in Arbeit', $null, $null),
    @('51.2-0185/25', 'Mängelmeldung Jugendhaus Dach', (D 2025 6 4), (D 2025 6 11), 'Sch', 'erledigt', (D 2025 6 11), 'an Gebäudemanagement'),
    @('51.3-0192/25', 'Evaluation Schulsozialarbeit', (D 2025 6 10), (D 2025 9 30), 'Ba', 'offen', $null, $null),
    @('51.2-0198/25', 'Antrag Inklusionsassistenz', (D 2025 6 16), (D 2025 7 1), 'Mü', 'in Bearb.', $null, $null),
    @('51.1-0203/25', 'Anfrage Presse Ferienbetreuung', (D 2025 6 23), (D 2025 6 25), 'Kl', 'erl.', (D 2025 6 24), $null),
    @('51.2-0207/25', 'Vertrag Reinigung Kita Nord', (D 2025 6 25), (D 2025 8 31), 'Sch', 'offen', $null, $null),
    @('51.3-0211/25', 'Bedarfsanalyse Spielplätze', (D 2025 6 27), (D 2025 10 15), 'Ba', 'offen', $null, $null),
    @('51.2-0215/25', 'Zuschuss Sprachkurs Eltern', (D 2025 6 30), (D 2025 7 18), 'Mü', 'offen', $null, $null)
)
[void](Put $ws 1 1 $daten)
Txt $ws 5 3 '12.3.25'
Txt $ws 12 7 's. Vermerk'
Format $ws 'C2:D23' 'dd.mm.yyyy'
Format $ws 'G2:G23' 'dd.mm.yyyy'
# Ampel auf der Frist (Stand 01.07.2025): rot ueberfaellig, gelb in den naechsten drei Wochen, gruen sonst bzw. erledigt
$ampel = @{ 2 = $GRUEN; 3 = $GRUEN; 4 = $GRUEN; 5 = $ROT; 6 = $GRUEN; 7 = $ROT; 8 = $ROT; 9 = $GRUEN; 10 = $ROT; 11 = $ROT; 12 = $GRUEN; 13 = $ROT;
            14 = $GELB; 15 = $GRUEN; 16 = $GRUEN; 17 = $GRUEN; 18 = $GRUEN; 19 = $GELB; 20 = $GRUEN; 21 = $GRUEN; 22 = $GRUEN; 23 = $GELB }
foreach ($r in $ampel.Keys) { $ws.Cells.Item($r, 4).Interior.Color = $ampel[$r] }
Kopf $ws.Range('A1:H1') $GRAU
Linien $ws.Range('A1:H23')
Breiten $ws @(13, 40, 11, 11, 8, 13, 12, 26)
[void]$ws.Range('A1:H23').AutoFilter()
Fixieren $wb $ws 1
Speichern $wb '02_Fristenliste_Wiedervorlage.xlsx'

#==============================================================================
# 03 Haushalt: Gruppenueberschriften, Zwischensummen, Gesamtsumme, Formeln, "k.A.", negativer Wert
#==============================================================================
$wb = NeueMappe @('Übersicht')
$ws = $wb.Worksheets.Item(1)
Titel $ws 'A1' 'Mittelabfluss Teilhaushalt 5 Jugend und Familie 2025'
$ws.Range('A2').Value2 = 'Beträge in Euro, Stand 30.06.2025'
[void](Put $ws 4 1 @(@('Konto', 'Bezeichnung', 'Ansatz 2025', 'Ist bis 30.06.', 'Rest', 'Abfluss in %'), @($null)))
$zeilen = @(
    @(5, 'G', 'Produkt 36.1 Förderung von Kindern in Tageseinrichtungen'),
    @(6, 'K', 4311000, 'Zuweisungen vom Land für Kitas', 12500000, 6100000),
    @(7, 'K', 5318100, 'Zuschüsse an freie Träger', 18400000, 9050000),
    @(8, 'K', 5241000, 'Bewirtschaftung Gebäude', 950000, 'k.A.'),
    @(9, 'K', 5431000, 'Geschäftsaufwendungen', 42000, 18350.4),
    @(10, 'S', 'Zwischensumme Produkt 36.1', 6, 9),
    @(12, 'G', 'Produkt 36.2 Kinder- und Jugendarbeit'),
    @(13, 'K', 5318200, 'Zuschüsse an Jugendverbände', 310000, 142000),
    @(14, 'K', 5422000, 'Ferienfreizeiten', 85000, 61200.75),
    @(15, 'K', 5291000, 'Honorare', 40000, -1200),
    @(16, 'S', 'Zwischensumme Produkt 36.2', 13, 15),
    @(18, 'G', 'Produkt 36.3 Hilfen zur Erziehung'),
    @(19, 'K', 5331000, 'Heimerziehung', 9800000, 5420000),
    @(20, 'K', 5332000, 'Sozialpädagogische Familienhilfe', 2100000, 1180000),
    @(21, 'K', 5333000, 'Eingliederungshilfe', 3500000, 2016500),
    @(22, 'S', 'Zwischensumme Produkt 36.3', 19, 21)
)
foreach ($z in $zeilen) {
    $r = $z[0]
    switch ($z[1]) {
        'G' { $ws.Cells.Item($r, 1).Value2 = $z[2]; $ws.Range("A${r}:F${r}").Merge(); Kopf $ws.Range("A${r}:F${r}") $BLAU }
        'K' { [void](Put $ws $r 1 @(@($z[2], $z[3], $z[4], $z[5], "=C$r-D$r", "=IF(C$r=0,`"`",D$r/C$r)"), @($null))) }
        'S' {
            $a = $z[3]; $e = $z[4]
            [void](Put $ws $r 1 @(@($z[2], $null, "=SUM(C${a}:C$e)", "=SUM(D${a}:D$e)", "=C$r-D$r", "=IF(C$r=0,`"`",D$r/C$r)"), @($null)))
            $ws.Range("A${r}:B${r}").Merge(); $ws.Range("A${r}:F${r}").Font.Bold = $true
        }
    }
}
[void](Put $ws 24 1 @(@('Gesamt', $null, '=C10+C16+C22', '=D10+D16+D22', '=C24-D24', '=IF(C24=0,"",D24/C24)'), @($null)))
$ws.Range('A24:B24').Merge(); Kopf $ws.Range('A24:F24') $GRAU
Kopf $ws.Range('A4:F4') $GRAU
Format $ws 'C5:E24' '#,##0.00'
Format $ws 'F5:F24' '0.0%'
Linien $ws.Range('A4:F24')
Breiten $ws @(11, 38, 15, 15, 15, 12)
Speichern $wb '03_Haushalt_Mittelabfluss.xlsx'

#==============================================================================
# 04 Belegung: Kreuztabelle (Monate als Spalten), leere Zukunftsmonate, Text in Zahlenfeld, Summenzeile
#==============================================================================
$wb = NeueMappe @('Belegung 2025')
$ws = $wb.Worksheets.Item(1)
Titel $ws 'A1' 'Belegte Plätze je Einrichtung 2025 (jeweils zum Monatsersten)'
$kitas = @(
    @('Kita Sonnenblume', 75, @(72, 73, 74, 75, 75, 74, 70, 68, 75)),
    @('Kita Regenbogen', 50, @(48, 48, 49, 50, 50, 50, 47, 45, 49)),
    @('Kita Pusteblume', 45, @(44, 45, 45, 45, 43, 44, 'Umbau', 'Umbau', 30)),
    @('Kita Kunterbunt', 100, @(96, 97, 99, 100, 100, 98, 94, 90, 99)),
    @('Kita Wirbelwind', 60, @(58, 59, 60, 60, 60, 59, 55, 54, 60)),
    @('Kita Löwenzahn', 120, @(115, 116, 118, 119, 120, 120, 112, 108, 118)),
    @('Kita Abenteuerland', 55, @(50, 51, 53, 54, 55, 55, 50, 49, 52)),
    @('Kita Sternschnuppe', 80, @(78, 79, 80, 80, 80, 79, 74, 72, 80)),
    @('Hort Kleine Strolche', 40, @(38, 38, 39, 40, 40, 40, 25, 22, 37)),
    @('Familienzentrum Nord', 35, @(30, 31, 33, 34, 35, 35, 31, 30, 33))
)
$rows = @(, @('Einrichtung', 'Plätze lt. Betriebserlaubnis', 'Jan', 'Feb', 'Mrz', 'Apr', 'Mai', 'Jun', 'Jul', 'Aug', 'Sep', 'Okt', 'Nov', 'Dez'))
foreach ($k in $kitas) { $rows += , (@($k[0], $k[1]) + $k[2] + @($null, $null, $null)) }
$summe = @('Summe')
foreach ($c in 'B', 'C', 'D', 'E', 'F', 'G', 'H', 'I', 'J', 'K', 'L', 'M', 'N') { $summe += "=SUM(${c}4:${c}13)" }
$rows += , $summe
[void](Put $ws 3 1 $rows)
Kopf $ws.Range('A3:N3') $GRAU
$ws.Range('B3').WrapText = $true
Kopf $ws.Range('A14:N14') $HELLGRAU
Linien $ws.Range('A3:N14')
Breiten $ws @(24, 18, 7, 7, 7, 7, 7, 7, 7, 7, 7, 7, 7, 7)
$ws.Range('C3:N14').HorizontalAlignment = -4108
Fixieren $wb $ws 3 1
Speichern $wb '04_Belegung_Kitas_Monate.xlsx'

#==============================================================================
# 05 Adressliste: Gruppenzeilen, Namen in verschiedenen Schreibweisen, Telefonformate, PLZ mit fuehrender Null
#==============================================================================
$wb = NeueMappe @('Telefonliste')
$ws = $wb.Worksheets.Item(1)
Titel $ws 'A1' 'Ansprechpersonen Netzwerk Frühe Hilfen'
$ws.Range('A2').Value2 = 'Stand: Mai 2025'
TextSpalte $ws 'C5:C20'
TextSpalte $ws 'F5:F20'
$liste = @(
    @('G', 'Stadtverwaltung'),
    @('Becker, Anna', 'Koordination Frühe Hilfen', '0123 45-3456', 'anna.becker@stadt.example', 'Rathausplatz 1', '55116', 'Musterstadt'),
    @('Dr. Martin Keller', 'Amtsleitung Jugendamt', '0123/45-2000', 'martin.keller@stadt.example', 'Rathausplatz 1', '55116', 'Musterstadt'),
    @('Fr. Özdemir', 'Sozialer Dienst', '0123 45-3480', 'sozialer.dienst@stadt.example', 'Am Markt 5', '55116', 'Musterstadt'),
    @('G', 'Freie Träger'),
    @('Lukas Brandt', 'Leitung Familienzentrum Nord', '0123 987654', 'l.brandt@familienzentrum-nord.example', 'Lindenstraße 12', '55122', 'Musterstadt'),
    @('Hebammenpraxis Sonnenschein', 'Familienhebammen', '0171 2345678', 'info@hebammen-sonnenschein.example', 'Gartenweg 3', '55124', 'Musterstadt'),
    @('Weber, Jonas', 'Erziehungsberatung', '+49 123 555012', 'beratung@eb-musterstadt.example', 'Kirchgasse 8', '55116', 'Musterstadt'),
    @('G', 'Gesundheitswesen'),
    @('Dr. med. Clara Vogt', 'Kinderärztin', '0123 44 55 66', 'praxis.vogt@aerzte.example', 'Bahnhofstraße 20', '55116', 'Musterstadt'),
    @('Gesundheitsamt, Frühförderung', 'Frühförderstelle', '0123 45-4400', 'fruehfoerderung@stadt.example', 'Am Markt 7', '55116', 'Musterstadt'),
    @('G', 'Externe Fachstellen'),
    @('Sandra Lorenz', 'Landeskoordination', '0351 123456', 's.lorenz@land.example', 'Elbufer 4', '01067', 'Dresden'),
    @('Fachstelle Kinderschutz', 'Beratung durch insoweit erfahrene Fachkraft', '0611 98 76 54', 'kinderschutz@fachstelle.example', 'Rheinallee 30', '65185', 'Wiesbaden')
)
[void](Put $ws 4 1 @(@('Name', 'Funktion', 'Telefon', 'E-Mail', 'Straße', 'PLZ', 'Ort'), @($null)))
$r = 5
foreach ($e in $liste) {
    if ($e[0] -eq 'G') { $ws.Cells.Item($r, 1).Value2 = $e[1]; $ws.Range("A${r}:G${r}").Merge(); Kopf $ws.Range("A${r}:G${r}") $BLAU }
    else { [void](Put $ws $r 1 @(, $e + @(, @($null)))) }
    $r++
}
Kopf $ws.Range('A4:G4') $GRAU
Linien $ws.Range("A4:G$($r - 1)")
Breiten $ws @(30, 40, 16, 38, 18, 7, 13)
Speichern $wb '05_Adressliste_Ansprechpersonen.xlsx'

#==============================================================================
# 06 Inventar: senkrecht verbundene Zellen, Musternummern, Seriennummern mit fuehrender Null,
#    Ja/Nein in Varianten, Synonyme (Laptop/Notebook), ausgeblendete Hilfsspalte
#==============================================================================
$wb = NeueMappe @('Inventar')
$ws = $wb.Worksheets.Item(1)
TextSpalte $ws 'B2:B18'
TextSpalte $ws 'F2:F18'
$daten = @(
    @('Gebäude', 'Raum', 'Inventar-Nr.', 'Gerät', 'Hersteller / Modell', 'Seriennummer', 'Anschaffung', 'Wert', 'defekt', 'Garantie bis', 'intern'),
    @('Rathaus', '1.12', 'INV-2023-0042', 'Laptop', 'Dell Latitude 5440', '5CG3421XQZ', (D 2023 3 14), 1149, 'nein', (D 2026 3 14), 'L-01'),
    @($null, '1.12', 'INV-2023-0043', 'Monitor', 'Dell P2423', 'CN0X7R2', (D 2023 3 14), 219.9, $null, (D 2026 3 14), 'M-01'),
    @($null, '1.14', 'INV-2022-0118', 'Drucker', 'HP LaserJet M507', '000731', (D 2022 9 2), 689, 'ja', (D 2024 9 2), 'D-03'),
    @($null, '1.14', 'INV-2024-0007', 'Notebook', 'Lenovo ThinkPad T14', 'PF4K2M9', (D 2024 1 22), 1289, 'nein', (D 2027 1 22), 'L-07'),
    @($null, '2.03', 'INV-2021-0350', 'Telefon', 'Snom D735', '00041A2B3C', (D 2021 6 30), 189, 'x', 'abgelaufen', 'T-12'),
    @($null, '2.03', 'INV-2024-0021', 'Dockingstation', 'Dell WD19S', 'DS19-88231', (D 2024 2 5), 179, $null, (D 2027 2 5), 'Z-02'),
    @($null, '2.05', 'INV-2023-0101', 'Laptop', 'HP EliteBook 840', '5CD3317KLM', (D 2023 8 17), 1199, 'nein', (D 2026 8 17), 'L-03'),
    @('Bürgerhaus', 'Saal', 'INV-2020-0210', 'Beamer', 'Epson EB-992F', 'X5KJ8800123', (D 2020 11 12), 749, 'ja', (D 2022 11 12), 'B-01'),
    @($null, 'Saal', 'INV-2022-0066', 'Lautsprecher', 'JBL EON 712', 'JB712-00456', (D 2022 4 4), 599, 'nein', (D 2024 4 4), 'A-02'),
    @($null, 'Büro', 'INV-2021-0402', 'Drucker', 'Brother MFC-L8900', 'E78123456', (D 2021 10 1), 529, $null, (D 2023 10 1), 'D-05'),
    @($null, 'Büro', 'INV-2024-0033', 'Laptop', 'Dell Latitude 5440', '5CG4110ABC', (D 2024 3 11), 1149, 'nein', (D 2027 3 11), 'L-09'),
    @($null, 'Foyer', 'INV-2019-0151', 'Monitor', 'Samsung S24', '0BFH3ZAN', (D 2019 5 20), 149, 'x', 'abgelaufen', 'M-06'),
    @('Kita Sonnenblume', 'Büro', 'INV-2023-0210', 'Laptop', 'Lenovo ThinkPad E14', 'PF3XA81', (D 2023 9 5), 899, 'nein', (D 2026 9 5), 'L-11'),
    @($null, 'Büro', 'INV-2023-0211', 'Drucker', 'HP OfficeJet Pro 9010', 'TH12345678', (D 2023 9 5), 249, $null, (D 2025 9 5), 'D-08'),
    @($null, 'Gruppe 1', 'INV-2024-0102', 'Tablet', 'Apple iPad 10', 'DMPX1234ABCD', (D 2024 7 1), 449, 'nein', (D 2025 7 1), 'T-20'),
    @($null, 'Gruppe 2', 'INV-2024-0103', 'Tablet', 'Apple iPad 10', 'DMPX5678EFGH', (D 2024 7 1), 449, 'ja', (D 2025 7 1), 'T-21'),
    @($null, 'Küche', 'INV-2018-0012', 'Telefon', 'Gigaset C430', '000098765', (D 2018 2 14), 59.9, 'nein', 'abgelaufen', 'T-02')
)
[void](Put $ws 1 1 $daten)
foreach ($a in 'A2:A8', 'A9:A13', 'A14:A18') { $ws.Range($a).Merge(); $ws.Range($a).VerticalAlignment = -4160 }
Format $ws 'G2:G18' 'dd.mm.yyyy'
Format $ws 'J2:J18' 'dd.mm.yyyy'
Format $ws 'H2:H18' '#,##0.00 "€"'
Kopf $ws.Range('A1:K1') $GRAU
Linien $ws.Range('A1:K18')
Breiten $ws @(17, 9, 15, 15, 22, 15, 12, 11, 7, 12, 7)
$ws.Columns.Item(11).Hidden = $true
Fixieren $wb $ws 1
Speichern $wb '06_Inventar_IT.xlsx'

#==============================================================================
# 07 Massnahmenplan: viel Freitext, Mehrfachwerte in einer Zelle mit wechselnden Trennzeichen,
#    unscharfe Zeitraeume und Betraege als Text, Zeilenumbrueche
#==============================================================================
$wb = NeueMappe @('Maßnahmen')
$ws = $wb.Worksheets.Item(1)
Titel $ws 'A1' 'Maßnahmenplan Sozialraum Nord 2025/2026'
TextSpalte $ws 'E4:F13'
$daten = @(
    @('Maßnahme', 'Beschreibung', 'Themen', 'Zielgruppe', 'Zeitraum', 'Kosten', 'verantwortlich', 'Stand'),
    @('Bewegungsfest im Park', 'Ein Nachmittag mit Mitmachstationen für Familien aus dem Quartier.', 'Sport, Gesundheit', 'Familien', 'Sommerferien 2025', 'ca. 5.000 €', 'Fr. Yilmaz', "Genehmigung Grünflächenamt liegt vor.`nFlyer in Arbeit."),
    @('Lesepaten an Grundschulen', 'Ehrenamtliche lesen wöchentlich mit Kindern der Klassen 1 und 2.', 'Bildung; Sprache', 'Kinder', 'laufend', '1.200 €', 'Herr Okafor', '8 Patinnen und Paten aktiv.'),
    @('Jugendtreff am Abend', 'Öffnung des Jugendhauses an zwei zusätzlichen Abenden.', 'Freizeit / Prävention', 'Jugendliche', 'ab Herbst 2025', '3.000 bis 4.000 €', 'Team Nord', 'Personal noch nicht gefunden.'),
    @('Sprachcafé für Eltern', 'Offener Treff mit Kinderbetreuung, Deutsch im Alltag.', 'Sprache, Integration, Bildung', 'Eltern', '03/2025 bis 06/2025', 'k.A.', 'Fr. Yilmaz', 'abgeschlossen, Fortsetzung geplant'),
    @('Seniorenfrühstück', 'Monatliches Frühstück im Bürgerhaus mit Themenimpulsen.', 'Begegnung', 'Senioren', 'laufend', '800', 'Herr Okafor', 'gut besucht (30 bis 40 Personen)'),
    @('Inklusive Ferienspiele', 'Ferienprogramm für Kinder mit und ohne Behinderung.', 'Inklusion, Freizeit', 'Kinder und Jugendliche', 'Sommerferien 2025', 'ca. 12.000 €', 'Team Nord', "Assistenzkräfte angefragt.`nBusunternehmen gebucht."),
    @('Quartiersspaziergang', 'Begehung mit Bewohnerinnen und Bewohnern zu Angsträumen.', 'Sicherheit; Stadtplanung', 'alle', 'Q2 2025', 'noch offen', 'Fr. Yilmaz', 'Termin abgestimmt'),
    @('Elternabend Medien', 'Informationsabend zur Mediennutzung in der Familie.', 'Medien, Bildung', 'Eltern', 'November 2025', '350 €', 'Herr Okafor', 'Referentin angefragt'),
    @('Kulturwerkstatt', 'Kreativangebote mit Künstlerinnen und Künstlern aus dem Stadtteil.', 'Kultur', 'Kinder, Jugendliche', 'ab 09/2025', 'ca. 6.500 €', 'Team Nord', 'Förderantrag gestellt'),
    @('Mittagstisch Schule', 'Gemeinsames Mittagessen mit Hausaufgabenhilfe.', 'Bildung / Ernährung', 'Kinder', 'laufend', '15.000 € pro Jahr', 'Team Nord', 'läuft')
)
[void](Put $ws 3 1 $daten)
Kopf $ws.Range('A3:H3') $GRAU
$ws.Range('A4:H13').WrapText = $true
$ws.Range('A4:H13').VerticalAlignment = -4160
Linien $ws.Range('A3:H13')
Breiten $ws @(24, 42, 22, 20, 18, 16, 15, 34)
[void]$ws.Range('A4:H13').Rows.AutoFit()
Speichern $wb '07_Massnahmenplan.xlsx'

#==============================================================================
# 08 Bedarfsmeldung je Kita: mehrere Blaetter mit (fast) gleichem Aufbau, Hinweisblatt,
#    Spaltensynonym (Menge/Anzahl), Zusatzspalte auf einem Blatt, Einheiten-Schreibweisen
#==============================================================================
$blaetter = @('Hinweise', 'Kita Sonnenblume', 'Kita Regenbogen', 'Kita Pusteblume', 'Kita Kunterbunt')
$wb = NeueMappe $blaetter
$ws = $wb.Worksheets.Item(1)
Titel $ws 'A1' 'Bedarfsmeldung Spiel- und Lernmaterial 2026'
$ws.Range('A3').Value2 = 'Bitte je Einrichtung ein Blatt ausfüllen und bis 15.10.2025 zurücksenden.'
$ws.Range('A4').Value2 = 'Preise bitte brutto angeben. Bei mehreren Lieferanten bitte den günstigsten wählen.'
$ws.Columns.Item(1).ColumnWidth = 90
$meldungen = @{
    'Kita Sonnenblume' = @(
        @('Artikel', 'Menge', 'Einheit', 'geschätzter Preis', 'Dringlichkeit', 'Begründung'),
        @('Bausteine Holz', 2, 'Set', 89.9, 'hoch', 'Bestand stark abgenutzt'),
        @('Bilderbücher Sprache', 15, 'Stück', 12.5, 'mittel', $null),
        @('Sandspielzeug', 1, 'Packung', 34, 'gering', $null),
        @('Turnmatten', 4, 'Stk.', 129, 'sehr hoch', 'Unfallgefahr durch Risse'),
        @('Malpapier A3', 10, 'Packung', 6.8, 'mittel', $null))
    'Kita Regenbogen' = @(
        @('Artikel', 'Anzahl', 'Einheit', 'geschätzter Preis', 'Dringlichkeit', 'Begründung'),
        @('Bilderbücher Sprache', 20, 'Stück', 12.5, 'hoch', 'Sprachkita'),
        @('Musikinstrumente Rhythmus', 1, 'Set', 149, 'mittel', $null),
        @('Kreide', 5, 'Packung', 4.2, 'gering', $null),
        @('Puppenhaus', 1, 'Stück', 219, 'mittel', 'Spende fällt weg'))
    'Kita Pusteblume' = @(
        @('Artikel', 'Menge', 'Einheit', 'geschätzter Preis', 'Dringlichkeit', 'Begründung'),
        @('Fahrzeuge Außengelände', 3, 'Stück', 79, 'hoch', $null),
        @('Bausteine Holz', 1, 'Set', 89.9, 'mittel', $null),
        @('Experimentierkasten', 2, 'Set', 59, 'mittel', 'Projekt Forscherkinder'),
        @('Malpapier A3', 8, 'Pck.', 6.8, 'gering', $null))
    'Kita Kunterbunt' = @(
        @('Artikel', 'Menge', 'Einheit', 'Lieferant', 'geschätzter Preis', 'Dringlichkeit', 'Begründung'),
        @('Turnmatten', 2, 'Stück', 'Sportversand', 129, 'hoch', $null),
        @('Bilderbücher Sprache', 10, 'Stück', 'Buchhandlung am Markt', 12.5, 'mittel', $null),
        @('Knete', 6, 'Dose', 'Bastelbedarf', 3.9, 'gering', $null))
}
for ($i = 2; $i -le 5; $i++) {
    $ws = $wb.Worksheets.Item($i)
    $name = $blaetter[$i - 1]
    Titel $ws 'A1' ("Bedarfsmeldung 2026: " + $name) 12
    $d = $meldungen[$name]
    [void](Put $ws 3 1 $d)
    $breit = @($d[0]).Count
    $ende = 2 + $d.Count
    $lc = [char](64 + $breit)
    Kopf $ws.Range("A3:${lc}3") $GRAU
    Linien $ws.Range("A3:${lc}$ende")
    $preisSpalte = if ($breit -eq 7) { 'E' } else { 'D' }
    Format $ws "${preisSpalte}4:${preisSpalte}$ende" '#,##0.00 "€"'
    if ($breit -eq 7) { Breiten $ws @(26, 8, 9, 22, 16, 13, 28) } else { Breiten $ws @(26, 8, 9, 16, 13, 28) }
}
Speichern $wb '08_Bedarfsmeldung_je_Kita.xlsx'

#==============================================================================
# 09 Projekte und Aufgaben: zwei zusammengehoerige Blaetter, Verknuepfung ueber den Projektnamen
#    (mit abweichenden Schreibweisen)
#==============================================================================
$wb = NeueMappe @('Projekte', 'Aufgaben')
$ws = $wb.Worksheets.Item(1)
$daten = @(
    @('Nr.', 'Projekt', 'Leitung', 'Beginn', 'Ende', 'Budget', 'Status'),
    @(1, 'Umbau Stadtteilbibliothek', 'A. Berger', (D 2025 8 4), (D 2025 11 21), 17000, 'aktiv'),
    @(2, 'Digitalisierung Antragswesen', 'M. Kaya', (D 2025 8 27), (D 2025 12 2), 26000, 'aktiv'),
    @(3, 'Sommerprogramm Jugend', 'S. Nowak', (D 2025 6 1), (D 2025 9 10), 23000, 'abgeschlossen'),
    @(4, 'Neues Leitsystem Rathaus', 'J. Lindqvist', (D 2025 10 2), (D 2026 2 27), 8000, 'geplant'),
    @(5, 'Barrierefreie Website', 'F. Hartmann', (D 2025 9 1), (D 2026 3 31), 35000, 'aktiv'),
    @(6, 'Energieberatung Quartier Süd', 'T. Okafor', (D 2025 10 28), (D 2026 1 12), 25000, 'aktiv')
)
[void](Put $ws 1 1 $daten)
Format $ws 'D2:E7' 'dd.mm.yyyy'
Format $ws 'F2:F7' '#,##0 "€"'
Kopf $ws.Range('A1:G1') $GRAU
Linien $ws.Range('A1:G7')
Breiten $ws @(5, 32, 13, 11, 11, 11, 14)
$ws = $wb.Worksheets.Item(2)
$daten = @(
    @('Aufgabe', 'Projekt', 'zuständig', 'Frist', 'erledigt'),
    @('Ausschreibung vorbereiten', 'Umbau Stadtteilbibliothek', 'T. Okafor', (D 2025 8 18), 'x'),
    @('Angebote einholen', 'Umbau Stadtteilbibliothek', 'J. Lindqvist', (D 2025 8 24), 'x'),
    @('Möblierung bestellen', 'Umbau Bibliothek', 'J. Lindqvist', (D 2025 9 30), $null),
    @('Formulare erfassen', 'Digitalisierung Antragswesen', 'M. Kaya', (D 2025 9 15), 'x'),
    @('Schnittstelle testen', 'Digitalisierung Antragswesen', 'M. Kaya', (D 2025 10 20), $null),
    @('Schulung Sachbearbeitung', 'Digitalisierung Antragswesen', 'S. Nowak', (D 2025 11 25), $null),
    @('Programmheft drucken', 'Sommerprogramm Jugend', 'S. Nowak', (D 2025 5 20), 'x'),
    @('Abrechnung Ferienfahrten', 'Sommerprogramm Jugend', 'S. Nowak', (D 2025 9 30), 'x'),
    @('Schilderplan abstimmen', 'Neues Leitsystem Rathaus', 'J. Lindqvist', (D 2025 10 30), $null),
    @('Prüfung nach BITV', 'Website', 'F. Hartmann', (D 2025 11 15), $null),
    @('Texte in Leichter Sprache', 'Barrierefreie Website', 'F. Hartmann', (D 2026 1 31), $null),
    @('Infoabend planen', 'Energieberatung Quartier Süd', 'T. Okafor', (D 2025 11 5), $null),
    @('Flyer Energieberatung', 'Energieberatung Quartier Süd', 'T. Okafor', (D 2025 11 1), 'x'),
    @('Raum für Beratung buchen', 'Energieberatung Quartier Süd', 'M. Kaya', 'vor Infoabend', $null)
)
[void](Put $ws 1 1 $daten)
Format $ws 'D2:D15' 'dd.mm.yyyy'
$ws.Range('E2:E15').HorizontalAlignment = -4108
Kopf $ws.Range('A1:E1') $GRAU
Linien $ws.Range('A1:E15')
Breiten $ws @(28, 32, 13, 13, 9)
Speichern $wb '09_Projekte_und_Aufgaben.xlsx'

#==============================================================================
# 10 Beschlusskontrolle: Datum mit Uhrzeit, gemischte TOP-Angaben, Abkuerzungen (JHA), vage Fristen, lange Texte
#==============================================================================
$wb = NeueMappe @('Beschlüsse')
$ws = $wb.Worksheets.Item(1)
TextSpalte $ws 'B2:B9'
$daten = @(
    @('Sitzung', 'TOP', 'Gremium', 'Beschluss', 'zuständig', 'Umsetzung bis', 'Stand der Umsetzung'),
    @((D 2025 2 13 17 0), 'TOP 4', 'Jugendhilfeausschuss', 'Der Ausschuss beschließt die Fortschreibung der Kita-Bedarfsplanung für die Jahre 2025 bis 2030.', 'Amt 51', (D 2025 6 30), 'Entwurf erstellt, Beteiligung der Träger läuft.'),
    @((D 2025 2 13 17 0), '5.1', 'JHA', 'Der Zuschuss für das Familienzentrum Nord in Höhe von 45.000 € wird bewilligt.', 'Amt 51', (D 2025 3 31), 'erledigt, Bescheid versandt am 20.03.2025'),
    @((D 2025 3 20 16 30), 'TOP 2', 'Sozialausschuss', 'Die Verwaltung prüft ein Konzept für Wärmeräume im Winter.', 'Amt 50', 'nach den Sommerferien', 'Konzept in Abstimmung mit den Wohlfahrtsverbänden.'),
    @('Mai 2025', '7', 'Stadtrat', 'Der Stadtrat beauftragt die Verwaltung, Schulwegpläne für alle Grundschulen zu erstellen.', 'Amt 40, Amt 61', (D 2025 12 31), "Pilot an zwei Schulen gestartet.`nRest folgt 2026."),
    @((D 2025 6 5 17 0), 'TOP 3', 'Jugendhilfeausschuss', 'Die Richtlinie zur Förderung der Jugendverbandsarbeit wird geändert.', 'Amt 51', (D 2025 7 1), 'in Kraft seit 01.07.2025'),
    @((D 2025 6 5 17 0), 'TOP 6', 'JHA', 'Die Verwaltung berichtet halbjährlich über die Fallzahlen in den Hilfen zur Erziehung.', 'Amt 51', 'fortlaufend', 'erster Bericht im Dezember'),
    @((D 2025 7 10 18 0), 'TOP 1', 'Sozialausschuss', 'Eine Koordinierungsstelle Pflege wird eingerichtet.', 'Amt 50', (D 2026 1 31), 'Stellenausschreibung läuft.'),
    @((D 2025 9 18 17 0), 'TOP 2', 'Stadtrat', 'Die Gebühren für die Ganztagsbetreuung bleiben 2026 unverändert.', 'Amt 40', (D 2026 1 1), $null)
)
[void](Put $ws 1 1 $daten)
Format $ws 'A2:A9' 'dd.mm.yyyy hh:mm'
Txt $ws 5 1 'Mai 2025'
Format $ws 'F2:F9' 'dd.mm.yyyy'
Kopf $ws.Range('A1:G1') $GRAU
$ws.Range('D2:D9').WrapText = $true; $ws.Range('G2:G9').WrapText = $true
$ws.Range('A2:G9').VerticalAlignment = -4160
Linien $ws.Range('A1:G9')
Breiten $ws @(16, 7, 21, 50, 14, 22, 38)
[void]$ws.Range('A2:G9').Rows.AutoFit()
Speichern $wb '10_Beschlusskontrolle.xlsx'

#==============================================================================
# 12 Antrag: formularartiges Blatt (ein Datensatz untereinander) mit kleiner Kostentabelle
#==============================================================================
$wb = NeueMappe @('Antrag') 'Arial' 10
$ws = $wb.Worksheets.Item(1)
Titel $ws 'A1' 'Antrag auf Zuschuss für eine Ferienfreizeit' 14
TextSpalte $ws 'B5:B5'
[void](Put $ws 3 1 @(
    @('Antragsteller:', 'Jugendverein Musterstadt e. V.'),
    @('Ansprechperson:', 'Jonas Weber'),
    @('Telefon:', '0123 456789'),
    @('E-Mail:', 'vorstand@jugendverein.example'),
    @('Maßnahme:', 'Zeltlager am See'),
    @('Zeitraum:', '21.07.2025 bis 01.08.2025'),
    @('Teilnehmende:', 24),
    @('davon mit Behinderung:', 2),
    @('Beantragter Zuschuss:', 2400)
))
Titel $ws 'A13' 'Kostenplan' 11
[void](Put $ws 14 1 @(
    @('Position', 'Betrag'),
    @('Unterkunft', 3200),
    @('Verpflegung', 1800),
    @('Fahrtkosten', 900),
    @('Programm und Material', 600),
    @('Summe', '=SUM(B15:B18)')
))
$ws.Range('A3:A11').Font.Bold = $true
Format $ws 'B11' '#,##0.00 "€"'
Format $ws 'B15:B19' '#,##0.00 "€"'
Kopf $ws.Range('A14:B14') $GRAU
$ws.Range('A19:B19').Font.Bold = $true
Linien $ws.Range('A14:B19')
$ws.Range('A21').Value2 = 'Musterstadt, 03.03.2025'
$ws.Range('A23').Value2 = 'Unterschrift: ______________________'
$ws.Range('B3:B11').HorizontalAlignment = -4131
Breiten $ws @(26, 36)
Speichern $wb '12_Antrag_Formular.xlsx'

#==============================================================================
# 13 Raumbelegung: Wochenraster (Uhrzeiten x Wochentage), Buchungen als verbundene, gefaerbte Zellen
#==============================================================================
$wb = NeueMappe @('KW 20')
$ws = $wb.Worksheets.Item(1)
Titel $ws 'A1' 'Raumbelegung Bürgerhaus, Saal: KW 20 (12. bis 16. Mai 2025)'
TextSpalte $ws 'A4:A14'
$rows = @(, @('Uhrzeit', 'Montag', 'Dienstag', 'Mittwoch', 'Donnerstag', 'Freitag'))
for ($h = 8; $h -le 18; $h++) { $rows += , @(('{0:00}:00' -f $h), $null, $null, $null, $null, $null) }
[void](Put $ws 3 1 $rows)
# Buchungen: Spalte, erste Zeile, letzte Zeile, Text, Farbe (Zeile 4 = 08:00)
$buchungen = @(
    @(2, 5, 6, 'Eltern-Kind-Gruppe', $GRUEN),
    @(2, 13, 14, 'Chorprobe', $LILA),
    @(3, 10, 11, 'Seniorentanz', $ORANGE),
    @(4, 6, 8, 'Sprachkurs A1', $TUERKIS),
    @(4, 12, 12, 'Elternberatung', $GELB),
    @(5, 4, 5, 'Reinigung', $HELLGRAU),
    @(5, 14, 14, 'Ortsbeirat', $BLAU),
    @(6, 11, 13, 'Jugendband Probe', $ROT)
)
foreach ($b in $buchungen) {
    $rg = $ws.Range($ws.Cells.Item($b[1], $b[0]), $ws.Cells.Item($b[2], $b[0]))
    $ws.Cells.Item($b[1], $b[0]).Value2 = $b[3]
    $rg.Merge()
    $rg.Interior.Color = $b[4]
    $rg.HorizontalAlignment = -4108; $rg.VerticalAlignment = -4108; $rg.WrapText = $true
}
Kopf $ws.Range('A3:F3') $GRAU
$ws.Range('A3:F3').HorizontalAlignment = -4108
Linien $ws.Range('A3:F14')
Breiten $ws @(9, 18, 18, 18, 18, 18)
for ($r = 4; $r -le 14; $r++) { $ws.Rows.Item($r).RowHeight = 22 }
Speichern $wb '13_Raumbelegung_Wochenplan.xlsx'

#==============================================================================
# 2ask: Nachbau des echten Exports. Aufbau geprueft mit einer Testbefragung (alle Fragearten
# des Testtarifs, November 2026), Originale in testdaten\importlisten\2ask_echt.
# XLSX: Blatt "results" mit vier Kopfzeilen (Fragetext, Unterbeschriftung, Import-JSON,
#       Variablenname), ab Zeile 4 eine Excel-Tabelle "results"; 24 Metadatenspalten;
#       Blatt "code_plan" (Langformat) nur beim kodierten Export mit Codeplan.
# CSV:  UTF-8 mit BOM, Semikolon, Datensaetze mit LF getrennt, Kopfzeilen in anderer Reihenfolge
#       (Variablenname, Fragetext, Unterbeschriftung, Import-JSON), Dezimalkomma, Texte und Codes
#       in Anfuehrungszeichen, ganze Zahlen der Metadaten ohne, leere Felder ganz leer.
# Nicht beantwortet = leer (Standard). 24 Pretest-Teilnehmende mit erfundenen Antworten.
#==============================================================================
$rnd = New-Object System.Random 2025
function Zufall([int]$a, [int]$b) { return $rnd.Next($a, $b + 1) }
function Chance([double]$p) { return ($rnd.NextDouble() -lt $p) }
function Pick([object[]]$arr) { return $arr[$rnd.Next($arr.Count)] }
function Kennung([int]$n) { $z = '0123456789ABCDEFGHJKMNPQRSTVWXYZ'; $s = ''; for ($i = 0; $i -lt $n; $i++) { $s += $z[$rnd.Next($z.Length)] }; return $s }
function B64([string]$s) { return [Convert]::ToBase64String([Text.Encoding]::UTF8.GetBytes($s)) }
function Buchstabe([int]$n) { $s = ''; while ($n -gt 0) { $m = ($n - 1) % 26; $s = [string][char](65 + $m) + $s; $n = [int](($n - $m - 1) / 26) }; return $s }
function JsonId([string]$pfad, [string]$extraName = '', [string]$extraWert = '') {
    $j = '{"import_id":"' + (B64 $pfad) + '"'
    if ($extraName) { $j += ',"' + $extraName + '":"' + $extraWert + '"' }
    return $j + '}'
}
function Pfad([string]$qid, [string]$aid) { return '$.survey.questions."' + $qid + '".answers."' + $aid + '"' }
$spalten = New-Object System.Collections.Generic.List[object]
$codeplan = New-Object System.Collections.Generic.List[object]
function NeueSpalte([string]$key, [string]$text, [string]$sub, [string]$json) { $spalten.Add([pscustomobject]@{ Key = $key; Text = $text; Sub = $sub; Json = $json }) }
function Plan($key, $typ, $sub, $daten, $code, $text) { $codeplan.Add(@($key, $typ, $sub, $daten, $code, $text)) }

# --- Metadaten (Schluessel, Beschriftung, JSON-Pfad, Datentyp) ---
$meta = @(
    @('survey_session_id', 'ID', '$.survey_session_id', 'string'),
    @('survey_session_type', 'Type', '', 'string'),
    @('started_at', 'Started at', '$.started_at', 'date'),
    @('ended_at', 'Ended at', '$.ended_at', 'date'),
    @('last_interaction_at', 'Last interaction at', '$.last_interaction_at', 'date'),
    @('progress_percent', 'Progress', '$.survey.progress_percent', 'number'),
    @('duration_in_seconds', 'Duration in seconds', '$.duration_in_seconds', 'number'),
    @('items_answered_count', 'Items answered count', '$.items_answered_count', 'number'),
    @('items_answered_percent', 'Items answered percent', '$.items_answered_percent', 'number'),
    @('locale', 'Locale (Survey language)', '$.survey.locale', 'string'),
    @('source', 'Respondent source', '$.respondent.source', 'string'),
    @('country', 'Respondent country', '$.respondent.country', 'string'),
    @('device_type', 'Device', '$.respondent.device_type', 'string'),
    @('browser', 'Browser', '$.respondent.browser', 'string'),
    @('browser_locale', 'Browser Locale', '$.respondent.language', 'string'),
    @('browser_version', 'Browser Version', '$.respondent.browser_version', 'string'),
    @('operating_system', 'Operating System', '$.respondent.operating_system', 'string'),
    @('contact.id', 'Contact ID', '$.respondent.contact.contact_id', 'string'),
    @('contact.external_key', 'Contact external key', '$.respondent.contact.external_key', 'string'),
    @('contact.email', 'Contact email', '$.respondent.contact.email', 'string'),
    @('contact.first_name', 'Contact first name', '$.respondent.contact.first_name', 'string'),
    @('contact.last_name', 'Contact last name', '$.respondent.contact.last_name', 'string'),
    @('contact.language', 'Contact language', '$.respondent.contact.language', 'string'),
    @('contact.gender', 'Contact gender', '$.respondent.contact.gender', 'string')
)
foreach ($m in $meta) {
    $json = if ($m[2]) { JsonId $m[2] } else { '{}' }
    NeueSpalte $m[0] $m[1] '' $json
    Plan $m[0] 'standard_field' $null $m[3] $null $null
}

# --- Fragen (Texte und Optionen wie in der echten Testbefragung) ---
$q1 = @('Innenstadt', 'Nord', 'Ost', 'Süd', 'West')
$qt = 'In welchem Stadtbezirk liegt Ihre Einrichtung?'
NeueSpalte 'Q1' $qt '' (JsonId (Pfad (Kennung 13) (Kennung 13)))
Plan 'Q1' 'question' 'multiple_choice' 'number' $null $qt
for ($i = 0; $i -lt $q1.Count; $i++) { Plan 'Q1' 'answer_option' 'radio' 'number' ($i + 1) $q1[$i] }

$q2 = @('Fachberatung', 'Fortbildungen', 'Supervision', 'Netzwerktreffen', 'Sonstiges, und zwar')
$qt = 'Welche Unterstützungsangebote haben Sie 2025 genutzt?'
$qid = Kennung 13; $aid = Kennung 13; $oid = Kennung 13
NeueSpalte 'Q2t_05' $qt 'Answer-Option: Sonstiges, und zwar (Label)' (JsonId ('$.survey.questions."' + $qid + '".answer_options."' + $oid + '".text_entry'))
Plan 'Q2t_05' 'question' 'multiple_choice' $null $null $qt
Plan 'Q2t_05' 'text_entry' $null 'string' $null $null
for ($i = 0; $i -lt $q2.Count; $i++) {
    $cid = if ($i -eq 4) { $oid } else { Kennung 13 }
    $key = 'Q2_{0:00}' -f ($i + 1)
    NeueSpalte $key $qt ('Answer-Option: ' + $q2[$i]) (JsonId (Pfad $qid $aid) 'import_choice_id' $cid)
    Plan $key 'answer_option' 'checkbox' 'number' ($i + 1) $q2[$i]
}

$q3 = @('Kommunal', 'Kirchlich', 'Freier Träger', 'Elterninitiative')
$qt = 'Wer ist Träger Ihrer Einrichtung?'
NeueSpalte 'Q3' $qt '' (JsonId (Pfad (Kennung 13) (Kennung 13)))
Plan 'Q3' 'question' 'multiple_choice' 'number' $null $qt
for ($i = 0; $i -lt $q3.Count; $i++) { Plan 'Q3' 'answer_option' 'select' 'number' ($i + 1) $q3[$i] }

# Matrix Likert; die dritte Zeile wurde im Editor geloescht, daher Q4_1, Q4_2, Q4_4
$skala = @('trifft voll zu', 'trifft eher zu', 'teils-teils', 'trifft eher nicht zu', 'trifft gar nicht zu', 'kann ich nicht beurteilen')
$q4 = @(@(1, 'Die Fachberatung ist gut erreichbar.'), @(2, 'Die Fortbildungen sind praxisnah.'), @(4, 'Informationen kommen rechtzeitig an.'))
$qt = 'Wie sehr treffen die folgenden Aussagen zu?'
$qid = Kennung 13
foreach ($z in $q4) {
    $key = 'Q4_' + $z[0]
    NeueSpalte $key $qt ('Answer: ' + $z[1]) (JsonId (Pfad $qid (Kennung 13)))
    if ($z[0] -eq 1) { Plan $key 'question' 'matrix_table' $null $null $qt }
    Plan $key 'answer' 'radio' 'number' $null $z[1]
    for ($i = 0; $i -lt $skala.Count; $i++) { Plan $key 'answer_option' 'radio' 'number' ($i + 1) $skala[$i] }
}

# Matrix Mehrfachauswahl 3 x 3
$themen = @('Sprache', 'Bewegung', 'Inklusion')
$alter = @('unter 3 Jahren', '3 bis 6 Jahre', 'Schulkinder')
$qt = 'Welche Themen sind für welche Altersgruppe wichtig?'
$qid = Kennung 13; $cids = @((Kennung 13), (Kennung 13), (Kennung 13))
for ($r = 0; $r -lt 3; $r++) {
    $aid = Kennung 13
    for ($c = 0; $c -lt 3; $c++) {
        $key = 'Q5_{0}_{1:00}' -f ($r + 1), ($c + 1)
        NeueSpalte $key $qt ('Answer: ' + $themen[$r] + '; Answer-Option:' + $alter[$c]) (JsonId (Pfad $qid $aid) 'import_choice_id' $cids[$c])
        if ($r -eq 0 -and $c -eq 0) { Plan $key 'question' 'matrix_table' $null $null $qt }
        Plan $key 'answer_option' 'checkbox' 'number' ($c + 1) $alter[$c]
    }
}

# Matrix Textfelder 2 x 2
$bereiche = @('Personal', 'Räume')
$q6sp = @('Anliegen', 'Was würde helfen?')
$qt = 'Was ist Ihr wichtigstes Anliegen im jeweiligen Bereich?'
$qid = Kennung 13; $cids = @((Kennung 13), (Kennung 13))
for ($r = 0; $r -lt 2; $r++) {
    $aid = Kennung 13
    for ($c = 0; $c -lt 2; $c++) {
        $key = 'Q6_{0}_{1:00}' -f ($r + 1), ($c + 1)
        NeueSpalte $key $qt ('Answer: ' + $bereiche[$r] + '; Answer-Option:' + $q6sp[$c]) (JsonId (Pfad $qid $aid) 'import_choice_id' $cids[$c])
        if ($r -eq 0 -and $c -eq 0) { Plan $key 'question' 'matrix_table' $null $null $qt }
        Plan $key 'answer_option' 'input' 'string' ($c + 1) $q6sp[$c]
    }
}

# Matrix Dropdown
$oft = @('nie', 'selten', 'oft')
$angebote = @('Fachberatung', 'Fortbildungen')
$qt = 'Wie oft nutzen Sie die folgenden Angebote?'
$qid = Kennung 13
for ($r = 0; $r -lt 2; $r++) {
    $key = 'Q7_' + ($r + 1)
    NeueSpalte $key $qt ('Answer: ' + $angebote[$r]) (JsonId (Pfad $qid (Kennung 13)))
    if ($r -eq 0) { Plan $key 'question' 'matrix_table' $null $null $qt }
    Plan $key 'answer' 'select' 'number' $null $angebote[$r]
    for ($i = 0; $i -lt 3; $i++) { Plan $key 'answer_option' 'select' 'number' ($i + 1) $oft[$i] }
}

# Matrix Bipolar 1 bis 5 (Optionen ohne Beschriftung, unkodiert "[n]")
$pole = @(@('kompliziert', 'unkompliziert'), @('langsam', 'schnell'))
$qt = 'Die Zusammenarbeit mit dem Jugendamt ist ...'
$qid = Kennung 13
for ($r = 0; $r -lt 2; $r++) {
    $key = 'Q8_' + ($r + 1)
    NeueSpalte $key $qt ('Answer: ' + $pole[$r][0] + ' - ' + $pole[$r][1]) (JsonId (Pfad $qid (Kennung 13)))
    if ($r -eq 0) { Plan $key 'question' 'matrix_table' $null $null $qt }
    Plan $key 'answer' 'radio' 'number' $null $pole[$r][0]
    for ($i = 1; $i -le 5; $i++) { Plan $key 'answer_option' 'radio' 'number' $i ('[' + $i + ']') }
}

# Sterne (zwei Zeilen)
$sterne = @('Inhalte', 'Organisation')
$qt = 'Wie bewerten Sie die Fortbildungen?'
$qid = Kennung 13
for ($r = 0; $r -lt 2; $r++) {
    $key = 'Q9_' + ($r + 1)
    NeueSpalte $key $qt ('Answer: ' + $sterne[$r]) (JsonId (Pfad $qid (Kennung 13)))
    if ($r -eq 0) { Plan $key 'question' 'slider' $null $null $qt }
    Plan $key 'answer' 'range' 'number' $null $sterne[$r]
}

# Schieberegler 0 bis 100 (Schrittweite 10)
$qt = 'Wie zufrieden sind Sie insgesamt mit der Zusammenarbeit?'
NeueSpalte 'Q10' $qt '' (JsonId (Pfad (Kennung 13) (Kennung 13)))
Plan 'Q10' 'question' 'slider' 'number' $null $qt
Plan 'Q10' 'answer' 'range' 'number' $null 'Zufriedenheit'

# Formular mit Dropdown
$funktion = @('Leitung', 'stellvertretende Leitung', 'Fachkraft')
$qt = 'Ansprechperson für Rückfragen'
$qid = Kennung 13
NeueSpalte 'Q11_1' $qt 'Answer: Name' (JsonId (Pfad $qid (Kennung 13)))
NeueSpalte 'Q11_2' $qt 'Answer: Telefon' (JsonId (Pfad $qid (Kennung 13)))
NeueSpalte 'Q11_3' $qt 'Answer: Funktion' (JsonId (Pfad $qid (Kennung 13)))
Plan 'Q11_1' 'question' 'form' $null $null $qt
Plan 'Q11_1' 'answer' 'input' 'string' $null 'Name'
Plan 'Q11_2' 'answer' 'input' 'string' $null 'Telefon'
Plan 'Q11_3' 'answer' 'select_coded' 'number' $null 'Funktion'
for ($i = 0; $i -lt 3; $i++) { Plan 'Q11_3' 'answer_option' 'select_coded' 'number' ($i + 1) $funktion[$i] }

# Gruppenmatrix: zwei Zeilen, zwei Gruppen; Spalten gruppenweise
$wicht = @('sehr wichtig', 'wichtig', 'weniger wichtig')
$zufr = @('zufrieden', 'teils-teils', 'unzufrieden')
$gruppen = @(@('Wichtigkeit', $wicht), @('Zufriedenheit', $zufr))
$qt = 'Wie wichtig ist Ihnen das Angebot und wie zufrieden sind Sie damit?'
$qid = Kennung 13; $aids = @((Kennung 13), (Kennung 13)); $gids = @((Kennung 13), (Kennung 13))
for ($g = 0; $g -lt 2; $g++) {
    for ($r = 0; $r -lt 2; $r++) {
        $key = 'Q12_{0}_{1:00}' -f ($r + 1), ($g + 1)
        NeueSpalte $key $qt ('Answer: ' + $angebote[$r] + '; Group: ' + $gruppen[$g][0]) (JsonId (Pfad $qid $aids[$r]) 'import_group_id' $gids[$g])
        if ($g -eq 0 -and $r -eq 0) { Plan $key 'question' 'side_by_side' $null $null $qt }
        if ($r -eq 0) { Plan $key 'group' 'radio' $null $null $gruppen[$g][0] }
        Plan $key 'answer' 'radio' 'number' $null $angebote[$r]
        for ($i = 0; $i -lt 3; $i++) { Plan $key 'answer_option' 'radio' 'number' ($i + 1) $gruppen[$g][1][$i] }
    }
}

# Rangfolge: je Thema eine Spalte mit dem Rang
$rang = @('Personalgewinnung', 'Räume und Ausstattung', 'Finanzierung', 'Qualitätsentwicklung')
$qt = 'Bitte bringen Sie die Themen in eine Rangfolge.'
$qid = Kennung 13; $aid = Kennung 13
for ($i = 0; $i -lt 4; $i++) {
    $key = 'Q13_{0:00}' -f ($i + 1)
    NeueSpalte $key $qt ('Answer-Option: ' + $rang[$i]) (JsonId (Pfad $qid $aid) 'import_choice_id' (Kennung 13))
    if ($i -eq 0) { Plan $key 'question' 'rank_order' $null $null $qt }
    Plan $key 'answer_option' 'drag_drop' 'number' ($i + 1) $rang[$i]
}

# Summe (Arbeitszeit in Prozent)
$zeit = @('Leitung und Verwaltung', 'Arbeit mit Kindern', 'Elternarbeit')
$qt = 'Wie verteilt sich Ihre Arbeitszeit (in Prozent)?'
$qid = Kennung 13
for ($i = 0; $i -lt 3; $i++) {
    $key = 'Q14_' + ($i + 1)
    NeueSpalte $key $qt ('Answer: ' + $zeit[$i]) (JsonId (Pfad $qid (Kennung 13)))
    if ($i -eq 0) { Plan $key 'question' 'constant_sum' $null $null $qt }
    Plan $key 'answer' 'number' 'number' $null $zeit[$i]
}

# Eingabefelder: einzeilig, mehrzeilig, Zahl, Datum, Zeit, E-Mail
$eingabe = @(
    @('Q15', 'Name der Einrichtung', 'string'),
    @('Q16', 'Was wünschen Sie sich für die Zusammenarbeit im kommenden Jahr?', 'string'),
    @('Q17', 'Wie viele Kinder betreut Ihre Einrichtung?', 'number'),
    @('Q18', 'Seit wann leiten Sie die Einrichtung?', 'localdate'),
    @('Q19', 'Ab welcher Uhrzeit ist Ihre Einrichtung geöffnet?', 'localtime'),
    @('Q20', 'E-Mail-Adresse für Rückfragen (freiwillig)', 'string')
)
foreach ($e in $eingabe) {
    NeueSpalte $e[0] $e[1] '' (JsonId (Pfad (Kennung 13) (Kennung 13)))
    Plan $e[0] 'question' 'text_entry' $e[2] $null $e[1]
}

# --- 24 Teilnehmende ---
$sonstiges = @('Kollegiale Beratung', 'Fachtag Inklusion; Austausch mit Grundschule', 'Beratung durch den Träger', 'Hospitation in einer anderen Kita')
$texte6 = @{
    'Q6_1_01' = @('Fachkräfte finden', 'Vertretung bei Krankheit', 'mehr Stunden für die Leitung')
    'Q6_1_02' = @('Springerpool', 'mehr Ausbildungsplätze', 'Quereinstieg erleichtern')
    'Q6_2_01' = @('Schlafraum fehlt', 'Lärmschutz im Speiseraum', 'Außengelände sanieren')
    'Q6_2_02' = @('Akustikdecke', 'Container als Übergang', 'Mittel aus dem Bauprogramm')
}
$namen = @('Petra Busch', 'Jens Albrecht', 'Ayşe Kaya', 'Sabine Roth', 'Kemal Arslan', 'Marta Nowicka', 'Ute Scholz', 'Nils Hansen')
$telefone = @('0123 456789', '+49 123 555012', '0123 45-3456', '0171 2345678', '0123/987654')
$einrichtungen = @('Kita Sonnenblume', 'Kita Regenbogen', 'Kita Löwenzahn', 'Kita Pusteblume', 'Kita Kunterbunt', 'Kita Wirbelwind', 'Kita Sternschnuppe', 'Hort Kleine Strolche')
$mails = @('leitung@kita-sonnenblume.example', 'kontakt@kita-loewenzahn.example', 'info@kita-regenbogen.example', 'leitung@kita-pusteblume.example')
$wuensche = @(
    "Mehr Fortbildungen vor Ort.`nKürzere Wege bei Anträgen; schnellere Rückmeldung.",
    'Bitte "Elternbriefe" in einfacher Sprache; danke!',
    'Feste Ansprechperson in der Fachberatung.',
    "Regelmäßiger Austausch der Leitungen.`nRückmeldung zu unseren Meldungen.",
    'Schnellere Bearbeitung der Förderanträge.'
)
$teilnehmer = New-Object System.Collections.Generic.List[object]
# Wert setzen: kodiert und unkodiert (ohne dritten Wert gleich)
function W([string]$key, $kod, $txt = '__gleich__') {
    $kv[$key] = $kod
    if ($txt -is [string] -and $txt -eq '__gleich__') { $tv[$key] = $kod } else { $tv[$key] = $txt }
}
for ($n = 0; $n -lt 24; $n++) {
    $kv = @{}; $tv = @{}; $items = 0
    $beginn = New-Object DateTime 2025, 5, (12 + $n % 10), (Zufall 7 17), (Zufall 0 59), (Zufall 0 59)
    $dauer = Zufall 240 1500
    $ende = $beginn.AddSeconds($dauer)
    W 'survey_session_id' (Kennung 13); W 'survey_session_type' 'pretest'
    W 'started_at' $beginn; W 'ended_at' $ende; W 'last_interaction_at' $ende
    W 'progress_percent' 100; W 'duration_in_seconds' $dauer
    W 'locale' 'de'; W 'source' 'weblink'; W 'country' 'DE'; W 'device_type' 'Computer'; W 'browser' 'Chrome'
    W 'browser_locale' 'de'; W 'browser_version' '156.0.8078.75'; W 'operating_system' 'Windows 10'

    if (Chance 0.95) { $v = Zufall 1 5; W 'Q1' $v $q1[$v - 1]; $items++ }
    $irgendeins = $false
    for ($i = 1; $i -le 5; $i++) {
        $p = @(0.7, 0.55, 0.25, 0.45, 0.2)[$i - 1]
        if (Chance $p) { W ('Q2_{0:00}' -f $i) $i $q2[$i - 1]; $irgendeins = $true }
    }
    if ($kv.ContainsKey('Q2_05')) { W 'Q2t_05' (Pick $sonstiges) }
    if ($irgendeins) { $items++ }
    if (Chance 0.9) { $v = Pick @(1, 1, 1, 2, 2, 3, 3, 4); W 'Q3' $v $q3[$v - 1]; $items++ }
    foreach ($z in $q4) {
        if (Chance 0.94) { $v = Pick @(1, 1, 2, 2, 2, 3, 3, 4, 5, 6); W ('Q4_' + $z[0]) $v $skala[$v - 1]; $items++ }
    }
    for ($r = 1; $r -le 3; $r++) {
        $zeile = $false
        for ($c = 1; $c -le 3; $c++) { if (Chance 0.4) { W ('Q5_{0}_{1:00}' -f $r, $c) $c $alter[$c - 1]; $zeile = $true } }
        if ($zeile) { $items++ }
    }
    foreach ($key in 'Q6_1_01', 'Q6_1_02', 'Q6_2_01', 'Q6_2_02') { if (Chance 0.55) { W $key (Pick $texte6[$key]); $items++ } }
    for ($r = 1; $r -le 2; $r++) { if (Chance 0.9) { $v = Zufall 1 3; W ('Q7_' + $r) $v $oft[$v - 1]; $items++ } }
    for ($r = 1; $r -le 2; $r++) { if (Chance 0.9) { $v = Zufall 1 5; W ('Q8_' + $r) $v ('[' + $v + ']'); $items++ } }
    for ($r = 1; $r -le 2; $r++) { if (Chance 0.85) { W ('Q9_' + $r) ([double](Zufall 1 5)); $items++ } }
    if (Chance 0.85) { W 'Q10' ([double](10 * (Zufall 3 10))); $items++ }
    if (Chance 0.9) { W 'Q11_1' (Pick $namen); $items++ }
    if (Chance 0.7) { W 'Q11_2' (Pick $telefone); $items++ }
    if (Chance 0.9) { $v = Pick @(1, 1, 1, 2, 3); W 'Q11_3' $v $funktion[$v - 1]; $items++ }
    foreach ($key in 'Q12_1_01', 'Q12_2_01', 'Q12_1_02', 'Q12_2_02') {
        if (Chance 0.88) {
            $v = Zufall 1 3
            $lab = if ($key.EndsWith('_01')) { $wicht[$v - 1] } else { $zufr[$v - 1] }
            W $key $v $lab; $items++
        }
    }
    if (Chance 0.8) {
        $perm = @(1, 2, 3, 4)
        for ($a = 3; $a -gt 0; $a--) { $b = $rnd.Next($a + 1); $h = $perm[$a]; $perm[$a] = $perm[$b]; $perm[$b] = $h }
        for ($i = 0; $i -lt 4; $i++) { W ('Q13_{0:00}' -f ($i + 1)) ([double]$perm[$i]) }
        $items++
    }
    if (Chance 0.92) {
        $z1 = 5 * (Zufall 4 10); $z2 = 5 * (Zufall 6 10); $z3 = 100 - $z1 - $z2
        if ($z3 -lt 0) { $z2 = $z2 + $z3; $z3 = 0 }
        W 'Q14_1' ([double]$z1); W 'Q14_2' ([double]$z2); W 'Q14_3' ([double]$z3); $items++
    }
    if (Chance 0.95) { W 'Q15' (Pick $einrichtungen); $items++ }
    if (Chance 0.6) { W 'Q16' (Pick $wuensche); $items++ }
    if (Chance 0.9) { $v = [double](Zufall 18 120); if (Chance 0.1) { $v = $v + 0.5 }; W 'Q17' $v; $items++ }
    if (Chance 0.8) { W 'Q18' ((New-Object DateTime (Zufall 2005 2024), (Zufall 1 12), (Zufall 1 28)).ToString('dd.MM.yyyy')); $items++ }
    if (Chance 0.9) { W 'Q19' (Pick @('06:30', '07:00', '07:00', '07:15', '07:30')); $items++ }
    if (Chance 0.5) { W 'Q20' (Pick $mails); $items++ }
    W 'items_answered_count' $items
    W 'items_answered_percent' ([math]::Round($items / 35 * 100, 5))
    $teilnehmer.Add(@($kv, $tv))
}

# --- XLSX (kodiert mit Codeplan, unkodiert) ---
$schluessel = @($spalten | ForEach-Object { $_.Key })
$nSp = $schluessel.Count
$letzte = Buchstabe $nSp
$nZ = 4 + $teilnehmer.Count
foreach ($variante in 'kodiert', 'unkodiert') {
    $idx = if ($variante -eq 'kodiert') { 0 } else { 1 }
    $blaetter = if ($variante -eq 'kodiert') { @('results', 'code_plan') } else { @('results') }
    $wb = NeueMappe $blaetter
    $ws = $wb.Worksheets.Item(1)
    for ($c = 0; $c -lt $nSp; $c++) {
        $istText = $false
        foreach ($tn in $teilnehmer) { if ($tn[$idx][$schluessel[$c]] -is [string]) { $istText = $true; break } }
        if ($istText) { $b = Buchstabe ($c + 1); TextSpalte $ws "${b}5:${b}$nZ" }
    }
    $zeilen = @(
        , @($spalten | ForEach-Object { $_.Text })
        , @($spalten | ForEach-Object { $_.Sub })
        , @($spalten | ForEach-Object { $_.Json })
        , $schluessel
    )
    foreach ($tn in $teilnehmer) {
        $werte = $tn[$idx]
        $zeile = New-Object object[] $nSp
        for ($c = 0; $c -lt $nSp; $c++) {
            $v = $werte[$schluessel[$c]]
            if ($v -is [DateTime]) { $v = $v.ToOADate() }
            $zeile[$c] = $v
        }
        $zeilen += , $zeile
    }
    [void](Put $ws 1 1 $zeilen)
    Format $ws "C5:E$nZ" 'dd.mm.yyyy hh:mm:ss'
    $ws.UsedRange.WrapText = $false
    [void]$ws.UsedRange.Rows.AutoFit()
    $lo = $ws.ListObjects.Add(1, $ws.Range("A4:$letzte$nZ"), [Type]::Missing, 1)
    $lo.Name = 'results'
    $lo.TableStyle = 'TableStyleMedium2'   # wie im Original (Zeilenstreifen, keine Spaltenstreifen)
    for ($c = 1; $c -le $nSp; $c++) { $ws.Columns.Item($c).ColumnWidth = 19.53 }
    foreach ($p in @(@(1, 48.62), @(2, 23.05), @(3, 25.39), @(4, 25.39), @(5, 52.39), @(6, 58.58), @(10, 40.73), @(24, 63.32))) { $ws.Columns.Item($p[0]).ColumnWidth = $p[1] }
    if ($variante -eq 'kodiert') {
        $cp = $wb.Worksheets.Item(2)
        $z = @(, @('export_key', 'type', 'sub_type', 'data_type', 'coding', 'de'))
        foreach ($e in $codeplan) { $z += , @($e[0], $e[1], $e[2], $e[3], $e[4], $e[5]) }
        [void](Put $cp 1 1 $z)
        Breiten $cp @(23.44, 23.44, 23.44, 23.44, 11.72, 78.13)
    }
    Speichern $wb "2ask_Export_$variante.xlsx"
}

# --- CSV (kodiert, unkodiert) ---
$deDE = [Globalization.CultureInfo]'de-DE'
$ohneQuote = @('progress_percent', 'duration_in_seconds', 'items_answered_count')
function CsvZelle($v, [bool]$roh = $false) {
    if ($null -eq $v -or ($v -is [string] -and $v -eq '')) { return '' }
    if ($roh) { return [string]$v }
    if ($v -is [DateTime]) { $s = $v.ToString('dd.MM.yyyy HH:mm:ss') }
    elseif ($v -is [double]) { $s = $v.ToString('0.0####', $deDE) }
    else { $s = [string]$v }
    $s = $s.Replace("`r`n", "`n").Replace("`n", "`r`n")
    return '"' + $s.Replace('"', '""') + '"'
}
foreach ($variante in 'kodiert', 'unkodiert') {
    $idx = if ($variante -eq 'kodiert') { 0 } else { 1 }
    $out = New-Object System.Collections.Generic.List[string]
    $out.Add((($schluessel | ForEach-Object { CsvZelle $_ }) -join ';'))
    $out.Add((($spalten | ForEach-Object { CsvZelle $_.Text }) -join ';'))
    $out.Add((($spalten | ForEach-Object { CsvZelle $_.Sub }) -join ';'))
    $out.Add((($spalten | ForEach-Object { CsvZelle $_.Json }) -join ';'))
    foreach ($tn in $teilnehmer) {
        $werte = $tn[$idx]
        $out.Add((($schluessel | ForEach-Object { CsvZelle $werte[$_] ($ohneQuote -contains $_) }) -join ';'))
    }
    [IO.File]::WriteAllText((Join-Path $OutDir "2ask_Export_$variante.csv"), (($out -join "`n") + "`n"), (New-Object Text.UTF8Encoding $true))
    Write-Output "erstellt: 2ask_Export_$variante.csv"
}

#==============================================================================
# 11 Fachverfahren: technischer CSV-Export (Windows-1252, Semikolon, ISO-Datum, Dezimalpunkt,
#    Nummern mit fuehrenden Nullen, J/N-Kennzeichen, Feld mit Semikolon in Anfuehrungszeichen)
#==============================================================================
$csv = @(
    'FALLNR;EINGANG;LEISTUNGSART;KOSTENSTELLE;BETRAG;KENNZ_AKTIV;SACHB;BEMERKUNG',
    '0004711;2025-01-14;Erziehungsbeistand;36301;1450.00;J;MUELLERA;',
    '0004712;2025-01-20;Sozialpädagogische Familienhilfe;36302;2380.50;J;SCHULZT;Rückfrage beim Träger läuft',
    '0004713;2025-02-03;Heimerziehung;36303;5120.00;J;KLEINP;',
    '0004714;2025-02-11;Erziehungsbeistand;36301;980.00;N;MUELLERA;beendet zum 30.04.',
    '0004715;2025-02-28;Eingliederungshilfe;36304;3310.75;J;BRAUNT;',
    '0004716;2025-03-07;Tagesgruppe;36305;1890.00;J;SCHULZT;',
    '0004717;2025-03-19;Sozialpädagogische Familienhilfe;36302;2380.50;J;KLEINP;',
    '0004718;2025-04-02;Heimerziehung;36303;5120.00;N;BRAUNT;"Wechsel in Pflegefamilie; Akte an ASD"',
    '0004719;2025-04-15;Erziehungsbeistand;36301;1450.00;J;MUELLERA;',
    '0004720;2025-05-06;Eingliederungshilfe;36304;2950.00;J;BRAUNT;',
    '0004721;2025-05-22;Tagesgruppe;36305;1890.00;J;SCHULZT;',
    '0004722;2025-06-10;Sozialpädagogische Familienhilfe;36302;0.00;J;KLEINP;Kostenzusage fehlt'
)
[IO.File]::WriteAllText((Join-Path $OutDir '11_Fachverfahren_Export.csv'), (($csv -join "`r`n") + "`r`n"), [Text.Encoding]::GetEncoding(1252))
Write-Output 'erstellt: 11_Fachverfahren_Export.csv'
}
finally {
    foreach ($w in @($xl.Workbooks)) { $w.Close($false) }
    $xl.Quit()
    [void][Runtime.InteropServices.Marshal]::ReleaseComObject($xl)
    [GC]::Collect(); [GC]::WaitForPendingFinalizers()
}
