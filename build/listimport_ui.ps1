# Entwicklungshilfe: Sichtprüfung des Assistenten „Liste übernehmen“.
# Öffnet eine Kopie der Demodatei, bedient den Assistenten nicht-modal über mListImport.LiUi und nimmt jeden
# Schritt per PrintWindow auf (unabhängig von darüberliegenden Fenstern). Ergebnis: pruefung\screenshots_180\*.jpg
# Aufnahmen zeigen keinen Entwicklungstag: Die Kopie heißt MidPoint_Demo.xlsm (eigener Ordner je Lauf), Sicherungszeit
# und Untertitel übernommener Tabellen werden vor der Aufnahme auf einen Tag der Beispielwelt gesetzt.
# Bricht ab, wenn bereits Excel läuft (fremde Sitzungen werden nie beendet). Schließt am Ende nur die eigene Instanz.
$root = Split-Path $PSScriptRoot -Parent
if (Get-Process EXCEL -ErrorAction SilentlyContinue) { 'ABBRUCH: Excel ist bereits geöffnet.'; return }
Add-Type -AssemblyName System.Drawing
Add-Type -AssemblyName Microsoft.VisualBasic
if (-not ('UiCap3' -as [type])) {
Add-Type @"
using System; using System.Runtime.InteropServices;
public class UiCap3 {
  [StructLayout(LayoutKind.Sequential)] public struct RECT { public int L, T, R, B; }
  [StructLayout(LayoutKind.Sequential)] public struct POINT { public int X, Y; }
  [DllImport("user32.dll")] public static extern bool GetWindowRect(IntPtr h, out RECT r);
  [DllImport("user32.dll")] public static extern bool GetClientRect(IntPtr h, out RECT r);
  [DllImport("user32.dll")] public static extern bool ClientToScreen(IntPtr h, ref POINT p);
  [DllImport("user32.dll")] public static extern bool PrintWindow(IntPtr h, IntPtr hdc, uint f);
  [DllImport("user32.dll")] public static extern bool SetProcessDPIAware();
  [DllImport("user32.dll", CharSet=CharSet.Unicode)] public static extern IntPtr FindWindow(string c, string t);
  [DllImport("dwmapi.dll")] public static extern int DwmGetWindowAttribute(IntPtr h, int a, out RECT r, int size);
  [DllImport("user32.dll")] public static extern bool PostMessage(IntPtr h, uint m, IntPtr w, IntPtr l);
}
"@
}
[void][UiCap3]::SetProcessDPIAware()
$shots = Join-Path $root 'pruefung\screenshots_180'
New-Item -ItemType Directory -Force $shots | Out-Null
$dir = Join-Path $root 'pruefung\ui170'
New-Item -ItemType Directory -Force $dir | Out-Null
# eigener Ordner je Lauf (Excel merkt sich abgebrochene Dateien unter ihrem Pfad), Datei mit neutralem Namen
$run = Join-Path $dir ('lauf_' + (Get-Date -Format 'HHmmss'))
New-Item -ItemType Directory -Force $run | Out-Null
$wbName = 'MidPoint_Demo.xlsm'
$copy = Join-Path $run $wbName
Copy-Item (Join-Path $root 'dist\MidPoint_Demo.xlsm') $copy -Force
& powershell -ExecutionPolicy Bypass -File (Join-Path $PSScriptRoot 'devopen.ps1') -Path $copy | Out-Null
Start-Sleep 6
$book = [Microsoft.VisualBasic.Interaction]::GetObject($copy)
$app = $book.Application
# Testlisten für die Aufnahmen in einen neutralen Ordner kopieren (der Pfad steht in Schritt 1 im Bild;
# er soll weder Benutzernamen noch Projektordner des Entwicklungsrechners zeigen). Am Ende wieder entfernt.
$td = Join-Path $env:PUBLIC 'Documents\Beispiellisten'
$n = 1
while (Test-Path $td) { $n++; $td = Join-Path $env:PUBLIC ('Documents\Beispiellisten_' + $n) }
New-Item -ItemType Directory -Force $td | Out-Null
Copy-Item (Join-Path $root 'testdaten\importlisten\*') $td -Recurse
$jpg = [System.Drawing.Imaging.ImageCodecInfo]::GetImageEncoders() | Where-Object { $_.MimeType -eq 'image/jpeg' }
$ep = New-Object System.Drawing.Imaging.EncoderParameters 1
$ep.Param[0] = New-Object System.Drawing.Imaging.EncoderParameter ([System.Drawing.Imaging.Encoder]::Quality), 90L
$log = @()
# Tag der Beispielwelt für Aufnahmen nach einer Übernahme
$tag = Get-Date '2026-03-10'
$shift = ((Get-Date).Date - $tag).Days

function UI([string]$cmd) {
    $r = $app.Run("'" + $wbName + "'!mListImport.LiUi", $cmd)
    $script:log += ($cmd + '  ->  ' + $r)
    Start-Sleep -Milliseconds 700
    return $r
}
# Aufnahme eines Fensters: Rahmen und Titelleiste aus der DWM-Aufnahme (PW_RENDERFULLCONTENT), beschnitten auf den
# sichtbaren Rahmen. Beim Formular zeichnet sich der Inhalt zusätzlich selbst in die Aufnahme (PW_CLIENTONLY): Die
# DWM-Aufnahme eines verdeckten MSForms-Formulars ist oft nur teilweise gezeichnet, obwohl das Formular vollständig ist.
function Snap([string]$name, $hw = $null) {
    Start-Sleep -Milliseconds 800
    $isForm = ($null -eq $hw)
    if ($isForm) { $h = [UiCap3]::FindWindow('ThunderDFrame', 'Liste übernehmen') } else { $h = [IntPtr]$hw }
    if ($h -eq [IntPtr]::Zero) { $script:log += ("(" + $name + ": kein Fenster)"); return }
    $r = New-Object UiCap3+RECT
    [void][UiCap3]::GetWindowRect($h, [ref]$r)
    $w = $r.R - $r.L; $hh = $r.B - $r.T
    $bmp = New-Object System.Drawing.Bitmap $w, $hh
    $g = [System.Drawing.Graphics]::FromImage($bmp)
    $hdc = $g.GetHdc()
    [void][UiCap3]::PrintWindow($h, $hdc, 2)
    $g.ReleaseHdc($hdc)
    if ($isForm) {
        $cr = New-Object UiCap3+RECT
        [void][UiCap3]::GetClientRect($h, [ref]$cr)
        $pt = New-Object UiCap3+POINT
        [void][UiCap3]::ClientToScreen($h, [ref]$pt)
        $cb = New-Object System.Drawing.Bitmap ($cr.R - $cr.L), ($cr.B - $cr.T)
        $cg = [System.Drawing.Graphics]::FromImage($cb)
        $chdc = $cg.GetHdc()
        [void][UiCap3]::PrintWindow($h, $chdc, 1)
        $cg.ReleaseHdc($chdc)
        $g.DrawImageUnscaled($cb, ($pt.X - $r.L), ($pt.Y - $r.T))
        $cg.Dispose(); $cb.Dispose()
    }
    # auf den sichtbaren Fensterrahmen beschneiden (ohne die unsichtbaren Ziehränder)
    $vb = New-Object UiCap3+RECT
    $out = $bmp
    if ([UiCap3]::DwmGetWindowAttribute($h, 9, [ref]$vb, 16) -eq 0) {
        $rect = New-Object System.Drawing.Rectangle ($vb.L - $r.L), ($vb.T - $r.T), ($vb.R - $vb.L), ($vb.B - $vb.T)
        if ($rect.Width -gt 0 -and $rect.Height -gt 0 -and $rect.Right -le $w -and $rect.Bottom -le $hh) { $out = $bmp.Clone($rect, $bmp.PixelFormat) }
    }
    $out.Save((Join-Path $shots ($name + '.jpg')), $jpg, $ep)
    if (-not [object]::ReferenceEquals($out, $bmp)) { $out.Dispose() }
    $g.Dispose(); $bmp.Dispose()
    $script:log += ('Aufnahme ' + $name)
}
function Xl([string]$name) { Snap $name ([IntPtr]$app.Hwnd) }
# Übernommene Tabelle zeigen: Sicherungszeit und Untertitel auf den Tag der Beispielwelt, dann Seite öffnen und aufnehmen
function ShowImported([string]$table, [string]$file, [string]$name) {
    $bk = Join-Path $run 'MidPoint_Demo_Sicherungen'
    if (Test-Path $bk) { Get-ChildItem $bk -File | ForEach-Object { $_.LastWriteTime = $_.LastWriteTime.AddDays(-$shift) } }
    $tid = $app.Run("'" + $wbName + "'!mCfg.TableIdByName", $table)
    if (-not $tid) { $script:log += ('(' + $table + ': Tabelle fehlt)'); return }
    [void]$app.Run("'" + $wbName + "'!mStruct.UpdateTable", $tid, $table, ('Übernommen aus „' + $file + '“ am ' + $tag.ToString('dd.MM.yyyy')), '')
    [void]$app.Run("'" + $wbName + "'!mNav.RefreshNav")
    [void]$app.Run("'" + $wbName + "'!mNav.GoTable", $tid)
    Start-Sleep 2
    Xl $name
}

try {
    Xl '01_Start_mit_Liste_uebernehmen'
    # --- Sonstige Liste: Teilnehmerliste mit Kreuzspalten ---
    UI 'open' | Out-Null
    Snap '02_Schritt1_Art_und_Datei'
    UI ('set:path=' + (Join-Path $td '01_Teilnehmerliste_Fortbildung.xlsx')) | Out-Null
    UI 'click:read' | Out-Null
    Snap '03_Schritt1_Datei_gelesen'
    UI 'click:next' | Out-Null
    Snap '04_Schritt2_Tabelle_und_Vorschau'
    UI 'click:next' | Out-Null
    Snap '05_Schritt3_Rueckfragen'
    UI 'idx:qc2=1' | Out-Null
    Snap '06_Schritt3_Kreuze_als_JaNein'
    UI 'idx:qc2=0' | Out-Null
    UI 'click:next' | Out-Null
    Snap '07_Schritt4_Ziel_und_Felder'
    UI 'idx:flds=6' | Out-Null
    Snap '08_Schritt4_Werte_Verpflegung'
    UI 'idx:flds=0' | Out-Null
    UI 'set:incl=0' | Out-Null
    UI 'idx:flds=1' | Out-Null
    UI 'set:fname=Nachname' | Out-Null
    Snap '09_Schritt4_Feld_geaendert'
    UI 'click:next' | Out-Null
    Snap '10_Schritt5_Pruefung'
    UI 'quiet:1' | Out-Null
    UI 'click:next' | Out-Null
    UI 'quiet:0' | Out-Null
    UI 'close' | Out-Null
    ShowImported 'Anmeldungen' '01_Teilnehmerliste_Fortbildung.xlsx' '11_Tabelle_Anmeldungen_uebernommen'
    # --- Wiedervorlage: Werte zusammenlegen („in Bearb.“ als „in Arbeit“) ---
    UI 'open' | Out-Null
    UI ('set:path=' + (Join-Path $td '02_Fristenliste_Wiedervorlage.xlsx')) | Out-Null
    UI 'click:read' | Out-Null
    UI 'click:next' | Out-Null
    UI 'click:next' | Out-Null
    Snap '12_Wiedervorlage_Rueckfragen'
    UI 'click:next' | Out-Null
    UI 'idx:flds=5' | Out-Null
    UI 'idx:vals=4' | Out-Null
    UI 'set:vtarget=in Arbeit' | Out-Null
    UI 'click:vapply' | Out-Null
    Snap '13_Wiedervorlage_Wert_zusammengelegt'
    UI 'close' | Out-Null
    # --- Kreuztabelle Monate ---
    UI 'open' | Out-Null
    UI ('set:path=' + (Join-Path $td '04_Belegung_Kitas_Monate.xlsx')) | Out-Null
    UI 'click:read' | Out-Null
    UI 'click:next' | Out-Null
    UI 'click:next' | Out-Null
    Snap '14_Monate_Rueckfragen'
    UI 'click:next' | Out-Null
    Snap '15_Monate_Felder'
    UI 'close' | Out-Null
    # --- Mehrere Blätter zusammenführen ---
    UI 'open' | Out-Null
    UI ('set:path=' + (Join-Path $td '08_Bedarfsmeldung_je_Kita.xlsx')) | Out-Null
    UI 'click:read' | Out-Null
    UI 'click:next' | Out-Null
    Snap '16_Blaetter_zusammen'
    UI 'click:next' | Out-Null
    Snap '17_Blaetter_Rueckfragen'
    UI 'close' | Out-Null
    # --- Formular ---
    UI 'open' | Out-Null
    UI ('set:path=' + (Join-Path $td '12_Antrag_Formular.xlsx')) | Out-Null
    UI 'click:read' | Out-Null
    UI 'click:next' | Out-Null
    Snap '18_Formular_und_Kostenplan'
    UI 'close' | Out-Null
    # --- 2ask: echter kodierter Export mit Codeplan ---
    $askFile = 'Original_XLSX_kodiert_Codeplan_Umkodierung_99_98_97.xlsx'
    UI 'open' | Out-Null
    UI 'set:optAsk=1' | Out-Null
    UI ('set:path=' + (Join-Path $td ('2ask_echt\' + $askFile))) | Out-Null
    UI 'click:read' | Out-Null
    Snap '19_2ask_gelesen'
    UI 'click:next' | Out-Null
    Snap '20_2ask_Fragen'
    UI 'click:next' | Out-Null
    Snap '21_2ask_Rueckfragen'
    UI 'click:next' | Out-Null
    Snap '22_2ask_Felder'
    UI 'idx:flds=26' | Out-Null
    Snap '23_2ask_Mehrfachauswahl_Werte'
    UI 'click:next' | Out-Null
    Snap '24_2ask_Pruefung'
    UI 'quiet:1' | Out-Null
    UI 'click:next' | Out-Null
    UI 'quiet:0' | Out-Null
    UI 'close' | Out-Null
    ShowImported 'Befragung' $askFile '25_Tabelle_Befragung_uebernommen'
    # --- 2ask-Datei als „Sonstige Liste“: Vorschlag, den Modus zu wechseln (Rückfrage wird im Testmodus bestätigt) ---
    UI 'open' | Out-Null
    UI 'quiet:1' | Out-Null
    UI ('set:path=' + (Join-Path $td '2ask_Export_kodiert.xlsx')) | Out-Null
    UI 'click:read' | Out-Null
    UI 'quiet:0' | Out-Null
    Snap '26_2ask_als_Liste_erkannt'
    UI 'close' | Out-Null
    # --- Wochenraster: verständliche Ablehnung ---
    UI 'open' | Out-Null
    UI ('set:path=' + (Join-Path $td '13_Raumbelegung_Wochenplan.xlsx')) | Out-Null
    UI 'click:read' | Out-Null
    UI 'click:next' | Out-Null
    Snap '27_Wochenraster_erkannt'
    UI 'quiet:1' | Out-Null
    UI 'click:next' | Out-Null
    UI 'quiet:0' | Out-Null
    UI 'close' | Out-Null
    # --- 1.7.1: Feld auf Mehrfachauswahl umstellen und einem Wert mehrere Werte zuordnen ---
    UI 'open' | Out-Null
    UI ('set:path=' + (Join-Path $td '01_Teilnehmerliste_Fortbildung.xlsx')) | Out-Null
    UI 'click:read' | Out-Null
    UI 'click:next' | Out-Null
    UI 'click:next' | Out-Null
    UI 'click:next' | Out-Null
    UI 'idx:flds=7' | Out-Null
    UI 'idx:ftype=9' | Out-Null
    UI 'idx:flds=7' | Out-Null
    UI 'idx:vals=1' | Out-Null
    UI 'sel:vopts=3' | Out-Null
    UI 'set:vtarget=Rückfrage' | Out-Null
    Snap '28_Mehrfachauswahl_Haekchen'
    UI 'click:vapply' | Out-Null
    UI 'idx:vals=1' | Out-Null
    Snap '29_Mehrfachauswahl_zugeordnet'
    UI 'close' | Out-Null
    # --- 1.8.0: gewachsene Listen und zweite 2ask-Befragung ---
    UI 'open' | Out-Null
    UI ('set:path=' + (Join-Path $td '16_Fortbildungsbedarf_Rueckmeldungen.xlsx')) | Out-Null
    UI 'click:read' | Out-Null
    UI 'click:next' | Out-Null
    Snap '30_Fortbildung_drei_Blaetter'
    UI 'click:next' | Out-Null
    Snap '31_Fortbildung_Rueckfragen'
    UI 'close' | Out-Null
    UI 'open' | Out-Null
    UI ('set:path=' + (Join-Path $td '15_Veranstaltungsplanung_2026_Kopie.xlsx')) | Out-Null
    UI 'click:read' | Out-Null
    UI 'click:next' | Out-Null
    UI 'click:next' | Out-Null
    Snap '32_Veranstaltung_Rueckfragen'
    UI 'close' | Out-Null
    UI 'open' | Out-Null
    UI 'set:optAsk=1' | Out-Null
    UI ('set:path=' + (Join-Path $td '2ask_komplex\Buergerbefragung_kodiert_Codeplan.xlsx')) | Out-Null
    UI 'click:read' | Out-Null
    UI 'click:next' | Out-Null
    UI 'click:next' | Out-Null
    UI 'click:next' | Out-Null
    UI 'idx:flds=37' | Out-Null
    Snap '33_2ask_Formular_Datum_erkannt'
    UI 'close' | Out-Null
    # --- 1.7.1: Design „Mainz“ ---
    [void]$app.Run("'" + $wbName + "'!mDocStyle.ApplyDesign", 'Mainz')
    [void]$app.Run("'" + $wbName + "'!mDocStyle.ResetFontCache")
    [void]$app.Run("'" + $wbName + "'!mNav.RefreshNav", $true)
    [void]$app.Run("'" + $wbName + "'!mNav.NavClick", 'START'); Start-Sleep 2; Xl '34_Mainz_Start'
    [void]$app.Run("'" + $wbName + "'!mNav.NavClick", 'T:T0002'); Start-Sleep 2; Xl '35_Mainz_Tabelle'
    [void]$app.Run("'" + $wbName + "'!mNav.NavClick", 'DASH'); Start-Sleep 4
    # Stand des Dashboards auf den Tag der Beispielwelt (kein Entwicklungstag im Bild)
    $dws = $app.ActiveSheet
    foreach ($sp in @($dws.Shapes)) {
        if ($sp.Name -like '*stamp*') {
            $txt = 'Stand: 10.03.2026 08:15 · aktuell'
            try { $sp.TextFrame2.TextRange.Text = $txt } catch { $dws.Unprotect(); $sp.TextFrame2.TextRange.Text = $txt; $dws.Protect([Type]::Missing, $true, $true, $true, $true) }
        }
    }
    Start-Sleep 1; Xl '36_Mainz_Dashboard'
    [void]$app.Run("'" + $wbName + "'!mNav.NavClick", 'SETUP'); Start-Sleep 2; Xl '37_Mainz_Einrichten'
    $app.OnTime((Get-Date).AddSeconds(1), "'" + $wbName + "'!mCmd.Cmd_DocStyle")
    Start-Sleep 2
    $hd = [IntPtr]::Zero
    for ($k = 0; $k -lt 40 -and $hd -eq [IntPtr]::Zero; $k++) { $hd = [UiCap3]::FindWindow('ThunderDFrame', 'Dokumentstil'); if ($hd -eq [IntPtr]::Zero) { Start-Sleep -Milliseconds 250 } }
    if ($hd -ne [IntPtr]::Zero) { Snap '38_Mainz_Dokumentstil' ([IntPtr]$hd); [void][UiCap3]::PostMessage($hd, 0x10, [IntPtr]::Zero, [IntPtr]::Zero); Start-Sleep 2 } else { $log += '(Dokumentstil: kein Fenster)' }
} catch {
    $log += ('FEHLER: ' + $_.Exception.Message)
}
$res = Join-Path $dir 'Sichtpruefung_Ablauf.txt'
$log | Set-Content $res -Encoding UTF8
# eigene Instanz schließen (ohne Speichern, ohne Rückfrage der Anwendung)
try {
    # geplante Aufrufe abmelden (sonst öffnet Excel die Datei für den nächsten Zoom-Takt erneut)
    foreach ($m in @('mNav.StopWatch', 'mBackup.CancelAutoSave')) { try { [void]$app.Run("'" + $wbName + "'!" + $m) } catch { } }
    $app.EnableEvents = $false
    $app.DisplayAlerts = $false
    foreach ($w in @($app.Workbooks)) { $w.Close($false) }
    $app.Quit()
    [void][Runtime.InteropServices.Marshal]::ReleaseComObject($app)
} catch { }
# eigenen Ordner mit den kopierten Testlisten entfernen
for ($i = 0; $i -lt 20; $i++) { if (-not (Get-Process EXCEL -ErrorAction SilentlyContinue)) { break }; Start-Sleep 1 }
if ($td -like (Join-Path $env:PUBLIC 'Documents\Beispiellisten*')) { Remove-Item $td -Recurse -Force -ErrorAction SilentlyContinue }
$log
