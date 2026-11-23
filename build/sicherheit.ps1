param([string]$Workbook = '', [string]$OutDir = '')
# Sicherheitsprüfung vor jedem Release (Verwaltungs-IT): Quelltext und fertige Datei.
#  A  Quelltext (src\ oder vba\): Netz und Internet, Programme und Befehle, Windows-Schnittstellen, COM-Objekte,
#     Zugriff auf den eigenen VBA-Code, Makrosicherheit, Dateizugriffe, personenbezogene Daten.
#  B  Datei (Standard: erste Masterdatei in dist\, ohne Makros geöffnet): externe Verknüpfungen, Datenverbindungen,
#     Links nach außen, Dateieigenschaften, sehr ausgeblendete Datenblätter, Hinweisblatt vorn.
# Ergebnis: pruefung\sicherheit\Sicherheitspruefung.md; Abbruch mit Fehler, wenn ein Punkt NICHT BESTANDEN ist.
# Vorlage aus Soziale Infrastruktur 2.3.0 (Skill excel-vba-meister, references/11_sicherheitspruefung.md):
# nur den Block „Projekt“ anpassen: Name, Quellordner, Standarddatei, Test-Module, erlaubte Bibliotheken und
# COM-Objekte jeweils mit Zweck in Nutzersprache. Als UTF-8 mit BOM speichern (PowerShell 5.1).
$ErrorActionPreference = 'Stop'
$root = Split-Path $PSScriptRoot -Parent
#--- Projekt ---------------------------------------------------------------------------------------------
$product = 'MidPoint'
$sources = @('src', 'vba') | ForEach-Object { Join-Path $root $_ } | Where-Object { Test-Path -LiteralPath $_ } | Select-Object -First 1
if (-not $Workbook) { $Workbook = Get-ChildItem -LiteralPath (Join-Path $root 'dist') -Filter 'MidPoint_Master.xlsm' | Select-Object -First 1 -ExpandProperty FullName }
if (-not $OutDir) { $OutDir = Join-Path $root 'pruefung\sicherheit' }
$hintSheet = 'Hinweis'   # Blatt, das ohne Makros vorn steht
$dataSheets = '^(D_|_)'   # Datentabellen (D_T0001 …) und Konfigurationsblätter (_Tabellen, _Felder …)
$testModules = @('mSelfTest', 'mBuild')   # nur für Prüfung, Beispieldaten und Build, nicht im Alltag
$allowedLibs = [ordered]@{
 'user32' = 'Navigation als Kindfenster andocken und bewegen, Fenstergrößen und Bildschirmbereich lesen'
 'gdi32' = 'Bildschirmauflösung (DPI) lesen'
 'gdiplus' = 'Navigation zeichnen (Hintergrund, Kurven, Logo)'
 'oleaut32' = 'gezeichnetes Bild an das Formular übergeben'
}
$allowedCom = [ordered]@{
 'MSXML2.DOMDocument.6.0' = 'eingebettete Bilder (Logo) aus Base64 lesen, ohne Laden von Adressen'
}
# Öffnen mit dem zugeordneten Windows-Programm (FollowHyperlink) nur in diesen Prozeduren, jeweils auf ausdrücklichen
# Klick; die Anwendung selbst baut dabei keine Verbindung auf. Jede andere Stelle bleibt NICHT BESTANDEN.
$allowedOpen = [ordered]@{
 'frmRecord.OpenLink' = 'Felder „URL“ und „E-Mail“: Adresse im Browser bzw. E-Mail-Programm öffnen, erst nach Rückfrage; versendet wird nichts'
 'frmRecord.OpenFile' = 'Feld „Dateien und Medien“: lokale Datei mit dem zugeordneten Programm öffnen, nach Rückfrage'
 'frmRecord.OpenPath' = 'Feld „Lokaler Pfad“: lokale Datei oder lokalen Ordner öffnen, nach Rückfrage'
 'mCmd.Cmd_OpenBackupFolder' = 'Schaltfläche „Sicherungsordner öffnen“: Ordner neben der Datei im Explorer zeigen'
}
$writePlaces = @(
 'Fest: Ordner „<Dateiname>_Sicherungen“ neben der Datei (automatische und manuelle Sicherungen, Prüfung von Sicherungen)',
 'Fest: TEMP-Ordner des Benutzers (Logo- und Diagrammbilder für Formulare und die Diagrammauswahl)',
 'Selbst gewählt: Speichern-unter-, Ordner- und Öffnen-Dialoge (Exporte, Erfassungsdateien, Codeexport, leere Kopie, Import, Liste übernehmen, Rücklauf einlesen)',
 'Nur gelesen: Dateien für Import, „Liste übernehmen“ und Rücklauf werden schreibgeschützt und mit gesperrten Makros geöffnet und sofort wieder geschlossen'
)
$personalData = @(
 'Änderungsprotokoll und die Spalte „_Benutzer“ jedes Eintrags speichern die Windows-Kennung mit Zeitpunkt (mUtil.UserId).',
 'Fachdaten können personenbezogene Angaben enthalten (Personen-, E-Mail- und Telefonfelder; übernommene Listen und 2ask-Befragungen, wenn die Quelle solche Angaben enthält). Kontaktspalten eines 2ask-Exports sind nur vorgewählt, wenn sie gefüllt sind.'
)
#---------------------------------------------------------------------------------------------------------
$results = New-Object System.Collections.Generic.List[object]
function Add-Result($id, $topic, $state, $text, $details) { $results.Add([pscustomobject]@{ ID = $id; Topic = $topic; State = $state; Text = $text; Details = @($details) }) }
$files = Get-ChildItem -LiteralPath $sources -File -Recurse | Where-Object Extension -In '.bas', '.cls', '.vba' | Sort-Object Name
$lines = New-Object System.Collections.Generic.List[object]
foreach ($f in $files) {
 $n = 0
 foreach ($l in [IO.File]::ReadAllLines($f.FullName, [Text.Encoding]::UTF8)) {
  $n++
  $code = [regex]::Replace($l, '"[^"]*"', '""') -replace "'.*$", ''
  $lines.Add([pscustomobject]@{ Module = $f.BaseName; File = $f.Name; Line = $n; Text = $l.Trim(); Code = $code })
 }
}
function Find($pattern, [switch]$InStrings, [switch]$Runtime) {
 $lines | Where-Object { ($(if ($InStrings) { $_.Text -replace "^\s*'.*$", '' } else { $_.Code })) -match $pattern -and (-not $Runtime -or $testModules -notcontains $_.Module) }
}
function Where-Text($hits) { $hits | ForEach-Object { '{0}:{1}  {2}' -f $_.File, $_.Line, ($(if ($_.Text.Length -gt 150) { $_.Text.Substring(0, 150) + ' …' } else { $_.Text })) } }
# A1 Netz und Internet
$net = Find '(?i)XMLHTTP|WinHttp|URLDownloadToFile|InternetOpen|wininet|urlmon|winsock|WebClient|FollowHyperlink|QueryTables\.Add|Connections\.Add|Workbooks\.OpenXML|SendMail|MAPI|Outlook\.Application|CDO\.Message|\.Navigate|InternetExplorer'
$urls = Find '(?i)(https?|ftp)://|mailto:|\\\\\\\\[A-Za-z]' -InStrings
# Prozedur, in der eine Fundstelle steht (für die Liste erlaubter Öffnen-Aufrufe)
function ProcOf($hit) {
 $decl = $lines | Where-Object { $_.File -eq $hit.File -and $_.Line -le $hit.Line -and $_.Code -match '(?i)^\s*(Public\s+|Private\s+|Friend\s+)?(Static\s+)?(Sub|Function|Property\s+(Get|Let|Set))\s+\w+' } | Select-Object -Last 1
 $mod = $hit.Module -replace '(?i)\.frm$', ''
 if ($decl -and $decl.Code -match '(?i)(Sub|Function|Property\s+(Get|Let|Set))\s+(\w+)') { return $mod + '.' + $Matches[3] }
 return $mod + '.?'
}
$openHits = @(); $netOther = @()
foreach ($h in @($net)) { if ($h.Code -match '(?i)FollowHyperlink' -and $allowedOpen.Contains((ProcOf $h))) { $openHits += $h } else { $netOther += $h } }
$openText = @($openHits | ForEach-Object { $p = ProcOf $_; '{0}:{1}  {2}: {3}' -f $_.File, $_.Line, $p, $allowedOpen[$p] })
if (@($netOther).Count -eq 0) {
 $t = 'Kein eigener Netz- oder Internetzugriff im Code (keine HTTP-, Download-, Mail- oder Datenbankverbindung). '
 if (@($openHits).Count) { $t += 'Nur auf ausdrücklichen Klick übergibt die Anwendung eine eingetragene Adresse, Datei oder einen Ordner an das zugeordnete Windows-Programm (Browser, E-Mail-Programm, Explorer), bei Adressen und Dateien erst nach Rückfrage (' + @($openHits).Count + ' Stellen). ' }
 $t += $(if (@($urls).Count) { @($urls).Count.ToString() + ' Adresse(n) nur als Text, z. B. Lizenzhinweis; sie werden nie aufgerufen.' } else { 'Keine Adressen im Text.' })
 Add-Result 'A1' 'Netz und Internet' 'BESTANDEN' $t (@($openText) + @(Where-Text $urls))
}
else { Add-Result 'A1' 'Netz und Internet' 'NICHT BESTANDEN' 'Netzzugriff im Code gefunden.' (Where-Text $netOther) }
# A2 Programme und Befehle
$cmd = Find '(?i)\bShell\s*\(|WScript\.Shell|ShellExecute|SendKeys|ExecuteExcel4Macro|\bEvaluate\s*\(|\bMacScript\b|CreateProcess|WinExec'
if (@($cmd).Count -eq 0) { Add-Result 'A2' 'Programme und Befehle' 'BESTANDEN' 'Die Anwendung startet keine Programme, Befehlszeilen oder Skripte und sendet keine Tastendrücke.' @() }
else { Add-Result 'A2' 'Programme und Befehle' 'NICHT BESTANDEN' 'Start fremder Programme oder Tastendrücke gefunden.' (Where-Text $cmd) }
# A3 Windows-Schnittstellen
$decl = Find '(?i)\bDeclare\b'
$libs = @{}; $unknown = @()
foreach ($d in $decl) { if ($d.Text -match '(?i)Lib\s+"([^"]+)"') { $lib = ($Matches[1] -replace '(?i)\.dll$', '').ToLower(); $libs[$lib] = 1 + [int]$libs[$lib]; if (-not $allowedLibs.Contains($lib)) { $unknown += $d } } }
$libText = ($libs.Keys | Sort-Object | ForEach-Object { '{0} ({1}×): {2}' -f $_, $libs[$_], $(if ($allowedLibs.Contains($_)) { $allowedLibs[$_] } else { 'NICHT ERLAUBT' }) })
if ($unknown.Count -eq 0) { Add-Result 'A3' 'Windows-Schnittstellen' 'BESTANDEN' ('Nur Windows-Systembibliotheken: ' + (($libs.Keys | Sort-Object) -join ', ') + '. Keine fremden Bibliotheken oder Add-ins.') $libText }
else { Add-Result 'A3' 'Windows-Schnittstellen' 'NICHT BESTANDEN' 'Nicht freigegebene Bibliothek.' ((Where-Text $unknown) + $libText) }
# A4 COM-Objekte
$com = Find '(?i)\b(CreateObject|GetObject)\s*\(' -InStrings
$progs = @{}; $bad = @()
foreach ($c in $com) {
 foreach ($m in [regex]::Matches($c.Text, '(?i)(CreateObject|GetObject)\s*\(\s*"([^"]*)"')) { $p = $m.Groups[2].Value; $progs[$p] = @($progs[$p]) + ($c.Module); if (-not $allowedCom.Contains($p)) { $bad += $c } }
 if ($c.Text -notmatch '(?i)(CreateObject|GetObject)\s*\(\s*"') { $bad += $c }
}
$comText = ($progs.Keys | Sort-Object | ForEach-Object { '{0} ({1}): {2}' -f $_, ((@($progs[$_]) | Where-Object { $_ } | Sort-Object -Unique) -join ', '), $(if ($allowedCom.Contains($_)) { $allowedCom[$_] } else { 'NICHT ERLAUBT' }) })
$domLoad = Find '(?i)\.(Load|loadXML)\s*\(' -InStrings
if ($bad.Count -eq 0 -and @($domLoad).Count -eq 0) { Add-Result 'A4' 'COM-Objekte' 'BESTANDEN' ('Nur bekannte Windows- und Office-Bausteine (' + $progs.Count + '), jeder mit festem Zweck; kein Laden von Adressen.') $comText }
else { Add-Result 'A4' 'COM-Objekte' 'NICHT BESTANDEN' 'Unbekanntes oder dynamisch erzeugtes COM-Objekt.' ((Where-Text ($bad + @($domLoad))) + $comText) }
# A5 Zugriff auf den eigenen VBA-Code
$vbp = Find '(?i)\bVBProject\b|\bVBComponents\b|\bCodeModule\b|AccessVBOM' -Runtime
if (@($vbp).Count -eq 0) { Add-Result 'A5' 'VBA-Projekt' 'BESTANDEN' 'Die Anwendung verändert ihren Code nicht und braucht keinen „Zugriff auf das VBA-Projektobjektmodell“. Der Codekatalog liest eine beim Build gespeicherte Textkopie.' @() }
else { Add-Result 'A5' 'VBA-Projekt' 'NICHT BESTANDEN' 'Zugriff auf das VBA-Projekt zur Laufzeit.' (Where-Text $vbp) }
# A6 Makro- und Sicherheitseinstellungen
$sec = Find '(?i)AutomationSecurity|MacroSecurity|Registry|RegWrite|SaveSetting|HKEY_|TrustCenter|AccessVBOM' -InStrings
$lowered = @($sec | Where-Object { $_.Text -match '(?i)AutomationSecurity\s*=\s*(msoAutomationSecurityLow|1)\b' -or $_.Text -match '(?i)RegWrite|HKEY_|SaveSetting' })
if ($lowered.Count -eq 0) { Add-Result 'A6' 'Sicherheitseinstellungen' 'BESTANDEN' ('Keine Office-, Windows- oder Registry-Einstellung wird geändert. ' + $(if (@($sec).Count) { 'Beim Lesen von Import- und Quelldateien und beim Prüfen von Sicherungen werden Makros der geöffneten Datei ausdrücklich gesperrt (AutomationSecurity = ForceDisable) und danach der vorherige Wert wiederhergestellt.' } else { '' })) (Where-Text $sec) }
else { Add-Result 'A6' 'Sicherheitseinstellungen' 'NICHT BESTANDEN' 'Makrosicherheit gesenkt oder Registry geschrieben.' (Where-Text $lowered) }
# A7 Dateizugriffe
$writes = Find '(?i)\bKill\b|\bMkDir\b|\bRmDir\b|\bFileCopy\b|^\s*Name\s+.+\s+As\s+|\bOpen\b.+\bFor\s+(Output|Binary|Append|Random)\b|\.SaveAs\b|\.SaveCopyAs\b|Workbooks\.Open\b|GetSaveAsFilename|GetOpenFilename' -Runtime
$risky = @($writes | Where-Object { $_.Code -match '(?i)\bRmDir\b|\bKill\b[^:]*\*|^\s*Name\s+' })
$where = @($writePlaces)
$kills = @($writes | Where-Object { $_.Code -match '(?i)\bKill\b' })
$where += 'Löschen (' + $kills.Count + ' Stellen): ' + (($kills | ForEach-Object { $_.File + ':' + $_.Line }) -join ', ')
if ($risky.Count -eq 0) { Add-Result 'A7' 'Dateien und Laufwerke' 'BESTANDEN' ('Geschrieben wird nur an festen, eigenen Orten oder an Zielen, die die Person selbst wählt. Gelöscht werden nur eigene Sicherungen (Namensmuster) und Zwischendateien; keine Platzhalter, keine Ordner. ' + @($writes).Count + ' Fundstellen.') ($where + (Where-Text $writes)) }
else { Add-Result 'A7' 'Dateien und Laufwerke' 'NICHT BESTANDEN' 'Löschen mit Platzhalter, Ordner löschen oder Umbenennen gefunden.' (Where-Text $risky) }
# A8 Personenbezogene Daten
$pers = Find '(?i)Environ\$?\s*\(\s*"(USERNAME|USERDOMAIN|USERPROFILE|COMPUTERNAME)"|Application\.UserName' -InStrings
$fields = Find '(?i)"(Contact|Kontakt|Email|E-Mail|Telefon|Phone)"' -InStrings
$pd = @($personalData)
if (@($pers).Count) { $pd += 'Fundstellen der Windows-Kennung: ' + ((@($pers) | ForEach-Object { $_.File + ':' + $_.Line }) -join ', ') }
if (@($fields).Count) { $pd += 'Fundstellen von Kontaktfeldern: ' + ((@($fields) | ForEach-Object { $_.File + ':' + $_.Line } | Select-Object -First 6) -join ', ') }
$pd += 'Alle Daten bleiben in der Datei; nichts wird übertragen. Sicherungen enthalten dieselben Daten und liegen im Ordner „<Dateiname>_Sicherungen“.'
Add-Result 'A8' 'Personenbezogene Daten' 'HINWEIS' 'Die Datei speichert Windows-Kennungen der Bearbeitenden und, je nach Einrichtung, personenbezogene Fachdaten. Für das Verzeichnis der Verarbeitungstätigkeiten melden; Zugriff über die Ordnerrechte regeln.' $pd
# B Datei
if (Test-Path -LiteralPath $Workbook) {
 $xl = New-Object -ComObject Excel.Application
 $xl.Visible = $false; $xl.DisplayAlerts = $false; $xl.EnableEvents = $false; $xl.AutomationSecurity = 3
 try {
  $wb = $xl.Workbooks.Open($Workbook, 0, $true)
  $links = @($wb.LinkSources(1)) + @($wb.LinkSources(2)) | Where-Object { $_ }
  if ($links.Count -eq 0) { Add-Result 'B1' 'Externe Verknüpfungen' 'BESTANDEN' 'Keine Verknüpfungen zu anderen Arbeitsmappen oder OLE-Quellen.' @() } else { Add-Result 'B1' 'Externe Verknüpfungen' 'NICHT BESTANDEN' 'Externe Verknüpfung gefunden.' $links }
  $qt = 0; $ext = @(); $hl = 0; $veryHidden = @(); $visibleDb = @()
  foreach ($w in $wb.Worksheets) {
   $qt += $w.QueryTables.Count
   foreach ($lo in $w.ListObjects) { if ($lo.SourceType -ne 1) { $ext += ($w.Name + '!' + $lo.Name) } }
   foreach ($h in $w.Hyperlinks) { $hl++; if ($h.Address) { $ext += ($w.Name + ': Link nach ' + $h.Address) } }
   if ($w.Name -match $dataSheets) { if ($w.Visible -eq 2) { $veryHidden += $w.Name } else { $visibleDb += $w.Name } }
  }
  if ($wb.Connections.Count -eq 0 -and $qt -eq 0 -and $ext.Count -eq 0) { Add-Result 'B2' 'Datenverbindungen und Links' 'BESTANDEN' ("Keine Datenverbindungen, Abfragen oder Links nach außen; $hl Links führen nur innerhalb der Datei.") @() }
  else { Add-Result 'B2' 'Datenverbindungen und Links' 'NICHT BESTANDEN' 'Verbindung oder Link nach außen.' $ext }
  if ($visibleDb.Count -eq 0) { Add-Result 'B3' 'Datenblätter' 'BESTANDEN' ("$($veryHidden.Count) Datentabellen sind sehr ausgeblendet und geschützt; bearbeitet wird nur über die Formulare.") @() }
  else { Add-Result 'B3' 'Datenblätter' 'NICHT BESTANDEN' 'Datenblatt sichtbar.' $visibleDb }
  $props = @()
  # Dateieigenschaften direkt aus der Datei lesen (docProps/core.xml und app.xml): Excel meldete über COM leere Werte, obwohl ein Name
  # in der Datei stand (November 2026). Das Lieferpaket leert Autor und „zuletzt gespeichert von“ (build\entspuren.ps1).
  try {
   Add-Type -AssemblyName System.IO.Compression.FileSystem
   $zip = [System.IO.Compression.ZipFile]::OpenRead($Workbook)
   foreach ($part in 'docProps/core.xml', 'docProps/app.xml') {
    $en = $zip.GetEntry($part); if (-not $en) { continue }
    $sr = New-Object System.IO.StreamReader($en.Open()); $xml = $sr.ReadToEnd(); $sr.Close()
    foreach ($tag in 'dc:creator', 'cp:lastModifiedBy', 'Company', 'Manager') { foreach ($m in [regex]::Matches($xml, "<$tag>([^<]+)</$tag>")) { $props += "$tag = $($m.Groups[1].Value)" } }
   }
   $zip.Dispose()
  } catch { $props += "Dateieigenschaften nicht lesbar: $($_.Exception.Message)" }
  Add-Result 'B4' 'Dateieigenschaften' $(if ($props.Count) { 'HINWEIS' } else { 'BESTANDEN' }) $(if ($props.Count) { 'In den Dateieigenschaften steht ein Name (Excel trägt beim Speichern das Office-Konto ein). Das Lieferpaket entfernt ihn (build\entspuren.ps1).' } else { 'Keine Namen in den Dateieigenschaften.' }) $props
  $active = $wb.ActiveSheet.Name
  Add-Result 'B5' 'Ohne Makros' $(if ($active -eq $hintSheet) { 'BESTANDEN' } else { 'NICHT BESTANDEN' }) ('Ohne Makros erscheint das Blatt „' + $active + '“ mit der Anleitung zum Freischalten; Daten bleiben verborgen und geschützt.') @()
  $wb.Close($false)
 } finally { $xl.Quit(); [void][Runtime.InteropServices.Marshal]::ReleaseComObject($xl) }
} else { Add-Result 'B0' 'Datei' 'NICHT BESTANDEN' ('Datei nicht gefunden: ' + $Workbook) @() }
# Bericht
$version = ''
foreach ($c in (Get-ChildItem -LiteralPath $sources -File -Recurse -Filter '*.bas')) {
 $m = [regex]::Match([IO.File]::ReadAllText($c.FullName), '(APP|PROG)_VERSION As String = "([^"]+)"'); if ($m.Success) { $version = $m.Groups[2].Value; break }
}
$failed = @($results | Where-Object State -eq 'NICHT BESTANDEN').Count
$sb = New-Object Text.StringBuilder
[void]$sb.AppendLine("# Sicherheitsprüfung: $product $version")
[void]$sb.AppendLine('')
[void]$sb.AppendLine(('Stand: {0} · geprüft: {1} Quelldateien ({2} Zeilen) und {3}' -f (Get-Date).ToString('MMMM yyyy', [Globalization.CultureInfo]'de-DE'), $files.Count, $lines.Count, (Split-Path $Workbook -Leaf)))
[void]$sb.AppendLine('')
[void]$sb.AppendLine($(if ($failed -eq 0) { '**Ergebnis: bestanden.** Die Anwendung arbeitet vollständig lokal, ohne Netz- oder Internetzugriff, ohne fremde Programme und ohne Änderung von Sicherheitseinstellungen.' } else { "**Ergebnis: $failed Punkt(e) nicht bestanden, kein Release.**" }))
[void]$sb.AppendLine('')
[void]$sb.AppendLine('| Nr. | Prüfpunkt | Ergebnis | Befund |')
[void]$sb.AppendLine('|---|---|---|---|')
foreach ($r in $results) { [void]$sb.AppendLine(('| {0} | {1} | {2} | {3} |' -f $r.ID, $r.Topic, $r.State, ($r.Text -replace '\|', '/'))) }
foreach ($r in $results) {
 if (@($r.Details).Count -eq 0) { continue }
 [void]$sb.AppendLine(''); [void]$sb.AppendLine("## $($r.ID) $($r.Topic)"); [void]$sb.AppendLine('')
 foreach ($d in $r.Details) { if ($d) { [void]$sb.AppendLine('- ' + ($d -replace '\|', '/')) } }
}
New-Item -ItemType Directory -Force $OutDir | Out-Null
$report = Join-Path $OutDir 'Sicherheitspruefung.md'
[IO.File]::WriteAllText($report, $sb.ToString(), (New-Object Text.UTF8Encoding($true)))
$results | ForEach-Object { '{0}  {1,-30} {2}' -f $_.ID, $_.Topic, $_.State }
'Bericht: ' + $report
if ($failed -gt 0) { throw "Sicherheitsprüfung: $failed Punkt(e) nicht bestanden." }
