# Entwicklungshilfe: Hilfe und Technische Dokumentation aus der Masterdatei als Markdown exportieren
# (doku\Hilfe.md, doku\Technische_Dokumentation.md). Quelle sind mHelpText.HelpLines/TechLines.
# Liegt der Bericht der Sicherheitspruefung vor (pruefung\sicherheit\Sicherheitspruefung.md), steht seine
# Ergebnistabelle ganz oben in der Technischen Dokumentation.
. "$PSScriptRoot\lib.ps1"
$root = Split-Path $PSScriptRoot -Parent
$master = Join-Path $root 'dist\MidPoint_Master.xlsm'
$report = Join-Path $root 'pruefung\sicherheit\Sicherheitspruefung.md'
$xl = New-ExcelApp $false
$xl.AutomationSecurity = 1
try {
    $wb = $xl.Workbooks.Open($master, 0, $true)
    foreach ($job in @(@('mHelpText.HelpLines', 'Hilfe.md', '# Hilfe zu MidPoint (aus der Anwendung exportiert)'),
                       @('mHelpText.TechLines', 'Technische_Dokumentation.md', '# Technische Dokumentation (aus der Anwendung exportiert)'))) {
        $lines = $xl.Run("'" + $wb.Name + "'!" + $job[0])
        $sb = New-Object System.Text.StringBuilder
        [void]$sb.AppendLine($job[2]).AppendLine()
        $firstNote = $true
        foreach ($l in $lines) {
            $p = $l.IndexOf('|'); $kind = $l.Substring(0, $p); $txt = $l.Substring($p + 1)
            switch ($kind) {
                'H1' { [void]$sb.AppendLine().AppendLine('## ' + $txt).AppendLine() }
                'H2' { [void]$sb.AppendLine().AppendLine('### ' + $txt).AppendLine() }
                'B'  { [void]$sb.AppendLine('- ' + $txt) }
                'N'  {
                    [void]$sb.AppendLine().AppendLine('> ' + $txt).AppendLine()
                    # Technische Dokumentation: Ergebnistabelle der Sicherheitspruefung direkt unter dem Programmstand
                    if ($firstNote -and $job[1] -eq 'Technische_Dokumentation.md' -and (Test-Path -LiteralPath $report)) {
                        $rep = [IO.File]::ReadAllLines($report, [Text.Encoding]::UTF8)
                        [void]$sb.AppendLine('### Sicherheitsprüfung vor jedem Release').AppendLine()
                        foreach ($r in $rep) { if ($r -like 'Stand:*' -or $r -like '`*`*Ergebnis*') { [void]$sb.AppendLine($r).AppendLine() } }
                        foreach ($r in $rep) { if ($r -like '|*') { [void]$sb.AppendLine($r) } }
                        [void]$sb.AppendLine()
                    }
                    $firstNote = $false
                }
                'C'  { [void]$sb.AppendLine('`' + $txt + '`') }
                default { [void]$sb.AppendLine($txt).AppendLine() }
            }
        }
        $out = Join-Path $root ('doku\' + $job[1])
        [System.IO.File]::WriteAllText($out, ($sb.ToString() -replace "(`r?`n){3,}", "`r`n`r`n"), (New-Object System.Text.UTF8Encoding $false))
        "geschrieben: $out"
    }
    $wb.Close($false)
} finally {
    $xl.Quit()
}
