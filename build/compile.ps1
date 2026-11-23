# Entwicklungshilfe: VBA-Projekt der geoeffneten Arbeitsmappe kompilieren (Debuggen > Kompilieren).
# Laeuft als eigener Prozess mit Zeitlimit, damit ein Fehlerdialog den Aufrufer nicht blockiert.
param([string]$WorkbookName, [int]$TimeoutSec = 25)
$script = @"
`$xl = [Runtime.InteropServices.Marshal]::GetActiveObject('Excel.Application')
`$wb = `$xl.Workbooks.Item('$WorkbookName')
`$xl.VBE.ActiveVBProject = `$wb.VBProject
`$ctl = `$xl.VBE.CommandBars.FindControl(1, 578)
if (`$ctl.Enabled) { `$ctl.Execute() ; 'COMPILED' } else { 'ALREADY_COMPILED' }
"@
$f = Join-Path $env:TEMP 'dbk_compile_helper.ps1'
Set-Content -Path $f -Value $script -Encoding UTF8
$p = Start-Process powershell -ArgumentList '-NoProfile','-ExecutionPolicy','Bypass','-File',$f -PassThru -WindowStyle Hidden -RedirectStandardOutput (Join-Path $env:TEMP 'dbk_compile_out.txt')
$done = $p.WaitForExit($TimeoutSec * 1000)
Start-Sleep 2
# Kompilierfehler erscheinen als modaler Dialog des VBA-Editors (Klasse #32770)
Add-Type -Namespace DbkWin -Name U -MemberDefinition '[DllImport("user32.dll", CharSet=CharSet.Unicode)] public static extern System.IntPtr FindWindow(string c, string t);'
$dlg = [DbkWin.U]::FindWindow('#32770', 'Microsoft Visual Basic for Applications')
if ($dlg -ne [IntPtr]::Zero) { 'KOMPILIERFEHLER (Dialog im VBA-Editor sichtbar)'; if (-not $done) { Stop-Process -Id $p.Id -Force }; return }
if (-not $done) {
    Stop-Process -Id $p.Id -Force
    'TIMEOUT (vermutlich Kompilierfehler-Dialog sichtbar)'
} else {
    Get-Content (Join-Path $env:TEMP 'dbk_compile_out.txt')
}
