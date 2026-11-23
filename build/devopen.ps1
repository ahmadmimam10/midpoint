# Entwicklungshilfe: Datei in einer sichtbaren Excel-Instanz mit aktivierten Makros oeffnen (Workbook_Open laeuft).
param([string]$Path, [switch]$Fresh)
if ($Fresh) { Get-Process EXCEL -ErrorAction SilentlyContinue | Stop-Process -Force; Start-Sleep 2 }
$xl = New-Object -ComObject Excel.Application
$xl.Visible = $true
$xl.AutomationSecurity = 1
$xl.EnableEvents = $true
$xl.DisplayAlerts = $true
$xl.WindowState = -4137
$wb = $xl.Workbooks.Open($Path)
"geoeffnet: " + $wb.Name
