# Entwicklungshilfe: geaenderte Quellen in die geoeffnete Entwicklungsmappe neu laden und kompilieren.
param([string]$WorkbookName = "dev.xlsm", [string[]]$Names = @(), [switch]$All)
. "$PSScriptRoot\lib.ps1"
$root = Split-Path $PSScriptRoot -Parent
$src = Join-Path $root 'src'
$tmp = Join-Path $env:TEMP 'dbk_reload'
New-Item -ItemType Directory -Force $tmp | Out-Null
$xl = [Runtime.InteropServices.Marshal]::GetActiveObject('Excel.Application')
$wb = $xl.Workbooks.Item($WorkbookName)
$proj = $wb.VBProject
$files = @()
$files += Get-ChildItem (Join-Path $src 'modules') -Filter *.bas
$files += Get-ChildItem (Join-Path $src 'classes') -Filter *.cls
$files += Get-ChildItem (Join-Path $src 'forms') -Filter *.frm.vba
foreach ($f in $files) {
    $name = $f.Name -replace '\.frm\.vba$','' -replace '\.(bas|cls)$',''
    if (-not $All -and $Names.Count -gt 0 -and ($Names -notcontains $name)) { continue }
    if (-not $All -and $Names.Count -eq 0) { continue }
    if ($f.Name -like '*.frm.vba') {
        $code = [System.IO.File]::ReadAllText($f.FullName, [System.Text.Encoding]::UTF8) -replace "`r?`n", "`r`n"
        $cm = $proj.VBComponents.Item($name).CodeModule
        if ($cm.CountOfLines -gt 0) { $cm.DeleteLines(1, $cm.CountOfLines) }
        $cm.AddFromString($code)
    } else {
        $c = $proj.VBComponents.Item($name)
        $proj.VBComponents.Remove($c)
        [void]$proj.VBComponents.Import((Convert-To1252 $f.FullName $tmp))
    }
    "neu geladen: $name"
}
if ($Names -contains 'ThisWorkbook' -or $All) {
    $code = [System.IO.File]::ReadAllText((Join-Path $src 'ThisWorkbook.cls'), [System.Text.Encoding]::UTF8) -replace "`r?`n", "`r`n"
    $cm = $proj.VBComponents.Item($wb.CodeName).CodeModule
    if ($cm.CountOfLines -gt 0) { $cm.DeleteLines(1, $cm.CountOfLines) }
    $cm.AddFromString($code)
}
& "$PSScriptRoot\compile.ps1" -WorkbookName $WorkbookName
