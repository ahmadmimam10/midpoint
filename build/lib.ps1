# Gemeinsame Build-Hilfen: Excel per COM starten, VBA-Quellen importieren.
# Quellen liegen als UTF-8 vor; der VBA-Import erwartet Windows-1252.

$ErrorActionPreference = 'Stop'
$enc1252 = [System.Text.Encoding]::GetEncoding(1252)

function New-ExcelApp([bool]$visible = $true) {
    $xl = New-Object -ComObject Excel.Application
    $xl.Visible = $visible
    $xl.DisplayAlerts = $false
    $xl.EnableEvents = $false
    return $xl
}

function Convert-To1252([string]$src, [string]$tmpDir) {
    $text = [System.IO.File]::ReadAllText($src, [System.Text.Encoding]::UTF8)
    $text = $text -replace "`r?`n", "`r`n"
    $dst = Join-Path $tmpDir ([System.IO.Path]::GetFileName($src))
    [System.IO.File]::WriteAllText($dst, $text, $enc1252)
    return $dst
}

# Bilder (src\res\*.png) als Base64-Text auf das ausgeblendete Blatt "_Ressourcen" schreiben
# (je Datei eine Zeile: Name | Teil 1 | Teil 2 ...; Zellen fassen hoechstens 32.767 Zeichen)
function Embed-Resources($wb, [string]$srcRoot) {
    $resDir = Join-Path $srcRoot 'res'
    if (-not (Test-Path $resDir)) { return }
    $ws = $null
    foreach ($s in @($wb.Worksheets)) { if ($s.Name -eq '_Ressourcen') { $ws = $s } }
    if ($ws -eq $null) {
        $ws = $wb.Worksheets.Add([Type]::Missing, $wb.Worksheets.Item($wb.Worksheets.Count))
        $ws.Name = '_Ressourcen'
    }
    [void]$ws.Cells.Clear()
    $ws.Cells.NumberFormat = '@'
    $r = 0
    foreach ($f in Get-ChildItem $resDir -Filter *.png | Sort-Object Name) {
        $r++
        $b64 = [Convert]::ToBase64String([IO.File]::ReadAllBytes($f.FullName))
        $ws.Cells.Item($r, 1).Value2 = $f.Name
        $c = 2
        for ($i = 0; $i -lt $b64.Length; $i += 30000) {
            $ws.Cells.Item($r, $c).Value2 = $b64.Substring($i, [Math]::Min(30000, $b64.Length - $i))
            $c++
        }
    }
    $ws.Visible = 2
}

function Import-VbaSources($wb, [string]$srcRoot, [string]$tmpDir) {
    New-Item -ItemType Directory -Force $tmpDir | Out-Null
    $proj = $wb.VBProject
    foreach ($f in Get-ChildItem (Join-Path $srcRoot 'modules') -Filter *.bas) {
        [void]$proj.VBComponents.Import((Convert-To1252 $f.FullName $tmpDir))
    }
    foreach ($f in Get-ChildItem (Join-Path $srcRoot 'classes') -Filter *.cls) {
        [void]$proj.VBComponents.Import((Convert-To1252 $f.FullName $tmpDir))
    }
    # Formulare: leeres MSForm anlegen, Code einfuegen (Oberflaeche entsteht zur Laufzeit)
    foreach ($f in Get-ChildItem (Join-Path $srcRoot 'forms') -Filter *.frm.vba) {
        $name = $f.Name -replace '\.frm\.vba$', ''
        $comp = $proj.VBComponents.Add(3)
        $comp.Name = $name
        $code = [System.IO.File]::ReadAllText($f.FullName, [System.Text.Encoding]::UTF8)
        $code = $code -replace "`r?`n", "`r`n"
        $cm = $comp.CodeModule
        if ($cm.CountOfLines -gt 0) { $cm.DeleteLines(1, $cm.CountOfLines) }
        $cm.AddFromString($code)
        $comp.Properties.Item('Caption').Value = $name
        $comp.Properties.Item('Width').Value = 300
        $comp.Properties.Item('Height').Value = 200
    }
    $twb = Join-Path $srcRoot 'ThisWorkbook.cls'
    if (Test-Path $twb) {
        $code = [System.IO.File]::ReadAllText($twb, [System.Text.Encoding]::UTF8) -replace "`r?`n", "`r`n"
        $cm = $proj.VBComponents.Item($wb.CodeName).CodeModule
        if ($cm.CountOfLines -gt 0) { $cm.DeleteLines(1, $cm.CountOfLines) }
        $cm.AddFromString($code)
    }
}
