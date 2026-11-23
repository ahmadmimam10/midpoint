# Einfache statische Pruefung der VBA-Quellen (Entwicklungshilfe):
# doppelte Deklarationen je Prozedur (Gross-/Kleinschreibung egal), Parameter-Kollisionen,
# reservierte Woerter als Bezeichner, Zeilenfortsetzungen > 24.
$src = Join-Path (Split-Path $PSScriptRoot -Parent) 'src'
$reserved = @('alias','base','type','property','event','stop','end','rem','option','lib','declare','friend','global','implements','raiseevent','with','wend','until','resume','return','gosub','goto','spc','tab','print','open','close','get','put','input','output','binary','random','append','lock','unlock','seek','me','new','not','and','or','mod','is','like','eqv','imp','xor','true','false','nothing','null','empty','date','string','integer','long','single','double','currency','variant','object','byte','boolean','collection','error','to','step','then','else','next','loop','do','for','each','in','select','case','call','let','set','dim','redim','preserve','static','const','private','public','sub','function','exit','on','attribute','erase','debug','load','unload','line')
function Split-Decl([string]$s) {
    $parts = @(); $depth = 0; $cur = ''
    foreach ($ch in $s.ToCharArray()) {
        if ($ch -eq '(') { $depth++ } elseif ($ch -eq ')') { $depth-- }
        if ($ch -eq ',' -and $depth -eq 0) { $parts += $cur; $cur = '' } else { $cur += $ch }
    }
    if ($cur.Trim()) { $parts += $cur }
    return $parts
}
function Name-Of([string]$d) {
    $d = $d.Trim() -replace '^(Optional\s+)?(ByVal\s+|ByRef\s+)?(ParamArray\s+)?', '' -replace '^WithEvents\s+',''
    if ($d -match '^([A-Za-z_][A-Za-z0-9_]*)') { return $Matches[1] }
    return $null
}
foreach ($f in Get-ChildItem $src -Recurse -File) {
    $raw = Get-Content $f.FullName -Encoding UTF8
    # Fortsetzungen zusammenfuehren
    $lines = @(); $acc = ''; $cont = 0
    foreach ($l in $raw) {
        $code = $l
        if ($code -match ' _\s*$') { $acc += ($code -replace ' _\s*$', ' '); $cont++; if ($cont -gt 24) { "$($f.Name): mehr als 24 Fortsetzungen" }; continue }
        $lines += ($acc + $code); $acc = ''; $cont = 0
    }
    $inProc = $false; $names = @{}; $proc = ''; $modNames = @{}
    $n = 0
    foreach ($l in $lines) {
        $n++
        $s = ($l -replace "'.*$", '')  # Kommentare grob entfernen
        $s = [regex]::Replace($s, '"[^"]*"', '""')
        if ($s -match '^\s*(Public |Private |Friend )?(Static )?(Sub|Function|Property (Get|Let|Set))\s+([A-Za-z_][A-Za-z0-9_]*)\s*\((.*)\)') {
            $inProc = $true; $names = @{}; $proc = $Matches[5]
            $plist = $Matches[6]
            if ($plist -match '^(.*)\)\s*As\s') { }
            foreach ($p in (Split-Decl $plist)) { $nm = Name-Of $p; if ($nm) { $k = $nm.ToLower(); if ($names.ContainsKey($k)) { "$($f.Name) $proc : doppelter Parameter $nm" }; $names[$k] = 1; if ($reserved -contains $k) { "$($f.Name) $proc : reserviertes Wort als Parameter: $nm" } } }
            continue
        }
        if ($s -match '^\s*End (Sub|Function|Property)') { $inProc = $false; continue }
        foreach ($stmt in ($s -split ':(?=\s)')) {
            if ($stmt -match '^\s*(Dim|Static|Private|Public)\s+(?!Sub|Function|Property|Const|Declare|Enum|Type)(.+)$') {
                foreach ($d in (Split-Decl $Matches[2])) {
                    $nm = Name-Of $d
                    if (-not $nm) { continue }
                    $k = $nm.ToLower()
                    if ($reserved -contains $k) { "$($f.Name) $proc : reserviertes Wort: $nm (Zeile $n)" }
                    if ($inProc) {
                        if ($names.ContainsKey($k)) { "$($f.Name) $proc : Mehrfachdeklaration $nm (Zeile $n)" }
                        $names[$k] = 1
                    } else {
                        if ($modNames.ContainsKey($k)) { "$($f.Name) Modulebene: Mehrfachdeklaration $nm" }
                        $modNames[$k] = 1
                    }
                }
            }
        }
    }
}

# Zeichen ausserhalb von Windows-1252 gehen beim Import in den VBA-Editor verloren (werden zu "?")
$enc1252 = [Text.Encoding]::GetEncoding(1252)
foreach ($f in Get-ChildItem (Join-Path (Split-Path $PSScriptRoot -Parent) 'src') -Recurse -Include *.bas,*.cls,*.vba) {
    $n = 0
    foreach ($line in [IO.File]::ReadAllLines($f.FullName, [Text.Encoding]::UTF8)) {
        $n++
        if ($enc1252.GetString($enc1252.GetBytes($line)) -ne $line) { "$($f.Name) Zeile $n : Zeichen ausserhalb Windows-1252 (ChrW verwenden)" }
    }
    # Deklarationen nach der ersten Prozedur sind in VBA unzulaessig
    $seenProc = $false; $n = 0
    foreach ($line in [IO.File]::ReadAllLines($f.FullName, [Text.Encoding]::UTF8)) {
        $n++
        if ($line -match '^(Public |Private |Friend )?(Sub|Function|Property) ') { $seenProc = $true }
        elseif ($seenProc -and $line -match '^(Public|Private|Dim|Const|Global)\s+\w+\s+As\s' -and $line -notmatch '^(Public|Private) (Sub|Function|Property|Declare|Enum|Type)') { "$($f.Name) Zeile $n : Deklaration nach Prozedur: $($line.Trim())" }
    }
}

# Zwei Anweisungen versehentlich in einer Zeile verschmolzen (z. B. "props    Case ...")
foreach ($m in (Get-ChildItem (Join-Path (Split-Path $PSScriptRoot -Parent) 'src') -Recurse -Include *.bas,*.cls,*.vba | Select-String -Pattern '\S {4,}(Case |End If|End Sub|End Function|Dim |Set |If |Private |Public |Next|Loop|Exit |Else|\w+ = )')) {
    if ($m.Line -notmatch "^\s*'" -and $m.Line -notmatch '"[^"]* {4,}') { "$($m.Filename) Zeile $($m.LineNumber) : verschmolzene Zeile? $($m.Line.Trim())" }
}

# Zugriff auf cTableData.Data von aussen kopiert das ganze Array (Leistung) - Cell(r, c) verwenden
foreach ($m in (Get-ChildItem (Join-Path (Split-Path $PSScriptRoot -Parent) 'src') -Recurse -Include *.bas,*.cls,*.vba | Where-Object { $_.Name -ne 'cTableData.cls' } | Select-String -Pattern '\.Data\(')) {
    "$($m.Filename) Zeile $($m.LineNumber) : .Data( von aussen - .Cell( verwenden"
}
