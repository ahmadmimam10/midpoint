# Entwicklungshilfe: Logo-Dateien fuer die Hauptdatei aufbereiten (src\res).
# Quelle: Ordner Logos (MidPoint). Beschneiden, Stoerpixel entfernen, weisse Variante,
# feste Groessen. Ergebnis wird beim Build in die Hauptdatei eingebettet.
Add-Type -AssemblyName System.Drawing
$root = Split-Path $PSScriptRoot -Parent
$res = Join-Path $root 'src\res'
New-Item -ItemType Directory -Force $res | Out-Null

function Get-Cleaned([string]$path, [bool]$white) {
    $src = New-Object System.Drawing.Bitmap($path)
    $w = $src.Width; $h = $src.Height
    $out = New-Object System.Drawing.Bitmap($w, $h, [System.Drawing.Imaging.PixelFormat]::Format32bppArgb)
    $minX = $w; $minY = $h; $maxX = 0; $maxY = 0
    for ($y = 0; $y -lt $h; $y++) {
        for ($x = 0; $x -lt $w; $x++) {
            $c = $src.GetPixel($x, $y)
            if ($c.A -lt 40) { continue }
            # Stoerpixel (reines Rot ohne Blauanteil) gehoeren nicht zum Logo
            if ($c.G -lt 60 -and $c.B -lt 35 -and $c.R -gt 120) { continue }
            if ($white) { $n = [System.Drawing.Color]::FromArgb($c.A, 255, 255, 255) } else { $n = $c }
            $out.SetPixel($x, $y, $n)
            if ($x -lt $minX) { $minX = $x }; if ($x -gt $maxX) { $maxX = $x }
            if ($y -lt $minY) { $minY = $y }; if ($y -gt $maxY) { $maxY = $y }
        }
    }
    $src.Dispose()
    $rect = New-Object System.Drawing.Rectangle($minX, $minY, ($maxX - $minX + 1), ($maxY - $minY + 1))
    $crop = $out.Clone($rect, $out.PixelFormat)
    $out.Dispose()
    return $crop
}

function Save-Scaled($bmp, [int]$targetW, [string]$file) {
    $th = [int]($bmp.Height * $targetW / $bmp.Width)
    $dst = New-Object System.Drawing.Bitmap($targetW, $th, [System.Drawing.Imaging.PixelFormat]::Format32bppArgb)
    $g = [System.Drawing.Graphics]::FromImage($dst)
    $g.InterpolationMode = [System.Drawing.Drawing2D.InterpolationMode]::HighQualityBicubic
    $g.SmoothingMode = [System.Drawing.Drawing2D.SmoothingMode]::HighQuality
    $g.PixelOffsetMode = [System.Drawing.Drawing2D.PixelOffsetMode]::HighQuality
    $g.DrawImage($bmp, 0, 0, $targetW, $th)
    $g.Dispose()
    $dst.Save((Join-Path $res $file), [System.Drawing.Imaging.ImageFormat]::Png)
    $dst.Dispose()
    Write-Output ("  " + $file + " " + $targetW + "x" + $th)
}

# Symbole verkleinert verarbeiten (schneller), danach auf Zielgroesse
$tmp = Join-Path $env:TEMP 'midpoint_logo_tmp.png'
foreach ($spec in @(@('04_symbol_farbe.png', 'logo_symbol.png', $false, 256), @('04_symbol_farbe.png', 'logo_symbol_weiss.png', $true, 256),
                    @('01_horizontal_farbe.png', 'logo_horizontal.png', $false, 900), @('01_horizontal_farbe.png', 'logo_horizontal_weiss.png', $true, 900))) {
    $orig = New-Object System.Drawing.Bitmap((Join-Path $root ('Logos\' + $spec[0])))
    $sw = [Math]::Min(700, $orig.Width)
    $sh = [int]($orig.Height * $sw / $orig.Width)
    $small = New-Object System.Drawing.Bitmap($sw, $sh, [System.Drawing.Imaging.PixelFormat]::Format32bppArgb)
    $g = [System.Drawing.Graphics]::FromImage($small)
    $g.InterpolationMode = [System.Drawing.Drawing2D.InterpolationMode]::HighQualityBicubic
    $g.DrawImage($orig, 0, 0, $sw, $sh)
    $g.Dispose(); $orig.Dispose()
    $small.Save($tmp, [System.Drawing.Imaging.ImageFormat]::Png); $small.Dispose()
    $clean = Get-Cleaned $tmp $spec[2]
    Save-Scaled $clean $spec[3] $spec[1]
    $clean.Dispose()
}
Remove-Item $tmp -ErrorAction SilentlyContinue
