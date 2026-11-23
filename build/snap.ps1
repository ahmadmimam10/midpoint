# Entwicklungshilfe: Bildschirmaufnahme (ohne Taskleiste) als JPG fuer Pruefbericht/Paket.
param([string]$Name, [string]$Dir = "")
Add-Type -AssemblyName System.Drawing, System.Windows.Forms
if ($Dir -eq "") { $Dir = Join-Path (Split-Path $PSScriptRoot -Parent) 'pruefung\screenshots' }
New-Item -ItemType Directory -Force $Dir | Out-Null
$b = [System.Windows.Forms.Screen]::PrimaryScreen.Bounds
$h = $b.Height - 48
$bmp = New-Object System.Drawing.Bitmap $b.Width, $h
$g = [System.Drawing.Graphics]::FromImage($bmp)
$g.CopyFromScreen(0, 0, 0, 0, $bmp.Size)
$codec = [System.Drawing.Imaging.ImageCodecInfo]::GetImageEncoders() | Where-Object { $_.MimeType -eq 'image/jpeg' }
$ep = New-Object System.Drawing.Imaging.EncoderParameters 1
$ep.Param[0] = New-Object System.Drawing.Imaging.EncoderParameter ([System.Drawing.Imaging.Encoder]::Quality), 88L
$out = Join-Path $Dir ($Name + '.jpg')
$bmp.Save($out, $codec, $ep)
$g.Dispose(); $bmp.Dispose()
$out
