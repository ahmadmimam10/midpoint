# Reproduzierbarer Build von MidPoint (Excel-Anwendung)
# Voraussetzung (nur Entwicklungsrechner): Excel 365 Desktop, Option
# "Zugriff auf das VBA-Projektobjektmodell vertrauen". Anwenderrechner brauchen das nicht.
#
# Aufruf:
#   powershell -ExecutionPolicy Bypass -File build\build.ps1 -Target master|demo|stress|gruen|dev
#   powershell -ExecutionPolicy Bypass -File build\build.ps1 -Update "<Pfad zur Fachdatei>"
param(
    [string]$Target = "master",
    [string]$OutFile = "",
    [string]$Update = "",
    [switch]$KeepOpen
)
. "$PSScriptRoot\lib.ps1"
$root = Split-Path $PSScriptRoot -Parent
$src = Join-Path $root 'src'
$dist = Join-Path $root 'dist'
$tmp = Join-Path $env:TEMP ('dbk_build_' + [guid]::NewGuid().ToString('N'))
New-Item -ItemType Directory -Force $dist | Out-Null

function Run-Macro($xl, $wb, [string]$name) {
    $ref = "'" + $wb.Name + "'!" + $name
    Write-Output ("  Makro " + $name)
    [void]$xl.Run($ref)
}

$xl = New-ExcelApp $true
try {
    if ($Update -ne "") {
        # Programmcode einer bestehenden Fachdatei aktualisieren (Daten bleiben erhalten)
        $copy = [System.IO.Path]::ChangeExtension($Update, $null).TrimEnd('.') + "_aktualisiert.xlsm"
        Copy-Item $Update $copy -Force
        $xl.AutomationSecurity = 3
        $wb = $xl.Workbooks.Open($copy)
        $proj = $wb.VBProject
        foreach ($c in @($proj.VBComponents)) { if ($c.Type -ne 100) { $proj.VBComponents.Remove($c) } }
        # entfernte Komponenten sind erst nach Speichern/Schliessen frei (gleichnamige Formulare)
        $wb.Save(); $wb.Close($false)
        # Makros erlaubt (fuer Katalog/Abschluss); Workbook_Open laeuft nicht, da Ereignisse aus sind
        $xl.AutomationSecurity = 1
        $wb = $xl.Workbooks.Open($copy)
        Import-VbaSources $wb $src $tmp
        Embed-Resources $wb $src
        Run-Macro $xl $wb "mBuild.Build_Logos"
        Run-Macro $xl $wb "mBuild.Build_CodeCatalog"
        Run-Macro $xl $wb "mBuild.Build_Finalize"
        # ohne Ereignisse speichern (sonst greift das kontrollierte Speichern der Anwendung)
        $xl.EnableEvents = $false
        $wb.Save()
        $wb.Close($false)
        Write-Output "Aktualisiert: $copy"
        return
    }
    $wb = $xl.Workbooks.Add(-4167)
    Import-VbaSources $wb $src $tmp
    # Gesamtes VBA-Projekt kompilieren; ein Fehlerdialog bricht den Build ab
    $xl.VBE.ActiveVBProject = $wb.VBProject
    $ctl = $xl.VBE.CommandBars.FindControl(1, 578)
    if ($ctl.Enabled) { $ctl.Execute() }
    Start-Sleep 2
    Add-Type -Namespace DbkWin -Name U -MemberDefinition '[DllImport("user32.dll", CharSet=CharSet.Unicode)] public static extern System.IntPtr FindWindow(string c, string t);' -ErrorAction SilentlyContinue
    if ([DbkWin.U]::FindWindow('#32770', 'Microsoft Visual Basic for Applications') -ne [IntPtr]::Zero) {
        Write-Output 'ABBRUCH: Kompilierfehler im VBA-Projekt (siehe VBA-Editor). Es wurde keine Datei erzeugt.'
        $KeepOpen = $true
        return
    }
    Write-Output '  VBA-Projekt fehlerfrei kompiliert'
    switch ($Target) {
        "dev" {
            $out = if ($OutFile) { $OutFile } else { Join-Path $root 'proto\dev.xlsm' }
            if (Test-Path $out) { Remove-Item $out -Force }
            $wb.SaveAs($out, 52)
            Write-Output "Entwicklungsstand (ohne Initialisierung): $out"
            $KeepOpen = $true
        }
        default {
            # zuerst als XLSM speichern, damit Makroaufrufe eindeutig adressiert sind
            $pre = Join-Path $tmp ("build_" + $Target + ".xlsm")
            $wb.SaveAs($pre, 52)
            Run-Macro $xl $wb "mBuild.Build_Init"
            Embed-Resources $wb $src
            Run-Macro $xl $wb "mBuild.Build_Logos"
            if ($Target -eq "demo") { Run-Macro $xl $wb "mBuild.Build_Demo" }
            if ($Target -eq "stress") { Run-Macro $xl $wb "mBuild.Build_Stress" }
            if ($Target -eq "gruen") { Run-Macro $xl $wb "mBuild.Build_Gruen" }
            Run-Macro $xl $wb "mBuild.Build_CodeCatalog"
            Run-Macro $xl $wb "mBuild.Build_Finalize"
            $name = switch ($Target) { "demo" { "MidPoint_Demo.xlsm" } "stress" { "MidPoint_Lasttest.xlsm" } "gruen" { "Beispiel_Gruenflaechenamt_Spielplaetze.xlsm" } default { "MidPoint_Master.xlsm" } }
            $out = if ($OutFile) { $OutFile } else { Join-Path $dist $name }
            if (Test-Path $out) { Remove-Item $out -Force }
            $xl.EnableEvents = $false
            $wb.SaveAs($out, 52)
            $cleanOut = $out
            Write-Output "Erstellt: $out"
        }
    }
}
finally {
    if (-not $KeepOpen) {
        $xl.DisplayAlerts = $false
        foreach ($w in @($xl.Workbooks)) { $w.Close($false) }
        $xl.Quit()
        [void][Runtime.InteropServices.Marshal]::ReleaseComObject($xl)
        # Speicherort und Windows-Benutzername (VBA-Verweis auf MSForms.exd) aus der Lieferdatei entfernen
        if ($cleanOut -and (Test-Path $cleanOut)) {
            & powershell -ExecutionPolicy Bypass -File (Join-Path $PSScriptRoot 'entspuren.ps1') -Path $cleanOut
        }
    } else {
        $xl.EnableEvents = $true
    }
    Remove-Item $tmp -Recurse -Force -ErrorAction SilentlyContinue
}
