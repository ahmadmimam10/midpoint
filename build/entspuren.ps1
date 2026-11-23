# Entwicklungshilfe: entfernt Spuren des Entwicklungsrechners aus Excel-Dateien (xlsx, xlsm).
#  1. xl/workbook.xml: Speicherort der Datei (x15ac:absPath), den Excel beim Speichern einträgt.
#  1b. docProps/core.xml: Autor und zuletzt gespeichert von (Excel trägt den Namen des Office-Kontos ein) werden geleert,
#      Erstellt und Geändert auf den Monatsersten gelegt (Entwicklungstage nur mit Monat und Jahr).
#  2. xl/vbaProject.bin: Der VBA-Editor legt den Verweis auf MSForms.exd mit dem Pfad
#     %TEMP%\VBE\MSForms.exd ab, also mit dem Windows-Benutzernamen (Streams dir, __SRP_n, _VBA_PROJECT).
#     Der Name wird durch x in gleicher Länge ersetzt; Aufbau und Größe der Datei bleiben gleich.
#     Excel legt MSForms.exd beim Öffnen ohnehin im eigenen TEMP-Ordner neu an.
#  Danach wird jede Datei erneut durchsucht: Speicherort, Profilpfad und der Benutzername als
#  eigenes Wort in allen XML-Teilen. Ein Rest beendet das Skript mit Exitcode 1.
#
# Aufruf: powershell -ExecutionPolicy Bypass -File build\entspuren.ps1 -Path <Datei oder Ordner> [...] [-Pruefen]
#   -Pruefen: nur suchen und melden, nichts ändern
param([Parameter(Mandatory = $true)][string[]]$Path, [switch]$Pruefen)
$ErrorActionPreference = 'Stop'
Add-Type -AssemblyName System.IO.Compression
Add-Type -AssemblyName System.IO.Compression.FileSystem

$code = @'
using System;
using System.Collections.Generic;
using System.IO;
using System.IO.Compression;
using System.Text;
using System.Text.RegularExpressions;

public static class MpEntspuren
{
    private class Entry { public string Name; public int Type; public uint L, R, C, Start; public long Size; public string Path; }

    // Verbunddokument (Compound File Binary) lesen; Streams werden ueber ihre Sektorketten adressiert
    private class Cfb
    {
        public byte[] B; public int SS, MSS; public uint Cutoff;
        public List<uint> Fat = new List<uint>(), MiniFat = new List<uint>(), RootChain;
        public List<Entry> Entries = new List<Entry>(), Streams = new List<Entry>();
        private HashSet<uint> seen = new HashSet<uint>();

        public Cfb(byte[] b)
        {
            B = b;
            if (b.Length < 512 || BitConverter.ToUInt64(b, 0) != 0xE11AB1A1E011CFD0UL) throw new Exception("vbaProject.bin ist kein Verbunddokument");
            SS = 1 << BitConverter.ToUInt16(b, 0x1E);
            MSS = 1 << BitConverter.ToUInt16(b, 0x20);
            uint nfat = BitConverter.ToUInt32(b, 0x2C), dir1 = BitConverter.ToUInt32(b, 0x30);
            Cutoff = BitConverter.ToUInt32(b, 0x38);
            uint mf1 = BitConverter.ToUInt32(b, 0x3C), dif1 = BitConverter.ToUInt32(b, 0x44), ndif = BitConverter.ToUInt32(b, 0x48);
            var difat = new List<uint>();
            for (int i = 0; i < 109; i++) difat.Add(BitConverter.ToUInt32(b, 0x4C + 4 * i));
            uint s = dif1;
            for (uint k = 0; k < ndif; k++)
            {
                long off = ((long)s + 1) * SS; int n = SS / 4;
                for (int i = 0; i < n - 1; i++) difat.Add(BitConverter.ToUInt32(b, (int)(off + 4 * i)));
                s = BitConverter.ToUInt32(b, (int)(off + 4 * (n - 1)));
            }
            for (int i = 0; i < nfat; i++) AddSector(Fat, difat[i]);
            foreach (uint sec in Chain(dir1, Fat))
            {
                long off = ((long)sec + 1) * SS;
                for (int e = 0; e < SS / 128; e++)
                {
                    int eo = (int)(off + e * 128);
                    var en = new Entry();
                    int nl = BitConverter.ToUInt16(b, eo + 64);
                    en.Name = nl >= 2 ? Encoding.Unicode.GetString(b, eo, Math.Min(nl, 64) - 2) : "";
                    en.Type = b[eo + 66];
                    en.L = BitConverter.ToUInt32(b, eo + 68); en.R = BitConverter.ToUInt32(b, eo + 72); en.C = BitConverter.ToUInt32(b, eo + 76);
                    en.Start = BitConverter.ToUInt32(b, eo + 116);
                    en.Size = (long)BitConverter.ToUInt64(b, eo + 120);
                    if (SS == 512) en.Size &= 0xFFFFFFFFL;
                    Entries.Add(en);
                }
            }
            foreach (uint sec in Chain(mf1, Fat)) AddSector(MiniFat, sec);
            RootChain = Chain(Entries[0].Start, Fat);
            Walk(Entries[0].C, "");
        }

        private void AddSector(List<uint> t, uint sec)
        {
            long off = ((long)sec + 1) * SS;
            for (int j = 0; j < SS / 4; j++) t.Add(BitConverter.ToUInt32(B, (int)(off + 4 * j)));
        }

        public List<uint> Chain(uint start, List<uint> t)
        {
            var r = new List<uint>();
            while (start < 0xFFFFFFFA)
            {
                if (start >= t.Count || r.Count > 4000000) throw new Exception("vbaProject.bin: Sektorkette defekt");
                r.Add(start); start = t[(int)start];
            }
            return r;
        }

        private void Walk(uint i, string path)
        {
            if (i == 0xFFFFFFFF || i >= Entries.Count || !seen.Add(i)) return;
            var e = Entries[(int)i];
            Walk(e.L, path); Walk(e.R, path);
            e.Path = path + "/" + e.Name;
            if (e.Type == 2) Streams.Add(e);
            else if (e.Type == 1) Walk(e.C, e.Path);
        }

        // Dateiposition jedes Stream-Bytes
        public long[] Map(Entry e)
        {
            var map = new long[e.Size];
            if (e.Size < Cutoff)
            {
                var ch = Chain(e.Start, MiniFat);
                for (long k = 0; k < e.Size; k++)
                {
                    long mo = (long)ch[(int)(k / MSS)] * MSS + k % MSS;
                    map[k] = ((long)RootChain[(int)(mo / SS)] + 1) * SS + mo % SS;
                }
            }
            else
            {
                var ch = Chain(e.Start, Fat);
                for (long k = 0; k < e.Size; k++) map[k] = ((long)ch[(int)(k / SS)] + 1) * SS + k % SS;
            }
            return map;
        }
    }

    // MS-OVBA-Entpacken; src (falls angegeben) haelt je Ausgabebyte die Position des Klartextbytes, sonst -1
    private static byte[] Decompress(byte[] d, List<int> src)
    {
        var o = new List<byte>(d.Length * 3);
        if (d.Length == 0 || d[0] != 1) throw new Exception("dir-Stream: unbekanntes Format");
        int pos = 1;
        while (pos < d.Length - 1)
        {
            int h = d[pos] | (d[pos + 1] << 8);
            int size = (h & 0x0FFF) + 3; bool packed = (h & 0x8000) != 0;
            int cs = pos + 2, ce = Math.Min(pos + size, d.Length), ds = o.Count;
            if (!packed)
            {
                for (int k = 0; k < 4096 && cs + k < d.Length; k++) { o.Add(d[cs + k]); if (src != null) src.Add(cs + k); }
                pos = cs + 4096; continue;
            }
            int p = cs;
            while (p < ce)
            {
                int fl = d[p++];
                for (int bit = 0; bit < 8 && p < ce; bit++)
                {
                    if (((fl >> bit) & 1) == 0) { o.Add(d[p]); if (src != null) src.Add(p); p++; }
                    else
                    {
                        int tok = d[p] | (d[p + 1] << 8); p += 2;
                        int diff = o.Count - ds, bc = 4;
                        while ((1 << bc) < diff) bc++;
                        int len = (tok & (0xFFFF >> bc)) + 3, from = o.Count - ((tok >> (16 - bc)) + 1);
                        for (int k = 0; k < len; k++) { o.Add(o[from + k]); if (src != null) src.Add(-1); }
                    }
                }
            }
            pos = ce;
        }
        return o.ToArray();
    }

    private static byte Low(byte c) { return (c >= 65 && c <= 90) ? (byte)(c + 32) : c; }

    // Fundstellen von "\users\<name>\" ohne Gross-/Kleinschreibung (einfach oder UTF-16LE); Ergebnis: Beginn des Namens
    private static List<int> FindProfile(byte[] d, string name, bool wide)
    {
        var res = new List<int>();
        string pat = "\\users\\" + name + "\\";
        int w = wide ? 2 : 1, n = pat.Length * w;
        for (int i = 0; i + n <= d.Length; i++)
        {
            bool ok = true;
            for (int k = 0; k < pat.Length && ok; k++)
            {
                if (Low(d[i + k * w]) != (byte)pat[k]) ok = false;
                else if (wide && d[i + k * w + 1] != 0) ok = false;
            }
            if (ok) res.Add(i + 7 * w);
        }
        return res;
    }

    private static int CleanVba(byte[] b, string[] names, bool apply, List<string> log)
    {
        var cfb = new Cfb(b);
        int hits = 0;
        foreach (var e in cfb.Streams)
        {
            if (e.Size == 0) continue;
            long[] map = cfb.Map(e);
            var data = new byte[e.Size];
            for (long k = 0; k < e.Size; k++) data[k] = b[map[k]];
            bool isDir = e.Path.EndsWith("/VBA/dir", StringComparison.OrdinalIgnoreCase);
            List<int> src = isDir ? new List<int>() : null;
            byte[] plain = isDir ? Decompress(data, src) : data;
            byte[] expect = isDir ? (byte[])plain.Clone() : null;
            int found = 0;
            foreach (string nm in names)
                foreach (bool wide in new bool[] { false, true })
                    foreach (int at in FindProfile(plain, nm, wide))
                    {
                        found++;
                        log.Add("vbaProject.bin" + e.Path + (wide ? " (UTF-16)" : "") + ": Profilpfad mit Benutzername");
                        if (!apply) continue;
                        int w = wide ? 2 : 1;
                        for (int k = 0; k < nm.Length; k++)
                        {
                            int q = at + k * w;
                            if (isDir)
                            {
                                if (src[q] < 0) throw new Exception("dir-Stream: Pfad liegt nicht als Klartext vor, bitte von Hand pruefen");
                                b[map[src[q]]] = (byte)'x'; expect[q] = (byte)'x';
                            }
                            else b[map[q]] = (byte)'x';
                        }
                    }
            if (apply && isDir && found > 0)
            {
                var neu = new byte[e.Size];
                for (long k = 0; k < e.Size; k++) neu[k] = b[map[k]];
                byte[] check = Decompress(neu, null);
                if (check.Length != expect.Length) throw new Exception("dir-Stream nach Ersetzung fehlerhaft (Laenge)");
                for (int k = 0; k < check.Length; k++)
                    if (check[k] != expect[k]) throw new Exception("dir-Stream nach Ersetzung fehlerhaft (Inhalt)");
            }
            hits += found;
        }
        return hits;
    }

    private static readonly Regex AbsPath = new Regex(
        "<mc:AlternateContent[^>]*>\\s*<mc:Choice Requires=\"x15\">\\s*<x15ac:absPath[^>]*/>\\s*</mc:Choice>\\s*</mc:AlternateContent>",
        RegexOptions.Singleline);

    private static readonly Regex CoreName = new Regex("<(dc:creator|cp:lastModifiedBy)>([^<]+)</(dc:creator|cp:lastModifiedBy)>");
    private static readonly Regex CoreDate = new Regex("<(dcterms:created|dcterms:modified)([^>]*)>(\\d{4})-(\\d{2})-(\\d{2})T([^<]*)</(dcterms:created|dcterms:modified)>");

    private static byte[] ReadAll(ZipArchiveEntry en)
    {
        using (var s = en.Open())
        using (var ms = new MemoryStream()) { s.CopyTo(ms); return ms.ToArray(); }
    }

    // apply = true: bereinigen; Rueckgabe: Fundstellen (bei apply die bereinigten)
    public static string[] Process(string file, string[] names, bool apply)
    {
        var log = new List<string>();
        var changes = new Dictionary<string, byte[]>();
        var utf8 = new UTF8Encoding(false);
        using (var zip = ZipFile.OpenRead(file))
        {
            foreach (var en in zip.Entries)
            {
                string n = en.FullName;
                bool isVba = n.EndsWith("vbaProject.bin", StringComparison.OrdinalIgnoreCase);
                bool isXml = n.EndsWith(".xml", StringComparison.OrdinalIgnoreCase) || n.EndsWith(".rels", StringComparison.OrdinalIgnoreCase)
                             || n.EndsWith(".vml", StringComparison.OrdinalIgnoreCase);
                if (!isVba && !isXml) continue;
                byte[] data = ReadAll(en);
                if (isVba)
                {
                    if (CleanVba(data, names, apply, log) > 0 && apply) changes[n] = data;
                    continue;
                }
                bool bom = data.Length >= 3 && data[0] == 0xEF && data[1] == 0xBB && data[2] == 0xBF;
                string t = utf8.GetString(data, bom ? 3 : 0, data.Length - (bom ? 3 : 0));
                if (AbsPath.IsMatch(t))
                {
                    log.Add(n + ": Speicherort der Datei (absPath)");
                    if (apply)
                    {
                        t = AbsPath.Replace(t, "");
                        byte[] body = utf8.GetBytes(t);
                        byte[] all = new byte[body.Length + (bom ? 3 : 0)];
                        if (bom) { all[0] = 0xEF; all[1] = 0xBB; all[2] = 0xBF; }
                        Buffer.BlockCopy(body, 0, all, bom ? 3 : 0, body.Length);
                        changes[n] = all;
                    }
                }
                if (n.Equals("docProps/core.xml", StringComparison.OrdinalIgnoreCase))
                {
                    bool dirty = false;
                    foreach (Match m in CoreName.Matches(t)) { log.Add(n + ": " + m.Groups[1].Value + " = " + m.Groups[2].Value); dirty = true; }
                    foreach (Match m in CoreDate.Matches(t)) { if (m.Groups[5].Value != "01" || m.Groups[6].Value != "00:00:00Z") { log.Add(n + ": " + m.Groups[1].Value + " " + m.Groups[3].Value + "-" + m.Groups[4].Value + "-" + m.Groups[5].Value); dirty = true; } }
                    if (dirty && apply)
                    {
                        t = CoreName.Replace(t, "<$1></$1>");
                        t = CoreDate.Replace(t, "<$1$2>$3-$4-01T00:00:00Z</$1>");
                        byte[] body = utf8.GetBytes(t);
                        byte[] all = new byte[body.Length + (bom ? 3 : 0)];
                        if (bom) { all[0] = 0xEF; all[1] = 0xBB; all[2] = 0xBF; }
                        Buffer.BlockCopy(body, 0, all, bom ? 3 : 0, body.Length);
                        changes[n] = all;
                    }
                }
                if (apply) continue;
                foreach (string nm in names)
                {
                    var rx = new Regex("(?<![A-Za-z])" + Regex.Escape(nm) + "(?![A-Za-z])", RegexOptions.IgnoreCase);
                    foreach (Match m in rx.Matches(t))
                    {
                        int a = Math.Max(0, m.Index - 40);
                        string ctx = t.Substring(a, Math.Min(t.Length - a, m.Length + 80)).Replace("\r", " ").Replace("\n", " ");
                        log.Add(n + ": Benutzername im Text: " + ctx);
                    }
                }
            }
        }
        if (apply && changes.Count > 0)
        {
            using (var zip = ZipFile.Open(file, ZipArchiveMode.Update))
            {
                foreach (var kv in changes)
                {
                    var en = zip.GetEntry(kv.Key);
                    DateTimeOffset t = en.LastWriteTime;
                    using (var s = en.Open()) { s.SetLength(0); s.Write(kv.Value, 0, kv.Value.Length); }
                    en.LastWriteTime = t;
                }
            }
        }
        return log.ToArray();
    }
}
'@
if (-not ('MpEntspuren' -as [type])) {
    $refs = @([System.IO.Compression.ZipArchive].Assembly.Location, [System.IO.Compression.ZipFile].Assembly.Location)
    Add-Type -TypeDefinition $code -ReferencedAssemblies $refs
}

$names = @($env:USERNAME, (Split-Path $env:USERPROFILE -Leaf)) | Where-Object { $_ } |
    ForEach-Object { $_.ToLowerInvariant() } | Select-Object -Unique
$files = foreach ($p in $Path) {
    if (Test-Path $p -PathType Container) { Get-ChildItem $p -Recurse -File -Include *.xlsx, *.xlsm | Where-Object { $_.Name -notlike '~$*' } }
    else { Get-Item $p }
}
$rest = 0; $fixedFiles = 0
foreach ($f in $files) {
    if (-not $Pruefen) {
        $done = [MpEntspuren]::Process($f.FullName, [string[]]$names, $true)
        if ($done.Count -gt 0) { $fixedFiles++; "  bereinigt: {0} ({1} Stellen)" -f $f.Name, $done.Count }
    }
    $left = [MpEntspuren]::Process($f.FullName, [string[]]$names, $false)
    foreach ($l in $left) { "  SPUR: {0}: {1}" -f $f.Name, $l; $rest++ }
}
"Entspuren: {0} Dateien durchsucht, {1} bereinigt, {2} Spuren übrig" -f @($files).Count, $fixedFiles, $rest
if ($rest -gt 0) { exit 1 }
