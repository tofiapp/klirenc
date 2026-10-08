# Kontrola Clearance - pracovní pomocník pro ruční procházení řádků z Excelu
# Spuštění: powershell.exe -NoProfile -STA -ExecutionPolicy Bypass -File .\klirenc.ps1
# Okno je ve WPF (součást Windows) - písmo se vykresluje hladce i při zvětšeném zobrazení.

$script:AppVersion = '25'   # zobrazuje se v titulku okna - podle ní se pozná, která verze běží

Add-Type -AssemblyName PresentationFramework
Add-Type -AssemblyName PresentationCore
Add-Type -AssemblyName WindowsBase
Add-Type -AssemblyName System.Drawing          # jen pro kreslení ikony KC

# Pomocný kód pro Windows (identita na liště, sledování Ctrl + kliknutí). Kompilace trvá několik sekund,
# proto se zkompilovaná knihovna uloží do %LOCALAPPDATA%\KontrolaClearance a při dalších spuštěních se jen načte.
$script:NativeSource = @'
using System;
using System.Diagnostics;
using System.Runtime.InteropServices;

public static class KlirencCtrlClick {
    private delegate IntPtr HookProc(int nCode, IntPtr wParam, IntPtr lParam);

    [StructLayout(LayoutKind.Sequential)]
    private struct POINT { public int X; public int Y; }

    [StructLayout(LayoutKind.Sequential)]
    private struct MSLLHOOKSTRUCT { public POINT pt; public uint mouseData; public uint flags; public uint time; public IntPtr extra; }

    [DllImport("user32.dll", SetLastError = true)] private static extern IntPtr SetWindowsHookEx(int idHook, HookProc fn, IntPtr hMod, uint threadId);
    [DllImport("user32.dll")] private static extern bool UnhookWindowsHookEx(IntPtr hook);
    [DllImport("user32.dll")] private static extern IntPtr CallNextHookEx(IntPtr hook, int nCode, IntPtr wParam, IntPtr lParam);
    [DllImport("kernel32.dll", CharSet = CharSet.Unicode)] private static extern IntPtr GetModuleHandle(string name);
    [DllImport("user32.dll")] private static extern IntPtr WindowFromPoint(POINT pt);
    [DllImport("user32.dll")] private static extern IntPtr GetAncestor(IntPtr hWnd, uint flags);
    [DllImport("user32.dll")] private static extern uint GetWindowThreadProcessId(IntPtr hWnd, out uint pid);
    [DllImport("user32.dll")] private static extern short GetAsyncKeyState(int vk);
    [DllImport("user32.dll")] private static extern void keybd_event(byte vk, byte scan, uint flags, UIntPtr extra);
    [DllImport("user32.dll")] private static extern void mouse_event(uint flags, int dx, int dy, uint data, UIntPtr extra);

    private const int WH_MOUSE_LL = 14;
    private const int WM_LBUTTONDOWN = 0x0201;
    private const int WM_LBUTTONUP = 0x0202;
    private const uint LLMHF_INJECTED = 0x01;
    private const byte VK_CONTROL = 0x11;
    private const uint KEYUP = 0x0002;

    private static HookProc proc = Callback;   // drží delegáta, aby ho neuklidil GC
    private static IntPtr hook = IntPtr.Zero;
    private static bool swallowedDown = false;
    private static readonly uint ownPid = (uint)Process.GetCurrentProcess().Id;

    public static volatile bool Pending = false;
    public static volatile bool Enabled = true;

    public static bool Start() {
        if (hook != IntPtr.Zero) return true;
        using (Process p = Process.GetCurrentProcess())
        using (ProcessModule m = p.MainModule) {
            hook = SetWindowsHookEx(WH_MOUSE_LL, proc, GetModuleHandle(m.ModuleName), 0);
        }
        return hook != IntPtr.Zero;
    }

    public static void Stop() {
        if (hook != IntPtr.Zero) { UnhookWindowsHookEx(hook); hook = IntPtr.Zero; }
    }

    // Leží bod v okně jiné aplikace než tato aplikace?
    private static bool PointIsOther(POINT pt) {
        IntPtr w = GetAncestor(WindowFromPoint(pt), 2);
        if (w == IntPtr.Zero) return false;
        uint pid;
        GetWindowThreadProcessId(w, out pid);
        return pid != 0 && pid != ownPid;
    }

    // Ctrl + levé kliknutí v jiné aplikaci: původní kliknutí se zahodí (aby se neprovedla akce Ctrl+klik)
    // a po puštění tlačítka se vloží aktuální údaj
    private static IntPtr Callback(int nCode, IntPtr wParam, IntPtr lParam) {
        try {
            if (nCode >= 0 && Enabled) {
                int msg = wParam.ToInt32();
                if (msg == WM_LBUTTONDOWN || msg == WM_LBUTTONUP) {
                    MSLLHOOKSTRUCT info = (MSLLHOOKSTRUCT)Marshal.PtrToStructure(lParam, typeof(MSLLHOOKSTRUCT));
                    if ((info.flags & LLMHF_INJECTED) == 0) {
                        if (msg == WM_LBUTTONDOWN && (GetAsyncKeyState(VK_CONTROL) & 0x8000) != 0 && PointIsOther(info.pt)) {
                            swallowedDown = true;
                            return (IntPtr)1;
                        }
                        if (msg == WM_LBUTTONUP && swallowedDown) {
                            swallowedDown = false;
                            Pending = true;
                            return (IntPtr)1;
                        }
                    }
                }
            }
        } catch { }
        return CallNextHookEx(hook, nCode, wParam, lParam);
    }

    private static void CtrlKey(byte vk) {
        keybd_event(VK_CONTROL, 0, 0, UIntPtr.Zero);
        keybd_event(vk, 0, 0, UIntPtr.Zero);
        keybd_event(vk, 0, KEYUP, UIntPtr.Zero);
        keybd_event(VK_CONTROL, 0, KEYUP, UIntPtr.Zero);
    }

    // Pustit Ctrl (uživatel ho ještě drží), obyčejně kliknout na místo kurzoru (aktivuje pole),
    // označit obsah pole (Ctrl+A) a vložit (Ctrl+V) - obsah pole se nahradí
    public static void ClickSelectAllPaste() {
        const uint LEFTDOWN = 0x0002, LEFTUP = 0x0004;
        keybd_event(VK_CONTROL, 0, KEYUP, UIntPtr.Zero);
        System.Threading.Thread.Sleep(20);
        mouse_event(LEFTDOWN, 0, 0, 0, UIntPtr.Zero);
        mouse_event(LEFTUP, 0, 0, 0, UIntPtr.Zero);
        System.Threading.Thread.Sleep(80);
        CtrlKey(0x41);   // A
        System.Threading.Thread.Sleep(30);
        CtrlKey(0x56);   // V
    }
}

public static class KcTaskbar {
    [DllImport("shell32.dll", CharSet = CharSet.Unicode)]
    public static extern int SetCurrentProcessExplicitAppUserModelID(string appId);
    [DllImport("user32.dll")]
    public static extern bool SetProcessDPIAware();
}
'@
function Import-NativeCode {
    if ('KlirencCtrlClick' -as [type]) { return }
    try {
        $sha = [System.Security.Cryptography.SHA1]::Create()
        $hash = -join ($sha.ComputeHash([System.Text.Encoding]::UTF8.GetBytes($script:NativeSource))[0..7] | ForEach-Object { $_.ToString('x2') })
        $dir = Join-Path $env:LOCALAPPDATA 'KontrolaClearance'
        $dll = Join-Path $dir "kc-native-$hash.dll"
        if (-not (Test-Path $dll)) {
            if (-not (Test-Path $dir)) { [void](New-Item -ItemType Directory -Path $dir) }
            Get-ChildItem -Path $dir -Filter 'kc-native-*.dll' -ErrorAction SilentlyContinue | Remove-Item -ErrorAction SilentlyContinue
            Add-Type -TypeDefinition $script:NativeSource -OutputAssembly $dll -OutputType Library -ErrorAction Stop
        }
        if (-not ('KlirencCtrlClick' -as [type])) { Add-Type -Path $dll -ErrorAction Stop }
    } catch {
        # záloha: kompilace jen v paměti (pomalejší, ale funguje i bez zápisu na disk)
        if (-not ('KlirencCtrlClick' -as [type])) { Add-Type -TypeDefinition $script:NativeSource }
    }
}
try {
    Import-NativeCode
    # vlastní identita procesu: okno se neseskupí s ostatními okny PowerShellu a na liště bude ikona KC
    [void][KcTaskbar]::SetCurrentProcessExplicitAppUserModelID('KontrolaClearance.App')
    [void][KcTaskbar]::SetProcessDPIAware()   # ostré vykreslení při zvětšeném zobrazení Windows
} catch { }

# ---------- Stav ----------
$script:Items = New-Object System.Collections.Generic.List[object]
$script:Index = 0
$script:Syncing = $false       # programové nastavení výběru v seznamu (neřešit jako klik)
$script:Done = $false          # všechny řádky rozhodnuty a zobrazuje se dokončení
$script:PastedCols = 0         # kolik sloupců mělo poslední vložení (před ořezáním)
$script:Collapsed = @{}         # složky (kategorie), které jsou sbalené
$script:GroupFilter = @{}       # filtr v každé složce: @{ S = stanice; L = délka } ('Vše' = bez filtru)
$script:Revision = ''           # číslo revize z 1. sloupce
$script:All = 'Vše'
$script:NotesFile = $null      # soubor, do kterého se poznámky ukládají (po Otevřít / Uložit jako)
$script:MyPath = $MyInvocation.MyCommand.Path
$win = $null

# Status: '' = nerozhodnuto, 'keep' = Ponechat, 'del' = Vymazat (poznámka zapsána), 'unsure' = vrátit se později
# Category = složka (3. sloupec), Station = stanice (9.), Length = délka (10.), MapUrl = odkaz na mapu (16.)
# Manual = záznam přidaný ručně; Num = pořadové číslo (pro zobrazení)
function New-Item2([string]$value, [string]$note, [bool]$manual = $false, [string]$category = '',
                   [string]$station = '', [string]$length = '', [string]$mapUrl = '') {
    [pscustomobject]@{ Value = $value; Note = $note; Status = ''; Manual = $manual; Num = 0; Category = $category
                       Station = $station; Length = $length; MapUrl = $mapUrl }
}

function Show-Msg([string]$text, [string]$icon) {
    if ($win) { [void][System.Windows.MessageBox]::Show($win, $text, 'Kontrola Clearance', 'OK', $icon) }
    else { [void][System.Windows.MessageBox]::Show($text, 'Kontrola Clearance', 'OK', $icon) }
}
function Show-Info([string]$text)  { Show-Msg $text 'Information' }
function Show-Warn([string]$text)  { Show-Msg $text 'Warning' }
function Show-Error([string]$text) { Show-Msg $text 'Error' }
function Ask-YesNo([string]$text, [string]$title, [string]$icon = 'Question') {
    ([System.Windows.MessageBox]::Show($win, $text, $title, 'YesNo', $icon)) -eq 'Yes'
}

# Bezpečné spuštění obsluhy události - uživateli se nezobrazí technická chyba
function Invoke-Safe([scriptblock]$action, [string]$message) {
    try { & $action } catch { Show-Error $message }
}

# Zápis do schránky s několika pokusy (schránku může chvilku držet jiná aplikace)
function Set-ClipboardText([string]$text) {
    for ($i = 0; $i -lt 6; $i++) {
        try { [System.Windows.Clipboard]::SetDataObject($text, $true); return $true } catch { Start-Sleep -Milliseconds 40 }
    }
    return $false
}

# ---------- Parsování vstupu ----------
# Rozdělí vložený sloupec na řádky; koncové prázdné řádky (Excel přidává konec řádku) se odříznou.
function Get-ColumnLines([string]$text) {
    if ([string]::IsNullOrEmpty($text)) { return ,@() }
    $lines = [System.Collections.Generic.List[string]]($text -split "\r\n|\n|\r")
    while ($lines.Count -gt 0 -and [string]::IsNullOrWhiteSpace($lines[$lines.Count - 1])) { $lines.RemoveAt($lines.Count - 1) }
    return ,$lines.ToArray()
}

# Sloupce z Excelu (list ZJISTENI, 37 sloupců), se kterými aplikace pracuje - číslováno od 1
$script:ColRevision = 1     # číslo revize (stejné pro celé načtení)
$script:ColCategory = 3     # složka (datum RRMMDD)
$script:ColStation  = 9     # stanice (filtr)
$script:ColLength   = 10    # délka (filtr)
$script:ColValue    = 12    # Údaj k ověření
$script:ColMap      = 16    # odkaz na mapu (dmwmap://…)
$script:ColNote     = 36    # poznámka při Vymazat
$script:ExcelBook   = '_kontrola_clearance_v4'
$script:ExcelSheet  = 'ZJISTENI'

# Z celého řádku (pole hodnot, index od 0) vybere potřebné údaje.
# Vrací $null pro prázdný řádek, jinak @{ Error } nebo @{ Fields = revize, složka, stanice, délka, údaj, mapa, poznámka }.
function Get-RowFields([string[]]$parts, [int]$rowNo) {
    if ([string]::IsNullOrWhiteSpace(($parts -join ''))) { return $null }
    if ($parts.Count -lt $script:ColNote) {
        return @{ Error = "Řádek $rowNo má jen $($parts.Count) sloupců, aplikace potřebuje aspoň $($script:ColNote). Nic nebylo načteno." }
    }
    $f = @($script:ColRevision, $script:ColCategory, $script:ColStation, $script:ColLength, $script:ColValue, $script:ColMap, $script:ColNote) |
         ForEach-Object { ([string]$parts[$_ - 1]).Replace("`t", ' ').Trim() }
    return @{ Fields = $f }
}

function New-ItemFromFields($f, [bool]$manual = $false) {
    New-Item2 $f[4] $f[6] $manual $f[1] $f[2] $f[3] $f[5]
}

# Ze vloženého textu (Ctrl+V z Excelu) ponechá v každém řádku jen potřebné sloupce.
# Vrací @{ Text; Columns } nebo @{ Error }.
function Reduce-PastedColumns([string]$text) {
    $lines = $text -split "\r\n|\n|\r"
    $columns = 0
    $out = New-Object System.Collections.Generic.List[string]
    for ($i = 0; $i -lt $lines.Count; $i++) {
        $line = $lines[$i]
        if ([string]::IsNullOrWhiteSpace($line.Replace("`t", ''))) { $out.Add(''); continue }
        $parts = $line.Split("`t")
        if ($columns -eq 0) { $columns = $parts.Count }
        elseif ($parts.Count -ne $columns) {
            return @{ Error = "Řádek $($i + 1) má $($parts.Count) sloupců, ale předchozí řádky mají $columns.`nZkopírujte z Excelu souvislý obdélníkový rozsah. Nic nebylo vloženo." }
        }
        $r = Get-RowFields $parts ($i + 1)
        if ($r.Error) { return @{ Error = $r.Error } }
        $out.Add($r.Fields -join "`t")
    }
    while ($out.Count -gt 0 -and $out[$out.Count - 1] -eq '') { $out.RemoveAt($out.Count - 1) }
    return @{ Text = (($out -join "`r`n") + "`r`n"); Columns = $columns }
}

# Z řádků po vložení přes Ctrl+V (7 vybraných sloupců) vytvoří záznamy.
# Vrací @{ Ok; Items; Revision; Error }. Při chybě nevrací žádné položky.
function ConvertFrom-Rows([string]$text, [bool]$manual = $false) {
    $lines = Get-ColumnLines $text
    if ($lines.Count -eq 0) {
        return @{ Ok = $false; Error = 'Vstup je prázdný. Načtěte data z Excelu nebo je vložte do pole (Ctrl+V).' }
    }
    $rows = New-Object System.Collections.Generic.List[object]
    for ($i = 0; $i -lt $lines.Count; $i++) {
        if ([string]::IsNullOrWhiteSpace($lines[$i].Replace("`t", ''))) { continue }
        $parts = $lines[$i].Split("`t")
        if ($parts.Count -ne 7) {
            return @{ Ok = $false; Error = "Řádek $($i + 1) nemá očekávaný tvar. Zkopírujte řádky z Excelu (všechny sloupce) a vložte je do pole přes Ctrl+V.`nNic nebylo načteno." }
        }
        $rows.Add(@{ Fields = $parts; RowNo = $i + 1 })
    }
    return (ConvertFrom-FieldRows $rows $manual)
}

# Společné pro vložení i načtení z Excelu: z vybraných údajů vytvoří záznamy
function ConvertFrom-FieldRows($rows, [bool]$manual = $false) {
    $result = New-Object System.Collections.Generic.List[object]
    $rev = ''
    foreach ($r in $rows) {
        $f = $r.Fields
        if ($f[4] -eq '') {
            if ($f[1] -eq '') { continue }   # řádek bez údaje i složky (např. prázdný konec tabulky)
            return @{ Ok = $false; Error = "Řádek $($r.RowNo) má prázdný $($script:ColValue). sloupec (Údaj k ověření).`nNic nebylo načteno." }
        }
        if ($rev -eq '') { $rev = $f[0] }
        $result.Add((New-ItemFromFields $f $manual))
    }
    return @{ Ok = $true; Items = $result; Revision = $rev }
}

# Text z buňky Excelu (čísla bez desetinných míst a bez formátování)
function ConvertTo-CellText($v) {
    if ($null -eq $v) { return '' }
    if ($v -is [double]) {
        if ([Math]::Floor($v) -eq $v) { return $v.ToString('0', [Globalization.CultureInfo]::InvariantCulture) }
        return $v.ToString([Globalization.CultureInfo]::CurrentCulture)
    }
    return [string]$v
}

# Načte záznamy přímo z otevřeného Excelu: sešit _kontrola_clearance_v4 (uložený lokálně), list ZJISTENI.
# 1. řádek listu je záhlaví a přeskočí se. Vrací stejný výsledek jako ConvertFrom-Rows.
function Read-FromExcel {
    try { $xl = [System.Runtime.InteropServices.Marshal]::GetActiveObject('Excel.Application') }
    catch { return @{ Ok = $false; Error = "Excel není spuštěný.`nOtevřete soubor $($script:ExcelBook), načtěte data (makro) a zkuste to znovu." } }
    $wb = $null
    foreach ($w in $xl.Workbooks) {
        $name = [System.IO.Path]::GetFileNameWithoutExtension([string]$w.Name)
        if ($name -ieq $script:ExcelBook -and ([string]$w.FullName) -notmatch '^https?://') { $wb = $w; break }
    }
    if (-not $wb) {
        return @{ Ok = $false; Error = "V Excelu není otevřený soubor $($script:ExcelBook) uložený v počítači.`n(Online verze se nenačítá.)" }
    }
    $ws = $null
    foreach ($sh in $wb.Worksheets) { if (([string]$sh.Name).Trim() -ieq $script:ExcelSheet) { $ws = $sh; break } }
    if (-not $ws) { return @{ Ok = $false; Error = "V souboru $($script:ExcelBook) chybí list $($script:ExcelSheet)." } }

    $ur = $ws.UsedRange
    $vals = $ur.Value2
    if ($null -eq $vals -or $vals -isnot [array]) { return @{ Ok = $false; Error = "List $($script:ExcelSheet) neobsahuje žádná data." } }
    $firstRow = [int]$ur.Row; $firstCol = [int]$ur.Column
    $r0 = $vals.GetLowerBound(0); $r1 = $vals.GetUpperBound(0)
    $c0 = $vals.GetLowerBound(1); $c1 = $vals.GetUpperBound(1)
    $lastCol = $firstCol + ($c1 - $c0)
    $width = [Math]::Max($lastCol, $script:ColNote)
    if ($lastCol -lt $script:ColNote) {
        return @{ Ok = $false; Error = "List $($script:ExcelSheet) má jen $lastCol sloupců, aplikace potřebuje aspoň $($script:ColNote)." }
    }
    # čtou se jen potřebné sloupce (rychlejší)
    $need = @($script:ColRevision, $script:ColCategory, $script:ColStation, $script:ColLength, $script:ColValue, $script:ColMap, $script:ColNote)
    $rows = New-Object System.Collections.Generic.List[object]
    for ($r = $r0; $r -le $r1; $r++) {
        $sheetRow = $firstRow + ($r - $r0)
        if ($sheetRow -le 1) { continue }   # záhlaví
        $parts = [string[]]::new($width)
        for ($k = 0; $k -lt $width; $k++) { $parts[$k] = '' }
        foreach ($col in $need) {
            $c = $c0 + ($col - $firstCol)
            if ($c -ge $c0 -and $c -le $c1) { $parts[$col - 1] = ConvertTo-CellText $vals[$r, $c] }
        }
        $g = Get-RowFields $parts $sheetRow
        if ($null -eq $g) { continue }
        if ($g.Error) { return @{ Ok = $false; Error = $g.Error } }
        # datum RRMMDD uložené jako číslo ztratí úvodní nulu (např. 050101 -> 50101)
        if ($g.Fields[1] -match '^\d{5}$') { $g.Fields[1] = '0' + $g.Fields[1] }
        $rows.Add(@{ Fields = $g.Fields; RowNo = $sheetRow })
    }
    if ($rows.Count -eq 0) { return @{ Ok = $false; Error = "List $($script:ExcelSheet) neobsahuje žádné záznamy (pod záhlavím)." } }

    # Odkaz na mapu: pokud buňka obsahuje hypertextový odkaz nebo vzorec HYPERLINK, vezme se jeho adresa
    $links = @{}
    try {
        foreach ($h in $ws.Hyperlinks) {
            try {
                if ([int]$h.Range.Column -eq $script:ColMap) {
                    $a = [string]$h.Address
                    if ([string]$h.SubAddress -ne '') { $a = $a + '#' + [string]$h.SubAddress }
                    if ($a -ne '') { $links[[int]$h.Range.Row] = $a }
                }
            } catch { }
        }
    } catch { }
    try {
        $lastRow = $firstRow + ($r1 - $r0)
        $fr = $ws.Range($ws.Cells($firstRow, $script:ColMap), $ws.Cells($lastRow, $script:ColMap)).Formula
        if ($fr -is [array]) {
            $fr0 = $fr.GetLowerBound(0); $fc0 = $fr.GetLowerBound(1)
            for ($k = $fr0; $k -le $fr.GetUpperBound(0); $k++) {
                $fx = [string]$fr[$k, $fc0]
                if ($fx -match '(?i)^=\s*HYPERLINK\(\s*"([^"]+)"') { $links[$firstRow + ($k - $fr0)] = $Matches[1] }
            }
        }
    } catch { }
    foreach ($row in $rows) {
        $f = $row.Fields
        if ($f[5] -notmatch '://' -and $links.ContainsKey([int]$row.RowNo)) { $f[5] = $links[[int]$row.RowNo] }
    }
    return (ConvertFrom-FieldRows $rows)
}

# Klíč pro řazení složek: datum RRMMDD z názvu složky (první šestice číslic). Bez data = $null.
function Get-CategoryDateKey([string]$category) {
    if ($category -match '(?<!\d)(\d{2})(\d{2})(\d{2})(?!\d)') {
        $m = [int]$Matches[2]; $d = [int]$Matches[3]
        if ($m -ge 1 -and $m -le 12 -and $d -ge 1 -and $d -le 31) { return $Matches[1] + $Matches[2] + $Matches[3] }
    }
    return $null
}

# Seřadí záznamy do složek. Složky jsou seřazené podle data RRMMDD v názvu od nejstarší;
# složky bez data jsou na konci v pořadí prvního výskytu. Uvnitř složky zůstává původní pořadí.
function Group-Items($list) {
    $groups = [ordered]@{}
    foreach ($it in $list) {
        if (-not $groups.Contains($it.Category)) { $groups[$it.Category] = New-Object System.Collections.Generic.List[object] }
        $groups[$it.Category].Add($it)
    }
    $pos = 0
    $keys = foreach ($k in $groups.Keys) {
        $dk = Get-CategoryDateKey $k
        [pscustomobject]@{ Name = $k; NoDate = [int]($null -eq $dk); Date = [string]$dk; Pos = $pos++ }
    }
    $out = New-Object System.Collections.Generic.List[object]
    foreach ($k in @($keys | Sort-Object NoDate, Date, Pos)) { $out.AddRange($groups[$k.Name]) }
    return ,$out
}

# ---------- Ikona KC ----------
function RGB([int]$r, [int]$g, [int]$b) { [System.Drawing.Color]::FromArgb($r, $g, $b) }
function New-KcBitmap([int]$size) {
    $bmp = New-Object System.Drawing.Bitmap($size, $size)
    $g = [System.Drawing.Graphics]::FromImage($bmp)
    $g.SmoothingMode = 'AntiAlias'
    $g.TextRenderingHint = 'AntiAliasGridFit'
    $g.Clear([System.Drawing.Color]::Transparent)
    $r = [Math]::Max(2, [int]($size * 0.22)); $d = 2 * $r; $w = $size - 1
    $path = New-Object System.Drawing.Drawing2D.GraphicsPath
    $path.AddArc(0, 0, $d, $d, 180, 90)
    $path.AddArc($w - $d, 0, $d, $d, 270, 90)
    $path.AddArc($w - $d, $w - $d, $d, $d, 0, 90)
    $path.AddArc(0, $w - $d, $d, $d, 90, 90)
    $path.CloseFigure()
    $rect = New-Object System.Drawing.Rectangle(0, 0, $size, $size)
    $brush = New-Object System.Drawing.Drawing2D.LinearGradientBrush($rect, (RGB 8 145 178), (RGB 22 78 99), 45.0)
    $g.FillPath($brush, $path)
    $font = New-Object System.Drawing.Font('Segoe UI', [float]($size * 0.40), [System.Drawing.FontStyle]::Bold, [System.Drawing.GraphicsUnit]::Pixel)
    $fmt = New-Object System.Drawing.StringFormat
    $fmt.Alignment = 'Center'; $fmt.LineAlignment = 'Center'
    $rf = New-Object System.Drawing.RectangleF(0, [float]($size * 0.02), $size, $size)
    $g.DrawString('KC', $font, [System.Drawing.Brushes]::White, $rf, $fmt)
    $font.Dispose(); $brush.Dispose(); $path.Dispose(); $g.Dispose()
    $bmp
}

# Ikona pro okno WPF (PNG s průhledností)
function Get-KcImageSource([int]$size) {
    $bmp = New-KcBitmap $size
    $ms = New-Object System.IO.MemoryStream
    $bmp.Save($ms, [System.Drawing.Imaging.ImageFormat]::Png)
    $bmp.Dispose()
    $ms.Position = 0
    $dec = New-Object System.Windows.Media.Imaging.PngBitmapDecoder($ms, 'PreservePixelFormat', 'OnLoad')
    $dec.Frames[0]
}

# Uloží ikonu KC jako .ico (obrázek PNG 256×256 uvnitř souboru ICO)
function Save-KcIco([string]$path) {
    $bmp = New-KcBitmap 256
    $ms = New-Object System.IO.MemoryStream
    $bmp.Save($ms, [System.Drawing.Imaging.ImageFormat]::Png)
    $png = $ms.ToArray()
    $ms.Dispose(); $bmp.Dispose()
    $fs = [System.IO.File]::Create($path)
    $bw = New-Object System.IO.BinaryWriter($fs)
    $bw.Write([UInt16]0); $bw.Write([UInt16]1); $bw.Write([UInt16]1)       # hlavička: typ ikona, 1 obrázek
    $bw.Write([byte]0); $bw.Write([byte]0); $bw.Write([byte]0); $bw.Write([byte]0)   # 256×256, bez palety
    $bw.Write([UInt16]1); $bw.Write([UInt16]32)                              # roviny, bitů na pixel
    $bw.Write([UInt32]$png.Length); $bw.Write([UInt32]22)                    # velikost dat, posun dat
    $bw.Write($png)
    $bw.Close()
}

# Zástupce „Kontrola Clearance“ s ikonou KC se udržuje automaticky při každém spuštění:
# vždy vede na tuto složku (i po přesunu nebo nahrání nové verze) a má aktuální ikonu.
# Ve složce aplikace a na ploše se vytvoří, připnutý zástupce na hlavním panelu se jen opraví.
function Update-Shortcuts {
    $dir = Split-Path -Parent $script:MyPath
    $ico = Join-Path $dir "kc-$($script:AppVersion).ico"     # název s verzí: Windows si ikony pamatují podle cesty
    if (-not (Test-Path $ico)) {
        Get-ChildItem -Path $dir -Filter 'kc*.ico' -ErrorAction SilentlyContinue | Remove-Item -ErrorAction SilentlyContinue
        Save-KcIco $ico
    }
    $shell = New-Object -ComObject WScript.Shell
    $name = 'Kontrola Clearance.lnk'
    $pinned = Join-Path $env:APPDATA "Microsoft\Internet Explorer\Quick Launch\User Pinned\TaskBar\$name"
    $targets = @(
        @{ Path = (Join-Path $dir $name); Create = $true },
        @{ Path = (Join-Path ([Environment]::GetFolderPath('Desktop')) $name); Create = $true },
        @{ Path = $pinned; Create = $false }
    )
    foreach ($t in $targets) {
        try {
            if (-not $t.Create -and -not (Test-Path $t.Path)) { continue }
            $lnk = $shell.CreateShortcut($t.Path)
            $lnk.TargetPath = Join-Path $env:SystemRoot 'System32\wscript.exe'
            $lnk.Arguments = '"' + (Join-Path $dir 'spustit.vbs') + '"'
            $lnk.WorkingDirectory = $dir
            $lnk.IconLocation = "$ico,0"
            $lnk.Description = 'Kontrola Clearance'
            $lnk.Save()
        } catch { }
    }
}

# ---------- Vzhled okna (XAML) ----------
[xml]$xaml = @'
<Window xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
        Title="Kontrola Clearance" Width="1200" Height="820" MinWidth="950" MinHeight="650"
        WindowStartupLocation="CenterScreen" Background="#F1F5F9"
        FontFamily="Segoe UI" FontSize="14" Foreground="#0F172A"
        UseLayoutRounding="True" SnapsToDevicePixels="True">
  <Window.Resources>
    <SolidColorBrush x:Key="Accent" Color="#0E7490"/>
    <SolidColorBrush x:Key="AccentBg" Color="#E0F7FA"/>
    <SolidColorBrush x:Key="Muted" Color="#64748B"/>
    <SolidColorBrush x:Key="Line" Color="#E2E8F0"/>

    <!-- Zaoblené ploché tlačítko -->
    <Style x:Key="Btn" TargetType="Button">
      <Setter Property="Background" Value="#E8EDF4"/>
      <Setter Property="Foreground" Value="#0F172A"/>
      <Setter Property="Padding" Value="16,8"/>
      <Setter Property="Margin" Value="0,0,8,0"/>
      <Setter Property="Cursor" Value="Hand"/>
      <Setter Property="FocusVisualStyle" Value="{x:Null}"/>
      <Setter Property="Tag" Value="8"/>
      <Setter Property="Template">
        <Setter.Value>
          <ControlTemplate TargetType="Button">
            <Border x:Name="Bd" Background="{TemplateBinding Background}" CornerRadius="8" Padding="{TemplateBinding Padding}">
              <ContentPresenter HorizontalAlignment="Center" VerticalAlignment="Center"/>
            </Border>
            <ControlTemplate.Triggers>
              <Trigger Property="IsMouseOver" Value="True"><Setter TargetName="Bd" Property="Opacity" Value="0.88"/></Trigger>
              <Trigger Property="IsPressed" Value="True"><Setter TargetName="Bd" Property="Opacity" Value="0.75"/></Trigger>
              <Trigger Property="IsEnabled" Value="False"><Setter TargetName="Bd" Property="Opacity" Value="0.4"/></Trigger>
            </ControlTemplate.Triggers>
          </ControlTemplate>
        </Setter.Value>
      </Setter>
    </Style>
    <Style x:Key="BtnPrimary" TargetType="Button" BasedOn="{StaticResource Btn}">
      <Setter Property="Background" Value="{StaticResource Accent}"/>
      <Setter Property="Foreground" Value="White"/>
      <Setter Property="FontWeight" Value="SemiBold"/>
    </Style>
    <Style x:Key="BtnBig" TargetType="Button" BasedOn="{StaticResource Btn}">
      <Setter Property="Foreground" Value="White"/>
      <Setter Property="FontSize" Value="24"/>
      <Setter Property="FontWeight" Value="SemiBold"/>
      <Setter Property="Margin" Value="6,4"/>
      <Setter Property="Template">
        <Setter.Value>
          <ControlTemplate TargetType="Button">
            <Border x:Name="Bd" Background="{TemplateBinding Background}" CornerRadius="14">
              <ContentPresenter HorizontalAlignment="Center" VerticalAlignment="Center"/>
            </Border>
            <ControlTemplate.Triggers>
              <Trigger Property="IsMouseOver" Value="True"><Setter TargetName="Bd" Property="Opacity" Value="0.9"/></Trigger>
              <Trigger Property="IsPressed" Value="True"><Setter TargetName="Bd" Property="Opacity" Value="0.78"/></Trigger>
              <Trigger Property="IsEnabled" Value="False"><Setter TargetName="Bd" Property="Opacity" Value="0.35"/></Trigger>
            </ControlTemplate.Triggers>
          </ControlTemplate>
        </Setter.Value>
      </Setter>
    </Style>
    <Style x:Key="BtnNav" TargetType="Button" BasedOn="{StaticResource Btn}">
      <Setter Property="Width" Value="40"/>
      <Setter Property="Height" Value="34"/>
      <Setter Property="Padding" Value="0"/>
      <Setter Property="Margin" Value="6,0,0,0"/>
      <Setter Property="Background" Value="#F1F5F9"/>
    </Style>

    <Style x:Key="Card" TargetType="Border">
      <Setter Property="Background" Value="White"/>
      <Setter Property="BorderBrush" Value="{StaticResource Line}"/>
      <Setter Property="BorderThickness" Value="1"/>
      <Setter Property="CornerRadius" Value="12"/>
      <Setter Property="Padding" Value="16,14"/>
      <Setter Property="Margin" Value="6"/>
    </Style>

    <Style x:Key="Field" TargetType="TextBox">
      <Setter Property="Background" Value="#F8FAFC"/>
      <Setter Property="BorderBrush" Value="{StaticResource Line}"/>
      <Setter Property="BorderThickness" Value="1"/>
      <Setter Property="Padding" Value="8,6"/>
    </Style>

    <!-- Výběr (stanice / délka) ve stylu aplikace -->
    <Style x:Key="PickItem" TargetType="ComboBoxItem">
      <Setter Property="Template">
        <Setter.Value>
          <ControlTemplate TargetType="ComboBoxItem">
            <Border x:Name="Bd" Background="Transparent" Padding="10,5" CornerRadius="4">
              <ContentPresenter/>
            </Border>
            <ControlTemplate.Triggers>
              <Trigger Property="IsHighlighted" Value="True"><Setter TargetName="Bd" Property="Background" Value="#F1F5F9"/></Trigger>
              <Trigger Property="IsSelected" Value="True">
                <Setter TargetName="Bd" Property="Background" Value="{StaticResource AccentBg}"/>
                <Setter Property="Foreground" Value="{StaticResource Accent}"/>
              </Trigger>
            </ControlTemplate.Triggers>
          </ControlTemplate>
        </Setter.Value>
      </Setter>
    </Style>
    <Style x:Key="Pick" TargetType="ComboBox">
      <Setter Property="Height" Value="28"/>
      <Setter Property="FontSize" Value="12"/>
      <Setter Property="Foreground" Value="#0F172A"/>
      <Setter Property="Cursor" Value="Hand"/>
      <Setter Property="FocusVisualStyle" Value="{x:Null}"/>
      <Setter Property="ItemContainerStyle" Value="{StaticResource PickItem}"/>
      <Setter Property="Template">
        <Setter.Value>
          <ControlTemplate TargetType="ComboBox">
            <Grid>
              <ToggleButton Focusable="False" ClickMode="Press"
                            IsChecked="{Binding IsDropDownOpen, Mode=TwoWay, RelativeSource={RelativeSource TemplatedParent}}">
                <ToggleButton.Template>
                  <ControlTemplate TargetType="ToggleButton">
                    <Border x:Name="B" Background="White" BorderBrush="#CBD5E1" BorderThickness="1" CornerRadius="6">
                      <TextBlock Text="▾" HorizontalAlignment="Right" VerticalAlignment="Center" Margin="0,0,9,0" Foreground="#64748B"/>
                    </Border>
                    <ControlTemplate.Triggers>
                      <Trigger Property="IsMouseOver" Value="True"><Setter TargetName="B" Property="BorderBrush" Value="#0E7490"/></Trigger>
                      <Trigger Property="IsChecked" Value="True"><Setter TargetName="B" Property="BorderBrush" Value="#0E7490"/></Trigger>
                    </ControlTemplate.Triggers>
                  </ControlTemplate>
                </ToggleButton.Template>
              </ToggleButton>
              <ContentPresenter IsHitTestVisible="False" Content="{TemplateBinding SelectionBoxItem}"
                                Margin="10,0,26,0" VerticalAlignment="Center" HorizontalAlignment="Left"/>
              <Popup IsOpen="{TemplateBinding IsDropDownOpen}" Placement="Bottom" AllowsTransparency="True" Focusable="False" PopupAnimation="Fade">
                <Border Background="White" BorderBrush="#CBD5E1" BorderThickness="1" CornerRadius="6" Padding="4" Margin="0,3,0,0"
                        MinWidth="{Binding ActualWidth, RelativeSource={RelativeSource TemplatedParent}}" MaxHeight="320">
                  <ScrollViewer VerticalScrollBarVisibility="Auto"><ItemsPresenter/></ScrollViewer>
                </Border>
              </Popup>
            </Grid>
          </ControlTemplate>
        </Setter.Value>
      </Setter>
    </Style>

    <!-- Řádek seznamu: aktuální řádek je podbarvený s proužkem vlevo -->
    <Style x:Key="Row" TargetType="ListBoxItem">
      <Setter Property="HorizontalContentAlignment" Value="Stretch"/>
      <Setter Property="FocusVisualStyle" Value="{x:Null}"/>
      <Setter Property="Template">
        <Setter.Value>
          <ControlTemplate TargetType="ListBoxItem">
            <Border x:Name="Bd" Background="Transparent" BorderThickness="4,0,0,0" BorderBrush="Transparent" Padding="8,5">
              <ContentPresenter/>
            </Border>
            <ControlTemplate.Triggers>
              <Trigger Property="IsMouseOver" Value="True"><Setter TargetName="Bd" Property="Background" Value="#F1F5F9"/></Trigger>
              <Trigger Property="IsSelected" Value="True">
                <Setter TargetName="Bd" Property="Background" Value="{StaticResource AccentBg}"/>
                <Setter TargetName="Bd" Property="BorderBrush" Value="{StaticResource Accent}"/>
              </Trigger>
            </ControlTemplate.Triggers>
          </ControlTemplate>
        </Setter.Value>
      </Setter>
    </Style>
  </Window.Resources>

  <DockPanel>
    <!-- Hlavička -->
    <Border DockPanel.Dock="Top" Background="White" BorderBrush="{StaticResource Line}" BorderThickness="0,0,0,1" Padding="20,12">
      <Grid>
        <Grid.ColumnDefinitions>
          <ColumnDefinition Width="Auto"/><ColumnDefinition Width="*"/><ColumnDefinition Width="Auto"/><ColumnDefinition Width="Auto"/>
        </Grid.ColumnDefinitions>
        <TextBlock Text="Kontrola Clearance" FontSize="22" FontWeight="SemiBold" Foreground="{StaticResource Accent}" VerticalAlignment="Center"/>
        <CheckBox x:Name="ChkCtrl" Grid.Column="2" Content="Vkládat Ctrl + kliknutím" IsChecked="True" VerticalAlignment="Center" Margin="0,0,28,0" Foreground="{StaticResource Muted}"/>
        <TextBlock x:Name="PosText" Grid.Column="3" Text="0 / 0" FontSize="22" FontWeight="SemiBold" VerticalAlignment="Center"/>
      </Grid>
    </Border>

    <Grid Margin="8">
      <Grid.ColumnDefinitions><ColumnDefinition Width="1*" MinWidth="280"/><ColumnDefinition Width="2.4*"/></Grid.ColumnDefinitions>

      <!-- Pravá část: vkládání, seznam a ovládání -->
      <Grid Grid.Column="1">
        <Grid.RowDefinitions>
          <RowDefinition Height="3*"/><RowDefinition Height="1.3*" MinHeight="150"/><RowDefinition Height="80"/><RowDefinition Height="Auto"/>
        </Grid.RowDefinitions>

        <!-- Vstup / seznam -->
        <Border Grid.Row="0" Style="{StaticResource Card}">
          <Grid>
            <Grid.RowDefinitions><RowDefinition Height="Auto"/><RowDefinition Height="*"/><RowDefinition Height="Auto"/></Grid.RowDefinitions>
            <TextBlock x:Name="RevisionText" Visibility="Collapsed" FontSize="16" FontWeight="SemiBold" Foreground="{StaticResource Accent}" Margin="2,0,0,8"/>
            <TextBox x:Name="InputBox" Grid.Row="1" Style="{StaticResource Field}" AcceptsReturn="True" AcceptsTab="True" TextWrapping="NoWrap"
                     FontFamily="Consolas" FontSize="13" VerticalScrollBarVisibility="Auto" HorizontalScrollBarVisibility="Auto"/>
            <ListBox x:Name="ItemsList" Grid.Row="1" Visibility="Collapsed" BorderThickness="1" BorderBrush="{StaticResource Line}" Background="#F8FAFC"
                     ItemContainerStyle="{StaticResource Row}" ScrollViewer.HorizontalScrollBarVisibility="Disabled"
                     VirtualizingStackPanel.IsVirtualizing="True">
              <ListBox.ItemTemplate>
                <DataTemplate>
                  <Grid>
                    <!-- záhlaví složky (kliknutím se rozbalí / sbalí) -->
                    <Border x:Name="HeaderRow" Visibility="Collapsed" Background="#E2E8F0" CornerRadius="8" Height="42" Padding="10,0" Margin="-8,6,0,2" Cursor="Hand">
                      <Grid VerticalAlignment="Center">
                        <Grid.ColumnDefinitions>
                          <ColumnDefinition Width="22"/><ColumnDefinition Width="*" MinWidth="70"/><ColumnDefinition Width="170"/><ColumnDefinition Width="110"/><ColumnDefinition Width="72"/>
                        </Grid.ColumnDefinitions>
                        <TextBlock Text="{Binding Arrow}" FontWeight="Bold" Foreground="#334155" VerticalAlignment="Center"/>
                        <TextBlock Grid.Column="1" Text="{Binding Title}" FontWeight="SemiBold" Foreground="#0F172A" TextTrimming="CharacterEllipsis" VerticalAlignment="Center"/>
                        <ComboBox Grid.Column="2" Tag="S" Style="{StaticResource Pick}" Margin="8,0,0,0" ToolTip="Stanice"
                                  ItemsSource="{Binding Stations}" SelectedItem="{Binding SelS, Mode=OneWay}"/>
                        <ComboBox Grid.Column="3" Tag="L" Style="{StaticResource Pick}" Margin="6,0,0,0" ToolTip="Délka"
                                  ItemsSource="{Binding Lengths}" SelectedItem="{Binding SelL, Mode=OneWay}"/>
                        <TextBlock Grid.Column="4" Text="{Binding Progress}" FontSize="12" Foreground="#475569" VerticalAlignment="Center" HorizontalAlignment="Right"/>
                      </Grid>
                    </Border>
                    <!-- řádek záznamu -->
                    <Grid x:Name="ItemRow" Margin="14,0,0,0">
                      <Grid.ColumnDefinitions>
                        <ColumnDefinition Width="44"/><ColumnDefinition Width="26"/><ColumnDefinition Width="45*"/><ColumnDefinition Width="55*"/>
                      </Grid.ColumnDefinitions>
                      <TextBlock Text="{Binding Num}" Foreground="#94A3B8" FontSize="12" VerticalAlignment="Center"/>
                      <TextBlock x:Name="Mark" Grid.Column="1" Text="" FontWeight="Bold" VerticalAlignment="Center"/>
                      <TextBlock Grid.Column="2" Text="{Binding Value}" TextTrimming="CharacterEllipsis" VerticalAlignment="Center" Margin="0,0,12,0"/>
                      <TextBlock x:Name="NoteTb" Grid.Column="3" Text="{Binding Note}" Foreground="{StaticResource Muted}" FontSize="12" TextTrimming="CharacterEllipsis" VerticalAlignment="Center"/>
                    </Grid>
                  </Grid>
                  <DataTemplate.Triggers>
                    <DataTrigger Binding="{Binding Kind}" Value="H">
                      <Setter TargetName="HeaderRow" Property="Visibility" Value="Visible"/>
                      <Setter TargetName="ItemRow" Property="Visibility" Value="Collapsed"/>
                    </DataTrigger>
                    <DataTrigger Binding="{Binding Status}" Value="keep"><Setter TargetName="Mark" Property="Text" Value="✓"/><Setter TargetName="Mark" Property="Foreground" Value="#16A34A"/></DataTrigger>
                    <DataTrigger Binding="{Binding Status}" Value="del"><Setter TargetName="Mark" Property="Text" Value="✗"/><Setter TargetName="Mark" Property="Foreground" Value="#E11D48"/></DataTrigger>
                    <DataTrigger Binding="{Binding Status}" Value="unsure"><Setter TargetName="Mark" Property="Text" Value="?"/><Setter TargetName="Mark" Property="Foreground" Value="#D97706"/></DataTrigger>
                    <DataTrigger Binding="{Binding Manual}" Value="True"><Setter TargetName="NoteTb" Property="FontStyle" Value="Italic"/></DataTrigger>
                  </DataTemplate.Triggers>
                </DataTemplate>
              </ListBox.ItemTemplate>
            </ListBox>
            <StackPanel Grid.Row="2" Orientation="Horizontal" Margin="0,12,0,0">
              <Button x:Name="BtnExcel" Style="{StaticResource BtnPrimary}" Content="Načíst z Excelu"/>
              <Button x:Name="BtnLoad" Style="{StaticResource Btn}" Content="Vytvořit seznam" ToolTip="Vytvoří seznam z řádků vložených do pole přes Ctrl+V"/>
              <Button x:Name="BtnClear" Style="{StaticResource Btn}" Content="Zrušit seznam"/>
              <TextBlock x:Name="CountText" Text="Řádků: 0" FontWeight="SemiBold" VerticalAlignment="Center" Margin="8,0,0,0"/>
            </StackPanel>
          </Grid>
        </Border>

        <!-- Aktuální údaj -->
        <Border Grid.Row="1" Style="{StaticResource Card}">
          <DockPanel>
            <Grid DockPanel.Dock="Top">
              <Grid.ColumnDefinitions><ColumnDefinition Width="*"/><ColumnDefinition Width="Auto"/><ColumnDefinition Width="Auto"/><ColumnDefinition Width="Auto"/></Grid.ColumnDefinitions>
              <TextBlock Text="AKTUÁLNÍ ÚDAJ  •  zkopírováno do schránky" FontSize="12" Foreground="{StaticResource Muted}" VerticalAlignment="Center"/>
              <Button x:Name="BtnMap" Grid.Column="1" Style="{StaticResource Btn}" Content="Zobrazit na mapě" Padding="14,6" Margin="0,0,6,0"/>
              <Button x:Name="BtnUp" Grid.Column="2" Style="{StaticResource BtnNav}" Content="▲" ToolTip="Předchozí řádek"/>
              <Button x:Name="BtnDown" Grid.Column="3" Style="{StaticResource BtnNav}" Content="▼" ToolTip="Další řádek"/>
            </Grid>
            <TextBlock x:Name="CurrentNote" DockPanel.Dock="Bottom" TextAlignment="Center" Foreground="{StaticResource Muted}" TextTrimming="CharacterEllipsis" Margin="0,6,0,0"/>
            <TextBlock x:Name="CurrentText" FontSize="26" FontWeight="SemiBold" TextAlignment="Center" TextWrapping="Wrap"
                       TextTrimming="CharacterEllipsis" VerticalAlignment="Center" HorizontalAlignment="Center"/>
          </DockPanel>
        </Border>

        <!-- Ponechat / Vymazat -->
        <Grid Grid.Row="2">
          <Grid.ColumnDefinitions><ColumnDefinition/><ColumnDefinition/></Grid.ColumnDefinitions>
          <Button x:Name="BtnYes" Style="{StaticResource BtnBig}" Background="#16A34A" Content="Ponechat"/>
          <Button x:Name="BtnNo" Grid.Column="1" Style="{StaticResource BtnBig}" Background="#E11D48" Content="Vymazat"/>
        </Grid>

        <!-- Vedlejší akce -->
        <StackPanel Grid.Row="3" Orientation="Horizontal" HorizontalAlignment="Center" Margin="0,6,0,4">
          <Button x:Name="BtnUnsure" Style="{StaticResource Btn}" Background="White" Foreground="#D97706" Content="?   Vrátit se později" Margin="0"/>
        </StackPanel>
      </Grid>

      <!-- Levá část: poznámky -->
      <Border Grid.Column="0" Style="{StaticResource Card}">
        <Grid>
          <Grid.RowDefinitions>
            <RowDefinition Height="Auto"/><RowDefinition Height="Auto"/><RowDefinition Height="Auto"/><RowDefinition Height="*"/><RowDefinition Height="Auto"/>
          </Grid.RowDefinitions>
          <TextBlock Text="Název souboru (.txt)" FontWeight="SemiBold"/>
          <TextBox x:Name="FileNameBox" Grid.Row="1" Style="{StaticResource Field}" Margin="0,6,0,14" FontSize="15"/>
          <TextBlock x:Name="NotesCaption" Grid.Row="2" Text="Poznámky  •  zatím neuloženo" FontWeight="SemiBold" TextTrimming="CharacterEllipsis" Margin="0,0,0,6"/>
          <TextBox x:Name="NotesBox" Grid.Row="3" Style="{StaticResource Field}" AcceptsReturn="True" TextWrapping="Wrap"
                   FontFamily="Consolas" FontSize="13" VerticalScrollBarVisibility="Auto"/>
          <StackPanel Grid.Row="4" Orientation="Horizontal" Margin="0,12,0,0">
            <Button x:Name="BtnOpen" Style="{StaticResource Btn}" Content="Otevřít soubor…"/>
            <Button x:Name="BtnSave" Style="{StaticResource BtnPrimary}" Content="Uložit"/>
            <Button x:Name="BtnSaveAs" Style="{StaticResource Btn}" Content="Uložit jako…"/>
          </StackPanel>
        </Grid>
      </Border>
    </Grid>
  </DockPanel>
</Window>
'@

$win = [System.Windows.Markup.XamlReader]::Load((New-Object System.Xml.XmlNodeReader $xaml))
foreach ($n in @('ChkCtrl','PosText','InputBox','ItemsList','BtnLoad','BtnClear','CountText',
                 'BtnUp','BtnDown','CurrentNote','CurrentText','BtnYes','BtnNo','BtnUnsure','BtnMap','BtnExcel','RevisionText',
                 'FileNameBox','NotesCaption','NotesBox','BtnOpen','BtnSave','BtnSaveAs')) {
    Set-Variable -Name $n -Value $win.FindName($n) -Scope Script
}
$win.Title = "Kontrola Clearance  –  verze $($script:AppVersion)"
try { $win.Icon = Get-KcImageSource 256 } catch { }

function Brush([string]$hex) { New-Object System.Windows.Media.SolidColorBrush ([System.Windows.Media.ColorConverter]::ConvertFromString($hex)) }
$bMuted = Brush '#64748B'; $bText = Brush '#0F172A'; $bDone = Brush '#15803D'; $bWarn = Brush '#B91C1C'

# ---------- Logika ----------
function Add-NoteLine([string]$line) {
    $t = $NotesBox.Text
    if ($t.Length -gt 0 -and -not $t.EndsWith("`n")) { $t += "`r`n" }
    $NotesBox.Text = $t + $line + "`r`n"
    $NotesBox.CaretIndex = $NotesBox.Text.Length
    $NotesBox.ScrollToEnd()
}

# Odstraní z poznámek poslední řádek přesně rovný $line (při změně Vymazat -> Ponechat / ?)
function Remove-NoteLine([string]$line) {
    $lines = [System.Collections.Generic.List[string]]($NotesBox.Text -split "\r?\n")
    for ($i = $lines.Count - 1; $i -ge 0; $i--) {
        if ($lines[$i].Trim() -eq $line) {
            $lines.RemoveAt($i)
            $NotesBox.Text = $lines -join "`r`n"
            return $true
        }
    }
    return $false
}

# Filtr složky (stanice / délka); 'Vše' = bez omezení
function Get-GroupFilter([string]$cat) {
    if (-not $script:GroupFilter.ContainsKey($cat)) { $script:GroupFilter[$cat] = @{ S = $script:All; L = $script:All } }
    $script:GroupFilter[$cat]
}
# Je záznam vidět (prochází filtrem své složky)?
function Test-Visible($it) {
    $f = Get-GroupFilter $it.Category
    ($f.S -eq $script:All -or $it.Station -eq $f.S) -and ($f.L -eq $script:All -or $it.Length -eq $f.L)
}
# Indexy viditelných záznamů v pořadí seznamu
function Get-VisibleIndexes {
    $v = New-Object System.Collections.Generic.List[int]
    for ($i = 0; $i -lt $script:Items.Count; $i++) { if (Test-Visible $script:Items[$i]) { $v.Add($i) } }
    return ,$v
}
function Get-Choices($items, [string]$prop) {
    # čísla (např. délka) se řadí podle hodnoty, text abecedně
    $vals = @($items | ForEach-Object { [string]$_.$prop } | Where-Object { $_ -ne '' } | Select-Object -Unique |
              Sort-Object { $d = 0.0; if ([double]::TryParse(($_ -replace ',', '.'), [Globalization.NumberStyles]::Float, [Globalization.CultureInfo]::InvariantCulture, [ref]$d)) { $d } else { [double]::MaxValue } }, { $_ })
    return ,(@($script:All) + $vals)
}

# Seznam místo vstupního pole (bez seznamu se zobrazí vstupní pole).
# Zobrazují se složky (záhlaví s výběrem stanice a délky) a pod nimi jejich viditelné záznamy.
$script:RowOfItem = @{}   # index záznamu -> index řádku v zobrazeném seznamu
function Update-ListBox {
    if ($script:Items.Count -eq 0) {
        $ItemsList.ItemsSource = $null
        $ItemsList.Visibility = 'Collapsed'
        $RevisionText.Visibility = 'Collapsed'
        $InputBox.Visibility = 'Visible'
        return
    }
    $script:Syncing = $true
    try {
        # složka aktuálního záznamu musí být rozbalená, aby byl vidět
        if (-not $script:Done -and $script:Index -lt $script:Items.Count) { $script:Collapsed.Remove($script:Items[$script:Index].Category) }
        $rows = New-Object System.Collections.Generic.List[object]
        $script:RowOfItem = @{}
        $i = 0
        while ($i -lt $script:Items.Count) {
            $cat = $script:Items[$i].Category
            $j = $i
            while ($j -lt $script:Items.Count -and $script:Items[$j].Category -eq $cat) { $j++ }
            $groupItems = @($script:Items[$i..($j - 1)])
            $f = Get-GroupFilter $cat
            $visible = @(); $done = 0
            for ($k = $i; $k -lt $j; $k++) {
                if (Test-Visible $script:Items[$k]) {
                    $visible += $k
                    if ($script:Items[$k].Status -eq 'keep' -or $script:Items[$k].Status -eq 'del') { $done++ }
                }
            }
            $collapsed = $script:Collapsed.ContainsKey($cat)
            $title = if ($cat -eq '') { '(bez složky)' } else { $cat }
            $rows.Add([pscustomobject]@{ Kind = 'H'; Category = $cat; Title = $title; Arrow = $(if ($collapsed) { '▸' } else { '▾' });
                                         Progress = "$done / $($visible.Count)"; Num = ''; Value = ''; Note = ''; Status = ''; Manual = $false
                                         Stations = (Get-Choices $groupItems 'Station'); Lengths = (Get-Choices $groupItems 'Length'); SelS = $f.S; SelL = $f.L })
            if (-not $collapsed) {
                foreach ($k in $visible) {
                    $it = $script:Items[$k]
                    $it.Num = $k + 1
                    $script:RowOfItem[$k] = $rows.Count
                    $rows.Add([pscustomobject]@{ Kind = 'I'; Category = $cat; Title = ''; Arrow = ''; Progress = '';
                                                 Num = $k + 1; Value = $it.Value; Note = $it.Note; Status = $it.Status; Manual = $it.Manual; Index = $k })
                }
            }
            $i = $j
        }
        $ItemsList.ItemsSource = $rows
        if (-not $script:Done -and $script:RowOfItem.ContainsKey($script:Index)) {
            $r = $script:RowOfItem[$script:Index]
            $ItemsList.SelectedIndex = $r
            $ItemsList.ScrollIntoView($rows[$r])
        } else {
            $ItemsList.SelectedIndex = -1
        }
        $RevisionText.Text = if ($script:Revision -ne '') { "Revize $($script:Revision)" } else { 'Revize –' }
        $RevisionText.Visibility = 'Visible'
        $InputBox.Visibility = 'Collapsed'
        $ItemsList.Visibility = 'Visible'
    } finally { $script:Syncing = $false }
}

function Update-View {
    $vis = Get-VisibleIndexes
    $count = $script:Items.Count
    $on = ($count -gt 0 -and -not $script:Done -and (Test-Visible $script:Items[$script:Index]))
    $BtnYes.IsEnabled = $on; $BtnNo.IsEnabled = $on; $BtnUnsure.IsEnabled = $on
    $BtnMap.IsEnabled = ($on -and $script:Items[$script:Index].MapUrl -ne '')
    $BtnUp.IsEnabled = ($vis.Count -gt 0); $BtnDown.IsEnabled = ($vis.Count -gt 0)
    $CurrentNote.Text = ''
    if ($count -eq 0) {
        $PosText.Text = '0 / 0'
        $CurrentText.Text = 'Načtěte data z Excelu'
        $CurrentText.Foreground = $bMuted
    } elseif ($script:Done) {
        $PosText.Text = "$($vis.Count) / $($vis.Count)"
        $CurrentText.Text = '✓ Hotovo – všechny zobrazené řádky jsou rozhodnuté'
        $CurrentText.Foreground = $bDone
        $CurrentNote.Text = 'Šipkami ▲ ▼ nebo kliknutím do seznamu se můžete k libovolnému řádku vrátit.'
    } else {
        $PosText.Text = "$($vis.IndexOf($script:Index) + 1) / $($vis.Count)"
        $item = $script:Items[$script:Index]
        $CurrentText.Text = $item.Value
        $CurrentText.Foreground = $bText
        $info = if ($item.Note -ne '') { "Při Vymazat se zapíše: $($item.Note)" } else { 'Při Vymazat se nic nezapíše (poznámka je prázdná)' }
        switch ($item.Status) {
            'keep'   { $info = '✓ Ponecháno   •   ' + $info }
            'del'    { $info = '✗ Vymazáno (zapsáno v poznámkách)   •   ' + $info }
            'unsure' { $info = '? Vrátit se později   •   ' + $info }
        }
        if ($item.Manual) { $info += '   •   vloženo ručně' }
        $CurrentNote.Text = $info
    }
    Update-ListBox
}

# Zobrazí aktuální řádek a zkopíruje jeho první hodnotu do schránky
function Show-Current {
    Update-View
    if (-not $script:Done -and $script:Index -lt $script:Items.Count -and (Test-Visible $script:Items[$script:Index])) {
        if (-not (Set-ClipboardText $script:Items[$script:Index].Value)) {
            Show-Warn 'Hodnotu se nepodařilo zkopírovat do schránky (schránka je možná obsazená jinou aplikací).'
        }
    }
}

# Přejde na další nerozhodnutý řádek (hledá od aktuálního dál, pak od začátku).
# Až nezbývá žádný nerozhodnutý, přijdou na řadu řádky označené „?“; když nejsou ani ty, dokončeno.
function Move-NextUndecided {
    $count = $script:Items.Count
    foreach ($wanted in @('', 'unsure')) {
        for ($k = 1; $k -le $count; $k++) {
            $i = ($script:Index + $k) % $count
            if ($script:Items[$i].Status -eq $wanted -and (Test-Visible $script:Items[$i])) {
                $script:Index = $i
                $script:Done = $false
                Show-Current
                return
            }
        }
    }
    $script:Done = $true
    Show-Current
}

# Ruční posun o $delta viditelných řádků
function Move-By([int]$delta) {
    $vis = Get-VisibleIndexes
    if ($vis.Count -eq 0) { return }
    if ($script:Done) { $script:Done = $false; $delta = 0 }
    $pos = $vis.IndexOf($script:Index)
    if ($pos -lt 0) {
        # aktuální řádek je schovaný filtrem: nejbližší viditelný za ním
        $pos = 0
        for ($p = 0; $p -lt $vis.Count; $p++) { if ($vis[$p] -gt $script:Index) { $pos = $p; break } }
        $delta = 0
    }
    $pos = [Math]::Max(0, [Math]::Min($vis.Count - 1, $pos + $delta))
    $script:Index = $vis[$pos]
    Show-Current
}

function Set-Decision([string]$status) {
    if ($script:Done -or $script:Index -ge $script:Items.Count) { return }
    $item = $script:Items[$script:Index]
    $note = $item.Note.Trim()
    if ($status -eq 'del' -and $item.Status -ne 'del') {
        if ($note -eq '') { Show-Warn 'Poznámka pro tento řádek je prázdná. Nic se nezapíše, pokračuje se dalším řádkem.' }
        else { Add-NoteLine $note }
    }
    if ($status -ne 'del' -and $item.Status -eq 'del' -and $note -ne '') {
        if (-not (Remove-NoteLine $note)) {
            Show-Warn "Řádek byl dříve vymazán, ale jeho poznámku „$note“ se v poznámkách nepodařilo najít. Zkontrolujte poznámky ručně."
        }
    }
    $item.Status = $status
    Move-NextUndecided
}

# ---------- Poznámky a soubor ----------
function Update-NotesCaption {
    $NotesCaption.Text = if ($script:NotesFile) { "Poznámky  •  $([System.IO.Path]::GetDirectoryName($script:NotesFile))" } else { 'Poznámky  •  zatím neuloženo' }
}

# Název z pole „Název souboru“ (bez .txt); prázdný = $null, neplatný = $false
function Get-WantedFileName {
    $name = $FileNameBox.Text.Trim()
    if ($name.ToLower().EndsWith('.txt')) { $name = $name.Substring(0, $name.Length - 4).Trim() }
    if ($name -eq '') { return $null }
    if ($name.IndexOfAny([System.IO.Path]::GetInvalidFileNameChars()) -ge 0) {
        Show-Warn 'Název souboru obsahuje nepovolené znaky (např. \ / : * ? " < > |).'
        return $false
    }
    return $name + '.txt'
}

# Uloží poznámky do $path. Pokud jde o přejmenování otevřeného souboru, starý soubor se po úspěšném uložení odstraní.
function Save-Notes([string]$path, [bool]$rename = $false) {
    try {
        $old = $script:NotesFile
        $renaming = ($rename -and $old -and ($old -ne $path) -and ([System.IO.Path]::GetDirectoryName($old) -eq [System.IO.Path]::GetDirectoryName($path)))
        if ($renaming -and [System.IO.File]::Exists($path) -and ($old.ToLower() -ne $path.ToLower())) {
            if (-not (Ask-YesNo "Soubor $([System.IO.Path]::GetFileName($path)) už existuje. Přepsat ho?" 'Uložit' 'Warning')) { return }
        }
        $enc = New-Object System.Text.UTF8Encoding($true)
        [System.IO.File]::WriteAllText($path, $NotesBox.Text, $enc)
        if ($renaming -and ($old.ToLower() -ne $path.ToLower()) -and [System.IO.File]::Exists($old)) {
            [System.IO.File]::Delete($old)
        }
        $script:NotesFile = $path
        $FileNameBox.Text = [System.IO.Path]::GetFileNameWithoutExtension($path)
        Update-NotesCaption
        Show-Info "Poznámky byly uloženy do:`n$path"
    } catch {
        Show-Error 'Poznámky se nepodařilo uložit. Zkontrolujte, zda je soubor dostupný a máte do složky právo zápisu.'
    }
}

function Save-NotesAs {
    $wanted = Get-WantedFileName
    if ($wanted -eq $false) { return }
    $sfd = New-Object Microsoft.Win32.SaveFileDialog
    $sfd.Filter = 'Textový soubor (*.txt)|*.txt'
    $sfd.DefaultExt = 'txt'
    $sfd.AddExtension = $true
    $sfd.FileName = if ($wanted) { $wanted } elseif ($script:NotesFile) { [System.IO.Path]::GetFileName($script:NotesFile) } else { 'poznamky.txt' }
    if ($script:NotesFile) { $sfd.InitialDirectory = [System.IO.Path]::GetDirectoryName($script:NotesFile) }
    if ($sfd.ShowDialog($win) -eq $true) { Save-Notes $sfd.FileName }
}

# Počet řádků ve vstupním poli (koncové prázdné řádky se nepočítají)
function Update-InputCount {
    $lines = Get-ColumnLines $InputBox.Text
    $rows = @($lines | Where-Object { -not [string]::IsNullOrWhiteSpace($_.Replace("`t", '')) })
    $text = "Řádků: $($rows.Count)"
    $CountText.Foreground = $bText
    if ($rows.Count -gt 0) {
        if ($rows[0].Split("`t").Count -ne 7) {
            $text = $text + '   •   vložte řádky z Excelu přes Ctrl+V'
            $CountText.Foreground = $bWarn
        } else {
            $cats = @($rows | ForEach-Object { $_.Split("`t")[1].Trim() } | Select-Object -Unique).Count
            $text = $text + "   •   složek: $cats"
        }
    }
    $CountText.Text = $text
}

# ---------- Obsluha ovládacích prvků ----------
$InputBox.Add_TextChanged({
    if ($InputBox.Text.Length -eq 0) { $script:PastedCols = 0 }
    Invoke-Safe { Update-InputCount } 'Počet řádků se nepodařilo spočítat.'
})

# Vložení (Ctrl+V, Shift+Insert, kontextová nabídka): jen čistý text a z každého řádku jen 1. a poslední sloupec
$InputBox.AddHandler([System.Windows.Input.CommandManager]::PreviewExecutedEvent,
    [System.Windows.Input.ExecutedRoutedEventHandler]{
        param($s, $e)
        if ($e.Command -ne [System.Windows.Input.ApplicationCommands]::Paste) { return }
        $e.Handled = $true
        Invoke-Safe {
            if (-not [System.Windows.Clipboard]::ContainsText()) { return }
            $r = Reduce-PastedColumns ([System.Windows.Clipboard]::GetText())
            if ($r.Error) { Show-Error $r.Error; return }
            $script:PastedCols = $r.Columns
            $InputBox.SelectedText = $r.Text
            $InputBox.CaretIndex = $InputBox.SelectionStart + $InputBox.SelectionLength
            $InputBox.SelectionLength = 0
        } 'Vložení ze schránky se nezdařilo.'
    })

# Založí nový seznam z načtených záznamů
function Start-List($parsed) {
    if ($parsed.Items.Count -eq 0) { Show-Error 'Nebyl nalezen žádný záznam.'; return }
    $script:Items = Group-Items $parsed.Items
    $script:Revision = [string]$parsed.Revision
    $script:Collapsed = @{}
    $script:GroupFilter = @{}
    $script:Index = 0
    $script:Done = $false
    Show-Current
}

$BtnExcel.Add_Click({ Invoke-Safe {
    if ($script:Items.Count -gt 0 -and -not (Ask-YesNo ('Seznam už je vytvořený. Nahradit ho daty z Excelu?' + "`n`n" + 'Poznámky zůstanou beze změny.') 'Načíst z Excelu')) { return }
    $win.Cursor = [System.Windows.Input.Cursors]::Wait
    try { $parsed = Read-FromExcel } finally { $win.Cursor = $null }
    if (-not $parsed.Ok) { Show-Error $parsed.Error; return }   # stávající seznam ani poznámky se nemění
    Start-List $parsed
} 'Data z Excelu se nepodařilo načíst. Zkontrolujte, že je soubor otevřený a data jsou načtená.' })

$BtnLoad.Add_Click({ Invoke-Safe {
    if ($script:Items.Count -gt 0) { Show-Warn 'Seznam už je vytvořený. Pro nové vložení ho nejdřív zrušte (Zrušit seznam).'; return }
    $parsed = ConvertFrom-Rows $InputBox.Text
    if (-not $parsed.Ok) { Show-Error $parsed.Error; return }
    Start-List $parsed
} 'Seznam se nepodařilo načíst. Zkuste znovu zkopírovat data z Excelu.' })

$BtnClear.Add_Click({ Invoke-Safe {
    if (-not (Ask-YesNo ('Opravdu zrušit načtený seznam a vymazat vstupní pole?' + "`n`n" + 'Poznámky zůstanou beze změny.') 'Zrušit seznam')) { return }
    $script:Items = New-Object System.Collections.Generic.List[object]
    $script:Index = 0
    $script:Done = $false
    $script:Revision = ''
    Update-View
    $InputBox.Clear()
} 'Seznam se nepodařilo vymazat.' })

# Kliknutí v seznamu: na záhlaví složky ji rozbalí / sbalí, na záznam ho udělá aktuálním
$ItemsList.Add_SelectionChanged({ param($s, $e)
    if ($script:Syncing) { return }
    if (-not [object]::ReferenceEquals($e.OriginalSource, $ItemsList)) { return }   # změna ve výběru stanice / délky
    Invoke-Safe {
        $row = $ItemsList.SelectedItem
        if ($null -eq $row) { return }
        if ($row.Kind -eq 'H') {
            if ($script:Collapsed.ContainsKey($row.Category)) { $script:Collapsed.Remove($row.Category) }
            elseif (-not $script:Done -and $script:Items[$script:Index].Category -eq $row.Category) {
                Show-Info 'Složku s aktuálním řádkem nejde sbalit. Nejdřív přejděte na řádek v jiné složce.'
            } else { $script:Collapsed[$row.Category] = $true }
            Update-ListBox
            return
        }
        if ($row.Index -ne $script:Index -or $script:Done) {
            $script:Index = $row.Index
            $script:Done = $false
            Show-Current
        }
    } 'Přechod na řádek se nezdařil.'
})

# Výběr stanice / délky v záhlaví složky: ostatní záznamy složky se schovají
$ItemsList.AddHandler([System.Windows.Controls.Primitives.Selector]::SelectionChangedEvent,
    [System.Windows.Controls.SelectionChangedEventHandler]{
        param($s, $e)
        $cb = $e.OriginalSource
        if ($cb -isnot [System.Windows.Controls.ComboBox]) { return }
        $e.Handled = $true
        if ($script:Syncing -or $null -eq $cb.SelectedItem) { return }
        $row = $cb.DataContext
        if ($null -eq $row -or $row.Kind -ne 'H') { return }
        $f = Get-GroupFilter $row.Category
        $val = [string]$cb.SelectedItem
        if ($f[$cb.Tag] -eq $val) { return }   # stejná hodnota (např. při vykreslení)
        $f[$cb.Tag] = $val
        # seznam se přestaví až po dokončení této události
        [void]$win.Dispatcher.BeginInvoke([Action]{ Invoke-Safe {
            # aktuální řádek se schoval (nebo bylo hotovo): přejít na další viditelný nerozhodnutý
            if ($script:Done -or -not (Test-Visible $script:Items[$script:Index])) { Move-NextUndecided }
            else { Update-View }
        } 'Filtr se nepodařilo použít.' })
    })

# Šipky na klávesnici v seznamu: o záznam nahoru / dolů (záhlaví složek se přeskakují)
$ItemsList.Add_PreviewKeyDown({ param($s, $e)
    if ($e.OriginalSource -is [System.Windows.Controls.ComboBox] -or $e.OriginalSource -is [System.Windows.Controls.ComboBoxItem]) { return }
    $d = switch ($e.Key) { 'Up' { -1 } 'Down' { 1 } 'PageUp' { -10 } 'PageDown' { 10 } default { 0 } }
    if ($d -ne 0) {
        $e.Handled = $true
        Invoke-Safe { Move-By $d } 'Přechod na řádek se nezdařil.'
    }
})

# Odkaz z textu buňky: pokud je v textu něco navíc, vezme se jen část „schéma://…“
function Get-LinkFromText([string]$t) {
    $t = $t.Trim().Trim('"')
    $m = [regex]::Match($t, '[A-Za-z][A-Za-z0-9+.\-]*://\S+')
    if ($m.Success) { return $m.Value }
    return $t
}

# Otevře odkaz v aplikaci zaregistrované ve Windows (dmwmap:// apod.); zkusí víc způsobů
function Open-Link([string]$url) {
    try {
        $psi = New-Object System.Diagnostics.ProcessStartInfo $url
        $psi.UseShellExecute = $true
        [void][System.Diagnostics.Process]::Start($psi)
        return $true
    } catch { }
    try { Start-Process -FilePath 'explorer.exe' -ArgumentList ('"' + $url + '"'); return $true } catch { }
    try { Start-Process -FilePath 'rundll32.exe' -ArgumentList 'url.dll,FileProtocolHandler', $url; return $true } catch { }
    return $false
}

# Zobrazit na mapě: otevře odkaz (dmwmap://…) aktuálního záznamu
$BtnMap.Add_Click({ Invoke-Safe {
    if ($script:Done -or $script:Index -ge $script:Items.Count) { return }
    $url = Get-LinkFromText $script:Items[$script:Index].MapUrl
    if ($url -eq '') { Show-Warn "Aktuální záznam nemá odkaz na mapu ($($script:ColMap). sloupec je prázdný)."; return }
    if ($url -notmatch '^[A-Za-z][A-Za-z0-9+.\-]*://') {
        Show-Warn "V $($script:ColMap). sloupci není odkaz (nezačíná dmwmap://):`n$url"
        return
    }
    if (-not (Open-Link $url)) {
        Show-Error "Odkaz se nepodařilo otevřít:`n$url`n`nZkontrolujte, že je v počítači aplikace pro odkazy dmwmap://."
    }
} 'Odkaz se nepodařilo otevřít.' })

$BtnYes.Add_Click({ Invoke-Safe { Set-Decision 'keep' } 'Přechod na další řádek se nezdařil.' })
$BtnNo.Add_Click({ Invoke-Safe { Set-Decision 'del' } 'Zápis poznámky se nezdařil.' })
$BtnUnsure.Add_Click({ Invoke-Safe { Set-Decision 'unsure' } 'Označení se nezdařilo.' })
$BtnUp.Add_Click({ Invoke-Safe { Move-By -1 } 'Přechod na řádek se nezdařil.' })
$BtnDown.Add_Click({ Invoke-Safe { Move-By 1 } 'Přechod na řádek se nezdařil.' })

$BtnOpen.Add_Click({ Invoke-Safe {
    if ($NotesBox.Text.Trim() -ne '') {
        if (-not (Ask-YesNo ('Otevřením souboru se nahradí aktuální obsah poznámek.' + "`n`n" + 'Pokud ho chcete zachovat, nejdřív ho uložte. Pokračovat?') 'Otevřít soubor')) { return }
    }
    $ofd = New-Object Microsoft.Win32.OpenFileDialog
    $ofd.Filter = 'Textový soubor (*.txt)|*.txt|Všechny soubory (*.*)|*.*'
    if ($ofd.ShowDialog($win) -eq $true) {
        try {
            $content = [System.IO.File]::ReadAllText($ofd.FileName, [System.Text.Encoding]::UTF8)
            $content = $content -replace "\r?\n", "`r`n"
            if ($content.Length -gt 0 -and -not $content.EndsWith("`n")) { $content += "`r`n" }
            $NotesBox.Text = $content
            $NotesBox.CaretIndex = $NotesBox.Text.Length
            $NotesBox.ScrollToEnd()
            $script:NotesFile = $ofd.FileName
            $FileNameBox.Text = [System.IO.Path]::GetFileNameWithoutExtension($ofd.FileName)
            Update-NotesCaption
        } catch {
            Show-Error 'Soubor se nepodařilo otevřít. Zkontrolujte, zda existuje a není otevřený jinou aplikací.'
        }
    }
} 'Soubor se nepodařilo otevřít.' })

$BtnSave.Add_Click({ Invoke-Safe {
    # Uložit: do otevřeného souboru; když se změnil název, soubor se přejmenuje (ve stejné složce)
    $wanted = Get-WantedFileName
    if ($wanted -eq $false) { return }
    if (-not $script:NotesFile) { Save-NotesAs; return }
    $dir = [System.IO.Path]::GetDirectoryName($script:NotesFile)
    $target = if ($wanted) { [System.IO.Path]::Combine($dir, $wanted) } else { $script:NotesFile }
    Save-Notes $target $true
} 'Ukládání se nezdařilo.' })

$BtnSaveAs.Add_Click({ Invoke-Safe { Save-NotesAs } 'Ukládání se nezdařilo.' })

# ---------- Vkládání Ctrl + kliknutím do jiné aplikace ----------
# Globální sledování myši (funkce Windows, bez instalace). Při Ctrl + levém kliknutí
# v jiném okně než této aplikace se do kliknutého pole vloží aktuální údaj (nahradí jeho obsah).
$script:MiddleHookOk = $false
try {
    $script:MiddleHookOk = [KlirencCtrlClick]::Start()
} catch {
    $script:MiddleHookOk = $false
}
if (-not $script:MiddleHookOk) {
    $ChkCtrl.IsChecked = $false
    $ChkCtrl.IsEnabled = $false
    $ChkCtrl.Content = 'Vkládání Ctrl + kliknutím není na tomto počítači dostupné'
}
$ChkCtrl.Add_Click({
    if ($script:MiddleHookOk) { [KlirencCtrlClick]::Enabled = [bool]$ChkCtrl.IsChecked }
})

# Po Ctrl + kliknutí: do schránky dát aktuální údaj a vložit ho do pole, kam se kliklo
$ctrlTimer = New-Object System.Windows.Threading.DispatcherTimer
$ctrlTimer.Interval = [TimeSpan]::FromMilliseconds(40)
$ctrlTimer.Add_Tick({
    try {
        if (-not $script:MiddleHookOk -or -not [KlirencCtrlClick]::Pending) { return }
        [KlirencCtrlClick]::Pending = $false
        if ($script:Done -or $script:Index -ge $script:Items.Count) { return }
        if (Set-ClipboardText $script:Items[$script:Index].Value) {
            Start-Sleep -Milliseconds 60
            [KlirencCtrlClick]::ClickSelectAllPaste()
        }
    } catch { }
})
if ($script:MiddleHookOk) { $ctrlTimer.Start() }

$win.Add_Loaded({
    # zástupce se aktualizuje až po zobrazení okna, aby nezdržoval start
    [void]$win.Dispatcher.BeginInvoke([Action]{ try { Update-Shortcuts } catch { } }, [System.Windows.Threading.DispatcherPriority]::ApplicationIdle)
    Update-NotesCaption
    Update-View
    [void]$InputBox.Focus()
})

try {
    [void]$win.ShowDialog()
} catch {
    Show-Error 'Aplikace narazila na neočekávaný problém a bude ukončena.'
} finally {
    try { $ctrlTimer.Stop(); if ($script:MiddleHookOk) { [KlirencCtrlClick]::Stop() } } catch { }
}
