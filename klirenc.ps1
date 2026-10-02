# Kontrola Clearance - pracovní pomocník pro ruční procházení řádků z Excelu
# Spuštění: powershell.exe -NoProfile -STA -ExecutionPolicy Bypass -File .\klirenc.ps1

# -VytvoritZastupce: jen vytvoří ikonu kc.ico a zástupce „Kontrola Clearance“ (ve složce a na ploše) a skončí
param([switch]$VytvoritZastupce)

Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing
# Ostré vykreslení při zvětšení zobrazení ve Windows (125 %, 150 % …): bez tohoto Windows aplikaci
# vykreslí v malém a roztáhnou ji, takže vypadá rozmazaně a „kostičkovaně“. Musí proběhnout před vytvořením oken.
try {
    if (-not ('KcDpi' -as [type])) {
        Add-Type -TypeDefinition @'
using System;
using System.Runtime.InteropServices;
public static class KcDpi {
    [DllImport("user32.dll")] public static extern bool SetProcessDPIAware();
    [DllImport("shcore.dll")] public static extern int SetProcessDpiAwareness(int value);
}
'@
    }
    try { [void][KcDpi]::SetProcessDpiAwareness(2) } catch { [void][KcDpi]::SetProcessDPIAware() }
} catch { }
[System.Windows.Forms.Application]::EnableVisualStyles()

# Měřítko obrazovky (1 = 100 %, 1.25 = 125 % …) pro ruční rozměry v pixelech
$script:S = 1.0
try { $gd = [System.Drawing.Graphics]::FromHwnd([IntPtr]::Zero); $script:S = $gd.DpiX / 96.0; $gd.Dispose() } catch { }
function Px([double]$n) { [int][Math]::Round($n * $script:S) }

# ---------- Stav ----------
$script:Items = New-Object System.Collections.Generic.List[object]
$script:Index = 0
$script:Highlighting = $false
$script:Syncing = $false       # programové nastavení výběru v seznamu (neřešit jako klik)
$script:Done = $false          # všechny řádky rozhodnuty a zobrazuje se dokončení
$script:PastedCols = 0      # kolik sloupců mělo poslední vložení (před ořezáním)
$script:NotesFile = $null   # soubor, do kterého se poznámky ukládají (po Otevřít / Uložit jako)

# Status: '' = nerozhodnuto, 'keep' = Ponechat, 'del' = Vymazat (poznámka zapsána)
# Manual = záznam přidaný přes Vložit mezi
function New-Item2([string]$value, [string]$note, [bool]$manual = $false) {
    [pscustomobject]@{ Value = $value; Note = $note; Status = ''; Manual = $manual }
}

function Show-Info([string]$text)  { [void][System.Windows.Forms.MessageBox]::Show($text, 'Kontrola Clearance', 'OK', 'Information') }
function Show-Warn([string]$text)  { [void][System.Windows.Forms.MessageBox]::Show($text, 'Kontrola Clearance', 'OK', 'Warning') }
function Show-Error([string]$text) { [void][System.Windows.Forms.MessageBox]::Show($text, 'Kontrola Clearance', 'OK', 'Error') }

# Bezpečné spuštění obsluhy události - uživateli se nezobrazí technická chyba
function Invoke-Safe([scriptblock]$action, [string]$message) {
    try { & $action } catch { Show-Error $message }
}

# ---------- Parsování vstupu ----------
# Rozdělí vložený sloupec na řádky; koncové prázdné řádky (Excel přidává konec řádku) se odříznou.
function Get-ColumnLines([string]$text) {
    if ([string]::IsNullOrEmpty($text)) { return ,@() }
    $lines = [System.Collections.Generic.List[string]]($text -split "\r\n|\n|\r")
    while ($lines.Count -gt 0 -and [string]::IsNullOrWhiteSpace($lines[$lines.Count - 1])) { $lines.RemoveAt($lines.Count - 1) }
    return ,$lines.ToArray()
}

# Ze vloženého textu ponechá v každém řádku jen první a poslední sloupec (prostřední zahodí hned při vložení).
# Vrací @{ Text; Columns } nebo @{ Error } při nestejném počtu sloupců.
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
        if ($parts.Count -ge 2) { $out.Add($parts[0] + "`t" + $parts[$parts.Count - 1]) } else { $out.Add($line) }
    }
    while ($out.Count -gt 0 -and $out[$out.Count - 1] -eq '') { $out.RemoveAt($out.Count - 1) }
    return @{ Text = (($out -join "`r`n") + "`r`n"); Columns = $columns }
}

# Z vložených řádků (libovolný počet sloupců oddělených tabulátorem) použije jen první a poslední sloupec.
# Vrací @{ Ok; Items; Error; Columns }. Při chybě nevrací žádné položky.
function ConvertFrom-Rows([string]$text) {
    $lines = Get-ColumnLines $text
    if ($lines.Count -eq 0) {
        return @{ Ok = $false; Error = 'Vstup je prázdný. Zkopírujte řádky z Excelu a vložte je do pole (Ctrl+V).' }
    }
    $result = New-Object System.Collections.Generic.List[object]
    $columns = -1
    for ($i = 0; $i -lt $lines.Count; $i++) {
        $rowNo = $i + 1
        $line = $lines[$i]
        if ([string]::IsNullOrWhiteSpace($line.Replace("`t", ''))) { continue }
        $parts = $line.Split("`t")
        if ($parts.Count -lt 2) {
            return @{ Ok = $false; Error = "Řádek $rowNo má jen jeden sloupec. Zkopírujte z Excelu celý rozsah sloupců (první až poslední).`nNic nebylo načteno." }
        }
        if ($columns -lt 0) { $columns = $parts.Count }
        elseif ($parts.Count -ne $columns) {
            return @{ Ok = $false; Error = "Řádek $rowNo má $($parts.Count) sloupců, ale předchozí řádky mají $columns.`nZkopírujte z Excelu souvislý obdélníkový rozsah.`nNic nebylo načteno." }
        }
        $v = $parts[0].Trim()
        $n = $parts[$parts.Count - 1].Trim()
        if ($v -eq '') {
            return @{ Ok = $false; Error = "Řádek $rowNo má prázdný první sloupec (Údaj k ověření).`nNic nebylo načteno." }
        }
        $result.Add((New-Item2 $v $n))
    }
    return @{ Ok = $true; Items = $result; Columns = $columns }
}

# ---------- GUI ----------
# Barevná paleta (světlý moderní vzhled)
function RGB([int]$r, [int]$g, [int]$b) { [System.Drawing.Color]::FromArgb($r, $g, $b) }
$cBg        = RGB 241 245 249   # pozadí okna
$cCard      = RGB 255 255 255   # karty
$cBorder    = RGB 226 232 240
$cText      = RGB 15 23 42
$cMuted     = RGB 100 116 139
$cAccent    = RGB 79 70 229     # indigo
$cAccentBg  = RGB 238 242 255   # zvýraznění aktuálního řádku
$cHeader    = RGB 255 255 255
$cYes       = RGB 22 163 74
$cNo        = RGB 225 29 72
$cUnsure    = RGB 217 119 6
$cWarn      = RGB 185 28 28
$cDone      = RGB 21 128 61
$cNeutral   = RGB 241 245 249
$cInput     = RGB 248 250 252
$cWhite     = [System.Drawing.Color]::White

$fontBase  = New-Object System.Drawing.Font('Segoe UI', 10)
$fontSmall = New-Object System.Drawing.Font('Segoe UI', 9)
$fontHead  = New-Object System.Drawing.Font('Segoe UI Semibold', 10.5)
$fontBig   = New-Object System.Drawing.Font('Segoe UI Semibold', 28)
$fontBtn   = New-Object System.Drawing.Font('Segoe UI Semibold', 18)
$fontState = New-Object System.Drawing.Font('Segoe UI Semibold', 10.5)
$fontMono  = New-Object System.Drawing.Font('Consolas', 10.5)
$fontMonoB = New-Object System.Drawing.Font('Consolas', 10.5, [System.Drawing.FontStyle]::Bold)

# Tmavší odstín barvy (pro najetí myší)
function Get-Shade($c, [double]$f) {
    [System.Drawing.Color]::FromArgb([int]($c.R * $f), [int]($c.G * $f), [int]($c.B * $f))
}

# Zaoblené rohy ovládacího prvku (poloměr v Tag)
$script:RoundHandler = {
    param($s, $e)
    try {
        $r = [int]([int]$s.Tag * $script:S)
        $w = $s.Width; $h = $s.Height
        if ($w -le 2 * $r -or $h -le 2 * $r) { return }
        $d = 2 * $r
        $path = New-Object System.Drawing.Drawing2D.GraphicsPath
        $path.AddArc(0, 0, $d, $d, 180, 90)
        $path.AddArc($w - $d, 0, $d, $d, 270, 90)
        $path.AddArc($w - $d, $h - $d, $d, $d, 0, 90)
        $path.AddArc(0, $h - $d, $d, $d, 90, 90)
        $path.CloseFigure()
        $s.Region = New-Object System.Drawing.Region($path)
    } catch { }
}
function Set-Rounded($ctrl, [int]$radius) {
    $ctrl.Tag = $radius
    $ctrl.Add_Resize($script:RoundHandler)
    & $script:RoundHandler $ctrl $null
}

function Set-FlatButton($btn, $back, $fore) {
    $btn.FlatStyle = 'Flat'
    $btn.FlatAppearance.BorderSize = 0
    $btn.FlatAppearance.MouseOverBackColor = Get-Shade $back 0.92
    $btn.FlatAppearance.MouseDownBackColor = Get-Shade $back 0.85
    $btn.BackColor = $back
    $btn.ForeColor = $fore
    $btn.Cursor = [System.Windows.Forms.Cursors]::Hand
    $btn.UseVisualStyleBackColor = $false
}

function New-Button([string]$text, $back, $fore) {
    $b = New-Object System.Windows.Forms.Button
    $b.Text = $text
    $b.AutoSize = $true
    $b.Height = 36
    $b.MinimumSize = New-Object System.Drawing.Size(0, 36)
    $b.Padding = New-Object System.Windows.Forms.Padding(12, 2, 12, 2)
    $b.Margin = New-Object System.Windows.Forms.Padding(0, 4, 8, 6)
    Set-FlatButton $b $back $fore
    Set-Rounded $b 8
    $b
}

# Velké tlačítko vyplňující buňku
function New-BigButton([string]$text, $back, $fore, $font) {
    $b = New-Object System.Windows.Forms.Button
    $b.Text = $text
    $b.Dock = 'Fill'
    $b.Font = $font
    $b.Margin = New-Object System.Windows.Forms.Padding(6, 4, 6, 4)
    Set-FlatButton $b $back $fore
    Set-Rounded $b 12
    $b
}

# Světlé vedlejší tlačítko
function New-GhostButton([string]$text, $back = $null) {
    if ($null -eq $back) { $back = RGB 232 237 244 }
    New-Button $text $back $cText
}

function New-Card {
    $outer = New-Object System.Windows.Forms.Panel
    $outer.Dock = 'Fill'
    $outer.BackColor = $cBorder
    $outer.Padding = New-Object System.Windows.Forms.Padding(1)
    $outer.Margin = New-Object System.Windows.Forms.Padding(6)
    $inner = New-Object System.Windows.Forms.Panel
    $inner.Dock = 'Fill'
    $inner.BackColor = $cCard
    $inner.Padding = New-Object System.Windows.Forms.Padding(16, 12, 16, 12)
    $outer.Controls.Add($inner)
    Set-Rounded $outer 12
    Set-Rounded $inner 11
    @{ Outer = $outer; Inner = $inner }
}

function New-Caption([string]$text) {
    $l = New-Object System.Windows.Forms.Label
    $l.Text = $text
    $l.Dock = 'Top'
    $l.Height = 26
    $l.Font = $fontHead
    $l.ForeColor = $cText
    $l
}

$form = New-Object System.Windows.Forms.Form
$form.Text = 'Kontrola Clearance'

# Ikona „KC“ kreslená přímo v aplikaci (bez souboru): zaoblený čtverec s bílým písmem
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
    $brush = New-Object System.Drawing.Drawing2D.LinearGradientBrush($rect, (RGB 99 102 241), (RGB 67 56 202), 45.0)
    $g.FillPath($brush, $path)
    $font = New-Object System.Drawing.Font('Segoe UI', [float]($size * 0.40), [System.Drawing.FontStyle]::Bold, [System.Drawing.GraphicsUnit]::Pixel)
    $fmt = New-Object System.Drawing.StringFormat
    $fmt.Alignment = 'Center'; $fmt.LineAlignment = 'Center'
    $rf = New-Object System.Drawing.RectangleF(0, [float]($size * 0.02), $size, $size)
    $g.DrawString('KC', $font, [System.Drawing.Brushes]::White, $rf, $fmt)
    $font.Dispose(); $brush.Dispose(); $path.Dispose(); $g.Dispose()
    $bmp
}
# Vlastní identita procesu pro hlavní panel Windows: okno se neseskupí s ostatními okny PowerShellu
# a na liště se zobrazí ikona KC
try {
    if (-not ('KcTaskbar' -as [type])) {
        Add-Type -TypeDefinition @'
using System.Runtime.InteropServices;
public static class KcTaskbar {
    [DllImport("shell32.dll", CharSet = CharSet.Unicode)]
    public static extern int SetCurrentProcessExplicitAppUserModelID(string appId);
}
'@
    }
    [void][KcTaskbar]::SetCurrentProcessExplicitAppUserModelID('KontrolaClearance.App')
} catch { }
try {
    $iconBmp = New-KcBitmap 256
    $form.Icon = [System.Drawing.Icon]::FromHandle($iconBmp.GetHicon())
} catch { }

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

if ($VytvoritZastupce) {
    try {
        $dir = Split-Path -Parent $MyInvocation.MyCommand.Path
        $ico = Join-Path $dir 'kc.ico'
        Save-KcIco $ico
        $shell = New-Object -ComObject WScript.Shell
        $targets = @((Join-Path $dir 'Kontrola Clearance.lnk'), (Join-Path ([Environment]::GetFolderPath('Desktop')) 'Kontrola Clearance.lnk'))
        foreach ($lnkPath in $targets) {
            $lnk = $shell.CreateShortcut($lnkPath)
            $lnk.TargetPath = Join-Path $env:SystemRoot 'System32\wscript.exe'
            $lnk.Arguments = '"' + (Join-Path $dir 'spustit.vbs') + '"'
            $lnk.WorkingDirectory = $dir
            $lnk.IconLocation = "$ico,0"
            $lnk.Description = 'Kontrola Clearance'
            $lnk.Save()
        }
        Show-Info "Zástupce „Kontrola Clearance“ s ikonou KC byl vytvořen ve složce aplikace a na ploše."
    } catch {
        Show-Error 'Zástupce se nepodařilo vytvořit. Zkontrolujte, zda máte do složky aplikace právo zápisu.'
    }
    return
}
$form.Size = New-Object System.Drawing.Size(1200, 820)
$form.MinimumSize = New-Object System.Drawing.Size(950, 650)
$form.StartPosition = 'CenterScreen'
$form.Font = $fontBase
# Pevné rozměry v pixelech (výšky řádků, okraje …) se přepočtou podle měřítka obrazovky
$form.AutoScaleDimensions = New-Object System.Drawing.SizeF(96, 96)
$form.AutoScaleMode = 'Dpi'
$form.BackColor = $cBg
$form.ForeColor = $cText

# --- Hlavička: název, stav a pozice ---
$header = New-Object System.Windows.Forms.TableLayoutPanel
$header.Dock = 'Top'
$header.Height = 60
$header.BackColor = $cHeader
$header.ColumnCount = 4
$header.Padding = New-Object System.Windows.Forms.Padding(16, 0, 16, 0)
[void]$header.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle('AutoSize')))
[void]$header.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle('Percent', 100)))
[void]$header.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle('AutoSize')))
[void]$header.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle('AutoSize')))
$lblApp = New-Object System.Windows.Forms.Label
$lblApp.Text = 'Kontrola Clearance'
$lblApp.AutoSize = $true
$lblApp.Anchor = 'Left'
$lblApp.Font = New-Object System.Drawing.Font('Segoe UI Semibold', 16)
$lblApp.ForeColor = $cAccent
$lblState = New-Object System.Windows.Forms.Label
$lblState.AutoSize = $true
$lblState.Anchor = 'Left'
$lblState.Margin = New-Object System.Windows.Forms.Padding(18, 0, 0, 0)
$lblState.Padding = New-Object System.Windows.Forms.Padding(12, 5, 12, 5)
$lblState.Font = $fontState
Set-Rounded $lblState 12
$lblPos = New-Object System.Windows.Forms.Label
$lblPos.AutoSize = $true
$lblPos.Anchor = 'Right'
$lblPos.Font = New-Object System.Drawing.Font('Segoe UI Semibold', 16)
$lblPos.ForeColor = $cText
$header.Controls.Add($lblApp, 0, 0)
$header.Controls.Add($lblState, 1, 0)
$chkMiddle = New-Object System.Windows.Forms.CheckBox
$chkMiddle.Text = 'Vkládat Ctrl + kliknutím'
$chkMiddle.AutoSize = $true
$chkMiddle.Anchor = 'Right'
$chkMiddle.Checked = $true
$chkMiddle.ForeColor = $cMuted
$chkMiddle.Margin = New-Object System.Windows.Forms.Padding(0, 0, 24, 0)
$header.Controls.Add($chkMiddle, 2, 0)
$header.Controls.Add($lblPos, 3, 0)

$split = New-Object System.Windows.Forms.SplitContainer
$split.Dock = 'Fill'
$split.Orientation = 'Vertical'
$split.BackColor = $cBg
$split.SplitterWidth = 6
$split.Padding = New-Object System.Windows.Forms.Padding(8)
$form.Controls.Add($split)
$headerLine = New-Object System.Windows.Forms.Panel
$headerLine.Dock = 'Top'
$headerLine.Height = 1
$headerLine.BackColor = $cBorder
$form.Controls.Add($headerLine)
$form.Controls.Add($header)

# --- Levá část: import a práce ---
$left = New-Object System.Windows.Forms.TableLayoutPanel
$left.Dock = 'Fill'
$left.ColumnCount = 1
$left.RowCount = 4
[void]$left.RowStyles.Add((New-Object System.Windows.Forms.RowStyle('Percent', 50)))    # import
[void]$left.RowStyles.Add((New-Object System.Windows.Forms.RowStyle('Percent', 50)))    # aktuální hodnota
[void]$left.RowStyles.Add((New-Object System.Windows.Forms.RowStyle('Absolute', 88)))   # Ponechat / Vymazat
[void]$left.RowStyles.Add((New-Object System.Windows.Forms.RowStyle('AutoSize')))       # vedlejší akce
$split.Panel1.Controls.Add($left)

# Import: jedno pole, do kterého se vloží řádky z Excelu (použije se 1. a poslední sloupec)
$importCard = New-Card
$left.Controls.Add($importCard.Outer, 0, 0)

$importGrid = New-Object System.Windows.Forms.TableLayoutPanel
$importGrid.Dock = 'Fill'
$importGrid.ColumnCount = 1
$importGrid.RowCount = 2
[void]$importGrid.RowStyles.Add((New-Object System.Windows.Forms.RowStyle('Percent', 100)))
[void]$importGrid.RowStyles.Add((New-Object System.Windows.Forms.RowStyle('AutoSize')))
$importCard.Inner.Controls.Add($importGrid)

# RichTextBox kvůli zvýraznění aktuálního řádku
$txtInput = New-Object System.Windows.Forms.RichTextBox
$txtInput.Multiline = $true
$txtInput.ScrollBars = 'Both'
$txtInput.WordWrap = $false
$txtInput.DetectUrls = $false
$txtInput.AcceptsTab = $true
$txtInput.Dock = 'Fill'
$txtInput.Font = $fontMono
$txtInput.BorderStyle = 'None'
$txtInput.BackColor = $cInput

$importPanel = New-Object System.Windows.Forms.FlowLayoutPanel
$importPanel.Dock = 'Fill'
$importPanel.AutoSize = $true
$importPanel.AutoSizeMode = 'GrowAndShrink'
$importPanel.WrapContents = $false
$importPanel.Padding = New-Object System.Windows.Forms.Padding(0, 4, 0, 0)
$btnLoad = New-Button 'Vytvořit seznam' $cAccent $cWhite
$btnLoad.Font = $fontState
$btnClear = New-GhostButton 'Zrušit seznam'
$lblCount = New-Object System.Windows.Forms.Label
$lblCount.AutoSize = $true
$lblCount.Padding = New-Object System.Windows.Forms.Padding(8, 9, 0, 0)
$lblCount.Font = $fontState
$lblCount.Text = 'Řádků: 0'
$importPanel.Controls.AddRange(@($btnLoad, $btnClear, $lblCount))
# Seznam řádků se stavem (zobrazí se místo vstupního pole po Vytvořit seznam)
$lstItems = New-Object System.Windows.Forms.ListBox
$lstItems.Dock = 'Fill'
$lstItems.DrawMode = 'OwnerDrawFixed'
$lstItems.ItemHeight = Px 28
$lstItems.IntegralHeight = $false
$lstItems.BorderStyle = 'None'
$lstItems.BackColor = $cInput
$lstItems.Font = $fontBase
$lstItems.Visible = $false
$inputHost = New-Object System.Windows.Forms.Panel
$inputHost.Dock = 'Fill'
$inputHost.Margin = New-Object System.Windows.Forms.Padding(0)
$inputHost.Controls.Add($txtInput)
$inputHost.Controls.Add($lstItems)
$importGrid.Controls.Add($inputHost, 0, 0)
$importGrid.Controls.Add($importPanel, 0, 1)

# Aktuální hodnota: horní lišta (popisek + šipky), velká hodnota, dole info o poznámce
$currentCard = New-Card
$left.Controls.Add($currentCard.Outer, 0, 1)

$currentTop = New-Object System.Windows.Forms.TableLayoutPanel
$currentTop.Dock = 'Top'
$currentTop.Height = 40
$currentTop.ColumnCount = 3
[void]$currentTop.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle('Percent', 100)))
[void]$currentTop.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle('Absolute', 46)))
[void]$currentTop.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle('Absolute', 46)))
$lblCurrentCap = New-Object System.Windows.Forms.Label
$lblCurrentCap.Text = 'AKTUÁLNÍ ÚDAJ  •  zkopírováno do schránky'
$lblCurrentCap.Dock = 'Fill'
$lblCurrentCap.TextAlign = 'MiddleLeft'
$lblCurrentCap.Font = $fontSmall
$lblCurrentCap.ForeColor = $cMuted
function New-NavButton([string]$text) {
    $b = New-Object System.Windows.Forms.Button
    $b.Text = $text
    $b.Dock = 'Fill'
    $b.Font = New-Object System.Drawing.Font('Segoe UI', 11)
    $b.Margin = New-Object System.Windows.Forms.Padding(4, 2, 0, 2)
    Set-FlatButton $b $cNeutral $cText
    Set-Rounded $b 8
    $b
}
$btnUp = New-NavButton '▲'
$btnDown = New-NavButton '▼'
$tips = New-Object System.Windows.Forms.ToolTip
$tips.SetToolTip($btnUp, 'Předchozí řádek')
$tips.SetToolTip($btnDown, 'Další řádek')
$currentTop.Controls.Add($lblCurrentCap, 0, 0)
$currentTop.Controls.Add($btnUp, 1, 0)
$currentTop.Controls.Add($btnDown, 2, 0)

$lblCurrentNote = New-Object System.Windows.Forms.Label
$lblCurrentNote.Dock = 'Bottom'
$lblCurrentNote.Height = 34
$lblCurrentNote.Font = $fontBase
$lblCurrentNote.ForeColor = $cMuted
$lblCurrentNote.TextAlign = 'MiddleCenter'
$lblCurrentNote.AutoEllipsis = $true
$lblCurrent = New-Object System.Windows.Forms.Label
$lblCurrent.Dock = 'Fill'
$lblCurrent.Font = $fontBig
$lblCurrent.TextAlign = 'MiddleCenter'
$lblCurrent.AutoEllipsis = $true
$currentCard.Inner.Controls.Add($lblCurrent)
$currentCard.Inner.Controls.Add($lblCurrentNote)
$currentCard.Inner.Controls.Add($currentTop)

# Hlavní rozhodnutí: dvě velká tlačítka vedle sebe
$actionPanel = New-Object System.Windows.Forms.TableLayoutPanel
$actionPanel.Dock = 'Fill'
$actionPanel.ColumnCount = 2
[void]$actionPanel.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle('Percent', 50)))
[void]$actionPanel.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle('Percent', 50)))
$btnYes = New-BigButton 'Ponechat' $cYes $cWhite $fontBtn
$btnNo  = New-BigButton 'Vymazat' $cNo $cWhite $fontBtn
$actionPanel.Controls.Add($btnYes, 0, 0)
$actionPanel.Controls.Add($btnNo, 1, 0)
$left.Controls.Add($actionPanel, 0, 2)

# Vedlejší akce: malá tlačítka uprostřed pod hlavními
$secondaryPanel = New-Object System.Windows.Forms.TableLayoutPanel
$secondaryPanel.Dock = 'Fill'
$secondaryPanel.AutoSize = $true
$secondaryPanel.AutoSizeMode = 'GrowAndShrink'
$secondaryPanel.ColumnCount = 4
[void]$secondaryPanel.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle('Percent', 50)))
[void]$secondaryPanel.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle('AutoSize')))
[void]$secondaryPanel.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle('AutoSize')))
[void]$secondaryPanel.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle('Percent', 50)))
$btnUnsure = New-GhostButton '?   Vrátit se později' $cWhite
$btnUnsure.ForeColor = $cUnsure
$btnInsert = New-GhostButton '+   Vložit mezi' $cWhite
$btnInsert.ForeColor = $cAccent
$btnUnsure.Anchor = 'None'
$btnInsert.Anchor = 'None'
$secondaryPanel.Controls.Add($btnUnsure, 1, 0)
$secondaryPanel.Controls.Add($btnInsert, 2, 0)
$left.Controls.Add($secondaryPanel, 0, 3)

# --- Pravá část: poznámky ---
$notesCard = New-Card
$split.Panel2.Controls.Add($notesCard.Outer)
$right = New-Object System.Windows.Forms.TableLayoutPanel
$right.Dock = 'Fill'
$right.ColumnCount = 1
$right.RowCount = 5
[void]$right.RowStyles.Add((New-Object System.Windows.Forms.RowStyle('Absolute', 26)))
[void]$right.RowStyles.Add((New-Object System.Windows.Forms.RowStyle('AutoSize')))
[void]$right.RowStyles.Add((New-Object System.Windows.Forms.RowStyle('Absolute', 30)))
[void]$right.RowStyles.Add((New-Object System.Windows.Forms.RowStyle('Percent', 100)))
[void]$right.RowStyles.Add((New-Object System.Windows.Forms.RowStyle('AutoSize')))
$notesCard.Inner.Controls.Add($right)

$lblTitle = New-Object System.Windows.Forms.Label
$lblTitle.Text = 'Název souboru'
$lblTitle.Dock = 'Fill'
$lblTitle.Font = $fontHead
$titleRow = New-Object System.Windows.Forms.TableLayoutPanel
$titleRow.Dock = 'Fill'
$titleRow.AutoSize = $true
$titleRow.AutoSizeMode = 'GrowAndShrink'
$titleRow.ColumnCount = 2
[void]$titleRow.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle('Percent', 100)))
[void]$titleRow.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle('AutoSize')))
$txtFileName = New-Object System.Windows.Forms.TextBox
$txtFileName.Dock = 'Fill'
$txtFileName.Font = New-Object System.Drawing.Font('Segoe UI', 11)
$txtFileName.Margin = New-Object System.Windows.Forms.Padding(0, 6, 4, 0)
$lblExt = New-Object System.Windows.Forms.Label
$lblExt.Text = '.txt'
$lblExt.AutoSize = $true
$lblExt.Anchor = 'Left'
$lblExt.ForeColor = $cMuted
$titleRow.Controls.Add($txtFileName, 0, 0)
$titleRow.Controls.Add($lblExt, 1, 0)
$lblNotes = New-Object System.Windows.Forms.Label
$lblNotes.Text = 'Poznámky'
$lblNotes.Dock = 'Fill'
$lblNotes.Font = $fontHead
$lblNotes.TextAlign = 'BottomLeft'
$lblNotes.AutoEllipsis = $true
$txtNotes = New-Object System.Windows.Forms.TextBox
$txtNotes.Multiline = $true
$txtNotes.ScrollBars = 'Vertical'
$txtNotes.WordWrap = $true
$txtNotes.Dock = 'Fill'
$txtNotes.Font = $fontMono
$txtNotes.BorderStyle = 'None'
$txtNotes.BackColor = $cInput
$saveRow = New-Object System.Windows.Forms.FlowLayoutPanel
$saveRow.Dock = 'Fill'
$saveRow.AutoSize = $true
$saveRow.AutoSizeMode = 'GrowAndShrink'
$saveRow.WrapContents = $false
$saveRow.Padding = New-Object System.Windows.Forms.Padding(0, 6, 0, 0)
$btnOpen = New-GhostButton 'Otevřít soubor…'
$btnSave = New-Button 'Uložit' $cAccent $cWhite
$btnSaveAs = New-GhostButton 'Uložit jako…'
$saveRow.Controls.AddRange(@($btnOpen, $btnSave, $btnSaveAs))
$right.Controls.Add($lblTitle, 0, 0)
$right.Controls.Add($titleRow, 0, 1)
$right.Controls.Add($lblNotes, 0, 2)
$right.Controls.Add($txtNotes, 0, 3)
$right.Controls.Add($saveRow, 0, 4)

# ---------- Logika ----------
function Add-NoteLine([string]$line) {
    $t = $txtNotes.Text
    if ($t.Length -gt 0 -and -not $t.EndsWith("`n")) { $t += "`r`n" }
    $txtNotes.Text = $t + $line + "`r`n"
    $txtNotes.SelectionStart = $txtNotes.Text.Length
    $txtNotes.ScrollToCaret()
}

function Set-State([string]$text, $back, $fore) {
    $lblState.Text = $text
    $lblState.BackColor = $back
    $lblState.ForeColor = $fore
}

# Odstraní z poznámek poslední řádek přesně rovný $line (při změně Vymazat -> Ponechat)
function Remove-NoteLine([string]$line) {
    $lines = [System.Collections.Generic.List[string]]($txtNotes.Text -split "\r?\n")
    for ($i = $lines.Count - 1; $i -ge 0; $i--) {
        if ($lines[$i].Trim() -eq $line) {
            $lines.RemoveAt($i)
            $txtNotes.Text = $lines -join "`r`n"
            return $true
        }
    }
    return $false
}

# Vykreslení jednoho řádku seznamu: číslo, značka stavu, údaj, poznámka
#   ✓ = Ponechat, ✗ = Vymazat (zapsáno do poznámek), ? = vrátit se později, bez značky = nerozhodnuto
$lstItems.Add_DrawItem({ param($s, $e)
    try {
        if ($e.Index -lt 0 -or $e.Index -ge $script:Items.Count) { return }
        $it = $script:Items[$e.Index]
        $g = $e.Graphics
        $r = $e.Bounds
        $isCur = ($e.Index -eq $script:Index -and -not $script:Done)
        $back = if ($isCur) { $cAccentBg } else { $cInput }
        $bb = New-Object System.Drawing.SolidBrush($back)
        $g.FillRectangle($bb, $r)
        $bb.Dispose()
        if ($isCur) {
            $ab = New-Object System.Drawing.SolidBrush($cAccent)
            $g.FillRectangle($ab, $r.X, $r.Y, (Px 4), $r.Height)
            $ab.Dispose()
        }
        $flags = [System.Windows.Forms.TextFormatFlags]'VerticalCenter, EndEllipsis, NoPrefix, SingleLine'
        $mark = ''; $markColor = $cMuted
        switch ($it.Status) {
            'keep'   { $mark = '✓'; $markColor = $cYes }
            'del'    { $mark = '✗'; $markColor = $cNo }
            'unsure' { $mark = '?'; $markColor = $cUnsure }
        }
        $wRest = [Math]::Max(10, $r.Width - (Px 80))
        $wVal = [int]($wRest * 0.45)
        $numRect  = New-Object System.Drawing.Rectangle(($r.X + (Px 8)), $r.Y, (Px 40), $r.Height)
        $markRect = New-Object System.Drawing.Rectangle(($r.X + (Px 48)), $r.Y, (Px 24), $r.Height)
        $valRect  = New-Object System.Drawing.Rectangle(($r.X + (Px 76)), $r.Y, $wVal, $r.Height)
        $noteRect = New-Object System.Drawing.Rectangle(($r.X + (Px 84) + $wVal), $r.Y, ($wRest - $wVal - (Px 8)), $r.Height)
        [System.Windows.Forms.TextRenderer]::DrawText($g, "$($e.Index + 1)", $fontSmall, $numRect, $cMuted, $flags)
        [System.Windows.Forms.TextRenderer]::DrawText($g, $mark, $fontState, $markRect, $markColor, $flags)
        $valFont = if ($isCur) { $fontHead } else { $fontBase }
        $valColor = if ($isCur) { $cAccent } else { $cText }
        [System.Windows.Forms.TextRenderer]::DrawText($g, $it.Value, $valFont, $valRect, $valColor, $flags)
        $note = $it.Note
        if ($it.Manual) { $note = "$note   (vloženo ručně)" }
        [System.Windows.Forms.TextRenderer]::DrawText($g, $note, $fontSmall, $noteRect, $cMuted, $flags)
    } catch { }
})

# Kliknutí nebo šipky v seznamu: vybraný řádek se stane aktuálním
$lstItems.Add_SelectedIndexChanged({
    if ($script:Syncing) { return }
    Invoke-Safe {
        $i = $lstItems.SelectedIndex
        if ($i -ge 0 -and $i -lt $script:Items.Count) {
            $script:Index = $i
            $script:Done = $false
            Show-Current
        }
    } 'Přechod na řádek se nezdařil.'
})

# Seznam místo vstupního pole (bez seznamu se zobrazí vstupní pole)
function Update-ListBox {
    if ($script:Items.Count -eq 0) {
        $lstItems.Visible = $false
        $txtInput.Visible = $true
        return
    }
    $script:Syncing = $true
    try {
        $lstItems.BeginUpdate()
        if ($lstItems.Items.Count -ne $script:Items.Count) {
            $lstItems.Items.Clear()
            for ($i = 0; $i -lt $script:Items.Count; $i++) { [void]$lstItems.Items.Add($i) }
        }
        $lstItems.SelectedIndex = if ($script:Done) { -1 } else { $script:Index }
        $lstItems.EndUpdate()
        $lstItems.Invalidate()
        $txtInput.Visible = $false
        $lstItems.Visible = $true
    } finally { $script:Syncing = $false }
}

function Update-View {
    $count = $script:Items.Count
    $decided = @($script:Items | Where-Object { $_.Status -eq 'keep' -or $_.Status -eq 'del' }).Count
    $btnOn = ($count -gt 0 -and -not $script:Done)
    # tlačítka se nevypínají (vypnutá vypadají nečitelně), jen se zesvětlí; obsluha si sama hlídá stav
    $btnYes.Cursor = if ($btnOn) { [System.Windows.Forms.Cursors]::Hand } else { [System.Windows.Forms.Cursors]::Default }
    $btnNo.Cursor = $btnYes.Cursor
    $btnUp.Enabled = ($count -gt 0); $btnDown.Enabled = ($count -gt 0)
    $btnYes.BackColor = if ($btnOn) { $cYes } else { RGB 187 222 199 }
    $btnNo.BackColor  = if ($btnOn) { $cNo }  else { RGB 240 190 202 }
    $lblCurrentNote.Text = ''
    if ($count -eq 0) {
        Set-State 'Bez seznamu' (RGB 226 232 240) $cMuted
        $lblPos.Text = '0 / 0'
        $lblCurrent.Text = 'Vložte řádky z Excelu a klikněte na Vytvořit seznam'
        $lblCurrent.ForeColor = $cMuted
    } elseif ($script:Done) {
        Set-State "✓ Dokončeno  ($decided / $count)" (RGB 220 252 231) $cDone
        $lblPos.Text = "$count / $count"
        $lblCurrent.Text = '✓ Hotovo – všechny řádky jsou rozhodnuté'
        $lblCurrent.ForeColor = $cDone
        $lblCurrentNote.Text = 'Šipkami ▲ ▼ nebo kliknutím do seznamu se můžete k libovolnému řádku vrátit.'
    } else {
        if ($decided -eq 0) { Set-State "Načteno  •  hotovo 0 / $count" $cAccentBg $cAccent }
        else { Set-State "Probíhá  •  hotovo $decided / $count" (RGB 254 243 199) (RGB 180 83 9) }
        $lblPos.Text = "$($script:Index + 1) / $count"
        $item = $script:Items[$script:Index]
        $lblCurrent.Text = $item.Value
        $lblCurrent.ForeColor = $cText
        $info = if ($item.Note -ne '') { "Při Vymazat se zapíše: $($item.Note)" } else { 'Při Vymazat se nic nezapíše (poznámka je prázdná)' }
        switch ($item.Status) {
            'keep' { $info = '✓ Ponecháno   •   ' + $info }
            'del'  { $info = '✗ Vymazáno (zapsáno v poznámkách)   •   ' + $info }
            'unsure' { $info = '? Vrátit se později   •   ' + $info }
        }
        if ($item.Manual) { $info += '   •   vloženo ručně' }
        $lblCurrentNote.Text = $info
    }
    Update-ListBox
}

# Zobrazí aktuální řádek a zkopíruje jeho první hodnotu do schránky
function Show-Current {
    Update-View
    if (-not $script:Done -and $script:Index -lt $script:Items.Count) {
        $value = $script:Items[$script:Index].Value
        try { [System.Windows.Forms.Clipboard]::SetText($value) }
        catch { Show-Warn 'Hodnotu se nepodařilo zkopírovat do schránky (schránka je možná obsazená jinou aplikací).' }
    }
}

# Přejde na další nerozhodnutý řádek (hledá od aktuálního dál, pak od začátku).
# Až nezbývá žádný nerozhodnutý, přijdou na řadu řádky označené „?“; když nejsou ani ty, dokončeno.
function Move-NextUndecided {
    $count = $script:Items.Count
    foreach ($wanted in @('', 'unsure')) {
        for ($k = 1; $k -le $count; $k++) {
            $i = ($script:Index + $k) % $count
            if ($script:Items[$i].Status -eq $wanted) {
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

# Ruční posun o $delta řádků
function Move-By([int]$delta) {
    $count = $script:Items.Count
    if ($count -eq 0) { return }
    if ($script:Done) { $script:Done = $false; $delta = 0 }
    $script:Index = [Math]::Max(0, [Math]::Min($count - 1, $script:Index + $delta))
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

function Update-NotesCaption {
    $lblNotes.Text = if ($script:NotesFile) { "Poznámky  •  $([System.IO.Path]::GetDirectoryName($script:NotesFile))" } else { 'Poznámky  •  zatím neuloženo' }
}

# Název z pole „Název souboru“ (bez .txt); prázdný = $null
function Get-WantedFileName {
    $name = $txtFileName.Text.Trim()
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
            $r = [System.Windows.Forms.MessageBox]::Show("Soubor $([System.IO.Path]::GetFileName($path)) už existuje. Přepsat ho?", 'Uložit', 'YesNo', 'Warning')
            if ($r -ne 'Yes') { return }
        }
        $enc = New-Object System.Text.UTF8Encoding($true)
        [System.IO.File]::WriteAllText($path, $txtNotes.Text, $enc)
        if ($renaming -and ($old.ToLower() -ne $path.ToLower()) -and [System.IO.File]::Exists($old)) {
            [System.IO.File]::Delete($old)
        }
        $script:NotesFile = $path
        $txtFileName.Text = [System.IO.Path]::GetFileNameWithoutExtension($path)
        Update-NotesCaption
        Show-Info "Poznámky byly uloženy do:`n$path"
    } catch {
        Show-Error 'Poznámky se nepodařilo uložit. Zkontrolujte, zda je soubor dostupný a máte do složky právo zápisu.'
    }
}

function Save-NotesAs {
    $sfd = New-Object System.Windows.Forms.SaveFileDialog
    $sfd.Filter = 'Textový soubor (*.txt)|*.txt'
    $sfd.DefaultExt = 'txt'
    $sfd.AddExtension = $true
    $wanted = Get-WantedFileName
    if ($wanted -eq $false) { return }
    $sfd.FileName = if ($wanted) { $wanted } elseif ($script:NotesFile) { [System.IO.Path]::GetFileName($script:NotesFile) } else { 'poznamky.txt' }
    if ($script:NotesFile) { $sfd.InitialDirectory = [System.IO.Path]::GetDirectoryName($script:NotesFile) }
    if ($sfd.ShowDialog($form) -eq 'OK') { Save-Notes $sfd.FileName }
    $sfd.Dispose()
}

# Počet řádků ve vstupním poli (koncové prázdné řádky se nepočítají)
function Update-InputCount {
    $lines = Get-ColumnLines $txtInput.Text
    $rows = @($lines | Where-Object { -not [string]::IsNullOrWhiteSpace($_.Replace("`t", '')) })
    $text = "Řádků: $($rows.Count)"
    $lblCount.ForeColor = $cText
    if ($rows.Count -gt 0 -and $rows[0].Split("`t").Count -lt 2) {
        $text = $text + '   •   jen 1 sloupec!'
        $lblCount.ForeColor = $cWarn
    }
    $lblCount.Text = $text
}

$txtInput.Add_TextChanged({
    if ($script:Highlighting) { return }
    if ($txtInput.TextLength -eq 0) { $script:PastedCols = 0 }
    Invoke-Safe { Update-InputCount } 'Počet řádků se nepodařilo spočítat.' })
# Ctrl+V / Shift+Insert vloží jen čistý text (bez formátování z Excelu)
$txtInput.Add_KeyDown({ param($s, $e)
    if (($e.Control -and $e.KeyCode -eq 'V') -or ($e.Shift -and $e.KeyCode -eq 'Insert')) {
        $e.SuppressKeyPress = $true
        $e.Handled = $true
        Invoke-Safe {
            if ($script:Items.Count -gt 0) { Show-Warn 'Seznam už je vytvořený. Pro nové vložení ho nejdřív zrušte (Zrušit seznam).'; return }
            if (-not [System.Windows.Forms.Clipboard]::ContainsText()) { return }
            $r = Reduce-PastedColumns ([System.Windows.Forms.Clipboard]::GetText())
            if ($r.Error) { Show-Error $r.Error; return }
            $script:PastedCols = $r.Columns
            $txtInput.SelectedText = $r.Text
        } 'Vložení ze schránky se nezdařilo.'
    }
})

$btnLoad.Add_Click({ Invoke-Safe {
    $parsed = ConvertFrom-Rows $txtInput.Text
    if (-not $parsed.Ok) { Show-Error $parsed.Error; return }   # stávající seznam ani poznámky se nemění
    if ($parsed.Items.Count -eq 0) { Show-Error 'Vstup neobsahuje žádný neprázdný řádek.'; return }

    $script:Items = $parsed.Items
    $script:Index = 0
    $script:Done = $false
    Show-Current
} 'Seznam se nepodařilo načíst. Zkuste znovu zkopírovat data z Excelu.' })

$btnOpen.Add_Click({ Invoke-Safe {
    if ($txtNotes.Text.Trim() -ne '') {
        $r = [System.Windows.Forms.MessageBox]::Show(
            'Otevřením souboru se nahradí aktuální obsah poznámek.' + "`n`n" + 'Pokud ho chcete zachovat, nejdřív ho uložte. Pokračovat?',
            'Otevřít soubor', 'YesNo', 'Question')
        if ($r -ne 'Yes') { return }
    }
    $ofd = New-Object System.Windows.Forms.OpenFileDialog
    $ofd.Filter = 'Textový soubor (*.txt)|*.txt|Všechny soubory (*.*)|*.*'
    if ($ofd.ShowDialog($form) -eq 'OK') {
        try {
            $content = [System.IO.File]::ReadAllText($ofd.FileName, [System.Text.Encoding]::UTF8)
            $content = $content -replace "\r?\n", "`r`n"
            if ($content.Length -gt 0 -and -not $content.EndsWith("`n")) { $content += "`r`n" }
            $txtNotes.Text = $content
            $txtNotes.SelectionStart = $txtNotes.Text.Length
            $txtNotes.ScrollToCaret()
            $script:NotesFile = $ofd.FileName
            $txtFileName.Text = [System.IO.Path]::GetFileNameWithoutExtension($ofd.FileName)
            Update-NotesCaption
        } catch {
            Show-Error 'Soubor se nepodařilo otevřít. Zkontrolujte, zda existuje a není otevřený jinou aplikací.'
        }
    }
    $ofd.Dispose()
} 'Soubor se nepodařilo otevřít.' })

$btnSave.Add_Click({ Invoke-Safe {
    # Uložit: do otevřeného souboru; když se změnil název, soubor se přejmenuje (ve stejné složce)
    $wanted = Get-WantedFileName
    if ($wanted -eq $false) { return }
    if (-not $script:NotesFile) { Save-NotesAs; return }
    $dir = [System.IO.Path]::GetDirectoryName($script:NotesFile)
    $target = if ($wanted) { [System.IO.Path]::Combine($dir, $wanted) } else { $script:NotesFile }
    Save-Notes $target $true
} 'Ukládání se nezdařilo.' })

$btnSaveAs.Add_Click({ Invoke-Safe { Save-NotesAs } 'Ukládání se nezdařilo.' })

$btnYes.Add_Click({ Invoke-Safe { Set-Decision 'keep' } 'Přechod na další řádek se nezdařil.' })
$btnUp.Add_Click({ Invoke-Safe { Move-By -1 } 'Přechod na řádek se nezdařil.' })
$btnDown.Add_Click({ Invoke-Safe { Move-By 1 } 'Přechod na řádek se nezdařil.' })

$btnNo.Add_Click({ Invoke-Safe { Set-Decision 'del' } 'Zápis poznámky se nezdařil.' })
$btnUnsure.Add_Click({ Invoke-Safe { Set-Decision 'unsure' } 'Označení se nezdařilo.' })

$btnInsert.Add_Click({ Invoke-Safe {
    if ($script:Index -ge $script:Items.Count) { return }

    $dlg = New-Object System.Windows.Forms.Form
    $dlg.Text = 'Vložit mezi'
    $dlg.FormBorderStyle = 'FixedDialog'
    $dlg.MaximizeBox = $false; $dlg.MinimizeBox = $false
    $dlg.StartPosition = 'CenterParent'
    $dlg.ClientSize = New-Object System.Drawing.Size(460, 190)
    $dlg.Font = $fontBase

    $l1 = New-Object System.Windows.Forms.Label
    $l1.Text = 'Údaj k ověření (povinné):'; $l1.SetBounds(12, 12, 430, 22)
    $t1 = New-Object System.Windows.Forms.TextBox
    $t1.SetBounds(12, 36, 434, 26)
    $l2 = New-Object System.Windows.Forms.Label
    $l2.Text = 'Poznámka při Vymazat (volitelné):'; $l2.SetBounds(12, 70, 430, 22)
    $t2 = New-Object System.Windows.Forms.TextBox
    $t2.SetBounds(12, 94, 434, 26)
    $ok = New-Object System.Windows.Forms.Button
    $ok.Text = 'Vložit'; $ok.SetBounds(256, 140, 90, 32)
    $cancel = New-Object System.Windows.Forms.Button
    $cancel.Text = 'Zrušit'; $cancel.SetBounds(356, 140, 90, 32)
    $cancel.DialogResult = 'Cancel'
    $ok.Add_Click({
        if ($t1.Text.Trim() -eq '') {
            Show-Warn 'Vyplňte Údaj k ověření.'
        } else {
            $dlg.DialogResult = 'OK'
            $dlg.Close()
        }
    })
    $dlg.Controls.AddRange(@($l1, $t1, $l2, $t2, $ok, $cancel))
    $dlg.AcceptButton = $ok
    $dlg.CancelButton = $cancel

    if ($dlg.ShowDialog($form) -eq 'OK') {
        $script:Items.Insert($script:Index + 1, (New-Item2 $t1.Text.Trim() $t2.Text.Trim() $true))
        Update-View   # aktuální řádek zůstává, mění se jen celkový počet
    }
    $dlg.Dispose()
} 'Záznam se nepodařilo vložit.' })

$btnClear.Add_Click({ Invoke-Safe {
    $r = [System.Windows.Forms.MessageBox]::Show(
        'Opravdu zrušit načtený seznam a vymazat vstupní pole?' + "`n`n" + 'Nadpis a poznámky zůstanou beze změny.',
        'Zrušit seznam', 'YesNo', 'Question')
    if ($r -ne 'Yes') { return }
    $script:Items = New-Object System.Collections.Generic.List[object]
    $script:Index = 0
    $script:Done = $false
    Update-View          # odemkne vstupní pole
    $txtInput.Clear()
} 'Seznam se nepodařilo vymazat.' })

# ---------- Vkládání Ctrl + kliknutím do jiné aplikace ----------
# Globální sledování myši (funkce Windows, bez instalace). Při Ctrl + levém kliknutí
# v jiném okně než této aplikace se do kliknutého pole vloží aktuální údaj (nahradí jeho obsah).
$script:MiddleHookOk = $false
try {
    if (-not ('KlirencCtrlClick' -as [type])) {
        Add-Type -TypeDefinition @'
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
'@
    }
    $script:MiddleHookOk = [KlirencCtrlClick]::Start()
} catch {
    $script:MiddleHookOk = $false
}
if (-not $script:MiddleHookOk) {
    $chkMiddle.Checked = $false
    $chkMiddle.Enabled = $false
    $chkMiddle.Text = 'Vkládání Ctrl + kliknutím není na tomto počítači dostupné'
}

$chkMiddle.Add_CheckedChanged({
    if ($script:MiddleHookOk) { [KlirencCtrlClick]::Enabled = $chkMiddle.Checked }
})

# Po Ctrl + kliknutí: do schránky dát aktuální údaj a poslat Ctrl+V do okna, kde se kliklo
$middleTimer = New-Object System.Windows.Forms.Timer
$middleTimer.Interval = 40
$middleTimer.Add_Tick({
    try {
        if (-not $script:MiddleHookOk -or -not [KlirencCtrlClick]::Pending) { return }
        [KlirencCtrlClick]::Pending = $false
        if ($script:Done -or $script:Index -ge $script:Items.Count) { return }
        [System.Windows.Forms.Clipboard]::SetDataObject($script:Items[$script:Index].Value, $true, 5, 50)
        Start-Sleep -Milliseconds 60
        [KlirencCtrlClick]::ClickSelectAllPaste()
    } catch { }
})
if ($script:MiddleHookOk) { $middleTimer.Start() }

$form.Add_Shown({
    $split.SplitterDistance = [int]($form.ClientSize.Width * 0.6)
    Update-NotesCaption
    Update-View
})

try {
    [void]$form.ShowDialog()
} catch {
    Show-Error 'Aplikace narazila na neočekávaný problém a bude ukončena.'
} finally {
    try { $middleTimer.Stop(); if ($script:MiddleHookOk) { [KlirencCtrlClick]::Stop() } } catch { }
    $form.Dispose()
}
