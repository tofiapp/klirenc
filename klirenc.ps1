# klirenc - pracovní pomocník pro ruční procházení řádků z Excelu
# Spuštění: powershell.exe -NoProfile -STA -ExecutionPolicy Bypass -File .\klirenc.ps1

Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing
[System.Windows.Forms.Application]::EnableVisualStyles()

# ---------- Stav ----------
$script:Items = New-Object System.Collections.Generic.List[object]
$script:Index = 0
$script:Highlighting = $false

# Line = číslo řádku ve vstupních polích (-1 u záznamů přidaných přes Vložit mezi)
function New-Item2([string]$value, [string]$note, [int]$line = -1) {
    [pscustomobject]@{ Value = $value; Note = $note; Line = $line }
}

function Show-Info([string]$text)  { [void][System.Windows.Forms.MessageBox]::Show($text, 'klirenc', 'OK', 'Information') }
function Show-Warn([string]$text)  { [void][System.Windows.Forms.MessageBox]::Show($text, 'klirenc', 'OK', 'Warning') }
function Show-Error([string]$text) { [void][System.Windows.Forms.MessageBox]::Show($text, 'klirenc', 'OK', 'Error') }

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

# Spáruje dva samostatně vložené sloupce po řádcích.
# Vrací @{ Ok; Items; Error }. Při chybě nevrací žádné položky.
function ConvertFrom-Columns([string]$text1, [string]$text2) {
    $a = Get-ColumnLines $text1
    $b = Get-ColumnLines $text2
    if ($a.Count -eq 0) {
        return @{ Ok = $false; Error = 'Sloupec Údaj k ověření je prázdný. Vložte ho ze schránky.' }
    }
    if ($b.Count -gt $a.Count) {
        return @{ Ok = $false; Error = "Počty řádků nesouhlasí: Údaj k ověření má $($a.Count), Poznámka při NE má $($b.Count).`n`nZkopírujte oba sloupce ze stejného rozsahu řádků.`nNic nebylo načteno." }
    }
    if ($b.Count -lt $a.Count -and $b.Count -gt 0) {
        # kratší druhý sloupec je v pořádku jen tehdy, když chybějící konec tvoří prázdné buňky - to ale nepoznáme, proto upozorníme
        return @{ Ok = $false; Error = "Počty řádků nesouhlasí: Údaj k ověření má $($a.Count), Poznámka při NE má $($b.Count).`n`nPokud jsou poslední poznámky prázdné, označte v Excelu i je (stejný rozsah řádků jako u prvního sloupce).`nNic nebylo načteno." }
    }
    $result = New-Object System.Collections.Generic.List[object]
    for ($i = 0; $i -lt $a.Count; $i++) {
        $rowNo = $i + 1
        $v = $a[$i]
        $n = if ($i -lt $b.Count) { $b[$i] } else { '' }
        if ($v.Contains("`t") -or $n.Contains("`t")) {
            return @{ Ok = $false; Error = "Řádek $rowNo obsahuje více sloupců. Do každého pole vložte vždy jen jeden sloupec z Excelu.`nNic nebylo načteno." }
        }
        if ([string]::IsNullOrWhiteSpace($v) -and [string]::IsNullOrWhiteSpace($n)) { continue }
        if ([string]::IsNullOrWhiteSpace($v)) {
            return @{ Ok = $false; Error = "Řádek $rowNo má poznámku, ale prázdný Údaj k ověření.`nNic nebylo načteno." }
        }
        $result.Add((New-Item2 $v.Trim() $n.Trim() $i))
    }
    return @{ Ok = $true; Items = $result }
}

# ---------- GUI ----------
# Barevná paleta (světlý moderní vzhled)
function RGB([int]$r, [int]$g, [int]$b) { [System.Drawing.Color]::FromArgb($r, $g, $b) }
$cBg        = RGB 243 244 246   # pozadí okna
$cCard      = RGB 255 255 255   # karty
$cBorder    = RGB 226 232 240
$cText      = RGB 17 24 39
$cMuted     = RGB 107 114 128
$cAccent    = RGB 37 99 235     # modrá
$cAccentBg  = RGB 219 234 254   # světle modrá - zvýraznění aktuálního řádku
$cHeader    = RGB 30 41 59
$cYes       = RGB 22 163 74
$cNo        = RGB 220 38 38
$cWarn      = RGB 185 28 28
$cDone      = RGB 21 128 61
$cNeutral   = RGB 229 231 235

$fontBase  = New-Object System.Drawing.Font('Segoe UI', 10)
$fontSmall = New-Object System.Drawing.Font('Segoe UI', 9)
$fontHead  = New-Object System.Drawing.Font('Segoe UI Semibold', 10.5)
$fontBig   = New-Object System.Drawing.Font('Segoe UI Semibold', 26)
$fontBtn   = New-Object System.Drawing.Font('Segoe UI', 20, [System.Drawing.FontStyle]::Bold)
$fontState = New-Object System.Drawing.Font('Segoe UI Semibold', 11)
$fontMono  = New-Object System.Drawing.Font('Consolas', 10.5)
$fontMonoB = New-Object System.Drawing.Font('Consolas', 10.5, [System.Drawing.FontStyle]::Bold)

function Set-FlatButton($btn, $back, $fore) {
    $btn.FlatStyle = 'Flat'
    $btn.FlatAppearance.BorderSize = 0
    $btn.BackColor = $back
    $btn.ForeColor = $fore
    $btn.Cursor = [System.Windows.Forms.Cursors]::Hand
    $btn.UseVisualStyleBackColor = $false
}

function New-Button([string]$text, $back, $fore) {
    $b = New-Object System.Windows.Forms.Button
    $b.Text = $text
    $b.AutoSize = $true
    $b.Height = 34
    $b.Padding = New-Object System.Windows.Forms.Padding(10, 2, 10, 2)
    $b.Margin = New-Object System.Windows.Forms.Padding(0, 3, 8, 3)
    Set-FlatButton $b $back $fore
    $b
}

# Bílá "karta" s okrajem 1 px
function New-Card {
    $outer = New-Object System.Windows.Forms.Panel
    $outer.Dock = 'Fill'
    $outer.BackColor = $cBorder
    $outer.Padding = New-Object System.Windows.Forms.Padding(1)
    $outer.Margin = New-Object System.Windows.Forms.Padding(6)
    $inner = New-Object System.Windows.Forms.Panel
    $inner.Dock = 'Fill'
    $inner.BackColor = $cCard
    $inner.Padding = New-Object System.Windows.Forms.Padding(12, 10, 12, 10)
    $outer.Controls.Add($inner)
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
$form.Text = 'klirenc'
$form.Size = New-Object System.Drawing.Size(1200, 820)
$form.MinimumSize = New-Object System.Drawing.Size(950, 650)
$form.StartPosition = 'CenterScreen'
$form.Font = $fontBase
$form.BackColor = $cBg
$form.ForeColor = $cText

# --- Hlavička: název, stav a pozice ---
$header = New-Object System.Windows.Forms.TableLayoutPanel
$header.Dock = 'Top'
$header.Height = 52
$header.BackColor = $cHeader
$header.ColumnCount = 3
$header.Padding = New-Object System.Windows.Forms.Padding(16, 0, 16, 0)
[void]$header.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle('AutoSize')))
[void]$header.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle('Percent', 100)))
[void]$header.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle('AutoSize')))
$lblApp = New-Object System.Windows.Forms.Label
$lblApp.Text = 'klirenc'
$lblApp.AutoSize = $true
$lblApp.Anchor = 'Left'
$lblApp.Font = New-Object System.Drawing.Font('Segoe UI Semibold', 15)
$lblApp.ForeColor = [System.Drawing.Color]::White
$lblState = New-Object System.Windows.Forms.Label
$lblState.AutoSize = $true
$lblState.Anchor = 'Left'
$lblState.Margin = New-Object System.Windows.Forms.Padding(18, 0, 0, 0)
$lblState.Padding = New-Object System.Windows.Forms.Padding(10, 4, 10, 4)
$lblState.Font = $fontState
$lblPos = New-Object System.Windows.Forms.Label
$lblPos.AutoSize = $true
$lblPos.Anchor = 'Right'
$lblPos.Font = New-Object System.Drawing.Font('Segoe UI Semibold', 15)
$lblPos.ForeColor = [System.Drawing.Color]::White
$header.Controls.Add($lblApp, 0, 0)
$header.Controls.Add($lblState, 1, 0)
$header.Controls.Add($lblPos, 2, 0)

$split = New-Object System.Windows.Forms.SplitContainer
$split.Dock = 'Fill'
$split.Orientation = 'Vertical'
$split.BackColor = $cBg
$split.SplitterWidth = 6
$split.Padding = New-Object System.Windows.Forms.Padding(8)
$form.Controls.Add($split)
$form.Controls.Add($header)

# --- Levá část: import a práce ---
$left = New-Object System.Windows.Forms.TableLayoutPanel
$left.Dock = 'Fill'
$left.ColumnCount = 1
$left.RowCount = 3
[void]$left.RowStyles.Add((New-Object System.Windows.Forms.RowStyle('Percent', 50)))    # import
[void]$left.RowStyles.Add((New-Object System.Windows.Forms.RowStyle('Percent', 50)))    # aktuální hodnota
[void]$left.RowStyles.Add((New-Object System.Windows.Forms.RowStyle('Absolute', 120)))  # Ano/Ne + Vložit mezi
$split.Panel1.Controls.Add($left)

# Import: dva samostatně vkládané sloupce vedle sebe, každý s počtem řádků
$importCard = New-Card
$left.Controls.Add($importCard.Outer, 0, 0)

$importGrid = New-Object System.Windows.Forms.TableLayoutPanel
$importGrid.Dock = 'Fill'
$importGrid.ColumnCount = 2
$importGrid.RowCount = 4
[void]$importGrid.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle('Percent', 50)))
[void]$importGrid.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle('Percent', 50)))
[void]$importGrid.RowStyles.Add((New-Object System.Windows.Forms.RowStyle('Absolute', 26)))
[void]$importGrid.RowStyles.Add((New-Object System.Windows.Forms.RowStyle('Percent', 100)))
[void]$importGrid.RowStyles.Add((New-Object System.Windows.Forms.RowStyle('Absolute', 42)))
[void]$importGrid.RowStyles.Add((New-Object System.Windows.Forms.RowStyle('Absolute', 46)))
$importCard.Inner.Controls.Add($importGrid)

function New-ColumnInput([string]$caption, [int]$col) {
    $lbl = New-Object System.Windows.Forms.Label
    $lbl.Text = $caption
    $lbl.Dock = 'Fill'
    $lbl.Font = $fontHead
    # RichTextBox kvůli zvýraznění aktuálního řádku
    $txt = New-Object System.Windows.Forms.RichTextBox
    $txt.Multiline = $true
    $txt.ScrollBars = 'Both'
    $txt.WordWrap = $false
    $txt.DetectUrls = $false
    $txt.Dock = 'Fill'
    $txt.Font = $fontMono
    $txt.BorderStyle = 'None'
    $txt.BackColor = RGB 249 250 251
    $txt.Margin = New-Object System.Windows.Forms.Padding(0, 0, 8, 0)
    $bar = New-Object System.Windows.Forms.FlowLayoutPanel
    $bar.Dock = 'Fill'
    $bar.WrapContents = $false
    $btn = New-Button 'Načíst ze schránky' $cNeutral $cText
    $cnt = New-Object System.Windows.Forms.Label
    $cnt.AutoSize = $true
    $cnt.Padding = New-Object System.Windows.Forms.Padding(4, 9, 0, 0)
    $cnt.Font = $fontState
    $cnt.Text = 'Řádků: 0'
    $bar.Controls.AddRange(@($btn, $cnt))
    $importGrid.Controls.Add($lbl, $col, 0)
    $importGrid.Controls.Add($txt, $col, 1)
    $importGrid.Controls.Add($bar, $col, 2)
    @{ Text = $txt; Button = $btn; CountLabel = $cnt }
}
$col1 = New-ColumnInput 'Údaj k ověření' 0
$col2 = New-ColumnInput 'Poznámka při NE' 1

$importPanel = New-Object System.Windows.Forms.FlowLayoutPanel
$importPanel.Dock = 'Fill'
$importPanel.Padding = New-Object System.Windows.Forms.Padding(0, 4, 0, 0)
$btnLoad = New-Button 'Vytvořit seznam' $cAccent ([System.Drawing.Color]::White)
$btnLoad.Font = $fontState
$btnClear = New-Button 'Vymazat seznam' $cNeutral $cText
$importPanel.Controls.AddRange(@($btnLoad, $btnClear))
$importGrid.Controls.Add($importPanel, 0, 3)
$importGrid.SetColumnSpan($importPanel, 2)

# Aktuální hodnota
$currentCard = New-Card
$left.Controls.Add($currentCard.Outer, 0, 1)
$lblCurrentCap = New-Caption 'Aktuální údaj k ověření (zkopírováno do schránky)'
$lblCurrentCap.ForeColor = $cMuted
$lblCurrentCap.Font = $fontSmall
$lblCurrentNote = New-Object System.Windows.Forms.Label
$lblCurrentNote.Dock = 'Bottom'
$lblCurrentNote.Height = 30
$lblCurrentNote.Font = $fontBase
$lblCurrentNote.ForeColor = $cMuted
$lblCurrentNote.TextAlign = 'MiddleCenter'
$lblCurrentNote.AutoEllipsis = $true
$lblCurrent = New-Object System.Windows.Forms.Label
$lblCurrent.Dock = 'Fill'
$lblCurrent.Font = $fontBig
$lblCurrent.TextAlign = 'MiddleCenter'
$lblCurrent.AutoEllipsis = $true
$currentCard.Inner.BackColor = $cCard
$currentCard.Inner.Controls.Add($lblCurrent)
$currentCard.Inner.Controls.Add($lblCurrentNote)
$currentCard.Inner.Controls.Add($lblCurrentCap)

$actionPanel = New-Object System.Windows.Forms.TableLayoutPanel
$actionPanel.Dock = 'Fill'
$actionPanel.ColumnCount = 3
$actionPanel.Padding = New-Object System.Windows.Forms.Padding(3, 0, 3, 3)
[void]$actionPanel.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle('Percent', 40)))
[void]$actionPanel.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle('Percent', 40)))
[void]$actionPanel.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle('Percent', 20)))
$btnYes = New-Object System.Windows.Forms.Button
$btnYes.Text = 'ANO'
$btnYes.Dock = 'Fill'
$btnYes.Font = $fontBtn
$btnYes.Margin = New-Object System.Windows.Forms.Padding(3)
Set-FlatButton $btnYes $cYes ([System.Drawing.Color]::White)
$btnNo = New-Object System.Windows.Forms.Button
$btnNo.Text = 'NE'
$btnNo.Dock = 'Fill'
$btnNo.Font = $fontBtn
$btnNo.Margin = New-Object System.Windows.Forms.Padding(3)
Set-FlatButton $btnNo $cNo ([System.Drawing.Color]::White)
$btnInsert = New-Object System.Windows.Forms.Button
$btnInsert.Text = 'Vložit mezi'
$btnInsert.Dock = 'Fill'
$btnInsert.Font = $fontState
$btnInsert.Margin = New-Object System.Windows.Forms.Padding(3)
Set-FlatButton $btnInsert $cCard $cAccent
$btnInsert.FlatAppearance.BorderSize = 1
$btnInsert.FlatAppearance.BorderColor = $cAccent
$actionPanel.Controls.Add($btnYes, 0, 0)
$actionPanel.Controls.Add($btnNo, 1, 0)
$actionPanel.Controls.Add($btnInsert, 2, 0)
$left.Controls.Add($actionPanel, 0, 2)

# --- Pravá část: poznámky ---
$notesCard = New-Card
$split.Panel2.Controls.Add($notesCard.Outer)
$right = New-Object System.Windows.Forms.TableLayoutPanel
$right.Dock = 'Fill'
$right.ColumnCount = 1
$right.RowCount = 5
[void]$right.RowStyles.Add((New-Object System.Windows.Forms.RowStyle('Absolute', 26)))
[void]$right.RowStyles.Add((New-Object System.Windows.Forms.RowStyle('Absolute', 36)))
[void]$right.RowStyles.Add((New-Object System.Windows.Forms.RowStyle('Absolute', 30)))
[void]$right.RowStyles.Add((New-Object System.Windows.Forms.RowStyle('Percent', 100)))
[void]$right.RowStyles.Add((New-Object System.Windows.Forms.RowStyle('Absolute', 46)))
$notesCard.Inner.Controls.Add($right)

$lblTitle = New-Object System.Windows.Forms.Label
$lblTitle.Text = 'Nadpis poznámek'
$lblTitle.Dock = 'Fill'
$lblTitle.Font = $fontHead
$txtTitle = New-Object System.Windows.Forms.TextBox
$txtTitle.Dock = 'Fill'
$txtTitle.Font = New-Object System.Drawing.Font('Segoe UI', 11)
$lblNotes = New-Object System.Windows.Forms.Label
$lblNotes.Text = 'Poznámky'
$lblNotes.Dock = 'Fill'
$lblNotes.Font = $fontHead
$lblNotes.TextAlign = 'BottomLeft'
$txtNotes = New-Object System.Windows.Forms.TextBox
$txtNotes.Multiline = $true
$txtNotes.ScrollBars = 'Both'
$txtNotes.WordWrap = $false
$txtNotes.Dock = 'Fill'
$txtNotes.Font = $fontMono
$txtNotes.BorderStyle = 'None'
$txtNotes.BackColor = RGB 249 250 251
$btnSave = New-Button 'Uložit poznámky jako…' $cAccent ([System.Drawing.Color]::White)
$btnSave.Margin = New-Object System.Windows.Forms.Padding(0, 8, 0, 0)
$right.Controls.Add($lblTitle, 0, 0)
$right.Controls.Add($txtTitle, 0, 1)
$right.Controls.Add($lblNotes, 0, 2)
$right.Controls.Add($txtNotes, 0, 3)
$right.Controls.Add($btnSave, 0, 4)

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

# Zvýrazní ve vstupním poli řádek $line (-1 = nic nezvýrazňovat)
function Set-LineHighlight($box, [int]$line) {
    $box.SuspendLayout()
    $selStart = $box.SelectionStart
    $box.SelectAll()
    $box.SelectionBackColor = $box.BackColor
    $box.SelectionColor = $cText
    $box.SelectionFont = $fontMono
    if ($line -ge 0 -and $line -lt $box.Lines.Count) {
        $start = $box.GetFirstCharIndexFromLine($line)
        $len = $box.Lines[$line].Length
        $box.Select($start, [Math]::Max($len, 0))
        $box.SelectionBackColor = $cAccentBg
        $box.SelectionColor = $cAccent
        $box.SelectionFont = $fontMonoB
        $box.Select($start, 0)
        $box.ScrollToCaret()
    } else {
        $box.Select([Math]::Min($selStart, $box.TextLength), 0)
    }
    $box.ResumeLayout()
}

function Update-Highlight {
    $line = -1
    if ($script:Index -lt $script:Items.Count) { $line = $script:Items[$script:Index].Line }
    $script:Highlighting = $true
    try {
        Set-LineHighlight $col1.Text $line
        Set-LineHighlight $col2.Text $line
    } finally { $script:Highlighting = $false }
}

function Update-View {
    $count = $script:Items.Count
    $btnOn = ($count -gt 0 -and $script:Index -lt $count)
    $btnYes.Enabled = $btnOn; $btnNo.Enabled = $btnOn; $btnInsert.Enabled = $btnOn
    $btnYes.BackColor = if ($btnOn) { $cYes } else { $cNeutral }
    $btnNo.BackColor  = if ($btnOn) { $cNo }  else { $cNeutral }
    $lblCurrentNote.Text = ''
    if ($count -eq 0) {
        Set-State 'Bez seznamu' (RGB 71 85 105) ([System.Drawing.Color]::White)
        $lblPos.Text = '0 / 0'
        $lblCurrent.Text = 'Vložte sloupce a klikněte na Vytvořit seznam'
        $lblCurrent.ForeColor = $cMuted
    } elseif ($script:Index -ge $count) {
        Set-State 'Dokončeno' $cDone ([System.Drawing.Color]::White)
        $lblPos.Text = "$count / $count"
        $lblCurrent.Text = '✓ Hotovo – všechny řádky prošly'
        $lblCurrent.ForeColor = $cDone
    } else {
        if ($script:Index -eq 0) { Set-State 'Načteno' $cAccent ([System.Drawing.Color]::White) }
        else { Set-State 'Probíhá' (RGB 202 138 4) ([System.Drawing.Color]::White) }
        $lblPos.Text = "$($script:Index + 1) / $count"
        $item = $script:Items[$script:Index]
        $lblCurrent.Text = $item.Value
        $lblCurrent.ForeColor = $cText
        $lblCurrentNote.Text = if ($item.Note -ne '') { "Při NE se zapíše: $($item.Note)" } else { 'Při NE se nic nezapíše (poznámka je prázdná)' }
        if ($item.Line -lt 0) { $lblCurrentNote.Text += '   •   vloženo ručně' }
    }
    Update-Highlight
}

# Zobrazí aktuální řádek a zkopíruje jeho první hodnotu do schránky
function Show-Current {
    Update-View
    if ($script:Index -lt $script:Items.Count) {
        $value = $script:Items[$script:Index].Value
        try { [System.Windows.Forms.Clipboard]::SetText($value) }
        catch { Show-Warn 'Hodnotu se nepodařilo zkopírovat do schránky (schránka je možná obsazená jinou aplikací).' }
    }
}

function Move-Next {
    if ($script:Index -lt $script:Items.Count) { $script:Index++ }
    Show-Current
}

# Počet řádků u každého vloženého sloupce (koncové prázdné řádky se nepočítají)
function Update-ColumnCount($c) {
    $lines = Get-ColumnLines $c.Text.Text
    $filled = @($lines | Where-Object { -not [string]::IsNullOrWhiteSpace($_) }).Count
    $c.CountLabel.Text = if ($filled -eq $lines.Count) { "Řádků: $($lines.Count)" } else { "Řádků: $($lines.Count) (vyplněných $filled)" }
    $a = (Get-ColumnLines $col1.Text.Text).Count
    $b = (Get-ColumnLines $col2.Text.Text).Count
    $color = if ($a -gt 0 -and $b -gt 0 -and $a -ne $b) { [System.Drawing.Color]::Firebrick } else { [System.Drawing.Color]::Black }
    $col1.CountLabel.ForeColor = $color; $col2.CountLabel.ForeColor = $color
}

foreach ($c in @($col1, $col2)) {
    $c.Text.Tag = $c
    $c.Button.Tag = $c
    $c.Text.Add_TextChanged({ param($s, $e)
        if ($script:Highlighting) { return }
        Invoke-Safe { Update-ColumnCount $s.Tag } 'Počet řádků se nepodařilo spočítat.' })
    # Ctrl+V / Shift+Insert vloží jen čistý text (bez formátování z Excelu)
    $c.Text.Add_KeyDown({ param($s, $e)
        if (($e.Control -and $e.KeyCode -eq 'V') -or ($e.Shift -and $e.KeyCode -eq 'Insert')) {
            $e.SuppressKeyPress = $true
            $e.Handled = $true
            Invoke-Safe {
                if ([System.Windows.Forms.Clipboard]::ContainsText()) { $s.SelectedText = [System.Windows.Forms.Clipboard]::GetText() }
            } 'Vložení ze schránky se nezdařilo.'
        }
    })
    $c.Button.Add_Click({ param($s, $e) Invoke-Safe {
        $text = ''
        if ([System.Windows.Forms.Clipboard]::ContainsText()) { $text = [System.Windows.Forms.Clipboard]::GetText() }
        if ([string]::IsNullOrWhiteSpace($text)) {
            Show-Warn 'Schránka neobsahuje text. V Excelu označte sloupec, stiskněte Ctrl+C a zkuste to znovu.'
            return
        }
        $s.Tag.Text.Text = $text
    } 'Vložení ze schránky se nezdařilo.' })
}

$btnLoad.Add_Click({ Invoke-Safe {
    $parsed = ConvertFrom-Columns $col1.Text.Text $col2.Text.Text
    if (-not $parsed.Ok) { Show-Error $parsed.Error; return }   # stávající seznam ani poznámky se nemění
    if ($parsed.Items.Count -eq 0) { Show-Error 'Vstup neobsahuje žádný neprázdný řádek.'; return }

    $script:Items = $parsed.Items
    $script:Index = 0

    $title = $txtTitle.Text.Trim()
    if ($title -ne '') {
        $firstLine = ($txtNotes.Text -split "\r?\n", 2)[0]
        if ($firstLine -ne $title) {
            $rest = $txtNotes.Text
            $txtNotes.Text = $title + "`r`n" + $rest
        }
    }
    Show-Current
} 'Seznam se nepodařilo načíst. Zkuste znovu zkopírovat data z Excelu.' })

$btnYes.Add_Click({ Invoke-Safe { Move-Next } 'Přechod na další řádek se nezdařil.' })

$btnNo.Add_Click({ Invoke-Safe {
    if ($script:Index -ge $script:Items.Count) { return }
    $note = $script:Items[$script:Index].Note.Trim()
    if ($note -eq '') {
        Show-Warn 'Poznámka při NE je pro tento řádek prázdná. Nic se nezapíše, pokračuje se dalším řádkem.'
    } else {
        Add-NoteLine $note
    }
    Move-Next
} 'Zápis poznámky se nezdařil.' })

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
    $l2.Text = 'Poznámka při NE (volitelné):'; $l2.SetBounds(12, 70, 430, 22)
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
        $script:Items.Insert($script:Index + 1, (New-Item2 $t1.Text.Trim() $t2.Text.Trim()))
        Update-View   # aktuální řádek zůstává, mění se jen celkový počet
    }
    $dlg.Dispose()
} 'Záznam se nepodařilo vložit.' })

$btnClear.Add_Click({ Invoke-Safe {
    $r = [System.Windows.Forms.MessageBox]::Show(
        'Opravdu vymazat načtený seznam a importní vstup?' + "`n`n" + 'Nadpis a poznámky zůstanou beze změny.',
        'Vymazat seznam', 'YesNo', 'Question')
    if ($r -ne 'Yes') { return }
    $script:Items = New-Object System.Collections.Generic.List[object]
    $script:Index = 0
    $col1.Text.Clear()
    $col2.Text.Clear()
    Update-View
} 'Seznam se nepodařilo vymazat.' })

$btnSave.Add_Click({ Invoke-Safe {
    $sfd = New-Object System.Windows.Forms.SaveFileDialog
    $sfd.Filter = 'Textový soubor (*.txt)|*.txt'
    $sfd.DefaultExt = 'txt'
    $sfd.AddExtension = $true
    $sfd.FileName = 'poznamky.txt'
    if ($sfd.ShowDialog($form) -eq 'OK') {
        try {
            $enc = New-Object System.Text.UTF8Encoding($true)
            [System.IO.File]::WriteAllText($sfd.FileName, $txtNotes.Text, $enc)
            Show-Info "Poznámky byly uloženy do:`n$($sfd.FileName)"
        } catch {
            Show-Error 'Poznámky se nepodařilo uložit. Zkontrolujte, zda je soubor dostupný a máte do složky právo zápisu.'
        }
    }
    $sfd.Dispose()
} 'Ukládání se nezdařilo.' })

$form.Add_Shown({
    $split.SplitterDistance = [int]($form.ClientSize.Width * 0.6)
    Update-View
})

try {
    [void]$form.ShowDialog()
} catch {
    Show-Error 'Aplikace narazila na neočekávaný problém a bude ukončena.'
} finally {
    $form.Dispose()
}
