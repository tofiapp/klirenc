# klirenc - pracovní pomocník pro ruční procházení řádků z Excelu
# Spuštění: powershell.exe -NoProfile -STA -ExecutionPolicy Bypass -File .\klirenc.ps1

Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing
[System.Windows.Forms.Application]::EnableVisualStyles()

# ---------- Stav ----------
$script:Items = New-Object System.Collections.Generic.List[object]
$script:Index = 0
$script:Highlighting = $false
$script:Done = $false          # všechny řádky rozhodnuty a zobrazuje se dokončení
$script:PastedCols = 0      # kolik sloupců mělo poslední vložení (před ořezáním)
$script:NotesFile = $null   # soubor, do kterého se poznámky ukládají (po Otevřít / Uložit jako)

# Status: '' = nerozhodnuto, 'keep' = Ponechat, 'del' = Vymazat (poznámka zapsána)
# Manual = záznam přidaný přes Vložit mezi
function New-Item2([string]$value, [string]$note, [bool]$manual = $false) {
    [pscustomobject]@{ Value = $value; Note = $note; Status = ''; Manual = $manual }
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
$left.RowCount = 4
[void]$left.RowStyles.Add((New-Object System.Windows.Forms.RowStyle('Percent', 50)))    # import
[void]$left.RowStyles.Add((New-Object System.Windows.Forms.RowStyle('Percent', 50)))    # aktuální hodnota
[void]$left.RowStyles.Add((New-Object System.Windows.Forms.RowStyle('Absolute', 120)))  # Ponechat/Vymazat + Vložit mezi
[void]$left.RowStyles.Add((New-Object System.Windows.Forms.RowStyle('Absolute', 40)))   # malé tlačítko ?
$split.Panel1.Controls.Add($left)

# Import: jedno pole, do kterého se vloží řádky z Excelu (použije se 1. a poslední sloupec)
$importCard = New-Card
$left.Controls.Add($importCard.Outer, 0, 0)

$importGrid = New-Object System.Windows.Forms.TableLayoutPanel
$importGrid.Dock = 'Fill'
$importGrid.ColumnCount = 1
$importGrid.RowCount = 3
[void]$importGrid.RowStyles.Add((New-Object System.Windows.Forms.RowStyle('Absolute', 26)))
[void]$importGrid.RowStyles.Add((New-Object System.Windows.Forms.RowStyle('Percent', 100)))
[void]$importGrid.RowStyles.Add((New-Object System.Windows.Forms.RowStyle('Absolute', 46)))
$importCard.Inner.Controls.Add($importGrid)

$lblInput = New-Object System.Windows.Forms.Label
$lblInput.Text = 'Řádky z Excelu (Ctrl+V) – použije se 1. sloupec (Údaj k ověření) a poslední sloupec (Poznámka při Vymazat)'
$lblInput.Dock = 'Fill'
$lblInput.Font = $fontHead
$lblInput.AutoEllipsis = $true
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
$txtInput.BackColor = RGB 249 250 251

$importPanel = New-Object System.Windows.Forms.FlowLayoutPanel
$importPanel.Dock = 'Fill'
$importPanel.WrapContents = $false
$importPanel.Padding = New-Object System.Windows.Forms.Padding(0, 4, 0, 0)
$btnLoad = New-Button 'Vytvořit seznam' $cAccent ([System.Drawing.Color]::White)
$btnLoad.Font = $fontState
$btnClear = New-Button 'Zrušit seznam' $cNeutral $cText
$lblCount = New-Object System.Windows.Forms.Label
$lblCount.AutoSize = $true
$lblCount.Padding = New-Object System.Windows.Forms.Padding(8, 9, 0, 0)
$lblCount.Font = $fontState
$lblCount.Text = 'Řádků: 0'
$importPanel.Controls.AddRange(@($btnLoad, $btnClear, $lblCount))
$importGrid.Controls.Add($lblInput, 0, 0)
$importGrid.Controls.Add($txtInput, 0, 1)
$importGrid.Controls.Add($importPanel, 0, 2)

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
$actionPanel.ColumnCount = 4
$actionPanel.Padding = New-Object System.Windows.Forms.Padding(3, 0, 3, 3)
[void]$actionPanel.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle('Absolute', 70)))
[void]$actionPanel.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle('Percent', 40)))
[void]$actionPanel.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle('Percent', 40)))
[void]$actionPanel.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle('Percent', 20)))
$btnYes = New-Object System.Windows.Forms.Button
$btnYes.Text = 'Ponechat'
$btnYes.Dock = 'Fill'
$btnYes.Font = $fontBtn
$btnYes.Margin = New-Object System.Windows.Forms.Padding(3)
Set-FlatButton $btnYes $cYes ([System.Drawing.Color]::White)
$btnNo = New-Object System.Windows.Forms.Button
$btnNo.Text = 'Vymazat'
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
# Ruční procházení seznamu
$navPanel = New-Object System.Windows.Forms.TableLayoutPanel
$navPanel.Dock = 'Fill'
$navPanel.RowCount = 2
$navPanel.Margin = New-Object System.Windows.Forms.Padding(0)
[void]$navPanel.RowStyles.Add((New-Object System.Windows.Forms.RowStyle('Percent', 50)))
[void]$navPanel.RowStyles.Add((New-Object System.Windows.Forms.RowStyle('Percent', 50)))
$btnUp = New-Object System.Windows.Forms.Button
$btnUp.Text = '▲'
$btnUp.Dock = 'Fill'
$btnUp.Font = $fontState
$btnUp.Margin = New-Object System.Windows.Forms.Padding(3)
Set-FlatButton $btnUp $cNeutral $cText
$btnDown = New-Object System.Windows.Forms.Button
$btnDown.Text = '▼'
$btnDown.Dock = 'Fill'
$btnDown.Font = $fontState
$btnDown.Margin = New-Object System.Windows.Forms.Padding(3)
Set-FlatButton $btnDown $cNeutral $cText
$navPanel.Controls.Add($btnUp, 0, 0)
$navPanel.Controls.Add($btnDown, 0, 1)
$actionPanel.Controls.Add($navPanel, 0, 0)
$actionPanel.Controls.Add($btnYes, 1, 0)
$actionPanel.Controls.Add($btnNo, 2, 0)
$actionPanel.Controls.Add($btnInsert, 3, 0)
$left.Controls.Add($actionPanel, 0, 2)

# Malé tlačítko „?“ – nejistý řádek, vrátit se k němu později
$unsurePanel = New-Object System.Windows.Forms.FlowLayoutPanel
$unsurePanel.Dock = 'Fill'
$unsurePanel.Padding = New-Object System.Windows.Forms.Padding(76, 0, 0, 0)
$btnUnsure = New-Button '?  Nevím – vrátit se později' $cNeutral $cText
$unsurePanel.Controls.Add($btnUnsure)
$left.Controls.Add($unsurePanel, 0, 3)

# --- Pravá část: poznámky ---
$notesCard = New-Card
$split.Panel2.Controls.Add($notesCard.Outer)
$right = New-Object System.Windows.Forms.TableLayoutPanel
$right.Dock = 'Fill'
$right.ColumnCount = 1
$right.RowCount = 5
[void]$right.RowStyles.Add((New-Object System.Windows.Forms.RowStyle('Absolute', 26)))
[void]$right.RowStyles.Add((New-Object System.Windows.Forms.RowStyle('Absolute', 40)))
[void]$right.RowStyles.Add((New-Object System.Windows.Forms.RowStyle('Absolute', 30)))
[void]$right.RowStyles.Add((New-Object System.Windows.Forms.RowStyle('Percent', 100)))
[void]$right.RowStyles.Add((New-Object System.Windows.Forms.RowStyle('Absolute', 46)))
$notesCard.Inner.Controls.Add($right)

$lblTitle = New-Object System.Windows.Forms.Label
$lblTitle.Text = 'Nadpis poznámek'
$lblTitle.Dock = 'Fill'
$lblTitle.Font = $fontHead
$titleRow = New-Object System.Windows.Forms.TableLayoutPanel
$titleRow.Dock = 'Fill'
$titleRow.ColumnCount = 2
[void]$titleRow.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle('Percent', 100)))
[void]$titleRow.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle('AutoSize')))
$txtTitle = New-Object System.Windows.Forms.TextBox
$txtTitle.Dock = 'Fill'
$txtTitle.Font = New-Object System.Drawing.Font('Segoe UI', 11)
$txtTitle.Margin = New-Object System.Windows.Forms.Padding(0, 4, 8, 0)
$btnAddTitle = New-Button '+ Přidat nadpis' $cNeutral $cText
$titleRow.Controls.Add($txtTitle, 0, 0)
$titleRow.Controls.Add($btnAddTitle, 1, 0)
$lblNotes = New-Object System.Windows.Forms.Label
$lblNotes.Text = 'Poznámky'
$lblNotes.Dock = 'Fill'
$lblNotes.Font = $fontHead
$lblNotes.TextAlign = 'BottomLeft'
$lblNotes.AutoEllipsis = $true
$txtNotes = New-Object System.Windows.Forms.TextBox
$txtNotes.Multiline = $true
$txtNotes.ScrollBars = 'Both'
$txtNotes.WordWrap = $false
$txtNotes.Dock = 'Fill'
$txtNotes.Font = $fontMono
$txtNotes.BorderStyle = 'None'
$txtNotes.BackColor = RGB 249 250 251
$saveRow = New-Object System.Windows.Forms.FlowLayoutPanel
$saveRow.Dock = 'Fill'
$saveRow.WrapContents = $false
$saveRow.Padding = New-Object System.Windows.Forms.Padding(0, 6, 0, 0)
$btnOpen = New-Button 'Otevřít soubor…' $cNeutral $cText
$btnSave = New-Button 'Uložit' $cAccent ([System.Drawing.Color]::White)
$btnSaveAs = New-Button 'Uložit jako…' $cNeutral $cText
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

# Po vytvoření seznamu zobrazí horní pole seznam se stavem každého řádku:
#   ✓ = Ponechat, ✗ = Vymazat (zapsáno do poznámek), ? = vrátit se později, prázdné = ještě nerozhodnuto
function Get-ListText {
    $sb = New-Object System.Text.StringBuilder
    for ($i = 0; $i -lt $script:Items.Count; $i++) {
        $it = $script:Items[$i]
        $mark = switch ($it.Status) { 'keep' { '✓' } 'del' { '✗' } 'unsure' { '?' } default { ' ' } }
        $line = "$mark  $($it.Value)`t$($it.Note)"
        if ($it.Manual) { $line += '   (vloženo ručně)' }
        if ($i -gt 0) { [void]$sb.Append("`n") }
        [void]$sb.Append($line)
    }
    $sb.ToString()
}

function Update-ListBox {
    $box = $txtInput
    $script:Highlighting = $true
    try {
        $box.SuspendLayout()
        if ($script:Items.Count -eq 0) {
            $box.ReadOnly = $false
            $box.SelectAll()
            $box.SelectionBackColor = $box.BackColor
            $box.SelectionColor = $cText
            $box.SelectionFont = $fontMono
            $box.Select(0, 0)
            return
        }
        $box.ReadOnly = $true
        $text = Get-ListText
        if ($box.Text -ne $text) { $box.Text = $text }
        $box.SelectAll()
        $box.SelectionBackColor = $box.BackColor
        $box.SelectionColor = $cText
        $box.SelectionFont = $fontMono
        $lines = $box.Lines
        for ($i = 0; $i -lt $script:Items.Count -and $i -lt $lines.Count; $i++) {
            $st = $script:Items[$i].Status
            if ($st -eq '' -and $i -ne $script:Index) { continue }
            $box.Select($box.GetFirstCharIndexFromLine($i), $lines[$i].Length)
            if ($st -eq 'keep') { $box.SelectionColor = $cYes }
            elseif ($st -eq 'del') { $box.SelectionColor = $cNo }
            elseif ($st -eq 'unsure') { $box.SelectionColor = RGB 202 138 4 }
            if ($i -eq $script:Index -and -not $script:Done) {
                $box.SelectionBackColor = $cAccentBg
                $box.SelectionFont = $fontMonoB
                if ($st -eq '') { $box.SelectionColor = $cAccent }
            }
        }
        $box.Select($box.GetFirstCharIndexFromLine([Math]::Min($script:Index, $script:Items.Count - 1)), 0)
        $box.ScrollToCaret()
    } finally {
        $box.ResumeLayout()
        $script:Highlighting = $false
    }
}

function Update-View {
    $count = $script:Items.Count
    $decided = @($script:Items | Where-Object { $_.Status -eq 'keep' -or $_.Status -eq 'del' }).Count
    $btnOn = ($count -gt 0 -and -not $script:Done)
    $btnYes.Enabled = $btnOn; $btnNo.Enabled = $btnOn; $btnInsert.Enabled = $btnOn; $btnUnsure.Enabled = $btnOn
    $btnUp.Enabled = ($count -gt 0); $btnDown.Enabled = ($count -gt 0)
    $btnYes.BackColor = if ($btnOn) { $cYes } else { $cNeutral }
    $btnNo.BackColor  = if ($btnOn) { $cNo }  else { $cNeutral }
    $lblCurrentNote.Text = ''
    if ($count -eq 0) {
        Set-State 'Bez seznamu' (RGB 71 85 105) ([System.Drawing.Color]::White)
        $lblPos.Text = '0 / 0'
        $lblCurrent.Text = 'Vložte řádky z Excelu a klikněte na Vytvořit seznam'
        $lblCurrent.ForeColor = $cMuted
    } elseif ($script:Done) {
        Set-State "Dokončeno  ($decided / $count)" $cDone ([System.Drawing.Color]::White)
        $lblPos.Text = "$count / $count"
        $lblCurrent.Text = '✓ Hotovo – všechny řádky jsou rozhodnuté'
        $lblCurrent.ForeColor = $cDone
        $lblCurrentNote.Text = 'Šipkami ▲ ▼ nebo kliknutím do seznamu se můžete k libovolnému řádku vrátit.'
    } else {
        if ($decided -eq 0) { Set-State "Načteno  (hotovo 0 / $count)" $cAccent ([System.Drawing.Color]::White) }
        else { Set-State "Probíhá  (hotovo $decided / $count)" (RGB 202 138 4) ([System.Drawing.Color]::White) }
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

# Přidá nadpis na konec poznámek (oddělený prázdným řádkem) a vyprázdní pole nadpisu
function Add-Title {
    $title = $txtTitle.Text.Trim()
    if ($title -eq '') { return $false }
    $t = $txtNotes.Text.TrimEnd()
    $txtNotes.Text = if ($t -eq '') { $title + "`r`n" } else { $t + "`r`n`r`n" + $title + "`r`n" }
    $txtNotes.SelectionStart = $txtNotes.Text.Length
    $txtNotes.ScrollToCaret()
    $txtTitle.Clear()
    return $true
}

function Update-NotesCaption {
    $lblNotes.Text = if ($script:NotesFile) { "Poznámky – $([System.IO.Path]::GetFileName($script:NotesFile))" } else { 'Poznámky (zatím neuloženo)' }
}

function Save-Notes([string]$path) {
    try {
        $enc = New-Object System.Text.UTF8Encoding($true)
        [System.IO.File]::WriteAllText($path, $txtNotes.Text, $enc)
        $script:NotesFile = $path
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
    $sfd.FileName = if ($script:NotesFile) { [System.IO.Path]::GetFileName($script:NotesFile) } else { 'poznamky.txt' }
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
    if ($rows.Count -gt 0) {
        $cols = $rows[0].Split("`t").Count
        if ($cols -ge 2 -and $script:PastedCols -gt 2) {
            $text = $text + "   •   vloženo sloupců: $($script:PastedCols), ponechán 1. a $($script:PastedCols)."
        } elseif ($cols -ge 2) {
            $text = $text + "   •   sloupců: $cols (použije se 1. a $cols.)"
        } else {
            $text = $text + '   •   jen 1 sloupec!'
            $lblCount.ForeColor = $cWarn
        }
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
    [void](Add-Title)   # vyplněný nadpis se přidá do poznámek
    Show-Current
} 'Seznam se nepodařilo načíst. Zkuste znovu zkopírovat data z Excelu.' })

$btnAddTitle.Add_Click({ Invoke-Safe {
    if (-not (Add-Title)) { Show-Warn 'Nejdřív napište nadpis.' }
} 'Nadpis se nepodařilo přidat.' })

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
            Update-NotesCaption
        } catch {
            Show-Error 'Soubor se nepodařilo otevřít. Zkontrolujte, zda existuje a není otevřený jinou aplikací.'
        }
    }
    $ofd.Dispose()
} 'Soubor se nepodařilo otevřít.' })

$btnSave.Add_Click({ Invoke-Safe {
    if ($script:NotesFile) { Save-Notes $script:NotesFile } else { Save-NotesAs }
} 'Ukládání se nezdařilo.' })

$btnSaveAs.Add_Click({ Invoke-Safe { Save-NotesAs } 'Ukládání se nezdařilo.' })

$btnYes.Add_Click({ Invoke-Safe { Set-Decision 'keep' } 'Přechod na další řádek se nezdařil.' })
$btnUp.Add_Click({ Invoke-Safe { Move-By -1 } 'Přechod na řádek se nezdařil.' })
$btnDown.Add_Click({ Invoke-Safe { Move-By 1 } 'Přechod na řádek se nezdařil.' })

# Kliknutím na řádek v seznamu nahoře se na něj přejde
$txtInput.Add_MouseUp({ Invoke-Safe {
    if ($script:Items.Count -eq 0) { return }
    if ($txtInput.SelectionLength -gt 0) { return }
    $line = $txtInput.GetLineFromCharIndex($txtInput.SelectionStart)
    if ($line -ge 0 -and $line -lt $script:Items.Count) {
        $script:Index = $line
        $script:Done = $false
        Show-Current
    }
} 'Přechod na řádek se nezdařil.' })

# Šipky na klávesnici v horním seznamu přepínají řádky
$txtInput.Add_KeyDown({ param($s, $e)
    if ($script:Items.Count -eq 0) { return }
    $delta = switch ($e.KeyCode) { 'Up' { -1 } 'Down' { 1 } 'PageUp' { -10 } 'PageDown' { 10 } default { 0 } }
    if ($delta -ne 0) {
        $e.Handled = $true
        $e.SuppressKeyPress = $true
        Invoke-Safe { Move-By $delta } 'Přechod na řádek se nezdařil.'
    }
})

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
    $form.Dispose()
}
