# klirenc - pracovní pomocník pro ruční procházení řádků z Excelu
# Spuštění: powershell.exe -NoProfile -STA -ExecutionPolicy Bypass -File .\klirenc.ps1

Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing
[System.Windows.Forms.Application]::EnableVisualStyles()

# ---------- Stav ----------
$script:Items = New-Object System.Collections.Generic.List[object]
$script:Index = 0

function New-Item2([string]$value, [string]$note) {
    [pscustomobject]@{ Value = $value; Note = $note }
}

function Show-Info([string]$text)  { [void][System.Windows.Forms.MessageBox]::Show($text, 'klirenc', 'OK', 'Information') }
function Show-Warn([string]$text)  { [void][System.Windows.Forms.MessageBox]::Show($text, 'klirenc', 'OK', 'Warning') }
function Show-Error([string]$text) { [void][System.Windows.Forms.MessageBox]::Show($text, 'klirenc', 'OK', 'Error') }

# Bezpečné spuštění obsluhy události - uživateli se nezobrazí technická chyba
function Invoke-Safe([scriptblock]$action, [string]$message) {
    try { & $action } catch { Show-Error $message }
}

# ---------- Parsování vstupu ----------
# Vrací @{ Ok; Items; Error }. Při chybě nevrací žádné položky.
function ConvertFrom-PastedText([string]$text) {
    $result = New-Object System.Collections.Generic.List[object]
    if ([string]::IsNullOrEmpty($text)) {
        return @{ Ok = $false; Error = 'Vstup je prázdný. Zkopírujte v Excelu dva sloupce a zkuste to znovu.' }
    }
    $lines = $text -split "\r\n|\n|\r"
    $lineNo = 0
    foreach ($line in $lines) {
        $lineNo++
        if ([string]::IsNullOrWhiteSpace($line)) { continue }
        $parts = $line.Split("`t")
        if ($parts.Count -ne 2) {
            return @{ Ok = $false; Error = "Řádek $lineNo nemá přesně dva sloupce oddělené tabulátorem (má $($parts.Count)).`n`nZkopírujte z Excelu právě dva sloupce: Údaj k ověření a Poznámka při NE.`nNic nebylo načteno." }
        }
        $value = $parts[0].Trim()
        if ($value -eq '') {
            return @{ Ok = $false; Error = "Řádek $lineNo má prázdný první sloupec (Údaj k ověření).`nNic nebylo načteno." }
        }
        $result.Add((New-Item2 $value $parts[1].Trim()))
    }
    if ($result.Count -eq 0) {
        return @{ Ok = $false; Error = 'Vstup neobsahuje žádný neprázdný řádek.' }
    }
    return @{ Ok = $true; Items = $result }
}

# ---------- GUI ----------
$fontBase  = New-Object System.Drawing.Font('Segoe UI', 10)
$fontBig   = New-Object System.Drawing.Font('Segoe UI', 22, [System.Drawing.FontStyle]::Bold)
$fontBtn   = New-Object System.Drawing.Font('Segoe UI', 20, [System.Drawing.FontStyle]::Bold)
$fontState = New-Object System.Drawing.Font('Segoe UI', 11, [System.Drawing.FontStyle]::Bold)

$form = New-Object System.Windows.Forms.Form
$form.Text = 'klirenc - procházení řádků'
$form.Size = New-Object System.Drawing.Size(1100, 760)
$form.MinimumSize = New-Object System.Drawing.Size(900, 600)
$form.StartPosition = 'CenterScreen'
$form.Font = $fontBase
$form.KeyPreview = $true

$split = New-Object System.Windows.Forms.SplitContainer
$split.Dock = 'Fill'
$split.Orientation = 'Vertical'
$form.Controls.Add($split)

# --- Levá část: import a práce ---
$left = New-Object System.Windows.Forms.TableLayoutPanel
$left.Dock = 'Fill'
$left.ColumnCount = 1
$left.RowCount = 6
$left.Padding = New-Object System.Windows.Forms.Padding(8)
[void]$left.RowStyles.Add((New-Object System.Windows.Forms.RowStyle('Absolute', 26)))   # popisek importu
[void]$left.RowStyles.Add((New-Object System.Windows.Forms.RowStyle('Percent', 45)))    # vstup
[void]$left.RowStyles.Add((New-Object System.Windows.Forms.RowStyle('Absolute', 44)))   # tlačítka importu
[void]$left.RowStyles.Add((New-Object System.Windows.Forms.RowStyle('Absolute', 34)))   # stav + pozice
[void]$left.RowStyles.Add((New-Object System.Windows.Forms.RowStyle('Percent', 55)))    # aktuální hodnota
[void]$left.RowStyles.Add((New-Object System.Windows.Forms.RowStyle('Absolute', 130)))  # Ano/Ne + Vložit mezi
$split.Panel1.Controls.Add($left)

$lblInput = New-Object System.Windows.Forms.Label
$lblInput.Text = 'Import (dva sloupce z Excelu: Údaj k ověření <TAB> Poznámka při NE):'
$lblInput.Dock = 'Fill'
$left.Controls.Add($lblInput, 0, 0)

$txtInput = New-Object System.Windows.Forms.TextBox
$txtInput.Multiline = $true
$txtInput.ScrollBars = 'Both'
$txtInput.WordWrap = $false
$txtInput.AcceptsTab = $true
$txtInput.Dock = 'Fill'
$txtInput.Font = New-Object System.Drawing.Font('Consolas', 10)
$left.Controls.Add($txtInput, 0, 1)

$importPanel = New-Object System.Windows.Forms.FlowLayoutPanel
$importPanel.Dock = 'Fill'
$btnLoad = New-Object System.Windows.Forms.Button
$btnLoad.Text = 'Načíst ze schránky'
$btnLoad.AutoSize = $true
$btnLoad.Height = 34
$btnClear = New-Object System.Windows.Forms.Button
$btnClear.Text = 'Vymazat seznam'
$btnClear.AutoSize = $true
$btnClear.Height = 34
$importPanel.Controls.AddRange(@($btnLoad, $btnClear))
$left.Controls.Add($importPanel, 0, 2)

$statusPanel = New-Object System.Windows.Forms.TableLayoutPanel
$statusPanel.Dock = 'Fill'
$statusPanel.ColumnCount = 2
[void]$statusPanel.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle('Percent', 60)))
[void]$statusPanel.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle('Percent', 40)))
$lblState = New-Object System.Windows.Forms.Label
$lblState.Dock = 'Fill'
$lblState.Font = $fontState
$lblState.TextAlign = 'MiddleLeft'
$lblPos = New-Object System.Windows.Forms.Label
$lblPos.Dock = 'Fill'
$lblPos.Font = $fontState
$lblPos.TextAlign = 'MiddleRight'
$statusPanel.Controls.Add($lblState, 0, 0)
$statusPanel.Controls.Add($lblPos, 1, 0)
$left.Controls.Add($statusPanel, 0, 3)

$lblCurrent = New-Object System.Windows.Forms.Label
$lblCurrent.Dock = 'Fill'
$lblCurrent.Font = $fontBig
$lblCurrent.TextAlign = 'MiddleCenter'
$lblCurrent.BorderStyle = 'FixedSingle'
$lblCurrent.BackColor = [System.Drawing.Color]::White
$lblCurrent.AutoEllipsis = $true
$left.Controls.Add($lblCurrent, 0, 4)

$actionPanel = New-Object System.Windows.Forms.TableLayoutPanel
$actionPanel.Dock = 'Fill'
$actionPanel.ColumnCount = 3
[void]$actionPanel.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle('Percent', 40)))
[void]$actionPanel.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle('Percent', 40)))
[void]$actionPanel.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle('Percent', 20)))
$btnYes = New-Object System.Windows.Forms.Button
$btnYes.Text = 'ANO'
$btnYes.Dock = 'Fill'
$btnYes.Font = $fontBtn
$btnYes.BackColor = [System.Drawing.Color]::FromArgb(46, 160, 67)
$btnYes.ForeColor = [System.Drawing.Color]::White
$btnYes.FlatStyle = 'Flat'
$btnNo = New-Object System.Windows.Forms.Button
$btnNo.Text = 'NE'
$btnNo.Dock = 'Fill'
$btnNo.Font = $fontBtn
$btnNo.BackColor = [System.Drawing.Color]::FromArgb(207, 34, 46)
$btnNo.ForeColor = [System.Drawing.Color]::White
$btnNo.FlatStyle = 'Flat'
$btnInsert = New-Object System.Windows.Forms.Button
$btnInsert.Text = 'Vložit mezi'
$btnInsert.Dock = 'Fill'
$actionPanel.Controls.Add($btnYes, 0, 0)
$actionPanel.Controls.Add($btnNo, 1, 0)
$actionPanel.Controls.Add($btnInsert, 2, 0)
$left.Controls.Add($actionPanel, 0, 5)

# --- Pravá část: poznámky ---
$right = New-Object System.Windows.Forms.TableLayoutPanel
$right.Dock = 'Fill'
$right.ColumnCount = 1
$right.RowCount = 5
$right.Padding = New-Object System.Windows.Forms.Padding(8)
[void]$right.RowStyles.Add((New-Object System.Windows.Forms.RowStyle('Absolute', 26)))
[void]$right.RowStyles.Add((New-Object System.Windows.Forms.RowStyle('Absolute', 34)))
[void]$right.RowStyles.Add((New-Object System.Windows.Forms.RowStyle('Absolute', 26)))
[void]$right.RowStyles.Add((New-Object System.Windows.Forms.RowStyle('Percent', 100)))
[void]$right.RowStyles.Add((New-Object System.Windows.Forms.RowStyle('Absolute', 44)))
$split.Panel2.Controls.Add($right)

$lblTitle = New-Object System.Windows.Forms.Label
$lblTitle.Text = 'Nadpis poznámek:'
$lblTitle.Dock = 'Fill'
$txtTitle = New-Object System.Windows.Forms.TextBox
$txtTitle.Dock = 'Fill'
$lblNotes = New-Object System.Windows.Forms.Label
$lblNotes.Text = 'Poznámky:'
$lblNotes.Dock = 'Fill'
$txtNotes = New-Object System.Windows.Forms.TextBox
$txtNotes.Multiline = $true
$txtNotes.ScrollBars = 'Both'
$txtNotes.WordWrap = $false
$txtNotes.Dock = 'Fill'
$txtNotes.Font = New-Object System.Drawing.Font('Consolas', 10)
$btnSave = New-Object System.Windows.Forms.Button
$btnSave.Text = 'Uložit poznámky jako…'
$btnSave.AutoSize = $true
$btnSave.Height = 34
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

function Update-View {
    $count = $script:Items.Count
    if ($count -eq 0) {
        $lblState.Text = 'Stav: bez seznamu'
        $lblState.ForeColor = [System.Drawing.Color]::Gray
        $lblPos.Text = '0 / 0'
        $lblCurrent.Text = 'Načtěte seznam ze schránky'
        $lblCurrent.ForeColor = [System.Drawing.Color]::Gray
        $btnYes.Enabled = $false; $btnNo.Enabled = $false; $btnInsert.Enabled = $false
        return
    }
    if ($script:Index -ge $count) {
        $lblState.Text = 'Stav: dokončeno'
        $lblState.ForeColor = [System.Drawing.Color]::DarkGreen
        $lblPos.Text = "$count / $count"
        $lblCurrent.Text = 'Hotovo - všechny řádky prošly.'
        $lblCurrent.ForeColor = [System.Drawing.Color]::DarkGreen
        $btnYes.Enabled = $false; $btnNo.Enabled = $false; $btnInsert.Enabled = $false
        return
    }
    $lblState.Text = if ($script:Index -eq 0) { 'Stav: načteno' } else { 'Stav: probíhá' }
    $lblState.ForeColor = [System.Drawing.Color]::DarkBlue
    $lblPos.Text = "$($script:Index + 1) / $count"
    $item = $script:Items[$script:Index]
    $lblCurrent.Text = $item.Value
    $lblCurrent.ForeColor = [System.Drawing.Color]::Black
    $btnYes.Enabled = $true; $btnNo.Enabled = $true; $btnInsert.Enabled = $true
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

$btnLoad.Add_Click({ Invoke-Safe {
    $text = ''
    if ([System.Windows.Forms.Clipboard]::ContainsText()) { $text = [System.Windows.Forms.Clipboard]::GetText() }
    if ([string]::IsNullOrWhiteSpace($text)) {
        Show-Warn 'Schránka neobsahuje text. V Excelu označte dva sloupce, stiskněte Ctrl+C a zkuste to znovu.'
        return
    }
    $parsed = ConvertFrom-PastedText $text
    if (-not $parsed.Ok) { Show-Error $parsed.Error; return }   # stávající seznam ani poznámky se nemění

    $txtInput.Text = $text
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
    $txtInput.Clear()
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
