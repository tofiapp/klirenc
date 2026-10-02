# Kontrola Clearance - pracovní pomocník pro ruční procházení řádků z Excelu
# Spuštění: powershell.exe -NoProfile -STA -ExecutionPolicy Bypass -File .\klirenc.ps1
# Okno je ve WPF (součást Windows) - písmo se vykresluje hladce i při zvětšeném zobrazení.

$script:AppVersion = '21'   # zobrazuje se v titulku okna - podle ní se pozná, která verze běží

Add-Type -AssemblyName PresentationFramework
Add-Type -AssemblyName PresentationCore
Add-Type -AssemblyName WindowsBase
Add-Type -AssemblyName System.Drawing          # jen pro kreslení ikony KC

# Vlastní identita procesu pro hlavní panel Windows: okno se neseskupí s ostatními okny PowerShellu
# a na liště se zobrazí ikona KC. Musí proběhnout před vytvořením okna.
try {
    if (-not ('KcTaskbar' -as [type])) {
        Add-Type -TypeDefinition @'
using System.Runtime.InteropServices;
public static class KcTaskbar {
    [DllImport("shell32.dll", CharSet = CharSet.Unicode)]
    public static extern int SetCurrentProcessExplicitAppUserModelID(string appId);
    [DllImport("user32.dll")]
    public static extern bool SetProcessDPIAware();
}
'@
    }
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
$script:NotesFile = $null      # soubor, do kterého se poznámky ukládají (po Otevřít / Uložit jako)
$script:MyPath = $MyInvocation.MyCommand.Path
$win = $null

# Status: '' = nerozhodnuto, 'keep' = Ponechat, 'del' = Vymazat (poznámka zapsána), 'unsure' = vrátit se později
# Category = složka podle 1. sloupce; Manual = záznam přidaný přes Vložit mezi; Num = pořadové číslo (pro zobrazení)
function New-Item2([string]$value, [string]$note, [bool]$manual = $false, [string]$category = '') {
    [pscustomobject]@{ Value = $value; Note = $note; Status = ''; Manual = $manual; Num = 0; Category = $category }
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

# Sloupce z Excelu, se kterými aplikace pracuje (číslováno od 1)
$script:ColCategory = 1     # složka (kategorie)
$script:ColValue    = 10    # Údaj k ověření
# poznámka při Vymazat = poslední sloupec

# Ze vloženého textu ponechá v každém řádku jen 3 sloupce: složka, údaj, poznámka (1., 10. a poslední).
# Ostatní sloupce zahodí hned při vložení. Vrací @{ Text; Columns } nebo @{ Error }.
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
        if ($parts.Count -lt $script:ColValue) {
            return @{ Error = "Řádek $($i + 1) má jen $($parts.Count) sloupců. Aplikace potřebuje aspoň $($script:ColValue) sloupců (1. = složka, $($script:ColValue). = údaj, poslední = poznámka).`nZkopírujte z Excelu celý rozsah sloupců. Nic nebylo vloženo." }
        }
        $out.Add($parts[$script:ColCategory - 1].Trim() + "`t" + $parts[$script:ColValue - 1] + "`t" + $parts[$parts.Count - 1])
    }
    while ($out.Count -gt 0 -and $out[$out.Count - 1] -eq '') { $out.RemoveAt($out.Count - 1) }
    return @{ Text = (($out -join "`r`n") + "`r`n"); Columns = $columns }
}

# Z řádků ve tvaru „složka <TAB> údaj <TAB> poznámka“ (po vložení přes Ctrl+V) vytvoří záznamy.
# Vrací @{ Ok; Items; Error }. Při chybě nevrací žádné položky.
function ConvertFrom-Rows([string]$text, [bool]$manual = $false) {
    $lines = Get-ColumnLines $text
    if ($lines.Count -eq 0) {
        return @{ Ok = $false; Error = 'Vstup je prázdný. Zkopírujte řádky z Excelu a vložte je do pole (Ctrl+V).' }
    }
    $result = New-Object System.Collections.Generic.List[object]
    for ($i = 0; $i -lt $lines.Count; $i++) {
        $rowNo = $i + 1
        $line = $lines[$i]
        if ([string]::IsNullOrWhiteSpace($line.Replace("`t", ''))) { continue }
        $parts = $line.Split("`t")
        if ($parts.Count -ne 3) {
            return @{ Ok = $false; Error = "Řádek $rowNo nemá očekávaný tvar. Zkopírujte řádky z Excelu (všechny sloupce) a vložte je do pole přes Ctrl+V.`nNic nebylo načteno." }
        }
        $c = $parts[0].Trim()
        $v = $parts[1].Trim()
        $n = $parts[2].Trim()
        if ($v -eq '') {
            return @{ Ok = $false; Error = "Řádek $rowNo má prázdný $($script:ColValue). sloupec (Údaj k ověření).`nNic nebylo načteno." }
        }
        $result.Add((New-Item2 $v $n $manual $c))
    }
    return @{ Ok = $true; Items = $result }
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
      <Grid.ColumnDefinitions><ColumnDefinition Width="3*"/><ColumnDefinition Width="2*"/></Grid.ColumnDefinitions>

      <!-- Levá část -->
      <Grid Grid.Column="0">
        <Grid.RowDefinitions>
          <RowDefinition Height="*"/><RowDefinition Height="*"/><RowDefinition Height="88"/><RowDefinition Height="Auto"/>
        </Grid.RowDefinitions>

        <!-- Vstup / seznam -->
        <Border Grid.Row="0" Style="{StaticResource Card}">
          <Grid>
            <Grid.RowDefinitions><RowDefinition Height="*"/><RowDefinition Height="Auto"/></Grid.RowDefinitions>
            <TextBox x:Name="InputBox" Style="{StaticResource Field}" AcceptsReturn="True" AcceptsTab="True" TextWrapping="NoWrap"
                     FontFamily="Consolas" FontSize="13" VerticalScrollBarVisibility="Auto" HorizontalScrollBarVisibility="Auto"/>
            <ListBox x:Name="ItemsList" Visibility="Collapsed" BorderThickness="1" BorderBrush="{StaticResource Line}" Background="#F8FAFC"
                     ItemContainerStyle="{StaticResource Row}" ScrollViewer.HorizontalScrollBarVisibility="Disabled"
                     VirtualizingStackPanel.IsVirtualizing="True">
              <ListBox.ItemTemplate>
                <DataTemplate>
                  <Grid>
                    <!-- záhlaví složky (kliknutím se rozbalí / sbalí) -->
                    <Border x:Name="HeaderRow" Visibility="Collapsed" Background="#E2E8F0" CornerRadius="6" Padding="8,5" Margin="-8,4,0,2" Cursor="Hand">
                      <Grid>
                        <Grid.ColumnDefinitions><ColumnDefinition Width="22"/><ColumnDefinition Width="*"/><ColumnDefinition Width="Auto"/></Grid.ColumnDefinitions>
                        <TextBlock Text="{Binding Arrow}" FontWeight="Bold" Foreground="#334155" VerticalAlignment="Center"/>
                        <TextBlock Grid.Column="1" Text="{Binding Title}" FontWeight="SemiBold" Foreground="#0F172A" TextTrimming="CharacterEllipsis" VerticalAlignment="Center"/>
                        <TextBlock Grid.Column="2" Text="{Binding Progress}" FontSize="12" Foreground="#475569" VerticalAlignment="Center" Margin="8,0,4,0"/>
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
            <StackPanel Grid.Row="1" Orientation="Horizontal" Margin="0,12,0,0">
              <Button x:Name="BtnLoad" Style="{StaticResource BtnPrimary}" Content="Vytvořit seznam"/>
              <Button x:Name="BtnClear" Style="{StaticResource Btn}" Content="Zrušit seznam"/>
              <TextBlock x:Name="CountText" Text="Řádků: 0" FontWeight="SemiBold" VerticalAlignment="Center" Margin="8,0,0,0"/>
            </StackPanel>
          </Grid>
        </Border>

        <!-- Aktuální údaj -->
        <Border Grid.Row="1" Style="{StaticResource Card}">
          <DockPanel>
            <Grid DockPanel.Dock="Top">
              <Grid.ColumnDefinitions><ColumnDefinition Width="*"/><ColumnDefinition Width="Auto"/><ColumnDefinition Width="Auto"/></Grid.ColumnDefinitions>
              <TextBlock Text="AKTUÁLNÍ ÚDAJ  •  zkopírováno do schránky" FontSize="12" Foreground="{StaticResource Muted}" VerticalAlignment="Center"/>
              <Button x:Name="BtnUp" Grid.Column="1" Style="{StaticResource BtnNav}" Content="▲" ToolTip="Předchozí řádek"/>
              <Button x:Name="BtnDown" Grid.Column="2" Style="{StaticResource BtnNav}" Content="▼" ToolTip="Další řádek"/>
            </Grid>
            <TextBlock x:Name="CurrentNote" DockPanel.Dock="Bottom" TextAlignment="Center" Foreground="{StaticResource Muted}" TextTrimming="CharacterEllipsis" Margin="0,6,0,0"/>
            <TextBlock x:Name="CurrentText" FontSize="34" FontWeight="SemiBold" TextAlignment="Center" TextWrapping="Wrap"
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
          <Button x:Name="BtnUnsure" Style="{StaticResource Btn}" Background="White" Foreground="#D97706" Content="?   Vrátit se později"/>
          <Button x:Name="BtnInsert" Style="{StaticResource Btn}" Background="White" Foreground="{StaticResource Accent}" Content="+   Vložit mezi" Margin="0"/>
        </StackPanel>
      </Grid>

      <!-- Pravá část: poznámky -->
      <Border Grid.Column="1" Style="{StaticResource Card}">
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
                 'BtnUp','BtnDown','CurrentNote','CurrentText','BtnYes','BtnNo','BtnUnsure','BtnInsert',
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

# Seznam místo vstupního pole (bez seznamu se zobrazí vstupní pole).
# Zobrazují se složky (záhlaví) a pod nimi jejich záznamy; sbalené složky záznamy skrývají.
$script:RowOfItem = @{}   # záznam -> index řádku v zobrazeném seznamu
function Update-ListBox {
    if ($script:Items.Count -eq 0) {
        $ItemsList.ItemsSource = $null
        $ItemsList.Visibility = 'Collapsed'
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
            $done = 0
            while ($j -lt $script:Items.Count -and $script:Items[$j].Category -eq $cat) {
                if ($script:Items[$j].Status -eq 'keep' -or $script:Items[$j].Status -eq 'del') { $done++ }
                $j++
            }
            $collapsed = $script:Collapsed.ContainsKey($cat)
            $title = if ($cat -eq '') { '(bez složky)' } else { $cat }
            $rows.Add([pscustomobject]@{ Kind = 'H'; Category = $cat; Title = $title; Arrow = $(if ($collapsed) { '▸' } else { '▾' });
                                         Progress = "$done / $($j - $i)"; Num = ''; Value = ''; Note = ''; Status = ''; Manual = $false })
            if (-not $collapsed) {
                for ($k = $i; $k -lt $j; $k++) {
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
        $InputBox.Visibility = 'Collapsed'
        $ItemsList.Visibility = 'Visible'
    } finally { $script:Syncing = $false }
}

function Update-View {
    $count = $script:Items.Count
    $on = ($count -gt 0 -and -not $script:Done)
    $BtnYes.IsEnabled = $on; $BtnNo.IsEnabled = $on; $BtnUnsure.IsEnabled = $on; $BtnInsert.IsEnabled = $on
    $BtnUp.IsEnabled = ($count -gt 0); $BtnDown.IsEnabled = ($count -gt 0)
    $CurrentNote.Text = ''
    if ($count -eq 0) {
        $PosText.Text = '0 / 0'
        $CurrentText.Text = 'Vložte řádky z Excelu a klikněte na Vytvořit seznam'
        $CurrentText.Foreground = $bMuted
    } elseif ($script:Done) {
        $PosText.Text = "$count / $count"
        $CurrentText.Text = '✓ Hotovo – všechny řádky jsou rozhodnuté'
        $CurrentText.Foreground = $bDone
        $CurrentNote.Text = 'Šipkami ▲ ▼ nebo kliknutím do seznamu se můžete k libovolnému řádku vrátit.'
    } else {
        $PosText.Text = "$($script:Index + 1) / $count"
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
    if (-not $script:Done -and $script:Index -lt $script:Items.Count) {
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
        if ($rows[0].Split("`t").Count -ne 3) {
            $text = $text + '   •   vložte řádky z Excelu přes Ctrl+V'
            $CountText.Foreground = $bWarn
        } else {
            $cats = @($rows | ForEach-Object { $_.Split("`t")[0].Trim() } | Select-Object -Unique).Count
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

$BtnLoad.Add_Click({ Invoke-Safe {
    if ($script:Items.Count -gt 0) { Show-Warn 'Seznam už je vytvořený. Pro nové vložení ho nejdřív zrušte (Zrušit seznam).'; return }
    $parsed = ConvertFrom-Rows $InputBox.Text
    if (-not $parsed.Ok) { Show-Error $parsed.Error; return }   # stávající seznam ani poznámky se nemění
    if ($parsed.Items.Count -eq 0) { Show-Error 'Vstup neobsahuje žádný neprázdný řádek.'; return }
    $script:Items = Group-Items $parsed.Items
    $script:Collapsed = @{}
    $script:Index = 0
    $script:Done = $false
    Show-Current
} 'Seznam se nepodařilo načíst. Zkuste znovu zkopírovat data z Excelu.' })

$BtnClear.Add_Click({ Invoke-Safe {
    if (-not (Ask-YesNo ('Opravdu zrušit načtený seznam a vymazat vstupní pole?' + "`n`n" + 'Poznámky zůstanou beze změny.') 'Zrušit seznam')) { return }
    $script:Items = New-Object System.Collections.Generic.List[object]
    $script:Index = 0
    $script:Done = $false
    Update-View
    $InputBox.Clear()
} 'Seznam se nepodařilo vymazat.' })

# Kliknutí v seznamu: na záhlaví složky ji rozbalí / sbalí, na záznam ho udělá aktuálním
$ItemsList.Add_SelectionChanged({
    if ($script:Syncing) { return }
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

# Šipky na klávesnici v seznamu: o záznam nahoru / dolů (záhlaví složek se přeskakují)
$ItemsList.Add_PreviewKeyDown({ param($s, $e)
    $d = switch ($e.Key) { 'Up' { -1 } 'Down' { 1 } 'PageUp' { -10 } 'PageDown' { 10 } default { 0 } }
    if ($d -ne 0) {
        $e.Handled = $true
        Invoke-Safe { Move-By $d } 'Přechod na řádek se nezdařil.'
    }
})

$BtnYes.Add_Click({ Invoke-Safe { Set-Decision 'keep' } 'Přechod na další řádek se nezdařil.' })
$BtnNo.Add_Click({ Invoke-Safe { Set-Decision 'del' } 'Zápis poznámky se nezdařil.' })
$BtnUnsure.Add_Click({ Invoke-Safe { Set-Decision 'unsure' } 'Označení se nezdařilo.' })
$BtnUp.Add_Click({ Invoke-Safe { Move-By -1 } 'Přechod na řádek se nezdařil.' })
$BtnDown.Add_Click({ Invoke-Safe { Move-By 1 } 'Přechod na řádek se nezdařil.' })

# Vložit mezi: vloží se řádky z Excelu stejně jako do hlavního pole (Ctrl+V), zařadí se za aktuální řádek
# do stejné složky jako aktuální řádek; aktuální řádek se nemění
$BtnInsert.Add_Click({ Invoke-Safe {
    if ($script:Done -or $script:Index -ge $script:Items.Count) { return }
    [xml]$dx = @'
<Window xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        Title="Vložit mezi" Width="720" Height="420" WindowStartupLocation="CenterOwner"
        ShowInTaskbar="False" FontFamily="Segoe UI" FontSize="14" Background="White" UseLayoutRounding="True">
  <Grid Margin="18">
    <Grid.RowDefinitions><RowDefinition Height="Auto"/><RowDefinition Height="*"/><RowDefinition Height="Auto"/></Grid.RowDefinitions>
    <TextBlock Text="Vložte řádky z Excelu (Ctrl+V) – zařadí se hned pod aktuální řádek (do jeho složky)." TextWrapping="Wrap" Foreground="#475569"/>
    <TextBox Name="Box" Grid.Row="1" Margin="0,10,0,10" AcceptsReturn="True" AcceptsTab="True" TextWrapping="NoWrap"
             FontFamily="Consolas" FontSize="13" Background="#F8FAFC" BorderBrush="#E2E8F0" Padding="8,6"
             VerticalScrollBarVisibility="Auto" HorizontalScrollBarVisibility="Auto"/>
    <DockPanel Grid.Row="2">
      <TextBlock Name="Cnt" Text="Řádků: 0" FontWeight="SemiBold" VerticalAlignment="Center"/>
      <StackPanel Orientation="Horizontal" HorizontalAlignment="Right">
        <Button Name="Ok" Content="Vložit" Width="110" Padding="0,7" Margin="0,0,8,0" Background="#0E7490" Foreground="White" BorderThickness="0"/>
        <Button Name="Cancel" Content="Zrušit" Width="110" Padding="0,7" IsCancel="True" Background="#E8EDF4" BorderThickness="0"/>
      </StackPanel>
    </DockPanel>
  </Grid>
</Window>
'@
    $dlg = [System.Windows.Markup.XamlReader]::Load((New-Object System.Xml.XmlNodeReader $dx))
    $dlg.Owner = $win
    $box = $dlg.FindName('Box'); $cnt = $dlg.FindName('Cnt')
    $box.AddHandler([System.Windows.Input.CommandManager]::PreviewExecutedEvent,
        [System.Windows.Input.ExecutedRoutedEventHandler]{
            param($s, $e)
            if ($e.Command -ne [System.Windows.Input.ApplicationCommands]::Paste) { return }
            $e.Handled = $true
            Invoke-Safe {
                if (-not [System.Windows.Clipboard]::ContainsText()) { return }
                $r = Reduce-PastedColumns ([System.Windows.Clipboard]::GetText())
                if ($r.Error) { Show-Error $r.Error; return }
                $box.SelectedText = $r.Text
                $box.CaretIndex = $box.SelectionStart + $box.SelectionLength
                $box.SelectionLength = 0
            } 'Vložení ze schránky se nezdařilo.'
        })
    $box.Add_TextChanged({
        $n = @(Get-ColumnLines $box.Text | Where-Object { -not [string]::IsNullOrWhiteSpace($_.Replace("`t", '')) }).Count
        $cnt.Text = "Řádků: $n"
    })
    $script:InsertParsed = $null
    $dlg.FindName('Ok').Add_Click({
        $p = ConvertFrom-Rows $box.Text $true
        if (-not $p.Ok) { Show-Error $p.Error; return }
        if ($p.Items.Count -eq 0) { Show-Warn 'Vložte aspoň jeden řádek.'; return }
        $script:InsertParsed = $p.Items
        $dlg.DialogResult = $true
    })
    [void]$box.Focus()
    if ($dlg.ShowDialog() -eq $true -and $script:InsertParsed) {
        $current = $script:Items[$script:Index]
        $list = New-Object System.Collections.Generic.List[object]
        $list.AddRange($script:Items)
        # vložené řádky patří do složky aktuálního řádku, aby se objevily přímo pod ním
        foreach ($it in $script:InsertParsed) { $it.Category = $current.Category }
        $list.InsertRange($script:Index + 1, $script:InsertParsed)
        $script:Items = Group-Items $list
        $script:Index = $script:Items.IndexOf($current)   # aktuální řádek zůstává stejný
        Update-View
    }
} 'Záznamy se nepodařilo vložit.' })

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
    try { Update-Shortcuts } catch { }
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
