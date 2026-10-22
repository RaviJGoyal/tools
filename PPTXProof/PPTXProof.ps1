<#
.SYNOPSIS
    PPTXProof - sets the proofing (spell-check) language of everything in PowerPoint files.

.DESCRIPTION
    Sets one language on all text in .pptx files: slides, speaker notes, tables, grouped
    shapes, charts, SmartArt, slide masters and layouts, notes and handout masters, and
    the defaults that new text inherits. The file is edited directly, so PowerPoint does
    not need to be running. If a file is open in PowerPoint, the window offers to save,
    close and reopen it.

    Without -NoGui a window opens where presentations can be picked or dropped.

.EXAMPLE
    .\PPTXProof.ps1
    Opens the window.

.EXAMPLE
    .\PPTXProof.ps1 .\Talk.pptx .\Slides -Language de-DE -NoGui
    Sets Talk.pptx and every presentation in the Slides folder to German, without a window.
#>
[CmdletBinding()]
param(
    # Presentations or folders to process (preselected when the window opens).
    [Parameter(Position = 0, ValueFromRemainingArguments = $true)]
    [string[]]$Path,

    # Language tag such as de-DE or en-US.
    [string]$Language,

    # Keep a copy of each original file next to it.
    [switch]$Backup,

    # Process the files directly instead of opening the window.
    [switch]$NoGui
)

$ErrorActionPreference = 'Stop'

# Languages offered in the window. Any language tag works with -Language.
$Languages = [ordered]@{
    'de-DE' = 'German (Germany)'
    'de-AT' = 'German (Austria)'
    'de-CH' = 'German (Switzerland)'
    'en-US' = 'English (United States)'
    'en-GB' = 'English (United Kingdom)'
}

$Extensions = '.pptx', '.pptm', '.ppsx', '.ppsm', '.potx', '.potm'

# DrawingML namespaces of transitional and strict Open XML files.
$NamespaceSets = @(
    @{ a = 'http://schemas.openxmlformats.org/drawingml/2006/main'; c = 'http://schemas.openxmlformats.org/drawingml/2006/chart' },
    @{ a = 'http://purl.oclc.org/ooxml/drawingml/main'; c = 'http://purl.oclc.org/ooxml/drawingml/chart' }
)

$AppPath      = $PSCommandPath
$AppDir       = $PSScriptRoot
$SettingsPath = Join-Path $env:APPDATA 'PPTXProof\settings.json'

Add-Type -AssemblyName System.IO.Compression, System.IO.Compression.FileSystem

#region Language engine

function Set-XmlLanguage {
    # Sets the language on all DrawingML text in one XML part.
    # Returns $true when the document was modified.
    param([xml]$Document, [string]$Lang, [hashtable]$Stats)

    $modified = $false
    foreach ($set in $NamespaceSets) {
        $ns = New-Object System.Xml.XmlNamespaceManager $Document.NameTable
        $ns.AddNamespace('a', $set.a)
        $ns.AddNamespace('c', $set.c)

        # Text without its own character properties falls back to PowerPoint's editing
        # language, so give every run, field and line break explicit properties.
        foreach ($node in @($Document.SelectNodes('//a:r[not(a:rPr)] | //a:fld[not(a:rPr)] | //a:br[not(a:rPr)]', $ns))) {
            [void]$node.PrependChild($Document.CreateElement($node.Prefix, 'rPr', $set.a))
        }
        # The end-of-paragraph properties decide the language of text typed there later.
        foreach ($node in @($Document.SelectNodes('//a:p[not(a:endParaRPr)]', $ns))) {
            [void]$node.AppendChild($Document.CreateElement($node.Prefix, 'endParaRPr', $set.a))
        }

        foreach ($node in $Document.SelectNodes('//a:rPr | //a:endParaRPr | //a:defRPr', $ns)) {
            $isRun = $node.LocalName -eq 'rPr' -and $node.ParentNode.LocalName -ne 'br'
            if ($isRun) { $Stats.Runs++ }
            $old = $node.GetAttribute('lang')
            if ($old -ceq $Lang) { continue }

            $node.SetAttribute('lang', $Lang)
            # Drop the cached spell-check result so PowerPoint checks the text again.
            $node.RemoveAttribute('err')
            $node.RemoveAttribute('dirty')
            $modified = $true
            if ($isRun) {
                $Stats.Changed++
                $key = if ($old) { $old } else { 'none' }
                $Stats.Previous[$key] = 1 + [int]$Stats.Previous[$key]
            }
        }

        # Charts carry their own editing language.
        foreach ($node in $Document.SelectNodes('//c:lang', $ns)) {
            if ($node.GetAttribute('val') -cne $Lang) {
                $node.SetAttribute('val', $Lang)
                $modified = $true
            }
        }
    }
    $modified
}

function Update-PackageLanguage {
    # Rewrites the language in all XML parts of a presentation package, in place.
    param([string]$File, [string]$Lang)

    $stats = @{ Runs = 0; Changed = 0; Previous = @{}; Parts = @() }
    $utf8 = New-Object System.Text.UTF8Encoding $false
    $zip = [System.IO.Compression.ZipFile]::Open($File, [System.IO.Compression.ZipArchiveMode]::Update)
    try {
        if (-not $zip.GetEntry('ppt/presentation.xml')) { throw 'This file is not a PowerPoint presentation.' }

        foreach ($entry in @($zip.Entries)) {
            $name = $entry.FullName -replace '\\', '/'
            if ($name -notmatch '^ppt/.+\.xml$' -or $name -like 'ppt/embeddings/*') { continue }

            $reader = New-Object System.IO.StreamReader($entry.Open(), $utf8, $true)
            try { $text = $reader.ReadToEnd() } finally { $reader.Dispose() }
            if (-not $text.Contains('drawingml')) { continue }

            $doc = New-Object System.Xml.XmlDocument
            $doc.PreserveWhitespace = $true
            $doc.XmlResolver = $null
            $doc.LoadXml($text)
            if (-not (Set-XmlLanguage $doc $Lang $stats)) { continue }

            $stream = $entry.Open()
            try {
                $stream.SetLength(0)
                $writer = New-Object System.IO.StreamWriter($stream, $utf8)
                $doc.Save($writer)
                $writer.Flush()
            } finally { $stream.Dispose() }
            $stats.Parts += $name
        }
    } finally { $zip.Dispose() }
    $stats
}

function Test-Package {
    # Re-reads the rewritten parts so a broken file never replaces the original.
    param([string]$File, [string[]]$Parts)

    $zip = [System.IO.Compression.ZipFile]::OpenRead($File)
    try {
        foreach ($part in $Parts) {
            $stream = $zip.GetEntry($part).Open()
            try {
                $doc = New-Object System.Xml.XmlDocument
                $doc.XmlResolver = $null
                $doc.Load($stream)
            } finally { $stream.Dispose() }
        }
    } finally { $zip.Dispose() }
}

function Test-FileLocked {
    param([string]$File)
    try {
        [System.IO.File]::Open($File, [System.IO.FileMode]::Open, [System.IO.FileAccess]::ReadWrite, [System.IO.FileShare]::None).Dispose()
        $false
    } catch [System.IO.IOException] {
        $true
    }
}

function Update-PresentationFile {
    # Sets the language of one presentation. Works on a temporary copy, checks the
    # result and only then replaces the original.
    param([string]$Path, [string]$Lang, [switch]$Backup)

    $file = Get-Item -LiteralPath $Path
    if ($file.Extension -notin $Extensions) {
        throw "Unsupported file type '$($file.Extension)'. Old .ppt files must be saved as .pptx first."
    }
    if ($file.IsReadOnly) { throw 'The file is read-only.' }
    if (Test-FileLocked $file.FullName) { throw 'The file is open in another program. Close it and try again.' }

    $tempDir = Join-Path ([System.IO.Path]::GetTempPath()) 'PPTXProof'
    [void](New-Item -ItemType Directory -Force -Path $tempDir)
    $temp = Join-Path $tempDir ([guid]::NewGuid().ToString('N') + $file.Extension)
    try {
        Copy-Item -LiteralPath $file.FullName -Destination $temp
        try {
            $result = Update-PackageLanguage $temp $Lang
        } catch [System.IO.InvalidDataException] {
            throw 'The file cannot be read. It may be password-protected or encrypted (sensitivity label).'
        }
        $result.Backup = $null
        if ($result.Parts.Count -gt 0) {
            Test-Package $temp $result.Parts
            if ($Backup) {
                $result.Backup = Join-Path $file.DirectoryName ('{0} (backup {1:yyyy-MM-dd HHmmss}){2}' -f $file.BaseName, (Get-Date), $file.Extension)
                Copy-Item -LiteralPath $file.FullName -Destination $result.Backup
            }
            [System.IO.File]::Copy($temp, $file.FullName, $true)
        }
        $result
    } finally {
        Remove-Item -LiteralPath $temp -Force -ErrorAction SilentlyContinue
    }
}

function Get-PresentationPaths {
    # Expands files and folders into the presentations that can be processed.
    param([string[]]$Paths)

    $files = New-Object System.Collections.Generic.List[string]
    $skipped = New-Object System.Collections.Generic.List[string]
    foreach ($p in $Paths) {
        if (-not $p) { continue }
        if (Test-Path -LiteralPath $p -PathType Container) {
            foreach ($item in Get-ChildItem -LiteralPath $p -Recurse -File -ErrorAction SilentlyContinue) {
                if ($item.Extension -in $Extensions -and $item.Name -notlike '~$*') { $files.Add($item.FullName) }
            }
        } elseif ((Test-Path -LiteralPath $p -PathType Leaf) -and ($item = Get-Item -LiteralPath $p).Extension -in $Extensions) {
            $files.Add($item.FullName)
        } else {
            $skipped.Add($p)
        }
    }
    @{ Files = $files; Skipped = $skipped }
}

function Get-LanguageName {
    param([string]$Lang)
    if ($Languages.Contains($Lang)) { return $Languages[$Lang] }
    try { [System.Globalization.CultureInfo]::GetCultureInfo($Lang).EnglishName } catch { $Lang }
}

function Resolve-Language {
    param([string]$Tag)
    try { $culture = [System.Globalization.CultureInfo]::GetCultureInfo($Tag) } catch { $culture = $null }
    if (-not $culture -or $culture.IsNeutralCulture -or -not $culture.Name) {
        throw "Unknown language '$Tag'. Use a tag such as $($Languages.Keys -join ', ')."
    }
    $culture.Name
}

function Format-Result {
    param($Result, [string]$Lang)

    $name = Get-LanguageName $Lang
    if ($Result.Parts.Count -eq 0) { return "Already $name, nothing to change." }
    if ($Result.Changed -eq 0) {
        $text = "Text was already $name; updated the defaults for new text."
    } else {
        $before = ($Result.Previous.GetEnumerator() | Sort-Object Value -Descending |
            ForEach-Object { '{0} {1:N0}' -f $_.Key, $_.Value }) -join ', '
        $text = '{0:N0} of {1:N0} text elements set to {2} (before: {3}).' -f $Result.Changed, $Result.Runs, $name, $before
    }
    if ($Result.Backup) { $text += " Backup: $([System.IO.Path]::GetFileName($Result.Backup))" }
    $text
}

function Get-ErrorText {
    param($ErrorRecord)
    $ErrorRecord.Exception.GetBaseException().Message
}

#endregion

#region PowerPoint integration

function Get-PowerPoint {
    # Attaches to a running PowerPoint; never starts one.
    try { [System.Runtime.InteropServices.Marshal]::GetActiveObject('PowerPoint.Application') } catch { $null }
}

function Resolve-LocalPath {
    # PowerPoint reports OneDrive/SharePoint files as https:// URLs; map them to the synced folder.
    param([string]$FullName)

    if ($FullName -notmatch '^https?://') {
        if ([System.IO.Path]::IsPathRooted($FullName) -and (Test-Path -LiteralPath $FullName -PathType Leaf)) { return $FullName }
        return $null  # never saved
    }
    $url = [Uri]::UnescapeDataString($FullName)
    foreach ($key in Get-ChildItem 'HKCU:\Software\SyncEngines\Providers\OneDrive' -ErrorAction SilentlyContinue) {
        $mount = $key.GetValue('MountPoint')
        $prefix = $key.GetValue('UrlNamespace')
        if (-not $mount -or -not $prefix) { continue }
        $prefix = [Uri]::UnescapeDataString($prefix).TrimEnd('/') + '/'
        if ($url.StartsWith($prefix, [StringComparison]::OrdinalIgnoreCase)) {
            $local = Join-Path $mount ($url.Substring($prefix.Length) -replace '/', '\')
            if (Test-Path -LiteralPath $local -PathType Leaf) { return $local }
        }
    }
    $null
}

function Get-OpenPresentations {
    param($App)
    if (-not $App) { return }
    foreach ($pres in $App.Presentations) {
        [pscustomobject]@{
            Presentation = $pres
            Name         = $pres.Name
            FullName     = $pres.FullName
            LocalPath    = Resolve-LocalPath $pres.FullName
            IsCloud      = $pres.FullName -match '^https?://'
        }
    }
}

function Find-OpenPresentation {
    param($App, [string]$Path)
    foreach ($open in Get-OpenPresentations $App) {
        if ($open.LocalPath) {
            if ($open.LocalPath -eq $Path) { return $open }
        } elseif ($open.IsCloud -and $open.Name -eq [System.IO.Path]::GetFileName($Path)) {
            return $open  # a cloud file we could not map; the same name is the best hint we have
        }
    }
}

function Wait-FileUnlocked {
    param([string]$Path, [int]$TimeoutMs = 10000)
    $watch = [System.Diagnostics.Stopwatch]::StartNew()
    while ((Test-FileLocked $Path) -and $watch.ElapsedMilliseconds -lt $TimeoutMs) { Start-Sleep -Milliseconds 200 }
}

#endregion

#region Window

$WindowXaml = @'
<Window xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
        Title="PPTXProof" Width="640" Height="600" MinWidth="520" MinHeight="480"
        WindowStartupLocation="CenterScreen" AllowDrop="True" Background="#F7F7F8"
        FontFamily="Segoe UI" FontSize="13" UseLayoutRounding="True">
  <Window.Resources>
    <Style TargetType="Button">
      <Setter Property="Padding" Value="12,5"/>
      <Setter Property="Margin" Value="0,0,8,0"/>
      <Setter Property="MinWidth" Value="72"/>
      <Setter Property="Foreground" Value="#1F1F1F"/>
      <Setter Property="Background" Value="White"/>
      <Setter Property="BorderBrush" Value="#D1D1D1"/>
      <Setter Property="Template">
        <Setter.Value>
          <ControlTemplate TargetType="Button">
            <Border x:Name="Chrome" Background="{TemplateBinding Background}" BorderBrush="{TemplateBinding BorderBrush}"
                    BorderThickness="1" CornerRadius="4" Padding="{TemplateBinding Padding}">
              <ContentPresenter HorizontalAlignment="Center" VerticalAlignment="Center"/>
            </Border>
            <ControlTemplate.Triggers>
              <Trigger Property="IsMouseOver" Value="True"><Setter TargetName="Chrome" Property="Background" Value="#F3F3F3"/></Trigger>
              <Trigger Property="IsPressed" Value="True"><Setter TargetName="Chrome" Property="Background" Value="#E6E6E6"/></Trigger>
              <Trigger Property="IsEnabled" Value="False">
                <Setter Property="Foreground" Value="#A6A6A6"/>
                <Setter TargetName="Chrome" Property="BorderBrush" Value="#E3E3E3"/>
              </Trigger>
            </ControlTemplate.Triggers>
          </ControlTemplate>
        </Setter.Value>
      </Setter>
    </Style>
    <Style x:Key="PrimaryButton" TargetType="Button">
      <Setter Property="Foreground" Value="White"/>
      <Setter Property="Background" Value="#0F6CBD"/>
      <Setter Property="FontWeight" Value="SemiBold"/>
      <Setter Property="Padding" Value="24,9"/>
      <Setter Property="Cursor" Value="Hand"/>
      <Setter Property="Template">
        <Setter.Value>
          <ControlTemplate TargetType="Button">
            <Border x:Name="Chrome" Background="{TemplateBinding Background}" CornerRadius="4" Padding="{TemplateBinding Padding}">
              <ContentPresenter HorizontalAlignment="Center" VerticalAlignment="Center"/>
            </Border>
            <ControlTemplate.Triggers>
              <Trigger Property="IsMouseOver" Value="True"><Setter TargetName="Chrome" Property="Background" Value="#115EA3"/></Trigger>
              <Trigger Property="IsPressed" Value="True"><Setter TargetName="Chrome" Property="Background" Value="#0C3B5E"/></Trigger>
              <Trigger Property="IsEnabled" Value="False"><Setter TargetName="Chrome" Property="Background" Value="#BDBDBD"/></Trigger>
            </ControlTemplate.Triggers>
          </ControlTemplate>
        </Setter.Value>
      </Setter>
    </Style>
  </Window.Resources>

  <Grid Margin="20,16,20,16">
    <Grid.RowDefinitions>
      <RowDefinition Height="Auto"/>
      <RowDefinition Height="*"/>
      <RowDefinition Height="Auto"/>
      <RowDefinition Height="Auto"/>
      <RowDefinition Height="Auto"/>
    </Grid.RowDefinitions>

    <StackPanel Margin="0,0,0,14">
      <TextBlock Text="Set the proofing language" FontSize="20" FontWeight="SemiBold" Foreground="#1F1F1F"/>
      <TextBlock Margin="0,4,0,0" TextWrapping="Wrap" Foreground="#5C5C5C"
                 Text="Applies one spell-check language to all text in the presentations below: slides, notes, tables, groups, charts, SmartArt, masters and layouts, plus the default for new text."/>
    </StackPanel>

    <Grid Grid.Row="1">
      <ListBox x:Name="FileList" SelectionMode="Extended" HorizontalContentAlignment="Stretch"
               ScrollViewer.HorizontalScrollBarVisibility="Disabled"
               Background="White" BorderBrush="#D6D6D6" Padding="2">
        <ListBox.ItemTemplate>
          <DataTemplate>
            <StackPanel Margin="4,4" ToolTip="{Binding Path}">
              <TextBlock Text="{Binding Name}" FontWeight="SemiBold" TextTrimming="CharacterEllipsis"/>
              <TextBlock Text="{Binding Folder}" Foreground="#767676" FontSize="12" TextTrimming="CharacterEllipsis"/>
              <TextBlock Text="{Binding Status}" Foreground="{Binding StatusColor}" Visibility="{Binding StatusVisibility}"
                         TextWrapping="Wrap" Margin="0,3,0,0"/>
            </StackPanel>
          </DataTemplate>
        </ListBox.ItemTemplate>
      </ListBox>
      <TextBlock x:Name="EmptyHint" HorizontalAlignment="Center" VerticalAlignment="Center" TextAlignment="Center"
                 Foreground="#8A8A8A" FontSize="14" IsHitTestVisible="False"
                 Text="Drop presentations or folders here&#10;or click &quot;Add files...&quot;"/>
    </Grid>

    <StackPanel Grid.Row="2" Orientation="Horizontal" Margin="0,8,0,0">
      <Button x:Name="AddButton" Content="Add files..."/>
      <Button x:Name="AddOpenButton" Content="Add from PowerPoint" ToolTip="Add the presentations that are open in PowerPoint right now"/>
      <Button x:Name="RemoveButton" Content="Remove"/>
      <Button x:Name="ClearButton" Content="Clear"/>
    </StackPanel>

    <Border Grid.Row="3" Margin="0,16,0,0" Padding="14,12" Background="White" BorderBrush="#E0E0E0" BorderThickness="1" CornerRadius="6">
      <Grid>
        <Grid.ColumnDefinitions>
          <ColumnDefinition Width="Auto"/>
          <ColumnDefinition Width="*"/>
        </Grid.ColumnDefinitions>
        <Grid.RowDefinitions>
          <RowDefinition Height="Auto"/>
          <RowDefinition Height="Auto"/>
        </Grid.RowDefinitions>
        <TextBlock Text="Language" FontWeight="SemiBold" VerticalAlignment="Center" Margin="0,0,14,0"/>
        <ComboBox x:Name="LanguageBox" Grid.Column="1" Width="260" HorizontalAlignment="Left" Padding="8,5"
                  DisplayMemberPath="Value" SelectedValuePath="Key"/>
        <CheckBox x:Name="BackupBox" Grid.Row="1" Grid.Column="1" Margin="0,10,0,0"
                  Content="Keep a backup copy of each original file"/>
      </Grid>
    </Border>

    <Grid Grid.Row="4" Margin="0,16,0,0">
      <Grid.ColumnDefinitions>
        <ColumnDefinition Width="*"/>
        <ColumnDefinition Width="Auto"/>
      </Grid.ColumnDefinitions>
      <StackPanel VerticalAlignment="Center" Margin="0,0,16,0">
        <TextBlock x:Name="StatusText" TextWrapping="Wrap" Foreground="#5C5C5C"/>
        <TextBlock FontSize="12" Margin="0,4,0,0">
          <Hyperlink x:Name="ShortcutLink">Create shortcuts (Start menu, Desktop, Send to)</Hyperlink>
        </TextBlock>
      </StackPanel>
      <Button x:Name="ApplyButton" Grid.Column="1" Style="{StaticResource PrimaryButton}" Content="Set language" IsDefault="True"
              VerticalAlignment="Center"/>
    </Grid>
  </Grid>
</Window>
'@

$StatusColors = @{ Busy = '#5C5C5C'; Info = '#5C5C5C'; Success = '#107C10'; Warning = '#9D5D00'; Error = '#C42B1C' }

function New-AppWindow {
    Add-Type -AssemblyName PresentationFramework, PresentationCore, WindowsBase

    $window = [System.Windows.Markup.XamlReader]::Load((New-Object System.Xml.XmlNodeReader ([xml]$WindowXaml)))
    $script:Ui = @{
        Window = $window
        Files  = New-Object 'System.Collections.ObjectModel.ObservableCollection[object]'
    }
    foreach ($name in 'FileList', 'EmptyHint', 'AddButton', 'AddOpenButton', 'RemoveButton', 'ClearButton',
                      'LanguageBox', 'BackupBox', 'StatusText', 'ApplyButton', 'ShortcutLink') {
        $Ui[$name] = $window.FindName($name)
    }

    $icon = Join-Path $AppDir 'app.ico'
    if (Test-Path -LiteralPath $icon) { $window.Icon = [System.Windows.Media.Imaging.BitmapFrame]::Create((New-Object Uri $icon)) }

    $settings = Get-Settings
    $Ui.FileList.ItemsSource = $Ui.Files
    $Ui.LanguageBox.ItemsSource = $Languages
    $Ui.LanguageBox.SelectedValue = if ($settings.Language -and $Languages.Contains($settings.Language)) { $settings.Language } else { 'de-DE' }
    $Ui.BackupBox.IsChecked = [bool]$settings.Backup

    $window.Add_DragOver({
        param($source, $e)
        $e.Effects = if ($e.Data.GetDataPresent([System.Windows.DataFormats]::FileDrop)) { 'Copy' } else { 'None' }
        $e.Handled = $true
    })
    $window.Add_Drop({
        param($source, $e)
        if ($e.Data.GetDataPresent([System.Windows.DataFormats]::FileDrop)) {
            Add-Files $e.Data.GetData([System.Windows.DataFormats]::FileDrop)
        }
    })
    $Ui.AddButton.Add_Click({
        $dialog = New-Object Microsoft.Win32.OpenFileDialog
        $dialog.Title = 'Choose presentations'
        $dialog.Filter = 'PowerPoint presentations|*' + ($Extensions -join ';*') + '|All files|*.*'
        $dialog.Multiselect = $true
        if ($dialog.ShowDialog($Ui.Window)) { Add-Files $dialog.FileNames }
    })
    $Ui.AddOpenButton.Add_Click({ Add-OpenPresentations })
    $Ui.RemoveButton.Add_Click({ Remove-SelectedFiles })
    $Ui.ClearButton.Add_Click({ $Ui.Files.Clear(); Set-StatusText ''; Update-ListState })
    $Ui.FileList.Add_SelectionChanged({ Update-ListState })
    $Ui.FileList.Add_KeyDown({
        param($source, $e)
        if ($e.Key -eq 'Delete') { Remove-SelectedFiles }
    })
    $Ui.ApplyButton.Add_Click({ Invoke-Apply })
    $Ui.ShortcutLink.Add_Click({ New-Shortcuts })

    Update-ListState
    $window
}

function New-FileItem {
    param([string]$Path)
    [pscustomobject]@{
        Path             = $Path
        Name             = [System.IO.Path]::GetFileName($Path)
        Folder           = [System.IO.Path]::GetDirectoryName($Path)
        Status           = ''
        StatusColor      = $StatusColors.Info
        StatusVisibility = 'Collapsed'
    }
}

function Add-Files {
    param([string[]]$Paths)
    $found = Get-PresentationPaths $Paths
    foreach ($file in $found.Files) {
        if (-not ($Ui.Files | Where-Object { $_.Path -eq $file })) { $Ui.Files.Add((New-FileItem $file)) }
    }
    if ($found.Skipped.Count) {
        Set-StatusText ("Ignored {0} file(s) that are not .pptx presentations. Old .ppt files must be saved as .pptx first." -f $found.Skipped.Count)
    }
    Update-ListState
}

function Add-OpenPresentations {
    param([switch]$ActiveOnly)

    $app = Get-PowerPoint
    if (-not $app) {
        if (-not $ActiveOnly) { Set-StatusText 'PowerPoint is not running.' }
        return
    }
    $open = @(Get-OpenPresentations $app)
    if ($ActiveOnly) {
        try { $active = $app.ActivePresentation.FullName } catch { return }
        $open = @($open | Where-Object { $_.FullName -eq $active })
    }
    Add-Files @($open | Where-Object LocalPath | ForEach-Object LocalPath)

    $missing = @($open | Where-Object { -not $_.LocalPath } | ForEach-Object Name)
    if ($missing.Count) {
        Set-StatusText ("No local file found for: {0}. Save it first, or add it with ""Add files..."" from your OneDrive folder." -f ($missing -join ', '))
    } elseif (-not $ActiveOnly -and -not $open.Count) {
        Set-StatusText 'No presentations are open in PowerPoint.'
    }
}

function Remove-SelectedFiles {
    foreach ($item in @($Ui.FileList.SelectedItems)) { [void]$Ui.Files.Remove($item) }
    Update-ListState
}

function Update-ListState {
    $count = $Ui.Files.Count
    $Ui.EmptyHint.Visibility = if ($count) { 'Collapsed' } else { 'Visible' }
    $Ui.ApplyButton.IsEnabled = $count -gt 0
    $Ui.ClearButton.IsEnabled = $count -gt 0
    $Ui.RemoveButton.IsEnabled = $Ui.FileList.SelectedItems.Count -gt 0
}

function Set-StatusText {
    param([string]$Text)
    $Ui.StatusText.Text = $Text
}

function Set-ItemStatus {
    param($Item, [string]$Text, [string]$Kind)
    $Item.Status = $Text
    $Item.StatusColor = $StatusColors[$Kind]
    $Item.StatusVisibility = 'Visible'
    $Ui.FileList.Items.Refresh()
    Update-Ui
}

function Update-Ui {
    # Lets the window repaint while files are being processed.
    $Ui.Window.Dispatcher.Invoke([Action] {}, [System.Windows.Threading.DispatcherPriority]::Background)
}

function Set-Busy {
    param([bool]$Busy)
    foreach ($name in 'AddButton', 'AddOpenButton', 'RemoveButton', 'ClearButton', 'LanguageBox', 'BackupBox', 'ApplyButton') {
        $Ui[$name].IsEnabled = -not $Busy
    }
    $Ui.Window.Cursor = if ($Busy) { [System.Windows.Input.Cursors]::Wait } else { $null }
    if (-not $Busy) { Update-ListState }
}

function Confirm-ClosePresentation {
    param($Open)
    $message = "'{0}' is open in PowerPoint.`n`nPPTXProof will save and close it, set the language, and open it again." -f $Open.Name
    [System.Windows.MessageBox]::Show($Ui.Window, $message, 'PPTXProof', 'OKCancel', 'Question') -eq 'OK'
}

function Invoke-Apply {
    $lang = $Ui.LanguageBox.SelectedValue
    $backup = [bool]$Ui.BackupBox.IsChecked
    Save-Settings $lang $backup

    Set-Busy $true
    Set-StatusText 'Working...'
    $app = Get-PowerPoint
    $failed = 0
    try {
        foreach ($item in @($Ui.Files)) {
            Set-ItemStatus $item 'Working...' 'Busy'
            $reopen = $false
            try {
                $open = if ($app) { Find-OpenPresentation $app $item.Path }
                if ($open) {
                    if (-not (Confirm-ClosePresentation $open)) {
                        Set-ItemStatus $item 'Skipped: still open in PowerPoint.' 'Warning'
                        continue
                    }
                    if (-not $open.Presentation.Saved) { $open.Presentation.Save() }
                    $open.Presentation.Close()
                    $reopen = $true
                    Wait-FileUnlocked $item.Path
                }
                $result = Update-PresentationFile -Path $item.Path -Lang $lang -Backup:$backup
                Set-ItemStatus $item (Format-Result $result $lang) $(if ($result.Parts.Count) { 'Success' } else { 'Info' })
            } catch {
                $failed++
                Set-ItemStatus $item (Get-ErrorText $_) 'Error'
            } finally {
                if ($reopen) {
                    try { [void]$app.Presentations.Open($item.Path) }
                    catch { Set-ItemStatus $item ($item.Status + ' Could not reopen it in PowerPoint.') 'Warning' }
                }
            }
        }
    } finally {
        Set-Busy $false
    }
    Set-StatusText $(if ($failed) { "Finished with $failed error(s)." } else { "Done. All presentations are set to $(Get-LanguageName $lang)." })
}

function New-Shortcuts {
    $shell = New-Object -ComObject WScript.Shell
    $powershell = Join-Path $env:SystemRoot 'System32\WindowsPowerShell\v1.0\powershell.exe'
    foreach ($folder in 'Programs', 'Desktop', 'SendTo') {
        $link = $shell.CreateShortcut((Join-Path ([Environment]::GetFolderPath($folder)) 'PPTXProof.lnk'))
        $link.TargetPath = $powershell
        $link.Arguments = '-NoProfile -ExecutionPolicy Bypass -STA -WindowStyle Hidden -File "{0}"' -f $AppPath
        $link.WorkingDirectory = $AppDir
        $link.WindowStyle = 7  # minimized, so no console flashes up
        $link.IconLocation = (Join-Path $AppDir 'app.ico') + ',0'
        $link.Description = 'Set the proofing language of PowerPoint presentations'
        $link.Save()
    }
    [void][System.Windows.MessageBox]::Show($Ui.Window,
        "Shortcuts created in the Start menu, on the Desktop and in the Send to menu.`n`nTip: right-click presentations in Explorer > Show more options > Send to > PPTXProof.`n`nIf you move the PPTXProof folder, create the shortcuts again.",
        'PPTXProof', 'OK', 'Information')
}

function Get-Settings {
    try { Get-Content -LiteralPath $SettingsPath -Raw | ConvertFrom-Json } catch { $null }
}

function Save-Settings {
    param([string]$Lang, [bool]$Backup)
    try {
        [void](New-Item -ItemType Directory -Force -Path (Split-Path $SettingsPath))
        @{ Language = $Lang; Backup = $Backup } | ConvertTo-Json | Set-Content -LiteralPath $SettingsPath -Encoding UTF8
    } catch {}
}

#endregion

function Invoke-NoGui {
    if (-not $Language) { throw 'Specify the language with -Language, for example -Language de-DE.' }
    $lang = Resolve-Language $Language
    $found = Get-PresentationPaths $Path
    foreach ($p in $found.Skipped) { Write-Warning "Skipped (not a .pptx presentation): $p" }

    $app = Get-PowerPoint
    $failed = $found.Skipped.Count
    foreach ($file in $found.Files) {
        try {
            if ($app -and (Find-OpenPresentation $app $file)) { throw 'The file is open in PowerPoint. Close it and try again.' }
            $result = Update-PresentationFile -Path $file -Lang $lang -Backup:$Backup
            Write-Host "$file`n  $(Format-Result $result $lang)"
        } catch {
            $failed++
            Write-Host "$file`n  Error: $(Get-ErrorText $_)" -ForegroundColor Red
        }
    }
    if (-not $found.Files.Count) { Write-Warning 'No presentations to process.' }
    exit [int]($failed -gt 0 -or -not $found.Files.Count)
}

function Start-App {
    if ($NoGui) { Invoke-NoGui }

    try {
        if ([System.Threading.Thread]::CurrentThread.ApartmentState -ne 'STA') {
            throw 'The window needs a single-threaded apartment. Start PowerShell with -STA.'
        }
        $window = New-AppWindow
        if ($Language -and $Languages.Contains($Language)) { $Ui.LanguageBox.SelectedValue = $Language }
        if ($Path) { Add-Files $Path } else { Add-OpenPresentations -ActiveOnly }
        [void]$window.ShowDialog()
    } catch {
        # The launcher hides the console, so errors must be shown in a message box.
        Add-Type -AssemblyName PresentationFramework
        [void][System.Windows.MessageBox]::Show("PPTXProof could not start:`n`n$(Get-ErrorText $_)", 'PPTXProof', 'OK', 'Error')
    }
}

# Dot-sourcing (". .\PPTXProof.ps1") only loads the functions.
if ($MyInvocation.InvocationName -ne '.') { Start-App }
