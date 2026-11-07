# RD Bar - lanceur RD au clavier, près de l'horloge.
#   Setup.cmd / Uninstall.cmd : installe (copie locale + démarrage à la connexion) / désinstalle.
#   Raccourci global (Hotkey du fichier de réglages, ou celui choisi sur ce PC) : ouvre la barre.
#   Entrée = action surlignée, Haut/Bas = déplacer, Ctrl+lettre/chiffre = actions, Alt+1..6 = liens, Échap = masquer.
#   Réglages : RD_Browser.config.json (relu à chaque ouverture). Mise à jour : copie partagée adoptée au démarrage si elle est valide.
param([switch]$Install, [switch]$Uninstall, [string]$SourceDir, [switch]$Welcome)

$scriptDir = $PSScriptRoot
if (-not $SourceDir) { $SourceDir = $scriptDir }
$psExe      = "$env:SystemRoot\System32\WindowsPowerShell\v1.0\powershell.exe"
$appDir     = Join-Path $env:LOCALAPPDATA 'RD_Browser'
$hotkeyFile = Join-Path $appDir 'hotkey.txt'
$themeFile  = Join-Path $appDir 'theme.txt'
$shortcuts  = 'Startup', 'Programs' | ForEach-Object { Join-Path ([Environment]::GetFolderPath($_)) 'RD Browser.lnk' }

function Get-RunArgs($script) { "-NoProfile -ExecutionPolicy Bypass -WindowStyle Hidden -STA -File `"$script`" -SourceDir `"$SourceDir`"" }

function Stop-Running {   # only this user's session, never other people's bars on a shared server
    $session = (Get-Process -Id $PID).SessionId
    Get-CimInstance Win32_Process -Filter "Name='powershell.exe' AND SessionId=$session" -EA 0 |
        Where-Object { $_.ProcessId -ne $PID -and $_.CommandLine -match 'RD_Bar\.ps1' -and $_.CommandLine -notmatch '-(un)?install' } |
        ForEach-Object { Stop-Process -Id $_.ProcessId -Force -EA 0 }
}

function Read-Config {
    $script:config = foreach ($d in $SourceDir, $scriptDir) {
        try { Get-Content -LiteralPath (Join-Path $d 'RD_Browser.config.json') -Raw -Encoding UTF8 -EA Stop | ConvertFrom-Json; break } catch { }
    }
    if (-not $config) {
        $script:config = [pscustomobject]@{ Buttons = @(
            [pscustomobject]@{ Label = 'Dossier'; Type = 'folder' },
            [pscustomobject]@{ Label = 'Manip'; Type = 'file'; Pattern = 'manip'; Shortcut = 'M' }) }
    }
    $script:root    = [IO.Path]::GetFullPath([IO.Path]::Combine($SourceDir, "$($config.RootPath)"))
    $script:buttons = @($config.Buttons)
    $script:defaultHotkey = if ($config.Hotkey) { "$($config.Hotkey)" } else { 'Ctrl+Alt+F' }
    $own = Get-Content -LiteralPath $hotkeyFile -EA 0 | Select-Object -First 1
    $script:hotkey = if ($own) { $own.Trim() } else { $defaultHotkey }
}

# ---------------------------------------------------------------- install / uninstall
if ($Install) {
    Write-Host 'Installation de RD Browser...'
    Stop-Running
    $null = New-Item -ItemType Directory -Force $appDir
    $local = Join-Path $appDir 'RD_Bar.ps1'
    Copy-Item -LiteralPath $PSCommandPath $local -Force
    Copy-Item -LiteralPath (Join-Path $SourceDir 'RD_Browser.config.json') $appDir -Force -EA 0
    $shell = New-Object -ComObject WScript.Shell
    foreach ($path in $shortcuts) {
        $sc = $shell.CreateShortcut($path)
        $sc.TargetPath = $psExe; $sc.Arguments = Get-RunArgs $local; $sc.WorkingDirectory = $appDir
        $sc.WindowStyle = 7; $sc.Hotkey = ''; $sc.IconLocation = "$env:SystemRoot\System32\shell32.dll,3"
        $sc.Description = 'RD Browser - tapez un numéro de RD, ouvrez son dossier ou son fichier Manip'
        $sc.Save()
    }
    Start-Process $psExe "$(Get-RunArgs $local) -Welcome" -WindowStyle Hidden
    Read-Config
    Write-Host "`nTerminé. RD Browser tourne maintenant près de l'horloge (en bas à droite)."
    Write-Host "  Appuyez sur $hotkey n'importe où pour l'ouvrir. Il démarre tout seul à chaque connexion."
    Write-Host "  Copie du programme : $appDir"
    Write-Host "  Installé depuis    : $SourceDir (réglages : RD_Browser.config.json)"
    return
}
if ($Uninstall) {
    Write-Host 'Désinstallation de RD Browser...'
    Stop-Running
    foreach ($p in $shortcuts + $appDir) { if (Test-Path -LiteralPath $p) { Remove-Item -LiteralPath $p -Recurse -Force; Write-Host "  supprimé : $p" } }
    Write-Host 'Terminé.'
    return
}

# ---------------------------------------------------------------- self-update (only a copy that parses is adopted)
$shared = Join-Path $SourceDir 'RD_Bar.ps1'
if ($SourceDir -ne $scriptDir -and (Test-Path -LiteralPath $shared)) {
    $candidate = Join-Path $appDir 'RD_Bar.new.ps1'
    try {
        if ((Get-FileHash -LiteralPath $shared).Hash -ne (Get-FileHash -LiteralPath $PSCommandPath).Hash) {
            Copy-Item -LiteralPath $shared $candidate -Force
            $errors = $null
            [void][Management.Automation.Language.Parser]::ParseFile($candidate, [ref]$null, [ref]$errors)
            if (-not $errors) {
                Move-Item -LiteralPath $candidate $PSCommandPath -Force
                Start-Process $psExe (Get-RunArgs $PSCommandPath) -WindowStyle Hidden
                return
            }
        }
    } catch { }
    Remove-Item -LiteralPath $candidate -Force -EA 0
}

# ---------------------------------------------------------------- single instance
$isFirst = $false
$mutex = New-Object Threading.Mutex($true, 'Local\RD_Bar_Tray', [ref]$isFirst)
if (-not $isFirst) { return }
Read-Config

# ---------------------------------------------------------------- RD lookup
function Get-RdName($text) { $d = $text -replace '\D'; if ($d -and $d.Length -le 6) { 'RD{0:D3}' -f [int]$d } }

function Find-RdFolder($rd) {
    $exact = Join-Path $root $rd
    if (Test-Path -LiteralPath $exact -PathType Container) { return Get-Item -LiteralPath $exact }
    Get-ChildItem -LiteralPath $root -Directory -EA 0 | Where-Object Name -match "^$rd(\D|$)" | Select-Object -First 1
}

function Find-Files($dir, $btn) {
    $ext = if ($btn.Extensions) { @($btn.Extensions) } else { '.xlsx', '.xlsm', '.xls' }
    $pattern = [regex]::Escape("$($btn.Pattern)")
    Get-ChildItem -LiteralPath $dir.FullName -File -EA 0 |
        Where-Object { $_.Name -notlike '~$*' -and $ext -contains $_.Extension -and $_.BaseName -match $pattern } |
        Sort-Object LastWriteTime -Descending
}

# ---------------------------------------------------------------- global shortcut
Add-Type -AssemblyName PresentationFramework, PresentationCore, WindowsBase, System.Windows.Forms, System.Drawing
Add-Type -ReferencedAssemblies System.Windows.Forms -TypeDefinition @'
using System; using System.Runtime.InteropServices; using System.Windows.Forms;
public class RdHotkey : NativeWindow, IDisposable {
    [DllImport("user32.dll")] static extern bool RegisterHotKey(IntPtr h, int id, uint mods, uint vk);
    [DllImport("user32.dll")] static extern bool UnregisterHotKey(IntPtr h, int id);
    public event EventHandler Pressed;
    public bool Registered;
    public RdHotkey(uint mods, uint vk) { CreateHandle(new CreateParams()); Registered = RegisterHotKey(Handle, 1, mods | 0x4000, vk); }
    protected override void WndProc(ref Message m) { if (m.Msg == 0x0312 && Pressed != null) Pressed(this, EventArgs.Empty); base.WndProc(ref m); }
    public void Dispose() { UnregisterHotKey(Handle, 1); DestroyHandle(); }
}
'@

# Returns the registered shortcut, or nothing if the text is invalid or another program already uses it
function Register-Hotkey($text) {
    $mods = 0; $vk = 0
    foreach ($part in "$text" -split '\+') {
        switch -regex ($part.Trim()) {
            '^(ctrl|control)$' { $mods = $mods -bor 2 }
            '^alt$'            { $mods = $mods -bor 1 }
            '^shift$'          { $mods = $mods -bor 4 }
            '^win$'            { $mods = $mods -bor 8 }
            default            { try { $vk = [int][Windows.Forms.Keys]($_ -replace '^(\d)$', 'D$1') } catch { } }
        }
    }
    if (-not ($mods -and $vk)) { return }
    $h = New-Object RdHotkey([uint32]$mods, [uint32]$vk)
    if (-not $h.Registered) { $h.Dispose(); return }
    $h.add_Pressed({ Show-Bar })
    $h
}

# ---------------------------------------------------------------- window (colours are theme resources set by Set-Theme)
$window = [Windows.Markup.XamlReader]::Parse(@'
<Window xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation" xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
        Title="RD Browser" Width="680" SizeToContent="Height" WindowStyle="None" AllowsTransparency="True" Background="Transparent"
        ResizeMode="NoResize" Topmost="True" ShowInTaskbar="False" FontFamily="Segoe UI">
  <Window.Resources>
    <Style x:Key="Ico" TargetType="Path">
      <Setter Property="Stretch" Value="Uniform"/><Setter Property="StrokeThickness" Value="1.8"/><Setter Property="VerticalAlignment" Value="Center"/>
      <Setter Property="StrokeLineJoin" Value="Round"/><Setter Property="StrokeStartLineCap" Value="Round"/><Setter Property="StrokeEndLineCap" Value="Round"/>
      <Setter Property="Stroke" Value="{DynamicResource Muted}"/>
    </Style>
    <Style x:Key="Hover" TargetType="Border">
      <Setter Property="Background" Value="Transparent"/><Setter Property="Cursor" Value="Hand"/>
      <Style.Triggers><Trigger Property="IsMouseOver" Value="True"><Setter Property="Background" Value="{DynamicResource Hover}"/></Trigger></Style.Triggers>
    </Style>
    <Style x:Key="Tool" TargetType="Border" BasedOn="{StaticResource Hover}">
      <Setter Property="Width" Value="34"/><Setter Property="Height" Value="34"/><Setter Property="CornerRadius" Value="6"/><Setter Property="Margin" Value="0,0,2,0"/>
    </Style>
    <Style x:Key="Kbd" TargetType="Border">
      <Setter Property="BorderBrush" Value="{DynamicResource KbdLine}"/><Setter Property="BorderThickness" Value="1"/>
      <Setter Property="CornerRadius" Value="4"/><Setter Property="Padding" Value="6,1"/><Setter Property="VerticalAlignment" Value="Center"/>
    </Style>
  </Window.Resources>
  <Grid x:Name="Root">
    <Border Margin="24" CornerRadius="12" BorderThickness="1" Background="{DynamicResource Card}" BorderBrush="{DynamicResource CardLine}">
      <Border.Effect><DropShadowEffect BlurRadius="28" ShadowDepth="6" Opacity="0.35"/></Border.Effect>
      <Grid>
        <Grid.ColumnDefinitions><ColumnDefinition Width="170"/><ColumnDefinition/></Grid.ColumnDefinitions>

        <!-- Links column: quick links, tool buttons at the bottom -->
        <Border Background="{DynamicResource Side}" BorderBrush="{DynamicResource Line}" BorderThickness="0,0,1,0" CornerRadius="12,0,0,12">
          <DockPanel Margin="8,12,8,8">
            <StackPanel DockPanel.Dock="Bottom" Orientation="Horizontal" Margin="0,10,0,0">
              <Border x:Name="ThemeBtn" Style="{StaticResource Tool}"><Path x:Name="ThemeIco" Style="{StaticResource Ico}" Width="18" Height="18"/></Border>
              <Border x:Name="KeyBtn" Style="{StaticResource Tool}"><Path x:Name="KeyIco" Style="{StaticResource Ico}" Width="18" Height="18"/></Border>
              <Border x:Name="Cog" Style="{StaticResource Tool}" ToolTip="Ouvrir les fichiers de RD Browser (réglages, scripts)"><Path x:Name="CogIco" Style="{StaticResource Ico}" Width="18" Height="18"/></Border>
            </StackPanel>
            <StackPanel>
              <TextBlock x:Name="LinksTitle" Text="LIENS" FontSize="10.5" FontWeight="SemiBold" Margin="8,2,0,6" Foreground="{DynamicResource Muted}"/>
              <ItemsControl x:Name="Links">
                <ItemsControl.ItemTemplate><DataTemplate>
                  <Border Style="{StaticResource Hover}" Height="32" Margin="0,1" Padding="8,0,6,0" CornerRadius="6" ToolTip="{Binding Tip}">
                    <DockPanel>
                      <Path Style="{StaticResource Ico}" Width="15" Height="15" Margin="0,0,8,0" Stroke="{DynamicResource Text}" Data="{Binding Geo}"/>
                      <Border DockPanel.Dock="Right" Style="{StaticResource Kbd}" CornerRadius="3" Padding="4,0">
                        <TextBlock Text="{Binding N}" FontSize="10" Foreground="{DynamicResource KbdText}"/>
                      </Border>
                      <TextBlock Text="{Binding Name}" FontSize="13" VerticalAlignment="Center" TextTrimming="CharacterEllipsis" Foreground="{DynamicResource Text}"/>
                    </DockPanel>
                  </Border>
                </DataTemplate></ItemsControl.ItemTemplate>
              </ItemsControl>
            </StackPanel>
          </DockPanel>
        </Border>

        <!-- Search box, RD header, message line, actions (actions share the remaining height) -->
        <Grid Grid.Column="1">
          <Grid.RowDefinitions><RowDefinition Height="Auto"/><RowDefinition Height="Auto"/><RowDefinition Height="Auto"/><RowDefinition Height="Auto"/><RowDefinition/></Grid.RowDefinitions>
          <DockPanel Height="58" Margin="18,0">
            <Path Style="{StaticResource Ico}" Width="20" Height="20" StrokeThickness="2" Data="M17,10 A7,7 0 1 1 3,10 A7,7 0 1 1 17,10 M15,15 L21,21"/>
            <Border DockPanel.Dock="Right" Style="{StaticResource Kbd}"><TextBlock Text="Échap" FontSize="11" Foreground="{DynamicResource KbdText}"/></Border>
            <Grid Margin="12,0">
              <TextBox x:Name="Box" BorderThickness="0" Background="Transparent" FontSize="20" VerticalContentAlignment="Center"
                       Foreground="{DynamicResource Text}" CaretBrush="{DynamicResource Text}"/>
              <TextBlock Text="RD000" FontSize="20" Margin="3,0,0,0" VerticalAlignment="Center" IsHitTestVisible="False" Foreground="{DynamicResource Dim}">
                <TextBlock.Style><Style TargetType="TextBlock"><Setter Property="Visibility" Value="Collapsed"/>
                  <Style.Triggers><DataTrigger Binding="{Binding Text, ElementName=Box}" Value=""><Setter Property="Visibility" Value="Visible"/></DataTrigger></Style.Triggers>
                </Style></TextBlock.Style>
              </TextBlock>
            </Grid>
          </DockPanel>
          <Border Grid.Row="1" Height="1" Background="{DynamicResource Line}"/>
          <StackPanel x:Name="Header" Grid.Row="2" Orientation="Horizontal" Margin="18,12,18,2">
            <TextBlock x:Name="RdName" FontSize="15" FontWeight="Bold"/>
            <TextBlock x:Name="RdPath" FontSize="12" Margin="10,0,0,1" VerticalAlignment="Bottom" TextTrimming="CharacterEllipsis" Foreground="{DynamicResource Muted}"/>
          </StackPanel>
          <TextBlock x:Name="Message" Grid.Row="3" Margin="18,8,18,2" FontSize="13" TextWrapping="Wrap"/>
          <ItemsControl x:Name="Rows" Grid.Row="4" Margin="8,6,8,10">
            <ItemsControl.ItemsPanel><ItemsPanelTemplate><UniformGrid Columns="1"/></ItemsPanelTemplate></ItemsControl.ItemsPanel>
            <ItemsControl.ItemTemplate><DataTemplate>
              <Border x:Name="R" Height="40" Margin="0,1" Padding="12,0" CornerRadius="8" Background="Transparent" Cursor="Hand">
                <DockPanel>
                  <Path x:Name="P" Style="{StaticResource Ico}" Width="18" Height="18" Margin="0,0,12,0" Stroke="{DynamicResource Text}" Data="{Binding Geo}"/>
                  <Border x:Name="K" DockPanel.Dock="Right" Style="{StaticResource Kbd}">
                    <TextBlock x:Name="KT" Text="{Binding Hint}" FontSize="11" Foreground="{DynamicResource KbdText}"/>
                  </Border>
                  <StackPanel Orientation="Horizontal" VerticalAlignment="Center">
                    <TextBlock x:Name="T" Text="{Binding Text}" FontSize="14" Foreground="{DynamicResource Text}"/>
                    <TextBlock Text="{Binding Sub}" FontSize="12" Margin="8,0,0,0" VerticalAlignment="Center" Foreground="{DynamicResource Muted}"/>
                  </StackPanel>
                </DockPanel>
              </Border>
              <DataTemplate.Triggers>
                <DataTrigger Binding="{Binding Sel}" Value="True">
                  <Setter TargetName="R" Property="Background" Value="{DynamicResource Sel}"/>
                  <Setter TargetName="P" Property="Stroke" Value="{DynamicResource SelText}"/>
                  <Setter TargetName="T" Property="Foreground" Value="{DynamicResource SelText}"/>
                  <Setter TargetName="T" Property="FontWeight" Value="SemiBold"/>
                  <Setter TargetName="K" Property="BorderBrush" Value="{DynamicResource SelKbd}"/>
                  <Setter TargetName="KT" Property="Foreground" Value="{DynamicResource SelText}"/>
                </DataTrigger>
                <DataTrigger Binding="{Binding On}" Value="False">
                  <Setter TargetName="R" Property="Cursor" Value="Arrow"/>
                  <Setter TargetName="P" Property="Stroke" Value="{DynamicResource Dim}"/>
                  <Setter TargetName="T" Property="Foreground" Value="{DynamicResource Dim}"/>
                </DataTrigger>
                <DataTrigger Binding="{Binding Hint}" Value=""><Setter TargetName="K" Property="Visibility" Value="Collapsed"/></DataTrigger>
              </DataTemplate.Triggers>
            </DataTemplate></ItemsControl.ItemTemplate>
          </ItemsControl>
        </Grid>
      </Grid>
    </Border>
  </Grid>
</Window>
'@)
foreach ($n in 'Box', 'Header', 'RdName', 'RdPath', 'Message', 'Rows', 'Links', 'LinksTitle', 'ThemeBtn', 'ThemeIco', 'KeyBtn', 'Cog') {
    Set-Variable $n $window.FindName($n)
}

$icons = @{}
@{
    folder = 'M3,7 A2,2 0 0 1 5,5 L9,5 L11,7 L19,7 A2,2 0 0 1 21,9 L21,17 A2,2 0 0 1 19,19 L5,19 A2,2 0 0 1 3,17 Z'
    sheet  = 'M6,3 L18,3 A2,2 0 0 1 20,5 L20,19 A2,2 0 0 1 18,21 L6,21 A2,2 0 0 1 4,19 L4,5 A2,2 0 0 1 6,3 Z M4,9 L20,9 M4,15 L20,15 M10,3 L10,21'
    globe  = 'M21,12 A9,9 0 1 1 3,12 A9,9 0 1 1 21,12 Z M3,12 L21,12 M12,3 C8,7 8,17 12,21 M12,3 C16,7 16,17 12,21'
    calc   = 'M6,2 L18,2 A2,2 0 0 1 20,4 L20,20 A2,2 0 0 1 18,22 L6,22 A2,2 0 0 1 4,20 L4,4 A2,2 0 0 1 6,2 Z M8,6 L16,6 L16,10 L8,10 Z M8,14 L8.01,14 M12,14 L12.01,14 M16,14 L16.01,14 M8,18 L8.01,18 M12,18 L12.01,18 M16,18 L16.01,18'
    doc    = 'M14,3 L7,3 A2,2 0 0 0 5,5 L5,19 A2,2 0 0 0 7,21 L17,21 A2,2 0 0 0 19,19 L19,8 Z M14,3 L14,8 L19,8 M9,13 L15,13 M9,17 L15,17'
    box    = 'M21,8 L12,3 L3,8 L3,16 L12,21 L21,16 Z M3,8 L12,13 L21,8 M12,13 L12,21'
    moon   = 'M21,12.79 A9,9 0 1 1 11.21,3 A7,7 0 0 0 21,12.79 Z'
    sun    = 'M16,12 A4,4 0 1 1 8,12 A4,4 0 1 1 16,12 Z M12,2 L12,4 M12,20 L12,22 M4.93,4.93 L6.34,6.34 M17.66,17.66 L19.07,19.07 M2,12 L4,12 M20,12 L22,12 M4.93,19.07 L6.34,17.66 M17.66,6.34 L19.07,4.93'
    key    = 'M3,6 L21,6 A1,1 0 0 1 22,7 L22,17 A1,1 0 0 1 21,18 L3,18 A1,1 0 0 1 2,17 L2,7 A1,1 0 0 1 3,6 Z M6,10 L6.01,10 M10,10 L10.01,10 M14,10 L14.01,10 M18,10 L18.01,10 M7,14 L17,14'
    cog    = 'M15,12 A3,3 0 1 1 9,12 A3,3 0 1 1 15,12 Z M19.4,15 a1.65,1.65 0 0 0 .33,1.82 l.06,.06 a2,2 0 1 1 -2.83,2.83 l-.06,-.06 a1.65,1.65 0 0 0 -1.82,-.33 a1.65,1.65 0 0 0 -1,1.51 V21 a2,2 0 1 1 -4,0 v-.09 A1.65,1.65 0 0 0 9,19.4 a1.65,1.65 0 0 0 -1.82,.33 l-.06,.06 a2,2 0 1 1 -2.83,-2.83 l.06,-.06 a1.65,1.65 0 0 0 .33,-1.82 a1.65,1.65 0 0 0 -1.51,-1 H3 a2,2 0 1 1 0,-4 h.09 A1.65,1.65 0 0 0 4.6,9 a1.65,1.65 0 0 0 -.33,-1.82 l-.06,-.06 a2,2 0 1 1 2.83,-2.83 l.06,.06 A1.65,1.65 0 0 0 9,4.6 a1.65,1.65 0 0 0 1,-1.51 V3 a2,2 0 1 1 4,0 v.09 a1.65,1.65 0 0 0 1,1.51 a1.65,1.65 0 0 0 1.82,-.33 l.06,-.06 a2,2 0 1 1 2.83,2.83 l-.06,.06 A1.65,1.65 0 0 0 19.4,9 a1.65,1.65 0 0 0 1.51,1 H21 a2,2 0 1 1 0,4 h-.09 a1.65,1.65 0 0 0 -1.51,1 Z'
}.GetEnumerator() | ForEach-Object { $icons[$_.Key] = [Windows.Media.Geometry]::Parse($_.Value) }
$window.FindName('KeyIco').Data = $icons.key
$window.FindName('CogIco').Data = $icons.cog
$linkIcons = @{ calculatrice = 'calc'; document = 'doc'; carton = 'box'; tableau = 'sheet'; web = 'globe'; dossier = 'folder' }

# ---------------------------------------------------------------- themes (each PC remembers its choice; first time follows Windows)
$Themes = @{
    light = @{ Card = '#FBFBFC'; CardLine = '#D5DAE0'; Side = '#F0F2F5'; Line = '#E4E7EB'; Text = '#1D232A'; Muted = '#5A6672'; Dim = '#8A939C'
               KbdLine = '#CFD5DC'; KbdText = '#4F5A66'; Hover = '#EEF1F5'; Sel = '#E3EDFB'; SelText = '#0F3D78'; SelKbd = '#A9C3E8'; Error = '#B3261E' }
    dark  = @{ Card = '#1F242B'; CardLine = '#3A424D'; Side = '#191D23'; Line = '#2E353E'; Text = '#E8EBEF'; Muted = '#A3ADB9'; Dim = '#6E7884'
               KbdLine = '#46505C'; KbdText = '#AEB8C3'; Hover = '#2B323B'; Sel = '#24364F'; SelText = '#CFE3FF'; SelKbd = '#4C6E9A'; Error = '#FF8A80' }
}
function Set-Theme($name) {
    $script:themeName = $name
    $conv = New-Object Windows.Media.BrushConverter
    foreach ($c in $Themes[$name].GetEnumerator()) { $window.Resources[$c.Key] = $conv.ConvertFromString($c.Value) }
    $ThemeIco.Data = $icons[$(if ($name -eq 'dark') { 'sun' } else { 'moon' })]
    $ThemeBtn.ToolTip = if ($name -eq 'dark') { 'Passer en mode clair' } else { 'Passer en mode sombre' }
}
function Save-Pref($file, $value) { try { $null = New-Item -ItemType Directory -Force $appDir; Set-Content -LiteralPath $file $value -Encoding ASCII } catch { } }

# ---------------------------------------------------------------- state and rendering
$S = @{ Sel = -1; Rd = $null; Folder = $null; Matches = @{}; Pick = $null; PickLabel = ''; Visible = $false; Quitting = $false; Capture = $false }
$fg = [Windows.Controls.TextBlock]::ForegroundProperty

function Set-Message($text, [switch]$IsError) {
    $Message.Text = $text
    $Message.Visibility = if ($text) { 'Visible' } else { 'Collapsed' }
    $Message.SetResourceReference($fg, $(if ($IsError) { 'Error' } else { 'Muted' }))
}

# Indexes of the rows that can be used right now
function Get-Live {
    if ($S.Pick) { return 0..($S.Pick.Count - 1) }
    if ($S.Folder) { for ($i = 0; $i -lt $buttons.Count; $i++) { if ($buttons[$i].Type -eq 'folder' -or $S.Matches[$i].Count) { $i } } }
}

function New-Row($i, $geo, $text, $sub, $on, $key) {
    $sel = $on -and $i -eq $S.Sel
    [pscustomobject]@{ I = $i; Geo = $geo; Text = $text; Sub = $sub; On = [bool]$on; Sel = [bool]$sel; Hint = $(if ($sel) { 'Entrée' } else { $key }) }
}

function Render {
    $Header.Visibility = if ($S.Rd) { 'Visible' } else { 'Collapsed' }
    $RdName.Text = if ($S.Folder) { $S.Folder.Name } else { "$($S.Rd) introuvable" }
    $RdName.SetResourceReference($fg, $(if ($S.Folder) { 'Text' } else { 'Error' }))
    $RdPath.Text = if ($S.Folder) { $S.Folder.FullName } else { "dans $root" }
    $live = @(Get-Live)
    if ($S.Pick) {
        Set-Message "$($S.Pick.Count) fichiers $($S.PickLabel) trouvés - choisissez-en un (Échap pour revenir)"
        $Rows.ItemsSource = @(for ($i = 0; $i -lt $S.Pick.Count; $i++) {
            New-Row $i $icons.sheet $S.Pick[$i].Name $S.Pick[$i].LastWriteTime.ToString('yyyy-MM-dd HH:mm') $true ''
        })
        return
    }
    Set-Message $(if (-not $S.Rd) { 'Tapez un numéro de RD, par ex. 7 ou RD007' })
    $Rows.ItemsSource = @(for ($i = 0; $i -lt $buttons.Count; $i++) {
        $b = $buttons[$i]; $m = $S.Matches[$i]; $on = $live -contains $i
        $key = if ($b.Shortcut) { "Ctrl+$("$($b.Shortcut)".ToUpper())" } else { "Ctrl+$($i + 1)" }
        switch ($b.Type) {   # buttons of any other type are not set up yet: hidden
            'folder' { New-Row $i $icons.folder 'Ouvrir le dossier' '' $on $key }
            'file'   {
                $sub = if (-not $S.Folder) { '' } elseif ($m.Count -gt 1) { "$($m.Count) fichiers" } elseif ($m) { $m[0].Name } else { 'aucun fichier trouvé' }
                New-Row $i $icons.sheet "Ouvrir $($b.Label)" $sub $on $key
            }
        }
    })
}

function Update-FromInput {
    $S.Pick = $null
    $S.Rd = Get-RdName $Box.Text
    $S.Folder = if ($S.Rd) { Find-RdFolder $S.Rd }
    $S.Matches = @{}
    if ($S.Folder) { for ($i = 0; $i -lt $buttons.Count; $i++) { if ($buttons[$i].Type -eq 'file') { $S.Matches[$i] = Find-Files $S.Folder $buttons[$i] } } }
    $S.Sel = (@(Get-Live) + -1)[0]
    Render
}

function Move-Sel($step) {
    $live = @(Get-Live)
    if (-not $live.Count) { return }
    $S.Sel = $live[([array]::IndexOf($live, $S.Sel) + $step + $live.Count) % $live.Count]
    Render
}

function Build-Links {
    $n = 0
    $Links.ItemsSource = @(foreach ($l in $config.QuickLinks) {
        $n++
        if (-not "$($l.Path)".Trim()) { continue }   # no address yet: hidden
        $ico = $linkIcons["$($l.Icon)".ToLower()]
        if (-not $ico) { $ico = if ($l.Path -match '^[a-z]+://') { 'globe' } else { 'folder' } }
        [pscustomobject]@{ N = $n; Name = $(if ($l.Name) { $l.Name } else { "Lien $n" }); Path = $l.Path; Tip = "$($l.Path)   (Alt+$n)"; Geo = $icons[$ico] }
    })
    $LinksTitle.Visibility = if ($Links.Items.Count) { 'Visible' } else { 'Collapsed' }
}

# ---------------------------------------------------------------- actions
function Hide-Bar {
    if ($S.Capture) { Complete-Capture '' }
    if ($S.Visible) { $S.Visible = $false; $window.Hide() }
}

function Show-Bar {
    if (-not $S.Visible) {
        Read-Config
        Build-Links
        if ($Box.Text) { $Box.Text = '' } else { Update-FromInput }
        $area = [Windows.SystemParameters]::WorkArea
        $window.Left = $area.Left + ($area.Width - $window.Width) / 2
        $window.Top  = $area.Top + $area.Height * 0.2
        $S.Visible = $true
        $window.Show()
    }
    [void]$window.Activate()
    [void]$Box.Focus()
}

function Open-Target($path) {
    $path = [Environment]::ExpandEnvironmentVariables($path)
    try { Start-Process -FilePath $path -EA Stop; Hide-Bar }
    catch { Set-Message "Impossible d'ouvrir $path - $($_.Exception.Message)" -IsError }
}

function Open-AppFiles {
    try { Start-Process explorer.exe "/select,`"$(Join-Path $SourceDir 'RD_Browser.config.json')`"" -EA Stop; Hide-Bar }
    catch { Set-Message "Impossible d'ouvrir $SourceDir" -IsError }
}

function Invoke-Row($i) {
    if ($S.Pick) { return Open-Target $S.Pick[$i].FullName }
    if (@(Get-Live) -notcontains $i) { return }
    $m = @($S.Matches[$i])
    if ($buttons[$i].Type -eq 'folder') { Open-Target $S.Folder.FullName }
    elseif ($m.Count -eq 1) { Open-Target $m[0].FullName }
    else { $S.Pick = $m; $S.PickLabel = $buttons[$i].Label; $S.Sel = 0; Render }
}

# ---------------------------------------------------------------- changing the shortcut (keyboard button)
function Update-HotkeyLabels {
    $KeyBtn.ToolTip = "Changer le raccourci (actuel : $hotkey)"
    if ($tray) { $tray.Text = "RD Browser ($hotkey)"; $miOpen.Text = "Ouvrir RD Browser`t$hotkey" }
}

function Start-Capture {
    $S.Capture = $true
    if ($hk) { $hk.Dispose(); $script:hk = $null }   # free it so the same keys can be pressed again
    $KeyBtn.SetResourceReference([Windows.Controls.Border]::BackgroundProperty, 'Sel')
    Set-Message "Appuyez sur le nouveau raccourci (actuel : $hotkey).`nÉchap = annuler · Retour arrière = par défaut ($defaultHotkey)"
}

function Complete-Capture($candidate, [switch]$Reset) {
    $S.Capture = $false
    $KeyBtn.ClearValue([Windows.Controls.Border]::BackgroundProperty)
    if ($candidate) {
        if ($h = Register-Hotkey $candidate) {
            $script:hk = $h
            $script:hotkey = $candidate
            if ($Reset) { Remove-Item -LiteralPath $hotkeyFile -EA 0 } else { Save-Pref $hotkeyFile $candidate }
            Update-HotkeyLabels
            return Set-Message "Raccourci enregistré : $candidate"
        }
        Set-Message "Le raccourci $candidate est déjà utilisé par un autre programme. Cliquez à nouveau sur le clavier pour en choisir un autre." -IsError
    }
    $script:hk = Register-Hotkey $hotkey   # put the current shortcut back
}

# ---------------------------------------------------------------- events
$Box.Add_TextChanged({ Update-FromInput })
$Rows.Add_MouseLeftButtonUp({ param($src, $e) $d = $e.OriginalSource.DataContext; if ($d.On) { Invoke-Row $d.I } })
$Links.Add_MouseLeftButtonUp({ param($src, $e) $d = $e.OriginalSource.DataContext; if ($d.Path) { Open-Target $d.Path } })
$Cog.Add_MouseLeftButtonUp({ Open-AppFiles })
$KeyBtn.Add_MouseLeftButtonUp({ if ($S.Capture) { Complete-Capture ''; Render } else { Start-Capture } })
$ThemeBtn.Add_MouseLeftButtonUp({ $t = if ($themeName -eq 'dark') { 'light' } else { 'dark' }; Set-Theme $t; Save-Pref $themeFile $t })
$window.Add_Deactivated({ Hide-Bar })   # clicking elsewhere hides the bar, like the Start menu
$window.Add_Closing({ param($src, $e) if (-not $S.Quitting) { $e.Cancel = $true; Hide-Bar } })   # Alt+F4 hides

$window.Add_PreviewKeyDown({
    param($src, $e)
    $key  = "$(if ($e.Key -eq 'System') { $e.SystemKey } else { $e.Key })"
    $mods = [Windows.Input.Keyboard]::Modifiers
    $e.Handled = $true

    if ($S.Capture) {   # recording a new shortcut: every key comes here
        if ($key -eq 'Escape') { Complete-Capture ''; return Render }
        if ($key -eq 'Back')   { return Complete-Capture $defaultHotkey -Reset }
        if ($key -match '(Ctrl|Alt|Shift|Win)$|^None$') { return }
        $parts = @(foreach ($m in 'Control', 'Alt', 'Shift', 'Windows') { if ($mods -band [Windows.Input.ModifierKeys]$m) { $m -replace 'Control', 'Ctrl' -replace 'Windows', 'Win' } })
        if (-not ($parts -match 'Ctrl|Alt|Win')) { return Set-Message 'Ajoutez Ctrl, Alt ou Win à la touche (ex. Ctrl+Alt+K). Échap = annuler.' -IsError }
        return Complete-Capture (($parts + ($key -replace '^D(\d)$', '$1')) -join '+')
    }

    switch ($key) {
        'Escape' { if ($S.Pick) { $S.Pick = $null; $S.Sel = (@(Get-Live) + -1)[0]; Render } else { Hide-Bar }; return }
        'Down'   { return Move-Sel 1 }
        'Up'     { return Move-Sel -1 }
        'Return' { if ($S.Sel -ge 0) { Invoke-Row $S.Sel }; return }
    }
    $digit = if ($key -match '^(D|NumPad)(\d)$') { [int]$Matches[2] }
    if ($digit -and $mods -eq 'Alt') {
        $ql = @($config.QuickLinks)
        if ($digit -le $ql.Count -and "$($ql[$digit - 1].Path)".Trim()) { Open-Target $ql[$digit - 1].Path }
        return
    }
    if ($mods -eq 'Control' -and -not $S.Pick) {
        for ($i = 0; $i -lt $buttons.Count; $i++) {
            if ($digit -eq $i + 1 -or ($buttons[$i].Shortcut -and $key -eq "$($buttons[$i].Shortcut)")) { return Invoke-Row $i }
        }
    }
    $e.Handled = $false
})

$saved = Get-Content -LiteralPath $themeFile -EA 0 | Select-Object -First 1
Set-Theme $(if ($saved -in 'dark', 'light') { $saved }
            elseif ((Get-ItemPropertyValue 'HKCU:\Software\Microsoft\Windows\CurrentVersion\Themes\Personalize' AppsUseLightTheme -EA 0) -eq 0) { 'dark' }
            else { 'light' })

# ---------------------------------------------------------------- tray icon and menu
[Windows.Forms.Application]::EnableVisualStyles()
$bmp = New-Object Drawing.Bitmap 32, 32
$g = [Drawing.Graphics]::FromImage($bmp)
$g.TextRenderingHint = 'AntiAliasGridFit'
$g.Clear([Drawing.Color]::FromArgb(11, 98, 196))
$fmt = New-Object Drawing.StringFormat
$fmt.Alignment = $fmt.LineAlignment = 'Center'
$g.DrawString('RD', (New-Object Drawing.Font('Segoe UI', 15, 'Bold', 'Pixel')), [Drawing.Brushes]::White, (New-Object Drawing.RectangleF(0, 1, 32, 32)), $fmt)
$g.Dispose()

$menu = New-Object Windows.Forms.ContextMenuStrip
function Add-MenuItem($text, $action) { $item = $menu.Items.Add($text); $item.Add_Click($action); $item }
$miOpen = Add-MenuItem 'Ouvrir RD Browser' { Show-Bar }
$miOpen.Font = New-Object Drawing.Font($miOpen.Font, 'Bold')
$null = Add-MenuItem "Ouvrir l'emplacement des dossiers RD" {
    Read-Config
    try { Start-Process -FilePath $root -EA Stop } catch { $tray.ShowBalloonTip(4000, 'RD Browser', "Impossible d'accéder à $root", 'Warning') }
}
$null = Add-MenuItem 'Ouvrir les fichiers de RD Browser (réglages)' { Open-AppFiles }
$null = $menu.Items.Add((New-Object Windows.Forms.ToolStripSeparator))
$null = Add-MenuItem "Quitter (jusqu'à la prochaine connexion)" {
    $S.Quitting = $true
    if ($hk) { $hk.Dispose() }
    $tray.Dispose()
    $window.Close()
    [Windows.Threading.Dispatcher]::CurrentDispatcher.InvokeShutdown()
}

$tray = New-Object Windows.Forms.NotifyIcon
$tray.Icon = [Drawing.Icon]::FromHandle($bmp.GetHicon())
$tray.ContextMenuStrip = $menu
$tray.Add_MouseClick({ param($src, $e) if ($e.Button -eq 'Left') { Show-Bar } })
$tray.Visible = $true

$hk = Register-Hotkey $hotkey
Update-HotkeyLabels
if (-not $hk) {
    $tray.ShowBalloonTip(6000, 'RD Browser', "Le raccourci $hotkey n'a pas pu être activé (un autre programme l'utilise peut-être). Cliquez plutôt sur l'icône RD près de l'horloge, ou changez-le avec le bouton clavier de la barre.", 'Warning')
} elseif ($Welcome) {
    $tray.ShowBalloonTip(5000, 'RD Browser est prêt', "Appuyez sur $hotkey n'importe où pour l'ouvrir. Il reste ici, près de l'horloge.", 'Info')
}

[System.Windows.Threading.Dispatcher]::Run()
