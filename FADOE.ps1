<#
.SYNOPSIS
    FADOE - Find A Damn Outlook Email. A small window for both tools in this folder.

.DESCRIPTION
    Find an email          The exact message in a reply chain where your words were written.
                           Read it right here, copy a search string for new Outlook, or open
                           just that one message in its own window.
    Catch up on a person   Every email with them over the last N days, browsable here and
                           copied to the clipboard as one document to paste into Claude.

    Uses Find-Email.ps1, FadoeIndex.ps1 and Get-EmailContext.ps1 from the same folder. The
    search index updates in the background when the window opens and stays loaded, so
    searches take well under a second. Read-only: nothing is ever sent, moved, or deleted.

.PARAMETER InstallShortcut
    Put "FADOE" shortcuts on your desktop and in the Start menu that open this window, then
    exit. Also writes FADOE.ico (the logo) next to this script for the shortcut icon.

.PARAMETER Theme
    auto (follow Windows light/dark mode), light, or dark.

.PARAMETER ScreenshotPath
    For testing: open the window off-screen, run ScreenshotTab with ScreenshotQuery, save a
    PNG of the result, and exit. Never touches the clipboard.

.EXAMPLE
    powershell -ExecutionPolicy Bypass -File FADOE.ps1 -InstallShortcut
#>

[CmdletBinding()]
param(
    [switch] $InstallShortcut,
    [ValidateSet('auto', 'light', 'dark')] [string] $Theme = 'auto',
    [string] $ScreenshotPath,
    [ValidateSet('find', 'catch', 'empty')] [string] $ScreenshotTab = 'find',
    [string] $ScreenshotQuery = '',
    [int]    $ScreenshotDays = 180
)

$FadoeVersion = '1.1.0'

$ErrorActionPreference = 'Stop'
$here      = $PSScriptRoot
$findTool  = Join-Path $here 'Find-Email.ps1'
$ctxTool   = Join-Path $here 'Get-EmailContext.ps1'
$ctxFolder = Join-Path $here 'context'
$dot       = ' ' + [char]0x00B7 + ' '

Add-Type -AssemblyName PresentationFramework, PresentationCore, WindowsBase

# ---------------------------------------------------------------- logo ----
# 128x128 design grid: gradient tile, envelope, and a magnifier whose lens enlarges the
# envelope's fold (it doubles as a "found it" checkmark), plus a sparkle.
$logoXaml = @'
<Canvas xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation" Width="128" Height="128">
  <Canvas.Clip><RectangleGeometry Rect="0,0,128,128" RadiusX="30" RadiusY="30"/></Canvas.Clip>
  <Rectangle Width="128" Height="128">
    <Rectangle.Fill>
      <LinearGradientBrush StartPoint="0,0" EndPoint="1,1">
        <GradientStop Color="#2F86FF" Offset="0"/>
        <GradientStop Color="#5A45E6" Offset="0.55"/>
        <GradientStop Color="#9B34D6" Offset="1"/>
      </LinearGradientBrush>
    </Rectangle.Fill>
  </Rectangle>
  <Ellipse Canvas.Left="-30" Canvas.Top="-70" Width="150" Height="120" Opacity="0.13" Fill="White"/>
  <Path Stroke="White" StrokeThickness="7" StrokeLineJoin="Round"
        Data="M 30,30 H 82 A 8,8 0 0 1 90,38 V 76 A 8,8 0 0 1 82,84 H 30 A 8,8 0 0 1 22,76 V 38 A 8,8 0 0 1 30,30 Z"/>
  <Path Stroke="White" StrokeThickness="7" StrokeLineJoin="Round" StrokeStartLineCap="Round" StrokeEndLineCap="Round"
        Data="M 27,37 L 56,59 L 85,37"/>
  <Ellipse Canvas.Left="62" Canvas.Top="60" Width="46" Height="46">
    <Ellipse.Fill>
      <RadialGradientBrush GradientOrigin="0.35,0.3">
        <GradientStop Color="#8A63F7" Offset="0"/>
        <GradientStop Color="#6A3FE0" Offset="1"/>
      </RadialGradientBrush>
    </Ellipse.Fill>
  </Ellipse>
  <Path Stroke="White" StrokeThickness="5" StrokeLineJoin="Round" StrokeStartLineCap="Round" StrokeEndLineCap="Round" Opacity="0.95"
        Data="M 66,74 L 85,88 L 104,74">
    <Path.Clip><EllipseGeometry Center="85,83" RadiusX="19" RadiusY="19"/></Path.Clip>
  </Path>
  <Ellipse Canvas.Left="62" Canvas.Top="60" Width="46" Height="46" Stroke="White" StrokeThickness="7.5"/>
  <Line X1="101" Y1="99" X2="114" Y2="112" Stroke="White" StrokeThickness="11" StrokeStartLineCap="Round" StrokeEndLineCap="Round"/>
  <Path Fill="White" Data="M 104,14 Q 105.5,22.5 114,24 Q 105.5,25.5 104,34 Q 102.5,25.5 94,24 Q 102.5,22.5 104,14 Z"/>
</Canvas>
'@

function New-Logo([double]$px) {
    $vb = New-Object System.Windows.Controls.Viewbox
    $vb.Width = $px; $vb.Height = $px
    $vb.Child = [Windows.Markup.XamlReader]::Parse($logoXaml)
    $vb
}

function Get-LogoPng([int]$px) {
    $vb = New-Logo $px
    $vb.Measure([System.Windows.Size]::new($px, $px))
    $vb.Arrange([System.Windows.Rect]::new(0, 0, $px, $px))
    $vb.UpdateLayout()
    $bmp = New-Object System.Windows.Media.Imaging.RenderTargetBitmap($px, $px, 96, 96, [System.Windows.Media.PixelFormats]::Pbgra32)
    $bmp.Render($vb)
    $enc = New-Object System.Windows.Media.Imaging.PngBitmapEncoder
    $enc.Frames.Add([System.Windows.Media.Imaging.BitmapFrame]::Create($bmp))
    $ms = New-Object System.IO.MemoryStream
    $enc.Save($ms)
    , $ms.ToArray()
}

# Classic 32-bit bitmap frame for an .ico: BITMAPINFOHEADER, bottom-up BGRA rows, empty AND mask.
function Get-LogoIcoDib([int]$px) {
    $vb = New-Logo $px
    $vb.Measure([System.Windows.Size]::new($px, $px))
    $vb.Arrange([System.Windows.Rect]::new(0, 0, $px, $px))
    $vb.UpdateLayout()
    $bmp = New-Object System.Windows.Media.Imaging.RenderTargetBitmap($px, $px, 96, 96, [System.Windows.Media.PixelFormats]::Pbgra32)
    $bmp.Render($vb)
    $conv = New-Object System.Windows.Media.Imaging.FormatConvertedBitmap($bmp, [System.Windows.Media.PixelFormats]::Bgra32, $null, 0)
    $stride = $px * 4
    $pixels = New-Object byte[] ($stride * $px)
    $conv.CopyPixels($pixels, $stride, 0)
    $maskRow = [int](([Math]::Ceiling($px / 32.0)) * 4)
    $ms = New-Object System.IO.MemoryStream
    $bw = New-Object System.IO.BinaryWriter($ms)
    $bw.Write([uint32]40); $bw.Write([int32]$px); $bw.Write([int32]($px * 2))
    $bw.Write([uint16]1); $bw.Write([uint16]32); $bw.Write([uint32]0)
    $bw.Write([uint32]($stride * $px + $maskRow * $px))
    $bw.Write([int32]0); $bw.Write([int32]0); $bw.Write([uint32]0); $bw.Write([uint32]0)
    for ($y = $px - 1; $y -ge 0; $y--) { $bw.Write($pixels, $y * $stride, $stride) }
    $bw.Write((New-Object byte[] ($maskRow * $px)))
    $bw.Flush()
    , $ms.ToArray()
}

# .ico for the shortcut: bitmap frames for the small sizes, PNG for 256 (the standard layout).
function Save-LogoIco([string]$path) {
    $sizes = 16, 20, 24, 32, 40, 48, 64, 256
    $frames = @(foreach ($s in $sizes) { if ($s -ge 256) { , (Get-LogoPng $s) } else { , (Get-LogoIcoDib $s) } })
    $ms = New-Object System.IO.MemoryStream
    $bw = New-Object System.IO.BinaryWriter($ms)
    $bw.Write([uint16]0); $bw.Write([uint16]1); $bw.Write([uint16]$sizes.Count)
    $offset = 6 + 16 * $sizes.Count
    for ($i = 0; $i -lt $sizes.Count; $i++) {
        $dim = [byte]($sizes[$i] % 256)   # 256 is written as 0
        $bw.Write($dim); $bw.Write($dim); $bw.Write([byte]0); $bw.Write([byte]0)
        $bw.Write([uint16]1); $bw.Write([uint16]32)
        $bw.Write([uint32]$frames[$i].Length); $bw.Write([uint32]$offset)
        $offset += $frames[$i].Length
    }
    foreach ($f in $frames) { $bw.Write($f) }
    $bw.Flush()
    [System.IO.File]::WriteAllBytes($path, $ms.ToArray())
}

# ------------------------------------------------------------ shortcut ----
if ($InstallShortcut) {
    $ico = Join-Path $here 'FADOE.ico'
    Save-LogoIco $ico
    $ws = New-Object -ComObject WScript.Shell
    foreach ($folder in 'Desktop', 'Programs') {
        $lnkPath = Join-Path ([Environment]::GetFolderPath($folder)) 'FADOE.lnk'
        $lnk = $ws.CreateShortcut($lnkPath)
        $lnk.TargetPath       = Join-Path $env:SystemRoot 'System32\WindowsPowerShell\v1.0\powershell.exe'
        $lnk.Arguments        = '-NoProfile -ExecutionPolicy Bypass -WindowStyle Hidden -File "' + $PSCommandPath + '"'
        $lnk.WorkingDirectory = $here
        $lnk.IconLocation     = $ico + ',0'
        $lnk.Description      = 'FADOE - Find A Damn Outlook Email, or catch up on a person'
        $lnk.Save()
        Write-Host "Shortcut created: $lnkPath"
    }
    return
}

# ------------------------------------------------------- single instance ----
# A second launch (double-click during a slow start, desktop + Start menu...) just brings the
# window that's already open to the front. Two copies would double the load on Outlook.
if (-not $ScreenshotPath) {
    $createdNew = $false
    $script:instanceLock = New-Object System.Threading.Mutex($true, 'Local\FADOE-FindADamnOutlookEmail', [ref]$createdNew)
    if (-not $createdNew) {
        try { [void](New-Object -ComObject WScript.Shell).AppActivate('FADOE - Find A Damn Outlook Email') } catch { }
        exit 0
    }
}

# -------------------------------------------------------------- splash ----
# Runs on its own thread so its progress bar keeps moving while the main window loads.
# It stays up at least ~1 second so it never just flickers.
$splashXaml = @'
<Window xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
        Title="FADOE" Width="560" Height="340" WindowStyle="None" AllowsTransparency="True" Background="Transparent"
        ResizeMode="NoResize" WindowStartupLocation="CenterScreen" Topmost="True" ShowInTaskbar="False">
  <Border Margin="20" CornerRadius="18">
    <Border.Effect><DropShadowEffect BlurRadius="26" ShadowDepth="5" Opacity="0.5" Color="Black"/></Border.Effect>
    <Border.Background>
      <LinearGradientBrush StartPoint="0,0" EndPoint="1,1">
        <GradientStop Color="#181B23" Offset="0"/>
        <GradientStop Color="#222739" Offset="1"/>
      </LinearGradientBrush>
    </Border.Background>
    <Grid Margin="40,38,40,30">
      <Grid.RowDefinitions>
        <RowDefinition Height="Auto"/><RowDefinition Height="*"/><RowDefinition Height="Auto"/><RowDefinition Height="Auto"/>
      </Grid.RowDefinitions>
      <StackPanel Orientation="Horizontal">
        <Viewbox x:Name="LogoBox" Width="92" Height="92"/>
        <StackPanel Margin="22,4,0,0" VerticalAlignment="Center">
          <StackPanel Orientation="Horizontal">
            <TextBlock Text="FAD" FontFamily="Segoe UI" FontWeight="Black" FontSize="48" Foreground="White"/>
            <Canvas Width="40" Height="48" Margin="2,0,1,0">
              <Ellipse Canvas.Left="3" Canvas.Top="13" Width="29" Height="29" Stroke="White" StrokeThickness="7"/>
              <Line X1="28" Y1="38" X2="37" Y2="47" Stroke="White" StrokeThickness="7.5" StrokeStartLineCap="Round" StrokeEndLineCap="Round"/>
            </Canvas>
            <TextBlock Text="E" FontFamily="Segoe UI" FontWeight="Black" FontSize="48" Foreground="White"/>
          </StackPanel>
          <TextBlock Text="Find A Damn Outlook Email" FontFamily="Segoe UI" FontSize="15" Foreground="#A9B1C6" Margin="2,-2,0,0"/>
        </StackPanel>
      </StackPanel>
      <Canvas Grid.Row="2" Height="4" ClipToBounds="True" Background="#2C3246">
        <Rectangle Canvas.Left="-140" Width="140" Height="4" RadiusX="2" RadiusY="2">
          <Rectangle.Fill>
            <LinearGradientBrush StartPoint="0,0" EndPoint="1,0">
              <GradientStop Color="#002F86FF" Offset="0"/>
              <GradientStop Color="#2F86FF" Offset="0.35"/>
              <GradientStop Color="#9B34D6" Offset="1"/>
            </LinearGradientBrush>
          </Rectangle.Fill>
          <Rectangle.Triggers>
            <EventTrigger RoutedEvent="FrameworkElement.Loaded">
              <BeginStoryboard>
                <Storyboard RepeatBehavior="Forever">
                  <DoubleAnimation Storyboard.TargetProperty="(Canvas.Left)" From="-140" To="440" Duration="0:0:1.3">
                    <DoubleAnimation.EasingFunction><SineEase EasingMode="EaseInOut"/></DoubleAnimation.EasingFunction>
                  </DoubleAnimation>
                </Storyboard>
              </BeginStoryboard>
            </EventTrigger>
          </Rectangle.Triggers>
        </Rectangle>
      </Canvas>
      <Grid Grid.Row="3" Margin="0,12,0,0">
        <TextBlock x:Name="Status" Text="Warming up the search engine..." FontFamily="Segoe UI" FontSize="12.5" Foreground="#8C94A8"/>
        <TextBlock x:Name="Version" Text="read-only" HorizontalAlignment="Right" FontFamily="Segoe UI" FontSize="12.5" Foreground="#5E6679"/>
      </Grid>
    </Grid>
  </Border>
</Window>
'@

$splashScript = {
    param($sync, $xaml, $logoXaml, $version)
    try {
        $w = [Windows.Markup.XamlReader]::Parse($xaml)
        $w.FindName('LogoBox').Child = [Windows.Markup.XamlReader]::Parse($logoXaml)
        $w.FindName('Version').Text = 'v' + $version + '  ' + [char]0x00B7 + '  read-only'
        $status = $w.FindName('Status')
        $lines = @('Warming up the search engine...', 'Untangling reply chains...',
                   'Ignoring everyone''s quoted replies...', 'Polishing the magnifying glass...')
        $clock = [Diagnostics.Stopwatch]::StartNew()
        $timer = New-Object System.Windows.Threading.DispatcherTimer
        $timer.Interval = [TimeSpan]::FromMilliseconds(50)
        $timer.Add_Tick({
            $status.Text = $lines[[int][Math]::Floor($clock.ElapsedMilliseconds / 1400) % $lines.Count]
            if (($sync.Close -and $clock.ElapsedMilliseconds -ge 900) -or $clock.ElapsedMilliseconds -ge 90000) {
                $timer.Stop(); $w.Close()
            }
        })
        $timer.Start()
        [void]$w.ShowDialog()
    } catch { }
}

$splash = [hashtable]::Synchronized(@{ Close = $false })
if (-not $ScreenshotPath) {
    try {
        $srs = [runspacefactory]::CreateRunspace()
        $srs.ApartmentState = 'STA'
        $srs.ThreadOptions = 'ReuseThread'
        $srs.Open()
        $sps = [powershell]::Create()
        $sps.Runspace = $srs
        [void]$sps.AddScript($splashScript).AddArgument($splash).AddArgument($splashXaml).AddArgument($logoXaml).AddArgument($FadoeVersion)
        [void]$sps.BeginInvoke()
    } catch { }
}

Add-Type -TypeDefinition @'
using System;
using System.Text;
using System.Runtime.InteropServices;
namespace FadoeNative {
    public static class Win {
        [DllImport("dwmapi.dll")] public static extern int DwmSetWindowAttribute(IntPtr hwnd, int attr, ref int value, int size);

        delegate bool EnumProc(IntPtr hwnd, IntPtr lParam);
        [DllImport("user32.dll")] static extern bool EnumWindows(EnumProc cb, IntPtr lParam);
        [DllImport("user32.dll")] static extern uint GetWindowThreadProcessId(IntPtr hwnd, out uint pid);
        [DllImport("user32.dll")] static extern bool IsWindowVisible(IntPtr hwnd);
        [DllImport("user32.dll", CharSet = CharSet.Unicode)] static extern int GetClassName(IntPtr hwnd, StringBuilder s, int n);
        [DllImport("user32.dll", CharSet = CharSet.Unicode)] static extern int GetWindowText(IntPtr hwnd, StringBuilder s, int n);

        // Title of a visible message box ("#32770" dialog) belonging to the process, or null.
        // Reminders, open messages and the main window are ordinary windows and don't count.
        public static string FindDialog(int pid) {
            string found = null;
            EnumWindows(delegate (IntPtr h, IntPtr l) {
                uint p; GetWindowThreadProcessId(h, out p);
                if (p != (uint)pid || !IsWindowVisible(h)) return true;
                StringBuilder c = new StringBuilder(64); GetClassName(h, c, 64);
                if (c.ToString() != "#32770") return true;
                StringBuilder t = new StringBuilder(256); GetWindowText(h, t, 256);
                found = t.ToString(); return false;
            }, IntPtr.Zero);
            return found;
        }
    }
}
'@

class FadoeRow {
    [string] $Who
    [string] $When
    [string] $Subject
    [string] $Where
    [double] $Dim = 1.0
    [object] $Data
}

# --------------------------------------------------------------- theme ----
$useDark = switch ($Theme) {
    'dark'  { $true }
    'light' { $false }
    default {
        try { (Get-ItemPropertyValue 'HKCU:\Software\Microsoft\Windows\CurrentVersion\Themes\Personalize' 'AppsUseLightTheme') -eq 0 }
        catch { $false }
    }
}
$palette = if ($useDark) {
    @{ Bg = '#1C1C1E'; Surface = '#262628'; Card2 = '#2F2F32'; Line = '#3A3A3F'; Text = '#EDEDEF'; Sub = '#A2A5AC'
       Accent = '#3B8EEA'; AccentText = '#FFFFFF'; Hover = '#303035'; Sel = '#1F3A58'; Hi = '#7A6400'; HiText = '#FFFFFF'
       Input = '#1F1F21'; Link = '#6CB6FF'; Warn = '#F2B84B'; Ok = '#5CC98A'; Scroll = '#56595F' }
} else {
    @{ Bg = '#F3F4F6'; Surface = '#FFFFFF'; Card2 = '#F6F7F9'; Line = '#E1E3E8'; Text = '#1B1C1F'; Sub = '#5F6570'
       Accent = '#0F6CBD'; AccentText = '#FFFFFF'; Hover = '#EEF1F5'; Sel = '#DCEAF8'; Hi = '#FFE36E'; HiText = '#1B1C1F'
       Input = '#FFFFFF'; Link = '#0F6CBD'; Warn = '#9A5B00'; Ok = '#107C41'; Scroll = '#C4C8CF' }
}

# ---------------------------------------------------------------- xaml ----
$xaml = @'
<Window xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
        Title="FADOE - Find A Damn Outlook Email" Width="1200" Height="800" MinWidth="920" MinHeight="580"
        WindowStartupLocation="CenterScreen" Background="{DynamicResource Bg}" Foreground="{DynamicResource Text}"
        FontFamily="Segoe UI Variable Text, Segoe UI" FontSize="13.5"
        UseLayoutRounding="True" TextOptions.TextFormattingMode="Display">
  <Window.Resources>
    <FontFamily x:Key="Icons">Segoe Fluent Icons, Segoe MDL2 Assets</FontFamily>

    <Style x:Key="Card" TargetType="Border">
      <Setter Property="Background" Value="{DynamicResource Surface}"/>
      <Setter Property="BorderBrush" Value="{DynamicResource Line}"/>
      <Setter Property="BorderThickness" Value="1"/>
      <Setter Property="CornerRadius" Value="10"/>
    </Style>

    <Style x:Key="Hint" TargetType="TextBlock">
      <Setter Property="Foreground" Value="{DynamicResource Sub}"/>
      <Setter Property="IsHitTestVisible" Value="False"/>
      <Setter Property="VerticalAlignment" Value="Center"/>
      <Setter Property="Margin" Value="13,0,10,0"/>
      <Setter Property="TextTrimming" Value="CharacterEllipsis"/>
    </Style>

    <Style x:Key="Icon" TargetType="TextBlock">
      <Setter Property="FontFamily" Value="{StaticResource Icons}"/>
      <Setter Property="VerticalAlignment" Value="Center"/>
      <Setter Property="Margin" Value="0,1,8,0"/>
    </Style>

    <Style x:Key="Meta" TargetType="TextBlock">
      <Setter Property="Foreground" Value="{DynamicResource Sub}"/>
      <Setter Property="TextWrapping" Value="Wrap"/>
      <Setter Property="Margin" Value="0,3,0,0"/>
    </Style>

    <Style TargetType="TextBox">
      <Setter Property="Foreground" Value="{DynamicResource Text}"/>
      <Setter Property="CaretBrush" Value="{DynamicResource Text}"/>
      <Setter Property="SelectionBrush" Value="{DynamicResource Accent}"/>
      <Setter Property="Padding" Value="10,9"/>
      <Setter Property="FontSize" Value="14"/>
      <Setter Property="Template">
        <Setter.Value>
          <ControlTemplate TargetType="TextBox">
            <Border x:Name="b" Background="{DynamicResource Input}" BorderBrush="{DynamicResource Line}" BorderThickness="1" CornerRadius="7">
              <ScrollViewer x:Name="PART_ContentHost" Margin="{TemplateBinding Padding}" VerticalAlignment="Center"/>
            </Border>
            <ControlTemplate.Triggers>
              <Trigger Property="IsKeyboardFocused" Value="True">
                <Setter TargetName="b" Property="BorderBrush" Value="{DynamicResource Accent}"/>
              </Trigger>
            </ControlTemplate.Triggers>
          </ControlTemplate>
        </Setter.Value>
      </Setter>
    </Style>

    <Style x:Key="Plain" TargetType="TextBox">
      <Setter Property="Foreground" Value="{DynamicResource Text}"/>
      <Setter Property="SelectionBrush" Value="{DynamicResource Accent}"/>
      <Setter Property="IsReadOnly" Value="True"/>
      <Setter Property="TextWrapping" Value="Wrap"/>
      <Setter Property="FontFamily" Value="Cascadia Mono, Consolas"/>
      <Setter Property="FontSize" Value="12.5"/>
      <Setter Property="Template">
        <Setter.Value>
          <ControlTemplate TargetType="TextBox">
            <ScrollViewer x:Name="PART_ContentHost"/>
          </ControlTemplate>
        </Setter.Value>
      </Setter>
    </Style>

    <Style x:Key="Primary" TargetType="Button">
      <Setter Property="Foreground" Value="{DynamicResource AccentText}"/>
      <Setter Property="Background" Value="{DynamicResource Accent}"/>
      <Setter Property="BorderBrush" Value="{DynamicResource Accent}"/>
      <Setter Property="BorderThickness" Value="1"/>
      <Setter Property="Padding" Value="16,9"/>
      <Setter Property="FontWeight" Value="SemiBold"/>
      <Setter Property="Cursor" Value="Hand"/>
      <Setter Property="Template">
        <Setter.Value>
          <ControlTemplate TargetType="Button">
            <Border x:Name="b" Background="{TemplateBinding Background}" BorderBrush="{TemplateBinding BorderBrush}"
                    BorderThickness="{TemplateBinding BorderThickness}" CornerRadius="7" Padding="{TemplateBinding Padding}">
              <ContentPresenter HorizontalAlignment="Center" VerticalAlignment="Center"/>
            </Border>
            <ControlTemplate.Triggers>
              <Trigger Property="IsMouseOver" Value="True"><Setter TargetName="b" Property="Opacity" Value="0.86"/></Trigger>
              <Trigger Property="IsPressed" Value="True"><Setter TargetName="b" Property="Opacity" Value="0.72"/></Trigger>
              <Trigger Property="IsEnabled" Value="False"><Setter TargetName="b" Property="Opacity" Value="0.45"/></Trigger>
            </ControlTemplate.Triggers>
          </ControlTemplate>
        </Setter.Value>
      </Setter>
    </Style>

    <Style x:Key="Secondary" TargetType="Button" BasedOn="{StaticResource Primary}">
      <Setter Property="Foreground" Value="{DynamicResource Text}"/>
      <Setter Property="Background" Value="{DynamicResource Card2}"/>
      <Setter Property="BorderBrush" Value="{DynamicResource Line}"/>
      <Setter Property="FontWeight" Value="Normal"/>
      <Setter Property="Padding" Value="14,8"/>
    </Style>

    <Style x:Key="Tab" TargetType="RadioButton">
      <Setter Property="Foreground" Value="{DynamicResource Sub}"/>
      <Setter Property="Cursor" Value="Hand"/>
      <Setter Property="FontSize" Value="14"/>
      <Setter Property="Template">
        <Setter.Value>
          <ControlTemplate TargetType="RadioButton">
            <Border x:Name="b" Background="Transparent" BorderBrush="Transparent" BorderThickness="1" CornerRadius="8" Padding="14,8">
              <ContentPresenter VerticalAlignment="Center"/>
            </Border>
            <ControlTemplate.Triggers>
              <Trigger Property="IsMouseOver" Value="True">
                <Setter TargetName="b" Property="Background" Value="{DynamicResource Hover}"/>
              </Trigger>
              <Trigger Property="IsChecked" Value="True">
                <Setter TargetName="b" Property="Background" Value="{DynamicResource Surface}"/>
                <Setter TargetName="b" Property="BorderBrush" Value="{DynamicResource Line}"/>
                <Setter Property="Foreground" Value="{DynamicResource Accent}"/>
                <Setter Property="FontWeight" Value="SemiBold"/>
              </Trigger>
            </ControlTemplate.Triggers>
          </ControlTemplate>
        </Setter.Value>
      </Setter>
    </Style>

    <Style TargetType="ListBoxItem">
      <Setter Property="Foreground" Value="{DynamicResource Text}"/>
      <Setter Property="HorizontalContentAlignment" Value="Stretch"/>
      <Setter Property="Cursor" Value="Hand"/>
      <Setter Property="Template">
        <Setter.Value>
          <ControlTemplate TargetType="ListBoxItem">
            <Border x:Name="b" Background="Transparent" CornerRadius="7" Margin="0,0,0,2">
              <ContentPresenter/>
            </Border>
            <ControlTemplate.Triggers>
              <Trigger Property="IsMouseOver" Value="True">
                <Setter TargetName="b" Property="Background" Value="{DynamicResource Hover}"/>
              </Trigger>
              <Trigger Property="IsSelected" Value="True">
                <Setter TargetName="b" Property="Background" Value="{DynamicResource Sel}"/>
              </Trigger>
            </ControlTemplate.Triggers>
          </ControlTemplate>
        </Setter.Value>
      </Setter>
    </Style>

    <Style TargetType="ProgressBar">
      <Setter Property="Height" Value="3"/>
      <Setter Property="Template">
        <Setter.Value>
          <ControlTemplate TargetType="ProgressBar">
            <Grid>
              <Border x:Name="PART_Track" Background="{DynamicResource Line}" CornerRadius="2"/>
              <Border x:Name="PART_Indicator" Background="{DynamicResource Accent}" CornerRadius="2" HorizontalAlignment="Left"/>
            </Grid>
          </ControlTemplate>
        </Setter.Value>
      </Setter>
    </Style>

    <Style TargetType="ScrollBar">
      <Setter Property="Width" Value="10"/>
      <Setter Property="MinWidth" Value="10"/>
      <Setter Property="Template">
        <Setter.Value>
          <ControlTemplate TargetType="ScrollBar">
            <Track x:Name="PART_Track" Orientation="Vertical" IsDirectionReversed="True">
              <Track.Thumb>
                <Thumb>
                  <Thumb.Template>
                    <ControlTemplate TargetType="Thumb">
                      <Border Background="{DynamicResource Scroll}" CornerRadius="4" Margin="2"/>
                    </ControlTemplate>
                  </Thumb.Template>
                </Thumb>
              </Track.Thumb>
            </Track>
          </ControlTemplate>
        </Setter.Value>
      </Setter>
      <Style.Triggers>
        <Trigger Property="Orientation" Value="Horizontal">
          <Setter Property="Width" Value="Auto"/>
          <Setter Property="MinWidth" Value="0"/>
          <Setter Property="Height" Value="10"/>
          <Setter Property="Template">
            <Setter.Value>
              <ControlTemplate TargetType="ScrollBar">
                <Track x:Name="PART_Track" Orientation="Horizontal" IsDirectionReversed="False">
                  <Track.Thumb>
                    <Thumb>
                      <Thumb.Template>
                        <ControlTemplate TargetType="Thumb">
                          <Border Background="{DynamicResource Scroll}" CornerRadius="4" Margin="2"/>
                        </ControlTemplate>
                      </Thumb.Template>
                    </Thumb>
                  </Track.Thumb>
                </Track>
              </ControlTemplate>
            </Setter.Value>
          </Setter>
        </Trigger>
      </Style.Triggers>
    </Style>

    <DataTemplate x:Key="RowTemplate">
      <StackPanel Margin="12,9,12,10" Opacity="{Binding Dim}">
        <Grid>
          <Grid.ColumnDefinitions><ColumnDefinition Width="*"/><ColumnDefinition Width="Auto"/></Grid.ColumnDefinitions>
          <TextBlock Text="{Binding Who}" FontWeight="SemiBold" TextTrimming="CharacterEllipsis"/>
          <TextBlock Grid.Column="1" Text="{Binding When}" Foreground="{DynamicResource Sub}" FontSize="12" Margin="10,1,0,0"/>
        </Grid>
        <TextBlock Text="{Binding Subject}" TextTrimming="CharacterEllipsis" Margin="0,3,0,0"/>
        <TextBlock Text="{Binding Where}" Foreground="{DynamicResource Sub}" FontSize="12" TextTrimming="CharacterEllipsis" Margin="0,3,0,0"/>
      </StackPanel>
    </DataTemplate>
  </Window.Resources>

  <Grid Background="{DynamicResource Bg}">
  <Grid Margin="24,18,24,16">
    <Grid.RowDefinitions>
      <RowDefinition Height="Auto"/>
      <RowDefinition Height="*"/>
      <RowDefinition Height="Auto"/>
    </Grid.RowDefinitions>

    <!-- ===================== header ===================== -->
    <Grid Margin="0,0,0,16">
      <Grid.ColumnDefinitions><ColumnDefinition Width="*"/><ColumnDefinition Width="Auto"/></Grid.ColumnDefinitions>
      <StackPanel Orientation="Horizontal" VerticalAlignment="Center">
        <ContentControl x:Name="HeaderLogo" Width="40" Height="40"/>
        <StackPanel Margin="12,0,0,0" VerticalAlignment="Center">
          <TextBlock Text="FADOE" FontSize="20" FontWeight="SemiBold"/>
          <TextBlock Text="Find A Damn Outlook Email" Foreground="{DynamicResource Sub}" FontSize="12"/>
        </StackPanel>
      </StackPanel>
      <StackPanel Grid.Column="1" Orientation="Horizontal" VerticalAlignment="Center">
        <RadioButton x:Name="TabFind" Style="{StaticResource Tab}" IsChecked="True" ToolTip="Ctrl+1">
          <StackPanel Orientation="Horizontal">
            <TextBlock Text="&#xE721;" Style="{StaticResource Icon}"/><TextBlock Text="Find an email"/>
          </StackPanel>
        </RadioButton>
        <RadioButton x:Name="TabCatch" Style="{StaticResource Tab}" Margin="6,0,0,0" ToolTip="Ctrl+2">
          <StackPanel Orientation="Horizontal">
            <TextBlock Text="&#xE77B;" Style="{StaticResource Icon}"/><TextBlock Text="Catch up on a person"/>
          </StackPanel>
        </RadioButton>
      </StackPanel>
    </Grid>

    <!-- ===================== find ===================== -->
    <Grid x:Name="FindPanel" Grid.Row="1">
      <Grid.RowDefinitions>
        <RowDefinition Height="Auto"/><RowDefinition Height="Auto"/><RowDefinition Height="*"/>
      </Grid.RowDefinitions>
      <Border Style="{StaticResource Card}" Padding="14">
        <Grid>
          <Grid.ColumnDefinitions>
            <ColumnDefinition Width="*"/><ColumnDefinition Width="250"/><ColumnDefinition Width="Auto"/>
          </Grid.ColumnDefinitions>
          <Grid>
            <TextBox x:Name="SearchBox" FontSize="15"/>
            <TextBlock x:Name="SearchHint" Style="{StaticResource Hint}" FontSize="15"
                       Text="What words do you remember?  Put &quot;quotes&quot; around a phrase."/>
          </Grid>
          <Grid Grid.Column="1" Margin="10,0,0,0">
            <TextBox x:Name="FromBox"/>
            <TextBlock x:Name="FromHint" Style="{StaticResource Hint}" Text="From (optional)"/>
          </Grid>
          <Button x:Name="SearchBtn" Grid.Column="2" Style="{StaticResource Primary}" Margin="10,0,0,0">
            <StackPanel Orientation="Horizontal">
              <TextBlock Text="&#xE721;" Style="{StaticResource Icon}"/><TextBlock Text="Search"/>
            </StackPanel>
          </Button>
        </Grid>
      </Border>
      <Grid Grid.Row="1" Margin="4,10,4,12">
        <TextBlock x:Name="FindStatus" Foreground="{DynamicResource Sub}" TextTrimming="CharacterEllipsis"
                   Text="Type what you remember and press Enter. Matches only what people actually wrote, not the quoted replies underneath."/>
        <ProgressBar x:Name="FindProgress" VerticalAlignment="Bottom" Margin="0,0,0,-7" Maximum="100" Visibility="Hidden"/>
      </Grid>
      <Grid Grid.Row="2">
        <Grid.ColumnDefinitions>
          <ColumnDefinition Width="400"/><ColumnDefinition Width="14"/><ColumnDefinition Width="*"/>
        </Grid.ColumnDefinitions>
        <Border Style="{StaticResource Card}" Padding="6,4,4,6">
          <DockPanel>
            <TextBlock x:Name="FindCount" DockPanel.Dock="Top" Foreground="{DynamicResource Sub}" FontSize="12" Margin="12,8,12,6" Text="Results"/>
            <ListBox x:Name="FindList" ItemTemplate="{StaticResource RowTemplate}" Background="Transparent"
                     BorderThickness="0" ScrollViewer.HorizontalScrollBarVisibility="Disabled"/>
          </DockPanel>
        </Border>
        <Border Grid.Column="2" Style="{StaticResource Card}" Padding="24,20,14,14">
          <Grid>
            <StackPanel x:Name="FindEmpty" HorizontalAlignment="Center" VerticalAlignment="Center">
              <TextBlock Text="&#xE715;" FontFamily="{StaticResource Icons}" FontSize="40" Foreground="{DynamicResource Line}" HorizontalAlignment="Center"/>
              <TextBlock Text="Pick a result to read the whole message here." Foreground="{DynamicResource Sub}" Margin="0,12,0,0" HorizontalAlignment="Center"/>
            </StackPanel>
            <DockPanel x:Name="FindDetail" Visibility="Collapsed">
              <StackPanel DockPanel.Dock="Top" Margin="0,0,10,0">
                <TextBlock x:Name="FSubject" FontSize="19" FontWeight="SemiBold" TextWrapping="Wrap"/>
                <TextBlock x:Name="FMeta" Style="{StaticResource Meta}" Margin="0,8,0,0"/>
                <TextBlock x:Name="FTo" Style="{StaticResource Meta}" TextWrapping="NoWrap" TextTrimming="CharacterEllipsis"/>
                <TextBlock x:Name="FWhere" Style="{StaticResource Meta}"/>
                <TextBlock x:Name="FNote" Style="{StaticResource Meta}" Foreground="{DynamicResource Warn}" Visibility="Collapsed"/>
                <WrapPanel Margin="0,14,0,0">
                  <Button x:Name="CopySearchBtn" Style="{StaticResource Primary}" Margin="0,0,8,6"
                          ToolTip="Paste into new Outlook's search box to bring up this message's conversation">
                    <StackPanel Orientation="Horizontal">
                      <TextBlock Text="&#xE8C8;" Style="{StaticResource Icon}"/><TextBlock Text="Copy Outlook search"/>
                    </StackPanel>
                  </Button>
                  <Button x:Name="OpenMsgBtn" Style="{StaticResource Secondary}" Margin="0,0,8,6"
                          ToolTip="Opens this one message in its own window (a classic Outlook window) - no conversation to dig through. You can reply or forward from there.">
                    <StackPanel Orientation="Horizontal">
                      <TextBlock Text="&#xE8A7;" Style="{StaticResource Icon}"/><TextBlock Text="Open just this message"/>
                    </StackPanel>
                  </Button>
                </WrapPanel>
                <Border Background="{DynamicResource Card2}" BorderBrush="{DynamicResource Line}" BorderThickness="1" CornerRadius="6" Padding="10,7" Margin="0,4,0,0">
                  <TextBox x:Name="FQuery" Style="{StaticResource Plain}"/>
                </Border>
                <TextBlock x:Name="FTip" Style="{StaticResource Meta}" FontSize="12" Margin="0,6,0,0"/>
                <Border Height="1" Background="{DynamicResource Line}" Margin="0,14,0,6"/>
              </StackPanel>
              <FlowDocumentScrollViewer x:Name="FBody" VerticalScrollBarVisibility="Auto" HorizontalScrollBarVisibility="Disabled" IsToolBarVisible="False"/>
            </DockPanel>
          </Grid>
        </Border>
      </Grid>
    </Grid>

    <!-- ===================== catch up ===================== -->
    <Grid x:Name="CatchPanel" Grid.Row="1" Visibility="Collapsed">
      <Grid.RowDefinitions>
        <RowDefinition Height="Auto"/><RowDefinition Height="Auto"/><RowDefinition Height="Auto"/><RowDefinition Height="*"/>
      </Grid.RowDefinitions>
      <Border Style="{StaticResource Card}" Padding="14">
        <Grid>
          <Grid.ColumnDefinitions>
            <ColumnDefinition Width="*"/><ColumnDefinition Width="Auto"/><ColumnDefinition Width="Auto"/><ColumnDefinition Width="Auto"/>
          </Grid.ColumnDefinitions>
          <Grid>
            <TextBox x:Name="PersonBox" FontSize="15"/>
            <TextBlock x:Name="PersonHint" Style="{StaticResource Hint}" FontSize="15"
                       Text="Their email address (or several, separated by commas)"/>
          </Grid>
          <TextBox x:Name="DaysBox" Grid.Column="1" Width="72" Margin="10,0,0,0" Text="180" TextAlignment="Center"/>
          <TextBlock Grid.Column="2" Text="days back" Foreground="{DynamicResource Sub}" VerticalAlignment="Center" Margin="8,0,0,0"/>
          <Button x:Name="GatherBtn" Grid.Column="3" Style="{StaticResource Primary}" Margin="14,0,0,0">
            <StackPanel Orientation="Horizontal">
              <TextBlock Text="&#xE77B;" Style="{StaticResource Icon}"/><TextBlock Text="Gather emails"/>
            </StackPanel>
          </Button>
        </Grid>
      </Border>
      <Grid Grid.Row="1" Margin="4,10,4,12">
        <TextBlock x:Name="CatchStatus" Foreground="{DynamicResource Sub}" TextTrimming="CharacterEllipsis"
                   Text="Collects every email sent to or from this person - full text and dates - ready to paste into Claude for a summary."/>
        <ProgressBar x:Name="CatchProgress" VerticalAlignment="Bottom" Margin="0,0,0,-7" Maximum="100" Visibility="Hidden"/>
      </Grid>
      <WrapPanel x:Name="CatchSuggest" Grid.Row="2" Margin="0,0,0,6" Visibility="Collapsed"/>
      <Border x:Name="CatchSummary" Grid.Row="2" Style="{StaticResource Card}" Padding="18,14" Margin="0,0,0,14" Visibility="Collapsed">
        <Grid>
          <Grid.ColumnDefinitions><ColumnDefinition Width="*"/><ColumnDefinition Width="Auto"/></Grid.ColumnDefinitions>
          <StackPanel VerticalAlignment="Center">
            <TextBlock x:Name="CTitle" FontSize="17" FontWeight="SemiBold" TextTrimming="CharacterEllipsis"/>
            <TextBlock x:Name="CSub" Style="{StaticResource Meta}" TextWrapping="NoWrap" TextTrimming="CharacterEllipsis"/>
          </StackPanel>
          <StackPanel Grid.Column="1" Orientation="Horizontal" VerticalAlignment="Center" Margin="14,0,0,0">
            <Button x:Name="CopyAllBtn" Style="{StaticResource Primary}"
                    ToolTip="Copy the whole document - every message, full text - to paste into Claude">
              <StackPanel Orientation="Horizontal">
                <TextBlock Text="&#xE8C8;" Style="{StaticResource Icon}"/><TextBlock Text="Copy all for Claude"/>
              </StackPanel>
            </Button>
            <Button x:Name="OpenFileBtn" Style="{StaticResource Secondary}" Margin="8,0,0,0">
              <StackPanel Orientation="Horizontal">
                <TextBlock Text="&#xE8E5;" Style="{StaticResource Icon}"/><TextBlock Text="Open file"/>
              </StackPanel>
            </Button>
            <Button x:Name="ShowFolderBtn" Style="{StaticResource Secondary}" Margin="8,0,0,0">
              <StackPanel Orientation="Horizontal">
                <TextBlock Text="&#xE838;" Style="{StaticResource Icon}"/><TextBlock Text="Show in folder"/>
              </StackPanel>
            </Button>
          </StackPanel>
        </Grid>
      </Border>
      <Grid Grid.Row="3">
        <Grid.ColumnDefinitions>
          <ColumnDefinition Width="400"/><ColumnDefinition Width="14"/><ColumnDefinition Width="*"/>
        </Grid.ColumnDefinitions>
        <Border Style="{StaticResource Card}" Padding="6,4,4,6">
          <DockPanel>
            <TextBlock x:Name="CatchCount" DockPanel.Dock="Top" Foreground="{DynamicResource Sub}" FontSize="12" Margin="12,8,12,6" Text="Messages"/>
            <ListBox x:Name="CatchList" ItemTemplate="{StaticResource RowTemplate}" Background="Transparent"
                     BorderThickness="0" ScrollViewer.HorizontalScrollBarVisibility="Disabled"/>
          </DockPanel>
        </Border>
        <Border Grid.Column="2" Style="{StaticResource Card}" Padding="24,20,14,14">
          <Grid>
            <StackPanel x:Name="CatchEmpty" HorizontalAlignment="Center" VerticalAlignment="Center">
              <TextBlock Text="&#xE77B;" FontFamily="{StaticResource Icons}" FontSize="40" Foreground="{DynamicResource Line}" HorizontalAlignment="Center"/>
              <TextBlock Text="Messages with this person show up here." Foreground="{DynamicResource Sub}" Margin="0,12,0,0" HorizontalAlignment="Center"/>
            </StackPanel>
            <DockPanel x:Name="CatchDetail" Visibility="Collapsed">
              <StackPanel DockPanel.Dock="Top" Margin="0,0,10,0">
                <TextBlock x:Name="CSubject" FontSize="19" FontWeight="SemiBold" TextWrapping="Wrap"/>
                <TextBlock x:Name="CMeta" Style="{StaticResource Meta}" Margin="0,8,0,0"/>
                <TextBlock x:Name="CTo" Style="{StaticResource Meta}" TextWrapping="NoWrap" TextTrimming="CharacterEllipsis"/>
                <TextBlock x:Name="CCc" Style="{StaticResource Meta}" TextWrapping="NoWrap" TextTrimming="CharacterEllipsis"/>
                <TextBlock x:Name="CAtt" Style="{StaticResource Meta}"/>
                <Border Height="1" Background="{DynamicResource Line}" Margin="0,14,0,6"/>
              </StackPanel>
              <FlowDocumentScrollViewer x:Name="CBody" VerticalScrollBarVisibility="Auto" HorizontalScrollBarVisibility="Disabled" IsToolBarVisible="False"/>
            </DockPanel>
          </Grid>
        </Border>
      </Grid>
    </Grid>

    <!-- ===================== footer ===================== -->
    <Grid Grid.Row="2" Margin="4,12,0,0">
      <TextBlock x:Name="Footer" Foreground="{DynamicResource Sub}" FontSize="12" VerticalAlignment="Center"
                 Text="Read-only. Reads your mailbox through classic Outlook in the background (about the last 12 months of mail)."/>
      <Border x:Name="Toast" HorizontalAlignment="Right" Background="{DynamicResource Surface}" BorderBrush="{DynamicResource Line}"
              BorderThickness="1" CornerRadius="14" Padding="12,5" Visibility="Hidden">
        <StackPanel Orientation="Horizontal">
          <TextBlock Text="&#xE73E;" Style="{StaticResource Icon}" Foreground="{DynamicResource Ok}"/>
          <TextBlock x:Name="ToastText" VerticalAlignment="Center"/>
        </StackPanel>
      </Border>
    </Grid>
  </Grid>
  </Grid>
</Window>
'@

# ---------------------------------------------------------- build window ----
try {
    $window = [Windows.Markup.XamlReader]::Parse($xaml)
} catch {
    $msg = $_.Exception.Message
    if ($_.Exception.InnerException) { $msg += "`n" + $_.Exception.InnerException.Message }
    $splash.Close = $true
    if ($ScreenshotPath) { Write-Host "FADOE couldn't start: $msg"; exit 1 }
    [System.Windows.MessageBox]::Show("FADOE couldn't start:`n`n$msg", 'FADOE') | Out-Null
    exit 1
}
foreach ($m in [regex]::Matches($xaml, 'x:Name="(\w+)"')) {
    $name = $m.Groups[1].Value
    if ($name -ne 'b' -and $name -notlike 'PART_*') { Set-Variable -Name $name -Value $window.FindName($name) -Scope Script }
}
foreach ($k in $palette.Keys) {
    $window.Resources[$k] = [System.Windows.Media.SolidColorBrush]::new([System.Windows.Media.ColorConverter]::ConvertFromString($palette[$k]))
}

$wa = [System.Windows.SystemParameters]::WorkArea
$window.Width  = [Math]::Min(1200, $wa.Width * 0.94)
$window.Height = [Math]::Min(800, $wa.Height * 0.94)

$HeaderLogo.Content = New-Logo 40
$Footer.Text = 'FADOE v' + $FadoeVersion + $dot + $Footer.Text
try {
    $window.Icon = [System.Windows.Media.Imaging.BitmapFrame]::Create((New-Object System.IO.MemoryStream(, (Get-LogoPng 64))))
} catch { }

$window.Add_SourceInitialized({
    if ($useDark) {
        try {
            $h = (New-Object System.Windows.Interop.WindowInteropHelper($window)).Handle
            $v = 1
            [void][FadoeNative.Win]::DwmSetWindowAttribute($h, 20, [ref]$v, 4)   # dark title bar
        } catch { }
    }
})

# ------------------------------------------------------------- helpers ----
function Update-Hint($box, $hint) { $hint.Visibility = if ($box.Text.Length) { 'Collapsed' } else { 'Visible' } }
function Get-FolderName([string]$path) { ($path -split '\\' | Where-Object { $_ } | Select-Object -Last 1) }
function Get-Name([string]$s) { ($s -replace '\s*<[^>]*>', '').Trim() }

$toastTimer = New-Object System.Windows.Threading.DispatcherTimer
$toastTimer.Interval = [TimeSpan]::FromSeconds(3.5)
$toastTimer.Add_Tick({ $Toast.Visibility = 'Hidden'; $toastTimer.Stop() })
function Show-Toast([string]$msg) {
    $ToastText.Text = $msg
    $Toast.Visibility = 'Visible'
    $toastTimer.Stop(); $toastTimer.Start()
}

function Copy-Text([string]$text, [string]$msg) {
    if (-not $ScreenshotPath) {
        try { [System.Windows.Clipboard]::SetText($text) }
        catch { try { Set-Clipboard -Value $text } catch { Show-Toast 'Could not reach the clipboard - try again'; return } }
    }
    Show-Toast $msg
}

# Message text -> FlowDocument: search words highlighted, links clickable.
$urlRx = '<(https?://[^>\s]+)>|(https?://[^\s<>"]+)'
function Add-Runs($inlines, [string]$text, [string]$hiRx) {
    if (-not $text) { return }
    if (-not $hiRx) { $inlines.Add([System.Windows.Documents.Run]::new($text)); return }
    $last = 0
    foreach ($m in [regex]::Matches($text, $hiRx, 'IgnoreCase')) {
        if ($m.Index -gt $last) { $inlines.Add([System.Windows.Documents.Run]::new($text.Substring($last, $m.Index - $last))) }
        $r = [System.Windows.Documents.Run]::new($m.Value)
        $r.Background = $window.Resources['Hi']
        $r.Foreground = $window.Resources['HiText']
        $r.FontWeight = [System.Windows.FontWeights]::SemiBold
        $inlines.Add($r)
        if (-not $script:firstHit) { $script:firstHit = $r }
        $last = $m.Index + $m.Length
    }
    if ($last -lt $text.Length) { $inlines.Add([System.Windows.Documents.Run]::new($text.Substring($last))) }
}
function Add-Inlines($inlines, [string]$line, [string]$hiRx) {
    $last = 0
    foreach ($m in [regex]::Matches($line, $urlRx)) {
        Add-Runs $inlines $line.Substring($last, $m.Index - $last) $hiRx
        $url = if ($m.Groups[1].Success) { $m.Groups[1].Value } else { $m.Groups[2].Value }
        try {
            $uri = [uri]$url
            $h = New-Object System.Windows.Documents.Hyperlink
            $h.Inlines.Add([System.Windows.Documents.Run]::new(' ' + [char]0x2197 + ' ' + ($uri.Host -replace '^www\.', '') + ' '))
            $h.NavigateUri = $uri
            $h.ToolTip = $url
            $h.FontSize = 12.5
            $h.Foreground = $window.Resources['Link']
            $h.Add_RequestNavigate({ param($s, $e) try { Start-Process $e.Uri.AbsoluteUri } catch { }; $e.Handled = $true })
            $inlines.Add($h)
        } catch { Add-Runs $inlines $url $hiRx }
        $last = $m.Index + $m.Length
    }
    Add-Runs $inlines $line.Substring($last) $hiRx
}
function Set-Doc($viewer, [string]$text, [string[]]$terms) {
    $doc = New-Object System.Windows.Documents.FlowDocument
    $doc.PagePadding = [System.Windows.Thickness]::new(0, 0, 12, 0)
    $doc.FontFamily = New-Object System.Windows.Media.FontFamily 'Segoe UI Variable Text, Segoe UI'
    $doc.FontSize = 14
    $doc.TextAlignment = 'Left'
    $doc.Foreground = $window.Resources['Text']
    $doc.Background = [System.Windows.Media.Brushes]::Transparent
    $script:firstHit = $null
    $hiRx = if ($terms) { ($terms | ForEach-Object { [regex]::Escape($_) }) -join '|' } else { $null }

    $t = ([string]$text) -replace "`r`n", "`n" -replace "`r", "`n"
    $lines = @($t.Trim() -split "`n")
    $p = $null
    for ($i = 0; $i -lt $lines.Count; $i++) {
        $line = $lines[$i]
        if (-not $line.Trim()) {
            if ($p) { $p.Margin = [System.Windows.Thickness]::new($p.Margin.Left, 0, 0, 12) }
            continue
        }
        # Outlook's plain text turns list bullets into a lone '*' (or 'o') line, or '*<tab>text'
        $bullet = $false
        if ($line.Trim() -match '^[\*o\u00B7\u2022]$') {
            $j = $i + 1
            while ($j -lt $lines.Count -and -not $lines[$j].Trim()) { $j++ }
            if ($j -lt $lines.Count) { $line = $lines[$j]; $i = $j; $bullet = $true }
        } else {
            $bm = [regex]::Match($line, '^\s*[\*\u00B7\u2022]\s+(.*)$')
            if ($bm.Success) { $line = $bm.Groups[1].Value; $bullet = $true }
        }
        $p = New-Object System.Windows.Documents.Paragraph
        if ($bullet) {
            $p.Margin = [System.Windows.Thickness]::new(18, 0, 0, 3)
            $p.TextIndent = -14
            $p.Inlines.Add([System.Windows.Documents.Run]::new([string][char]0x2022 + '  '))
            $line = $line.Trim()
        } else {
            $p.Margin = [System.Windows.Thickness]::new(0, 0, 0, 3)
        }
        Add-Inlines $p.Inlines $line $hiRx
        $doc.Blocks.Add($p)
    }
    $viewer.Document = $doc
    if ($script:firstHit) {
        [void]$viewer.Dispatcher.BeginInvoke([System.Windows.Threading.DispatcherPriority]::Loaded,
            [System.Action]{ try { $script:firstHit.BringIntoView() } catch { } })
    }
}

# ---- background work: the tools run in a separate runspace so the window stays responsive ----
$script:job = $null
$jobTimer = New-Object System.Windows.Threading.DispatcherTimer
$jobTimer.Interval = [TimeSpan]::FromMilliseconds(150)
$jobTimer.Add_Tick({
    $j = $script:job
    if (-not $j) { $jobTimer.Stop(); return }
    $prog = $j.PS.Streams.Progress
    if ($prog.Count -ne $j.SeenProgress) {
        $j.SeenProgress = $prog.Count
        $j.LastProgress = Get-Date
        $rec = $prog[$prog.Count - 1]
        if ($rec.PercentComplete -ge 0) { $j.Bar.Value = $rec.PercentComplete }
        $desc = [string]$rec.StatusDescription
        if ($desc -match '\\') { $j.Status.Text = '{0} {1} ...' -f $j.Label, (Get-FolderName $desc) }
        elseif ($desc) { $j.Status.Text = $desc + ' ...' }
    }
    # Watchdog: classic Outlook runs invisibly for FADOE, so if it stops to show a message
    # (e.g. "exhausted all shared resources") everything waits on a window nobody can see.
    if (-not $j.Warned -and ((Get-Date) - $j.LastProgress).TotalSeconds -ge 12 -and ((Get-Date) - $j.LastCheck).TotalSeconds -ge 3) {
        $j.LastCheck = Get-Date
        $stuck = Get-OutlookBlocker
        if ($stuck) {
            $j.Warned = $true
            $j.Status.SetResourceReference([System.Windows.Controls.TextBlock]::ForegroundProperty, 'Warn')
            $j.Status.Text = "Classic Outlook has paused to show a message ('" + $stuck.MainWindowTitle + "'). Click OK in that window. " +
                             "If it says 'exhausted all shared resources', close Outlook completely (new Outlook too), reopen it, then reopen FADOE."
            try { [void](New-Object -ComObject WScript.Shell).AppActivate($stuck.Id) } catch { }
        }
    }
    if ($j.Handle.IsCompleted) {
        $jobTimer.Stop()
        $out = $null; $errs = @()
        try { $out = $j.PS.EndInvoke($j.Handle) } catch { $errs += $_.Exception.Message }
        foreach ($e in $j.PS.Streams.Error) { $errs += $e.ToString() }
        try { $j.PS.Dispose() } catch { }
        $script:job = $null
        $j.Bar.Visibility = 'Hidden'
        $j.Status.SetResourceReference([System.Windows.Controls.TextBlock]::ForegroundProperty, 'Sub')
        & $j.OnDone $out $errs
        if (-not $script:job -and $script:pending) { $p = $script:pending; $script:pending = $null; & $p }
    }
})

# Classic Outlook runs invisibly for FADOE. If it stops on a message box (e.g. "Outlook Data
# File: exhausted all shared resources"), everything waits on a window nobody is looking at.
# Only real message boxes count -- reminders and reopened message windows don't block anything.
function Get-OutlookBlocker {
    foreach ($o in @(Get-Process OUTLOOK -ErrorAction SilentlyContinue)) {
        $t = $null
        try { $t = [FadoeNative.Win]::FindDialog($o.Id) } catch { }
        if ($null -ne $t) { return [pscustomobject]@{ Id = $o.Id; MainWindowTitle = $(if ($t) { $t } else { 'Microsoft Outlook' }) } }
    }
}

# One long-lived engine runspace (one thread, so Outlook's COM objects stay valid). The search
# index and the Outlook connection stay loaded in it between searches.
$script:engine = $null
function Get-Engine {
    if ($script:engine -and $script:engine.RunspaceStateInfo.State -eq 'Opened') { return $script:engine }
    $iss = [System.Management.Automation.Runspaces.InitialSessionState]::CreateDefault()
    $iss.ExecutionPolicy = 'Bypass'
    $rs = [runspacefactory]::CreateRunspace($iss)
    $rs.ApartmentState = 'STA'
    $rs.ThreadOptions = 'ReuseThread'
    $rs.Open()
    $script:engine = $rs
    $rs
}

function Start-Work([string]$scriptPath, [hashtable]$params, $bar, $status, [string]$label, [scriptblock]$onDone, [string]$kind = 'work') {
    $ps = [powershell]::Create()
    $ps.Runspace = Get-Engine
    [void]$ps.AddCommand($scriptPath)
    foreach ($k in $params.Keys) { [void]$ps.AddParameter($k, $params[$k]) }
    $bar.Value = 0
    $bar.Visibility = 'Visible'
    $status.Text = $label + ' ...'
    $now = Get-Date
    $script:job = @{ PS = $ps; Handle = $ps.BeginInvoke(); Bar = $bar; Status = $status; Label = $label; OnDone = $onDone; Kind = $kind
                     SeenProgress = 0; LastProgress = $now; LastCheck = $now; Warned = $false }
    $jobTimer.Start()
}

# The index updates in the background when the window opens. A search started meanwhile waits
# for it instead of being refused.
$script:pending = $null
function Wait-ForIndex([scriptblock]$then, $status) {
    if ($script:job -and $script:job.Kind -eq 'sync') {
        $script:pending = $then
        $status.Text = 'Finishing the search index update first - this will start right after ...'
        return $true
    }
    return [bool]$script:job
}

$onSyncDone = {
    param($out, $errs)
    $s = @($out | Where-Object { $_ -and $_.PSObject.Properties['NewMessages'] }) | Select-Object -First 1
    if ($s) {
        $FindStatus.Text = ('Search index ready: {0:N0} messages from the last year. Type what you remember and press Enter.' -f $s.Messages)
    } elseif ($errs.Count) {
        $FindStatus.Text = 'Could not update the search index: ' + $errs[0]
    }
}

function Start-IndexSync {
    if ($script:job) { return }
    $stuck = Get-OutlookBlocker
    if ($stuck) {
        $FindStatus.SetResourceReference([System.Windows.Controls.TextBlock]::ForegroundProperty, 'Warn')
        $FindStatus.Text = "Classic Outlook is showing a message ('" + $stuck.MainWindowTitle + "') and won't answer until it's closed. " +
                           "Click OK there; if it says 'exhausted all shared resources', close Outlook completely (new Outlook too), reopen it, then reopen FADOE."
        try { [void](New-Object -ComObject WScript.Shell).AppActivate($stuck.Id) } catch { }
        return
    }
    $first = -not (Test-Path (Join-Path $env:LOCALAPPDATA 'FADOE'))
    Start-Work $findTool @{ SyncOnly = $true } $FindProgress $FindStatus `
        $(if ($first) { 'Building the search index for the first time (a couple of minutes, once)' } else { 'Updating the search index' }) `
        $onSyncDone 'sync'
}

function Request-Snapshot {
    [void]$window.Dispatcher.BeginInvoke([System.Windows.Threading.DispatcherPriority]::ApplicationIdle, [System.Action]{
        $window.UpdateLayout()
        $root = $window.Content
        $w = [int]$root.ActualWidth; $h = [int]$root.ActualHeight
        $bmp = New-Object System.Windows.Media.Imaging.RenderTargetBitmap($w, $h, 96, 96, [System.Windows.Media.PixelFormats]::Pbgra32)
        $bmp.Render($root)
        $enc = New-Object System.Windows.Media.Imaging.PngBitmapEncoder
        $enc.Frames.Add([System.Windows.Media.Imaging.BitmapFrame]::Create($bmp))
        $fs = [System.IO.File]::Create($ScreenshotPath)
        try { $enc.Save($fs) } finally { $fs.Close() }
        $window.Close()
    })
}

# ------------------------------------------------------------ find tab ----
function Show-FindDetail($r) {
    $script:findSel = $r
    if (-not $r) { $FindDetail.Visibility = 'Collapsed'; $FindEmpty.Visibility = 'Visible'; return }
    $FindEmpty.Visibility = 'Collapsed'
    $FindDetail.Visibility = 'Visible'
    $FSubject.Text = if ($r.Subject) { $r.Subject } else { '(no subject)' }
    $from = $r.Sender
    if ($r.SenderAddr -and $r.SenderAddr -notlike '/o=*' -and $r.SenderAddr -ne $r.Sender) { $from += ' <' + $r.SenderAddr + '>' }
    $FMeta.Text = 'From ' + $from + $dot + $r.When.ToString('dddd, MMMM d, yyyy \a\t h:mm tt')
    $FTo.Text = if ($r.To) { 'To ' + $r.To } else { '' }
    $where = 'In ' + (Get-FolderName $r.Folder)
    if ($r.Position) { $where += $dot + ('message {0} of {1} in the conversation (oldest first)' -f $r.Position, $r.ThreadCount) }
    $FWhere.Text = $where
    if (-not $r.IsPrimary) {
        $FNote.Text = "This is in the shared mailbox '" + $r.Mailbox + "'. In new Outlook, click a folder in that mailbox before pasting the search."
        $FNote.Visibility = 'Visible'
    } else { $FNote.Visibility = 'Collapsed' }
    $FQuery.Text = $r.Query
    $FTip.Text = if ($r.ThreadCount -gt 1) {
        'In new Outlook: paste the search, open the conversation, and look for the message sent at ' + $r.When.ToString('h:mm tt') + ' - or use Open just this message.'
    } else { 'In new Outlook: paste this into the search box.' }
    Set-Doc $FBody $r.Text $r.Terms
}

$onFindDone = {
    param($out, $errs)
    $SearchBtn.IsEnabled = $true
    $rows = @($out | Where-Object { $_ -and $_.PSObject.Properties['Query'] })
    if ($rows.Count -eq 0) {
        $FindStatus.Text = if ($errs.Count) { 'Something went wrong: ' + $errs[0] }
                           else { "No messages found with: $script:lastSearch.  Try fewer or different words." }
        $FindCount.Text = 'Results'
        if ($ScreenshotPath) { Request-Snapshot }
        return
    }
    $items = foreach ($r in $rows) {
        $row = [FadoeRow]::new()
        $row.Who = $r.Sender
        $row.When = $r.When.ToString('MMM d, yyyy')
        $row.Subject = if ($r.Subject) { $r.Subject } else { '(no subject)' }
        $bits = @(Get-FolderName $r.Folder)
        if ($r.Position) { $bits += ('{0} of {1} in thread' -f $r.Position, $r.ThreadCount) }
        if (-not $r.IsPrimary) { $bits += $r.Mailbox }
        if ($r.IsBulk) { $bits += 'newsletter' }
        $row.Where = $bits -join $dot
        $row.Dim = if ($r.IsBulk) { 0.6 } else { 1.0 }
        $row.Data = $r
        $row
    }
    $FindList.ItemsSource = @($items)
    $FindCount.Text = '{0} conversations matched{1}best {2} shown' -f $rows[0].TotalConversations, $dot, $rows.Count
    $FindStatus.Text = 'The top result''s Outlook search is on your clipboard. Click any result to read it; double-click to copy its search.'
    Copy-Text $rows[0].Query 'Copied the top result''s Outlook search'
    $FindList.SelectedIndex = 0
    [void]$FindList.Focus()
    if ($ScreenshotPath) { Request-Snapshot }
}

function Start-Find {
    if (Wait-ForIndex { Start-Find } $FindStatus) { return }
    $words = $SearchBox.Text.Trim()
    if (-not $words) { [void]$SearchBox.Focus(); return }
    $params = @{ Search = $words; Top = 25; NoClipboard = $true; PassThru = $true }
    if ($FromBox.Text.Trim()) { $params.From = $FromBox.Text.Trim() }
    $script:lastSearch = $words
    $SearchBtn.IsEnabled = $false
    $FindList.ItemsSource = $null
    Show-FindDetail $null
    $FindCount.Text = 'Results'
    Start-Work $findTool $params $FindProgress $FindStatus 'Searching' $onFindDone
}

# -------------------------------------------------------- catch-up tab ----
function Show-CatchDetail($m) {
    if (-not $m) { $CatchDetail.Visibility = 'Collapsed'; $CatchEmpty.Visibility = 'Visible'; return }
    $CatchEmpty.Visibility = 'Collapsed'
    $CatchDetail.Visibility = 'Visible'
    $CSubject.Text = if ($m.Subject) { $m.Subject } else { '(no subject)' }
    $CMeta.Text = 'From ' + $m.From + $dot + $m.When.ToString('dddd, MMMM d, yyyy \a\t h:mm tt')
    $CTo.Text = if ($m.To) { 'To ' + $m.To } else { '' }
    $CCc.Text = if ($m.Cc) { 'Cc ' + $m.Cc } else { '' }
    $CCc.Visibility = if ($m.Cc) { 'Visible' } else { 'Collapsed' }
    $CAtt.Text = if ($m.Attachments) { 'Attachments: ' + $m.Attachments } else { '' }
    $CAtt.Visibility = if ($m.Attachments) { 'Visible' } else { 'Collapsed' }
    Set-Doc $CBody $m.Body @()
}

$onCatchDone = {
    param($out, $errs)
    $GatherBtn.IsEnabled = $true
    $res = @($out | Where-Object { $_ -and $_.PSObject.Properties['Messages'] }) | Select-Object -First 1
    $CatchSuggest.Children.Clear()
    $CatchSuggest.Visibility = 'Collapsed'
    if (-not $res -or @($res.Messages).Count -eq 0) {
        $CatchSummary.Visibility = 'Collapsed'
        # @( ) around the whole thing: a one-item result would otherwise unroll to a bare object with no .Count
        $sugg = @(if ($res -and $res.PSObject.Properties['Suggestions']) { $res.Suggestions | Where-Object { $_ } })
        if (-not $res -and $errs.Count) {
            $CatchStatus.Text = 'Something went wrong: ' + $errs[0]
        } elseif ($sugg.Count) {
            $CatchStatus.Text = 'No emails found with ' + ($script:ctxPeople -join ', ') + '.  Did you mean one of these people you''ve emailed?'
            foreach ($s in $sugg) {
                $b = New-Object System.Windows.Controls.Button
                $b.Style = $window.FindResource('Secondary')
                $b.Margin = [System.Windows.Thickness]::new(0, 0, 8, 8)
                $b.Tag = $s.Address
                $label = if ($s.Name -and $s.Name -ne $s.Address) { $s.Name + $dot + $s.Address } else { $s.Address }
                $b.Content = $label + $dot + ('{0} emails' -f $s.Messages)
                $b.ToolTip = 'Catch up on ' + $s.Address
                $b.Add_Click({ param($src, $e) $PersonBox.Text = [string]$src.Tag; Start-CatchUp })
                [void]$CatchSuggest.Children.Add($b)
            }
            $CatchSuggest.Visibility = 'Visible'
        } else {
            $CatchStatus.Text = 'No emails found with ' + ($script:ctxPeople -join ', ') + '. Try just their last name, or more days.'
        }
        if ($ScreenshotPath) { Request-Snapshot }
        return
    }
    $script:ctxResult = $res
    $msgs = @($res.Messages)
    $sent = @($msgs | Where-Object { $_.Direction -like 'SENT*' }).Count
    $first = ($msgs | Sort-Object When | Select-Object -First 1).When
    $last  = ($msgs | Sort-Object When | Select-Object -Last 1).When
    $len   = (Get-Item $res.OutFile).Length

    $CTitle.Text = '{0} messages with {1}' -f $msgs.Count, ($script:ctxPeople -join ', ')
    $CSub.Text = ($first.ToString('MMM d, yyyy') + ' - ' + $last.ToString('MMM d, yyyy')) + $dot +
                 ('{0} sent by you' -f $sent) + $dot + ('{0} received' -f ($msgs.Count - $sent)) + $dot +
                 ('about {0:N0}k tokens' -f [Math]::Max(1, $len / 4000))
    $CatchSummary.Visibility = 'Visible'

    $items = foreach ($m in ($msgs | Sort-Object When -Descending)) {
        $row = [FadoeRow]::new()
        $isSent = $m.Direction -like 'SENT*'
        if ($isSent) {
            $tos = @($m.To -split ';' | Where-Object { $_.Trim() })
            $row.Who = 'You ' + [char]0x2192 + ' ' + (Get-Name $tos[0])
            if ($tos.Count -gt 1) { $row.Who += ' +' + ($tos.Count - 1) }
        } else { $row.Who = Get-Name $m.From }
        $row.When = $m.When.ToString('MMM d, yyyy')
        $row.Subject = if ($m.Subject) { $m.Subject } else { '(no subject)' }
        $row.Where = $(if ($isSent) { 'Sent by you' } else { 'Received' }) + $dot + (Get-FolderName $m.Folder)
        $row.Data = $m
        $row
    }
    $CatchList.ItemsSource = @($items)
    $CatchCount.Text = '{0} messages{1}newest first' -f $msgs.Count, $dot
    Copy-Text ([System.IO.File]::ReadAllText($res.OutFile)) ('Copied all {0} messages - paste into Claude' -f $msgs.Count)
    $CatchStatus.Text = 'Done - the whole document is on your clipboard. Paste it into Claude and ask for a summary.'
    $CatchList.SelectedIndex = 0
    if ($ScreenshotPath) { Request-Snapshot }
}

function Start-CatchUp {
    if (Wait-ForIndex { Start-CatchUp } $CatchStatus) { return }
    $who = $PersonBox.Text.Trim()
    if (-not $who) { [void]$PersonBox.Focus(); return }
    $days = 180; $d = 0
    if ([int]::TryParse($DaysBox.Text.Trim(), [ref]$d) -and $d -gt 0) { $days = $d }
    $DaysBox.Text = [string]$days
    $script:ctxPeople = @($who -split ',' | ForEach-Object { $_.Trim() } | Where-Object { $_ })
    if (-not (Test-Path $ctxFolder)) { New-Item -ItemType Directory -Path $ctxFolder | Out-Null }
    $slug = ($script:ctxPeople[0] -replace '[^A-Za-z0-9]+', '-').Trim('-')
    $out = Join-Path $ctxFolder ('{0}_{1}.md' -f $slug, (Get-Date -Format 'yyyy-MM-dd'))
    if (Test-Path $out) { Remove-Item $out }   # same person, same day: regenerate
    $GatherBtn.IsEnabled = $false
    $CatchSummary.Visibility = 'Collapsed'
    $CatchSuggest.Visibility = 'Collapsed'
    $CatchList.ItemsSource = $null
    Show-CatchDetail $null
    $CatchCount.Text = 'Messages'
    Start-Work $ctxTool @{ Person = [string[]]$script:ctxPeople; Days = $days; OutFile = $out; PassThru = $true } `
        $CatchProgress $CatchStatus 'Reading' $onCatchDone
}

# -------------------------------------------------------------- wiring ----
$SearchBox.Add_TextChanged({ Update-Hint $SearchBox $SearchHint })
$FromBox.Add_TextChanged({ Update-Hint $FromBox $FromHint })
$PersonBox.Add_TextChanged({ Update-Hint $PersonBox $PersonHint })

$enterFind  = { param($s, $e) if ($e.Key -eq [System.Windows.Input.Key]::Return) { Start-Find } }
$enterCatch = { param($s, $e) if ($e.Key -eq [System.Windows.Input.Key]::Return) { Start-CatchUp } }
$SearchBox.Add_KeyDown($enterFind)
$FromBox.Add_KeyDown($enterFind)
$PersonBox.Add_KeyDown($enterCatch)
$DaysBox.Add_KeyDown($enterCatch)
$SearchBtn.Add_Click({ Start-Find })
$GatherBtn.Add_Click({ Start-CatchUp })

$FindList.Add_SelectionChanged({ if ($FindList.SelectedItem) { Show-FindDetail $FindList.SelectedItem.Data } })
$FindList.Add_MouseDoubleClick({
    if ($FindList.SelectedItem) { Copy-Text $FindList.SelectedItem.Data.Query 'Copied the Outlook search' }
})
$CatchList.Add_SelectionChanged({ if ($CatchList.SelectedItem) { Show-CatchDetail $CatchList.SelectedItem.Data } })

$CopySearchBtn.Add_Click({
    if ($script:findSel) { Copy-Text $script:findSel.Query 'Copied - paste it into new Outlook''s search box' }
})
$OpenMsgBtn.Add_Click({
    if (-not $script:findSel) { return }
    try {
        $ns = (New-Object -ComObject Outlook.Application).GetNamespace('MAPI')
        $ns.GetItemFromID($script:findSel.EntryID, $script:findSel.StoreID).Display()
    } catch { Show-Toast ('Could not open it: ' + $_.Exception.Message) }
})
$CopyAllBtn.Add_Click({
    if ($script:ctxResult) {
        Copy-Text ([System.IO.File]::ReadAllText($script:ctxResult.OutFile)) 'Copied everything - paste into Claude'
    }
})
$OpenFileBtn.Add_Click({
    if ($script:ctxResult) { Start-Process notepad.exe -ArgumentList ('"' + $script:ctxResult.OutFile + '"') }
})
$ShowFolderBtn.Add_Click({
    if ($script:ctxResult) { Start-Process explorer.exe -ArgumentList ('/select,"' + $script:ctxResult.OutFile + '"') }
})

$TabFind.Add_Checked({ $FindPanel.Visibility = 'Visible'; $CatchPanel.Visibility = 'Collapsed'; [void]$SearchBox.Focus() })
$TabCatch.Add_Checked({ $FindPanel.Visibility = 'Collapsed'; $CatchPanel.Visibility = 'Visible'; [void]$PersonBox.Focus() })
$window.Add_PreviewKeyDown({
    param($s, $e)
    if ([System.Windows.Input.Keyboard]::Modifiers -eq 'Control') {
        if ($e.Key -eq 'D1' -or $e.Key -eq 'NumPad1') { $TabFind.IsChecked = $true; $e.Handled = $true }
        if ($e.Key -eq 'D2' -or $e.Key -eq 'NumPad2') { $TabCatch.IsChecked = $true; $e.Handled = $true }
    }
})
$window.Add_Loaded({ [void]$SearchBox.Focus() })
$window.Add_ContentRendered({
    $splash.Close = $true
    [void]$window.Activate()
    if (-not $ScreenshotPath) { Start-IndexSync }
})
$window.Add_Closed({
    if ($script:job) { try { $script:job.PS.Stop() } catch { }; $script:forceExit = $true }
})

# ------------------------------------------------------------ test mode ----
if ($ScreenshotPath) {
    $window.WindowStartupLocation = 'Manual'
    $window.Left = -30000; $window.Top = 0
    $window.ShowInTaskbar = $false
    $window.ShowActivated = $false
    $window.Add_ContentRendered({
        switch ($ScreenshotTab) {
            'catch' { $TabCatch.IsChecked = $true; $PersonBox.Text = $ScreenshotQuery; $DaysBox.Text = [string]$ScreenshotDays; Start-CatchUp }
            'find'  { $SearchBox.Text = $ScreenshotQuery; Start-Find }
            default { Request-Snapshot }
        }
    })
}

[void]$window.ShowDialog()
$splash.Close = $true
if ($script:forceExit) { [Environment]::Exit(0) }
