#requires -Version 7.0

function New-StockExchangeCountdownDashboard {
    <#
.SYNOPSIS
    PSExchangeClock — Global Stock Exchange & Market Data Dashboard
.DESCRIPTION
    A real-time WPF dashboard for monitoring global stock exchange closing times and
    live market data. Built in PowerShell with a professional dark-themed interface.

    Features include:
    - Live countdown timers with color-coded status for 20 major exchanges worldwide
    - Interactive world map with clickable exchange markers (NASA Earth at Night imagery)
    - World clocks across multiple time zones
    - Market data sidebar: financial news (RSS), forex rates, cryptocurrency prices,
      stock indices, commodity prices, and individual stock quotes with profiles
    - Desktop notifications at configurable thresholds before exchange close
    - Secure API key management (Credential Manager / SecretManagement / CliXml)
    - Lunch break tracking for Tokyo, Shanghai, and Hong Kong
    - Holiday support via holidays.json
    - System tray integration, always-on-top toggle, sortable DataGrid
    - User preference persistence and multi-currency support
.EXAMPLE
    New-StockExchangeCountdownDashboard
    # Launch the dashboard
.EXAMPLE
    Start-PSExchangeClock
    # Launch via the convenience alias
.NOTES
    Author  : PowerShellYoungTeam
    Version : 1.0.0
    Project : https://github.com/PowerShellYoungTeam/PSExchangeClock
.LINK
    https://github.com/PowerShellYoungTeam/PSExchangeClock
.LINK
    Get-StockExchangeData
#>
    [CmdletBinding()]
    param()

    # ── Assemblies ────────────────────────────────────────────────
    Add-Type -AssemblyName PresentationFramework
    Add-Type -AssemblyName PresentationCore
    Add-Type -AssemblyName WindowsBase
    Add-Type -AssemblyName System.Windows.Forms

    # ── Load Exchange Data ────────────────────────────────────────
    # Resolve module root (Public/ is one level below the module directory)
    $moduleRoot = Split-Path $PSScriptRoot -Parent
    $dataDir = Join-Path $moduleRoot 'Data'

    $holidaysPath = Join-Path $dataDir 'holidays.json'
    $detailsPath = Join-Path $dataDir 'exchange-details.json'

    # User preferences stored in AppData (module dir may be read-only when installed)
    $appDataDir = Join-Path $env:APPDATA 'PSExchangeClock'
    if (-not (Test-Path $appDataDir)) {
        New-Item -Path $appDataDir -ItemType Directory -Force | Out-Null
    }
    $prefsPath = Join-Path $appDataDir 'user-preferences.json'
    # Seed from example if no user prefs exist yet
    if (-not (Test-Path $prefsPath)) {
        $examplePrefs = Join-Path $dataDir 'user-preferences.example.json'
        if (Test-Path $examplePrefs) {
            Copy-Item $examplePrefs $prefsPath -Force
        }
    }
    $mapImagePath = Join-Path $appDataDir 'worldmap.jpg'

    # Get-StockExchangeData is loaded as a module function
    try {
        $allExchanges = Get-StockExchangeData
    }
    catch {
        Write-Warning "Failed to load exchange data: $($_.Exception.Message). Using empty set."
        $allExchanges = @()
    }

    # Load holidays
    $holidays = @{}
    if (Test-Path $holidaysPath) {
        try {
            $holidayJson = Get-Content $holidaysPath -Raw | ConvertFrom-Json
            foreach ($prop in $holidayJson.PSObject.Properties) {
                $holidays[$prop.Name] = @($prop.Value)
            }
        }
        catch {
            Write-Warning "Failed to load holidays.json: $($_.Exception.Message)"
        }
    }

    # Load exchange details
    $script:exchangeDetails = @{}
    if (Test-Path $detailsPath) {
        try {
            $detailJson = Get-Content $detailsPath -Raw | ConvertFrom-Json
            foreach ($prop in $detailJson.PSObject.Properties) {
                $script:exchangeDetails[$prop.Name] = $prop.Value
            }
        }
        catch {
            Write-Warning "Failed to load exchange-details.json: $($_.Exception.Message)"
        }
    }

    # Load user preferences
    $script:userPrefs = $null
    if (Test-Path $prefsPath) {
        try {
            $script:userPrefs = Get-Content $prefsPath -Raw | ConvertFrom-Json
        }
        catch {
            Write-Warning "Failed to load user-preferences.json: $($_.Exception.Message)"
        }
    }

    # Download world map image if not present
    if (-not (Test-Path $mapImagePath)) {
        $mapUrl = 'https://eoimages.gsfc.nasa.gov/images/imagerecords/79000/79765/dnb_land_ocean_ice.2012.3600x1800.jpg'
        Write-Host 'Downloading NASA Earth at Night map image (first run only)...' -ForegroundColor Cyan
        try {
            Invoke-WebRequest -Uri $mapUrl -OutFile $mapImagePath -UseBasicParsing -TimeoutSec 60
            Write-Host 'Map image downloaded successfully.' -ForegroundColor Green
        }
        catch {
            Write-Warning "Failed to download map image: $($_.Exception.Message). Using polygon fallback."
        }
    }

    function Save-UserPreferences {
        $activeExchanges = @($script:exchangeRows | Where-Object { $_.IsActive } | ForEach-Object { $_.Code })
        $prefs = @{
            ActiveExchanges       = $activeExchanges
            NotifyThresholds      = @(30, 15, 5)
            AlwaysOnTop           = $false
            CommodityBaseCurrency = $script:commodityBaseCurrency
            SecretsBackend        = $script:secretsBackend
        }
        try {
            if ($chkNotify30) { $prefs.NotifyThresholds = @() }
            if ($chkNotify30 -and $chkNotify30.IsChecked) { $prefs.NotifyThresholds += 30 }
            if ($chkNotify15 -and $chkNotify15.IsChecked) { $prefs.NotifyThresholds += 15 }
            if ($chkNotify5 -and $chkNotify5.IsChecked) { $prefs.NotifyThresholds += 5 }
            if ($chkAlwaysOnTop) { $prefs.AlwaysOnTop = [bool]$chkAlwaysOnTop.IsChecked }
        }
        catch { }
        # API keys are stored via the secrets backend, not in JSON
        $prefs | ConvertTo-Json -Depth 3 | Set-Content -Path $prefsPath -Encoding UTF8
    }

    # ── Secure API Key Storage ────────────────────────────────────

    # Windows Credential Manager P/Invoke (used by CredentialManager backend)
    Add-Type -TypeDefinition @'
using System;
using System.Runtime.InteropServices;
using System.Text;

public static class CredManager {
    [DllImport("advapi32.dll", SetLastError = true, CharSet = CharSet.Unicode)]
    private static extern bool CredReadW(string target, int type, int flags, out IntPtr credential);

    [DllImport("advapi32.dll", SetLastError = true, CharSet = CharSet.Unicode)]
    private static extern bool CredWriteW(ref CREDENTIAL credential, int flags);

    [DllImport("advapi32.dll", SetLastError = true)]
    private static extern bool CredDeleteW(string target, int type, int flags);

    [DllImport("advapi32.dll")]
    private static extern void CredFree(IntPtr buffer);

    [StructLayout(LayoutKind.Sequential, CharSet = CharSet.Unicode)]
    private struct CREDENTIAL {
        public int Flags;
        public int Type;
        public string TargetName;
        public string Comment;
        public long LastWritten;
        public int CredentialBlobSize;
        public IntPtr CredentialBlob;
        public int Persist;
        public int AttributeCount;
        public IntPtr Attributes;
        public string TargetAlias;
        public string UserName;
    }

    private const int CRED_TYPE_GENERIC = 1;
    private const int CRED_PERSIST_LOCAL_MACHINE = 2;

    public static string Read(string target) {
        IntPtr credPtr;
        if (!CredReadW(target, CRED_TYPE_GENERIC, 0, out credPtr)) return null;
        try {
            CREDENTIAL cred = (CREDENTIAL)Marshal.PtrToStructure(credPtr, typeof(CREDENTIAL));
            if (cred.CredentialBlobSize > 0) {
                return Marshal.PtrToStringUni(cred.CredentialBlob, cred.CredentialBlobSize / 2);
            }
            return null;
        } finally { CredFree(credPtr); }
    }

    public static bool Write(string target, string secret, string userName) {
        byte[] bytes = Encoding.Unicode.GetBytes(secret);
        CREDENTIAL cred = new CREDENTIAL();
        cred.Type = CRED_TYPE_GENERIC;
        cred.TargetName = target;
        cred.CredentialBlobSize = bytes.Length;
        cred.CredentialBlob = Marshal.AllocHGlobal(bytes.Length);
        cred.Persist = CRED_PERSIST_LOCAL_MACHINE;
        cred.UserName = userName;
        try {
            Marshal.Copy(bytes, 0, cred.CredentialBlob, bytes.Length);
            return CredWriteW(ref cred, 0);
        } finally { Marshal.FreeHGlobal(cred.CredentialBlob); }
    }

    public static bool Delete(string target) {
        return CredDeleteW(target, CRED_TYPE_GENERIC, 0);
    }
}
'@ -ErrorAction SilentlyContinue

    function Get-SecureApiKey {
        param(
            [Parameter(Mandatory)][string]$ServiceName,
            [string]$Backend = $script:secretsBackend
        )
        $target = "CountdownDashboard_$ServiceName"
        switch ($Backend) {
            'CredentialManager' {
                try { return [CredManager]::Read($target) }
                catch { Write-Verbose "CredManager read failed: $_"; return $null }
            }
            'SecretManagement' {
                try {
                    if (-not (Get-Module -ListAvailable -Name Microsoft.PowerShell.SecretManagement)) { return $null }
                    Import-Module Microsoft.PowerShell.SecretManagement -ErrorAction Stop
                    $secret = Get-Secret -Name $target -AsPlainText -ErrorAction Stop
                    return $secret
                }
                catch { Write-Verbose "SecretManagement read failed: $_"; return $null }
            }
            'CliXml' {
                try {
                    $xmlPath = Join-Path $env:APPDATA 'CountdownDashboard\api-keys.clixml'
                    if (-not (Test-Path $xmlPath)) { return $null }
                    $store = Import-Clixml -Path $xmlPath
                    if ($store.ContainsKey($target)) {
                        $ss = $store[$target]
                        $bstr = [System.Runtime.InteropServices.Marshal]::SecureStringToBSTR($ss)
                        try { return [System.Runtime.InteropServices.Marshal]::PtrToStringBSTR($bstr) }
                        finally { [System.Runtime.InteropServices.Marshal]::ZeroFreeBSTR($bstr) }
                    }
                    return $null
                }
                catch { Write-Verbose "CliXml read failed: $_"; return $null }
            }
            default { return $null }
        }
    }

    function Set-SecureApiKey {
        param(
            [Parameter(Mandatory)][string]$ServiceName,
            [Parameter(Mandatory)][string]$ApiKey,
            [string]$Backend = $script:secretsBackend
        )
        if ([string]::IsNullOrWhiteSpace($ApiKey)) { return $false }
        $target = "CountdownDashboard_$ServiceName"
        switch ($Backend) {
            'CredentialManager' {
                try { return [CredManager]::Write($target, $ApiKey, $ServiceName) }
                catch { Write-Verbose "CredManager write failed: $_"; return $false }
            }
            'SecretManagement' {
                try {
                    if (-not (Get-Module -ListAvailable -Name Microsoft.PowerShell.SecretManagement)) { return $false }
                    Import-Module Microsoft.PowerShell.SecretManagement -ErrorAction Stop
                    Set-Secret -Name $target -Secret $ApiKey -ErrorAction Stop
                    return $true
                }
                catch { Write-Verbose "SecretManagement write failed: $_"; return $false }
            }
            'CliXml' {
                try {
                    $xmlDir = Join-Path $env:APPDATA 'CountdownDashboard'
                    $xmlPath = Join-Path $xmlDir 'api-keys.clixml'
                    if (-not (Test-Path $xmlDir)) { New-Item -Path $xmlDir -ItemType Directory -Force | Out-Null }
                    $store = @{}
                    if (Test-Path $xmlPath) { $store = Import-Clixml -Path $xmlPath }
                    $store[$target] = ConvertTo-SecureString -String $ApiKey -AsPlainText -Force
                    $store | Export-Clixml -Path $xmlPath -Force
                    return $true
                }
                catch { Write-Verbose "CliXml write failed: $_"; return $false }
            }
            default { return $false }
        }
    }

    function Remove-SecureApiKey {
        param(
            [Parameter(Mandatory)][string]$ServiceName,
            [string]$Backend = $script:secretsBackend
        )
        $target = "CountdownDashboard_$ServiceName"
        switch ($Backend) {
            'CredentialManager' {
                try { [CredManager]::Delete($target) | Out-Null } catch { }
            }
            'SecretManagement' {
                try {
                    if (Get-Module -ListAvailable -Name Microsoft.PowerShell.SecretManagement) {
                        Import-Module Microsoft.PowerShell.SecretManagement -ErrorAction Stop
                        Remove-Secret -Name $target -ErrorAction Stop
                    }
                }
                catch { }
            }
            'CliXml' {
                try {
                    $xmlPath = Join-Path $env:APPDATA 'CountdownDashboard\api-keys.clixml'
                    if (Test-Path $xmlPath) {
                        $store = Import-Clixml -Path $xmlPath
                        $store.Remove($target) | Out-Null
                        $store | Export-Clixml -Path $xmlPath -Force
                    }
                }
                catch { }
            }
        }
    }

    # ── Countdown Engine Functions ────────────────────────────────

    function Get-ExchangeLocalTime {
        param([PSCustomObject]$Exchange)
        try {
            $tz = [System.TimeZoneInfo]::FindSystemTimeZoneById($Exchange.TimeZoneId)
            return [System.TimeZoneInfo]::ConvertTimeFromUtc([DateTime]::UtcNow, $tz)
        }
        catch {
            return [DateTime]::UtcNow
        }
    }

    function Get-ExchangeStatus {
        param([PSCustomObject]$Exchange)

        $localNow = Get-ExchangeLocalTime -Exchange $Exchange
        $dayOfWeek = $localNow.DayOfWeek

        # Weekend check
        if ($dayOfWeek -eq 'Saturday' -or $dayOfWeek -eq 'Sunday') {
            return @{ Status = 'Closed'; Text = 'Weekend'; TimeRemaining = $null; CountdownTarget = $null }
        }

        # Holiday check
        $todayStr = $localNow.ToString('yyyy-MM-dd')
        if ($holidays.ContainsKey($Exchange.Code) -and $holidays[$Exchange.Code] -contains $todayStr) {
            return @{ Status = 'Holiday'; Text = 'Holiday - Closed'; TimeRemaining = $null; CountdownTarget = $null }
        }

        # Parse times
        $openParts = $Exchange.OpenTimeLocal -split ':'
        $closeParts = $Exchange.CloseTimeLocal -split ':'
        $openTime = $localNow.Date.AddHours([int]$openParts[0]).AddMinutes([int]$openParts[1])
        $closeTime = $localNow.Date.AddHours([int]$closeParts[0]).AddMinutes([int]$closeParts[1])

        # Lunch break
        $lunchStart = $null
        $lunchEnd = $null
        if ($Exchange.LunchBreakStart -and $Exchange.LunchBreakEnd) {
            $lsParts = $Exchange.LunchBreakStart -split ':'
            $leParts = $Exchange.LunchBreakEnd -split ':'
            $lunchStart = $localNow.Date.AddHours([int]$lsParts[0]).AddMinutes([int]$lsParts[1])
            $lunchEnd = $localNow.Date.AddHours([int]$leParts[0]).AddMinutes([int]$leParts[1])
        }

        # Pre-market
        if ($localNow -lt $openTime) {
            $remaining = $openTime - $localNow
            return @{ Status = 'PreMarket'; Text = "Opens in $("{0:D2}:{1:D2}:{2:D2}" -f [int]$remaining.TotalHours, $remaining.Minutes, $remaining.Seconds)"; TimeRemaining = $remaining; CountdownTarget = 'Open' }
        }

        # Post-close
        if ($localNow -ge $closeTime) {
            return @{ Status = 'Closed'; Text = 'Closed'; TimeRemaining = $null; CountdownTarget = $null }
        }

        # Lunch break
        if ($lunchStart -and $lunchEnd -and $localNow -ge $lunchStart -and $localNow -lt $lunchEnd) {
            $remaining = $lunchEnd - $localNow
            return @{ Status = 'LunchBreak'; Text = "Lunch - Reopens in $("{0:D2}:{1:D2}:{2:D2}" -f [int]$remaining.TotalHours, $remaining.Minutes, $remaining.Seconds)"; TimeRemaining = $remaining; CountdownTarget = 'Reopen' }
        }

        # Open — countdown to close (skip lunch if applicable)
        $remaining = $closeTime - $localNow
        $countdownStr = "{0:D2}:{1:D2}:{2:D2}" -f [int]$remaining.TotalHours, $remaining.Minutes, $remaining.Seconds

        if ($remaining.TotalMinutes -le 5) {
            return @{ Status = 'ClosingImminent'; Text = "Closes in $countdownStr"; TimeRemaining = $remaining; CountdownTarget = 'Close' }
        }
        elseif ($remaining.TotalMinutes -le 30) {
            return @{ Status = 'ClosingSoon'; Text = "Closes in $countdownStr"; TimeRemaining = $remaining; CountdownTarget = 'Close' }
        }
        else {
            return @{ Status = 'Open'; Text = "Closes in $countdownStr"; TimeRemaining = $remaining; CountdownTarget = 'Close' }
        }
    }

    function Get-StatusColor {
        param([string]$Status)
        switch ($Status) {
            'Open' { '#00CC66' }  # Green
            'ClosingSoon' { '#FFD700' }  # Yellow/Gold
            'ClosingImminent' { '#FF4444' }  # Red
            'Closed' { '#666666' }  # Gray
            'PreMarket' { '#4488FF' }  # Blue
            'LunchBreak' { '#FF8800' }  # Orange
            'Holiday' { '#9944CC' }  # Purple
            default { '#AAAAAA' }
        }
    }

    # ── XAML Definition ───────────────────────────────────────────

    $xaml = @'
<Window xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
        Title="PSExchangeClock — Global Market Dashboard"
        Height="750" Width="1100"
        MinHeight="500" MinWidth="800"
        Background="#1A1A2E"
        WindowStartupLocation="CenterScreen">
    <Window.Resources>
        <Style x:Key="TabStyle" TargetType="TabItem">
            <Setter Property="Foreground" Value="#CCCCCC"/>
            <Setter Property="FontSize" Value="13"/>
            <Setter Property="Padding" Value="12,6"/>
            <Setter Property="Template">
                <Setter.Value>
                    <ControlTemplate TargetType="TabItem">
                        <Border x:Name="tabBorder" Background="#16213E" BorderBrush="#0F3460" BorderThickness="1,1,1,0" CornerRadius="4,4,0,0" Margin="2,0">
                            <ContentPresenter x:Name="contentPresenter" ContentSource="Header" Margin="{TemplateBinding Padding}" HorizontalAlignment="Center"/>
                        </Border>
                        <ControlTemplate.Triggers>
                            <Trigger Property="IsSelected" Value="True">
                                <Setter TargetName="tabBorder" Property="Background" Value="#0F3460"/>
                                <Setter Property="Foreground" Value="#00CC66"/>
                            </Trigger>
                            <Trigger Property="IsMouseOver" Value="True">
                                <Setter TargetName="tabBorder" Property="Background" Value="#1A3A6E"/>
                            </Trigger>
                        </ControlTemplate.Triggers>
                    </ControlTemplate>
                </Setter.Value>
            </Setter>
        </Style>

        <!-- Dark-themed ComboBox ControlTemplate -->
        <ControlTemplate x:Key="DarkComboBoxToggleButton" TargetType="ToggleButton">
            <Grid>
                <Grid.ColumnDefinitions>
                    <ColumnDefinition/>
                    <ColumnDefinition Width="20"/>
                </Grid.ColumnDefinitions>
                <Border x:Name="Border" Grid.ColumnSpan="2" Background="#16213E" BorderBrush="#0F3460" BorderThickness="1" CornerRadius="3"/>
                <Border Grid.Column="0" Background="#16213E" BorderBrush="#0F3460" BorderThickness="0" CornerRadius="3,0,0,3" Margin="1"/>
                <Path x:Name="Arrow" Grid.Column="1" Fill="#E0E0E0" HorizontalAlignment="Center" VerticalAlignment="Center" Data="M 0 0 L 4 4 L 8 0 Z"/>
            </Grid>
            <ControlTemplate.Triggers>
                <Trigger Property="IsMouseOver" Value="True">
                    <Setter TargetName="Border" Property="Background" Value="#1A3A6E"/>
                </Trigger>
            </ControlTemplate.Triggers>
        </ControlTemplate>

        <Style x:Key="DarkComboBoxStyle" TargetType="ComboBox">
            <Setter Property="Foreground" Value="#E0E0E0"/>
            <Setter Property="Background" Value="#16213E"/>
            <Setter Property="BorderBrush" Value="#0F3460"/>
            <Setter Property="FontSize" Value="11"/>
            <Setter Property="Padding" Value="4,2"/>
            <Setter Property="MinWidth" Value="80"/>
            <Setter Property="Template">
                <Setter.Value>
                    <ControlTemplate TargetType="ComboBox">
                        <Grid>
                            <ToggleButton Name="ToggleButton" Template="{StaticResource DarkComboBoxToggleButton}"
                                          Focusable="False" IsChecked="{Binding Path=IsDropDownOpen, Mode=TwoWay, RelativeSource={RelativeSource TemplatedParent}}" ClickMode="Press"/>
                            <ContentPresenter Name="ContentSite" IsHitTestVisible="False"
                                              Content="{TemplateBinding SelectionBoxItem}" ContentTemplate="{TemplateBinding SelectionBoxItemTemplate}"
                                              Margin="6,3,23,3" VerticalAlignment="Center" HorizontalAlignment="Left"/>
                            <Popup Name="Popup" Placement="Bottom" IsOpen="{TemplateBinding IsDropDownOpen}" AllowsTransparency="True" Focusable="False" PopupAnimation="Slide">
                                <Grid Name="DropDown" SnapsToDevicePixels="True" MinWidth="{TemplateBinding ActualWidth}" MaxHeight="300">
                                    <Border x:Name="DropDownBorder" Background="#16213E" BorderBrush="#0F3460" BorderThickness="1" CornerRadius="0,0,3,3"/>
                                    <ScrollViewer Margin="2" SnapsToDevicePixels="True">
                                        <StackPanel IsItemsHost="True" KeyboardNavigation.DirectionalNavigation="Contained"/>
                                    </ScrollViewer>
                                </Grid>
                            </Popup>
                        </Grid>
                    </ControlTemplate>
                </Setter.Value>
            </Setter>
        </Style>

        <Style x:Key="DarkComboBoxItemStyle" TargetType="ComboBoxItem">
            <Setter Property="Foreground" Value="#E0E0E0"/>
            <Setter Property="Background" Value="#16213E"/>
            <Setter Property="Padding" Value="6,4"/>
            <Setter Property="FontSize" Value="11"/>
            <Setter Property="Template">
                <Setter.Value>
                    <ControlTemplate TargetType="ComboBoxItem">
                        <Border x:Name="Bd" Background="{TemplateBinding Background}" Padding="{TemplateBinding Padding}" SnapsToDevicePixels="True">
                            <ContentPresenter/>
                        </Border>
                        <ControlTemplate.Triggers>
                            <Trigger Property="IsHighlighted" Value="True">
                                <Setter TargetName="Bd" Property="Background" Value="#1A3A6E"/>
                                <Setter Property="Foreground" Value="#00CC66"/>
                            </Trigger>
                            <Trigger Property="IsMouseOver" Value="True">
                                <Setter TargetName="Bd" Property="Background" Value="#1A3A6E"/>
                                <Setter Property="Foreground" Value="#00CC66"/>
                            </Trigger>
                        </ControlTemplate.Triggers>
                    </ControlTemplate>
                </Setter.Value>
            </Setter>
        </Style>
    </Window.Resources>

    <Grid>
        <Grid.RowDefinitions>
            <RowDefinition Height="50"/>
            <RowDefinition Height="*"/>
            <RowDefinition Height="28"/>
        </Grid.RowDefinitions>

        <!-- Header -->
        <Border Grid.Row="0" Background="#16213E" BorderBrush="#0F3460" BorderThickness="0,0,0,1">
            <Grid Margin="15,0">
                <Grid.ColumnDefinitions>
                    <ColumnDefinition Width="*"/>
                    <ColumnDefinition Width="Auto"/>
                    <ColumnDefinition Width="Auto"/>
                </Grid.ColumnDefinitions>
                <TextBlock Grid.Column="0" Text="&#x1F4C8; Stock Exchange Countdown Dashboard" FontSize="20" FontWeight="Bold" Foreground="#E0E0E0" VerticalAlignment="Center"/>
                <TextBlock Grid.Column="1" x:Name="txtHeaderClock" Text="UTC: --:--:--" FontSize="14" Foreground="#00CC66" VerticalAlignment="Center" FontFamily="Consolas"/>
                <Button Grid.Column="2" x:Name="btnExit" Content="✕ Exit" Background="Transparent" Foreground="#AA6666" BorderThickness="0" FontSize="12" Cursor="Hand" Padding="10,4" Margin="12,0,0,0" VerticalAlignment="Center" ToolTip="Exit dashboard"/>
            </Grid>
        </Border>

        <!-- Tab Control -->
        <TabControl Grid.Row="1" x:Name="tabMain" Background="#1A1A2E" BorderBrush="#0F3460" Margin="5">

            <!-- ═══ Tab 1: Dashboard ═══ -->
            <TabItem Header="  Dashboard  " Style="{StaticResource TabStyle}">
                <Grid Background="#1A1A2E">
                    <Grid.ColumnDefinitions>
                        <ColumnDefinition Width="*"/>
                        <ColumnDefinition Width="200"/>
                    </Grid.ColumnDefinitions>
                    <Grid.RowDefinitions>
                        <RowDefinition Height="2*" MinHeight="200"/>
                        <RowDefinition Height="3"/>
                        <RowDefinition Height="3*" MinHeight="200"/>
                    </Grid.RowDefinitions>

                    <!-- World Map Area -->
                    <Border Grid.Row="0" Grid.Column="0" Background="#0D1B2A" CornerRadius="6" Margin="5" BorderBrush="#0F3460" BorderThickness="1">
                        <Grid>
                            <Grid.RowDefinitions>
                                <RowDefinition Height="Auto"/>
                                <RowDefinition Height="*"/>
                            </Grid.RowDefinitions>
                            <!-- Map Toolbar -->
                            <Border Grid.Row="0" Panel.ZIndex="20" Background="#16213E" CornerRadius="6,6,0,0" Padding="4,2" BorderBrush="#0F3460" BorderThickness="0,0,0,1">
                                <DockPanel>
                                    <StackPanel DockPanel.Dock="Left" Orientation="Horizontal" VerticalAlignment="Center">
                                        <TextBlock Text="Map:" Foreground="#888888" FontSize="10" VerticalAlignment="Center" Margin="4,0,4,0"/>
                                        <ComboBox x:Name="cmbMapStyle" Width="120" FontSize="10" Style="{StaticResource DarkComboBoxStyle}" ItemContainerStyle="{StaticResource DarkComboBoxItemStyle}" VerticalAlignment="Center"/>
                                        <TextBlock Text="  View:" Foreground="#888888" FontSize="10" VerticalAlignment="Center" Margin="8,0,4,0"/>
                                        <ComboBox x:Name="cmbMapProjection" Width="100" FontSize="10" Style="{StaticResource DarkComboBoxStyle}" ItemContainerStyle="{StaticResource DarkComboBoxItemStyle}" VerticalAlignment="Center"/>
                                    </StackPanel>
                                    <StackPanel DockPanel.Dock="Right" Orientation="Horizontal" HorizontalAlignment="Right" VerticalAlignment="Center">
                                        <ToggleButton x:Name="btnOverlays" Content=" Overlays ▾ " Background="#0F3460" Foreground="#E0E0E0" BorderThickness="0" FontSize="10" Cursor="Hand" Padding="6,2" Margin="4,0" VerticalAlignment="Center" ToolTip="Toggle map overlays"/>
                                        <Rectangle Width="1" Height="16" Fill="#0F3460" Margin="6,0"/>
                                        <Button x:Name="btnZoomIn" Content=" + " Background="#0F3460" Foreground="#E0E0E0" BorderThickness="0" FontSize="12" FontWeight="Bold" Cursor="Hand" Padding="4,0" Margin="2,0" VerticalAlignment="Center"/>
                                        <Button x:Name="btnZoomOut" Content=" − " Background="#0F3460" Foreground="#E0E0E0" BorderThickness="0" FontSize="12" FontWeight="Bold" Cursor="Hand" Padding="4,0" Margin="2,0" VerticalAlignment="Center"/>
                                        <Button x:Name="btnZoomReset" Content=" ↺ " Background="#0F3460" Foreground="#E0E0E0" BorderThickness="0" FontSize="12" Cursor="Hand" Padding="4,0" Margin="2,0" VerticalAlignment="Center" ToolTip="Reset zoom"/>
                                        <ToggleButton x:Name="btnGlobeSpin" Content=" ⟳ " Background="#0F3460" Foreground="#E0E0E0" BorderThickness="0" FontSize="12" Cursor="Hand" Padding="4,0" Margin="2,0" VerticalAlignment="Center" ToolTip="Auto-rotate globe"/>
                                    </StackPanel>
                                </DockPanel>
                            </Border>
                            <!-- Overlay selector popup -->
                            <Popup x:Name="popOverlays" Placement="Bottom" StaysOpen="False" AllowsTransparency="True">
                                <Border Background="#1A1A2E" BorderBrush="#0F3460" BorderThickness="1" CornerRadius="4" Padding="8,6">
                                    <StackPanel>
                                        <TextBlock Text="── Geo ──" Foreground="#5588AA" FontSize="9" FontWeight="Bold" Margin="0,0,0,3"/>
                                        <CheckBox x:Name="chkTerminator" Content="Day / Night" Foreground="#CCCCCC" FontSize="10" Margin="2,2" IsChecked="True"/>
                                        <CheckBox x:Name="chkTimeZoneBands" Content="TZ Bands" Foreground="#CCCCCC" FontSize="10" Margin="2,2"/>
                                        <CheckBox x:Name="chkPoliticalBounds" Content="Borders" Foreground="#CCCCCC" FontSize="10" Margin="2,2"/>
                                        <CheckBox x:Name="chkTZBoundaries" Content="TZ Boundaries" Foreground="#CCCCCC" FontSize="10" Margin="2,2"/>
                                        <TextBlock Text="── Live Data ──" Foreground="#55AA88" FontSize="9" FontWeight="Bold" Margin="0,6,0,3"/>
                                        <CheckBox x:Name="chkEarthquakes" Content="Earthquakes" Foreground="#CCCCCC" FontSize="10" Margin="2,2"/>
                                        <CheckBox x:Name="chkVolcanoes" Content="Volcanoes" Foreground="#CCCCCC" FontSize="10" Margin="2,2"/>
                                        <CheckBox x:Name="chkConflictZones" Content="Conflict Zones" Foreground="#CCCCCC" FontSize="10" Margin="2,2"/>
                                        <TextBlock Text="── Infrastructure ──" Foreground="#AA8855" FontSize="9" FontWeight="Bold" Margin="0,6,0,3"/>
                                        <CheckBox x:Name="chkSubmarineCables" Content="Submarine Cables" Foreground="#CCCCCC" FontSize="10" Margin="2,2"/>
                                        <CheckBox x:Name="chkPowerPlants" Content="Power Plants" Foreground="#CCCCCC" FontSize="10" Margin="2,2"/>
                                    </StackPanel>
                                </Border>
                            </Popup>
                            <!-- Clip container so zoomed canvas cannot overflow into toolbar -->
                            <Border Grid.Row="1" ClipToBounds="True">
                                <Canvas x:Name="canvasMap" Background="Transparent" ClipToBounds="True"/>
                            </Border>

                            <!-- Market Data Panel (shown when no exchange flyout is open) -->
                            <Border Grid.Row="1" Panel.ZIndex="10" x:Name="marketDataPanel" Background="#CC1A1A2E" BorderBrush="#0F3460" BorderThickness="1" CornerRadius="6"
                                    Width="300" HorizontalAlignment="Left" VerticalAlignment="Stretch" Margin="5,5,0,5"
                                    Visibility="Visible">
                                <Grid>
                                    <Grid.RowDefinitions>
                                        <RowDefinition Height="Auto"/>
                                        <RowDefinition Height="Auto"/>
                                        <RowDefinition Height="*"/>
                                    </Grid.RowDefinitions>
                                    <!-- Market Data Header -->
                                    <Border Grid.Row="0" Background="#0F3460" CornerRadius="6,6,0,0" Padding="8,6">
                                        <TextBlock x:Name="marketDataTitle" Text="Market Data" FontSize="13" FontWeight="Bold" Foreground="#E0E0E0" HorizontalAlignment="Center"/>
                                    </Border>
                                    <!-- Tab Buttons -->
                                    <WrapPanel Grid.Row="1" x:Name="marketTabBar" HorizontalAlignment="Center" Margin="4,6,4,2">
                                        <Button x:Name="btnTabNews" Content=" News " Background="#0F3460" Foreground="#00CC66" BorderBrush="#00CC66" BorderThickness="1" Margin="2" Padding="8,3" FontSize="10" Cursor="Hand"/>
                                        <Button x:Name="btnTabFx" Content=" FX Rates " Background="#16213E" Foreground="#AAAAAA" BorderBrush="#0F3460" BorderThickness="1" Margin="2" Padding="8,3" FontSize="10" Cursor="Hand"/>
                                        <Button x:Name="btnTabCrypto" Content=" Crypto " Background="#16213E" Foreground="#AAAAAA" BorderBrush="#0F3460" BorderThickness="1" Margin="2" Padding="8,3" FontSize="10" Cursor="Hand"/>
                                        <Button x:Name="btnTabIndices" Content=" Indices " Background="#16213E" Foreground="#AAAAAA" BorderBrush="#0F3460" BorderThickness="1" Margin="2" Padding="8,3" FontSize="10" Cursor="Hand" Visibility="Collapsed"/>
                                        <Button x:Name="btnTabCommodities" Content=" Commodities " Background="#16213E" Foreground="#AAAAAA" BorderBrush="#0F3460" BorderThickness="1" Margin="2" Padding="8,3" FontSize="10" Cursor="Hand" Visibility="Collapsed"/>
                                        <Button x:Name="btnTabStocks" Content=" Stocks " Background="#16213E" Foreground="#AAAAAA" BorderBrush="#0F3460" BorderThickness="1" Margin="2" Padding="8,3" FontSize="10" Cursor="Hand" Visibility="Collapsed"/>
                                    </WrapPanel>
                                    <!-- Tab Content Area -->
                                    <ScrollViewer Grid.Row="2" VerticalScrollBarVisibility="Auto" Margin="0">
                                        <StackPanel x:Name="marketDataContent" Margin="10,6"/>
                                    </ScrollViewer>
                                </Grid>
                            </Border>

                            <!-- Detail Flyout Panel (hidden by default, LEFT aligned to not cover Far East) -->
                            <Border Grid.Row="1" Panel.ZIndex="10" x:Name="flyoutPanel" Background="#1A1A2E" BorderBrush="#0F3460" BorderThickness="1" CornerRadius="6"
                                    Width="280" HorizontalAlignment="Left" VerticalAlignment="Stretch" Margin="5,5,0,5"
                                    Visibility="Collapsed">
                                <Grid>
                                    <Grid.RowDefinitions>
                                        <RowDefinition Height="Auto"/>
                                        <RowDefinition Height="*"/>
                                    </Grid.RowDefinitions>
                                    <!-- Flyout Header -->
                                    <Border Grid.Row="0" Background="#0F3460" CornerRadius="6,6,0,0" Padding="10,8">
                                        <Grid>
                                            <Grid.ColumnDefinitions>
                                                <ColumnDefinition Width="*"/>
                                                <ColumnDefinition Width="Auto"/>
                                            </Grid.ColumnDefinitions>
                                            <TextBlock Grid.Column="0" x:Name="flyoutTitle" Text="Exchange Details" FontSize="14" FontWeight="Bold" Foreground="#E0E0E0" VerticalAlignment="Center"/>
                                            <Button Grid.Column="1" x:Name="btnCloseFlyout" Content="  X  " Background="Transparent" Foreground="#AAAAAA" BorderThickness="0" FontSize="14" FontWeight="Bold" Cursor="Hand" VerticalAlignment="Center"/>
                                        </Grid>
                                    </Border>
                                    <!-- Flyout Content -->
                                    <ScrollViewer Grid.Row="1" VerticalScrollBarVisibility="Auto" Margin="0">
                                        <StackPanel x:Name="flyoutContent" Margin="12,10"/>
                                    </ScrollViewer>
                                </Grid>
                            </Border>
                        </Grid>
                    </Border>

                    <!-- World Clocks Sidebar -->
                    <Border Grid.Row="0" Grid.RowSpan="3" Grid.Column="1" Background="#16213E" CornerRadius="6" Margin="5" BorderBrush="#0F3460" BorderThickness="1">
                        <ScrollViewer VerticalScrollBarVisibility="Auto">
                            <StackPanel x:Name="panelWorldClocks" Margin="10,10">
                                <TextBlock Text="World Clocks" FontSize="14" FontWeight="Bold" Foreground="#E0E0E0" Margin="0,0,0,10"/>
                            </StackPanel>
                        </ScrollViewer>
                    </Border>

                    <!-- Splitter -->
                    <GridSplitter Grid.Row="1" Grid.Column="0" Height="3" HorizontalAlignment="Stretch" Background="#0F3460"/>

                    <!-- Countdown Cards -->
                    <Border Grid.Row="2" Grid.Column="0" Background="#0D1B2A" CornerRadius="6" Margin="5" BorderBrush="#0F3460" BorderThickness="1">
                        <ScrollViewer VerticalScrollBarVisibility="Auto" Padding="5">
                            <WrapPanel x:Name="panelCards" Margin="5"/>
                        </ScrollViewer>
                    </Border>
                </Grid>
            </TabItem>

            <!-- ═══ Tab 2: All Exchanges ═══ -->
            <TabItem Header="  All Exchanges  " Style="{StaticResource TabStyle}">
                <Grid Background="#1A1A2E" Margin="5">
                    <Grid.RowDefinitions>
                        <RowDefinition Height="40"/>
                        <RowDefinition Height="*"/>
                    </Grid.RowDefinitions>

                    <!-- Filter Bar -->
                    <Grid Grid.Row="0" Margin="5,5,5,0">
                        <Grid.ColumnDefinitions>
                            <ColumnDefinition Width="Auto"/>
                            <ColumnDefinition Width="250"/>
                            <ColumnDefinition Width="Auto"/>
                            <ColumnDefinition Width="Auto"/>
                            <ColumnDefinition Width="*"/>
                        </Grid.ColumnDefinitions>
                        <TextBlock Grid.Column="0" Text="Filter: " Foreground="#E0E0E0" VerticalAlignment="Center" Margin="5,0" FontSize="13"/>
                        <TextBox Grid.Column="1" x:Name="txtFilter" Background="#16213E" Foreground="#E0E0E0" BorderBrush="#0F3460" Padding="5,3" FontSize="13" VerticalAlignment="Center"/>
                        <Button Grid.Column="2" x:Name="btnSelectAll" Content="  Select All  " Background="#0F3460" Foreground="#E0E0E0" Padding="8,3" BorderBrush="#00CC66" Margin="10,0,5,0" FontSize="11" Cursor="Hand" VerticalAlignment="Center"/>
                        <Button Grid.Column="3" x:Name="btnDeselectAll" Content="  Deselect All  " Background="#0F3460" Foreground="#E0E0E0" Padding="8,3" BorderBrush="#FFD700" Margin="0,0,5,0" FontSize="11" Cursor="Hand" VerticalAlignment="Center"/>
                    </Grid>

                    <!-- DataGrid -->
                    <DataGrid Grid.Row="1" x:Name="gridExchanges"
                              AutoGenerateColumns="False"
                              CanUserAddRows="False" CanUserDeleteRows="False"
                              IsReadOnly="False"
                              Background="#16213E" Foreground="#E0E0E0"
                              RowBackground="#16213E" AlternatingRowBackground="#1A2744"
                              BorderBrush="#0F3460" GridLinesVisibility="Horizontal"
                              HorizontalGridLinesBrush="#0F3460"
                              HeadersVisibility="Column"
                              SelectionMode="Single"
                              ColumnHeaderHeight="32"
                              RowHeight="28"
                              Margin="5"
                              FontSize="12">
                        <DataGrid.ColumnHeaderStyle>
                            <Style TargetType="DataGridColumnHeader">
                                <Setter Property="Background" Value="#0F3460"/>
                                <Setter Property="Foreground" Value="#E0E0E0"/>
                                <Setter Property="Padding" Value="8,4"/>
                                <Setter Property="BorderBrush" Value="#0A2A50"/>
                                <Setter Property="BorderThickness" Value="0,0,1,1"/>
                                <Setter Property="FontWeight" Value="SemiBold"/>
                            </Style>
                        </DataGrid.ColumnHeaderStyle>
                        <DataGrid.Columns>
                            <DataGridCheckBoxColumn Header="Active" Binding="{Binding IsActive, Mode=TwoWay, UpdateSourceTrigger=PropertyChanged}" Width="55"/>
                            <DataGridTextColumn Header="Exchange" Binding="{Binding DisplayName}" Width="200" IsReadOnly="True"/>
                            <DataGridTextColumn Header="Code" Binding="{Binding Code}" Width="65" IsReadOnly="True"/>
                            <DataGridTextColumn Header="Country" Binding="{Binding Country}" Width="110" IsReadOnly="True"/>
                            <DataGridTextColumn Header="Open" Binding="{Binding OpenTime}" Width="60" IsReadOnly="True"/>
                            <DataGridTextColumn Header="Close" Binding="{Binding CloseTime}" Width="60" IsReadOnly="True"/>
                            <DataGridTextColumn Header="Status" Binding="{Binding StatusText}" Width="120" IsReadOnly="True"/>
                            <DataGridTextColumn Header="Countdown" Binding="{Binding CountdownText}" Width="100" IsReadOnly="True"/>
                            <DataGridTextColumn Header="Local Time" Binding="{Binding CurrentLocalTime}" Width="90" IsReadOnly="True"/>
                        </DataGrid.Columns>
                    </DataGrid>
                </Grid>
            </TabItem>

            <!-- ═══ Tab 3: Settings ═══ -->
            <TabItem Header="  Settings  " Style="{StaticResource TabStyle}">
                <ScrollViewer VerticalScrollBarVisibility="Auto" Background="#1A1A2E">
                    <StackPanel Margin="20">
                        <TextBlock Text="Notification Settings" FontSize="16" FontWeight="Bold" Foreground="#E0E0E0" Margin="0,0,0,10"/>

                        <CheckBox x:Name="chkNotify30" Content="Notify 30 minutes before close" Foreground="#E0E0E0" IsChecked="True" Margin="10,5" FontSize="13"/>
                        <CheckBox x:Name="chkNotify15" Content="Notify 15 minutes before close" Foreground="#E0E0E0" IsChecked="True" Margin="10,5" FontSize="13"/>
                        <CheckBox x:Name="chkNotify5" Content="Notify 5 minutes before close" Foreground="#E0E0E0" IsChecked="True" Margin="10,5" FontSize="13"/>

                        <TextBlock Text="Display Settings" FontSize="16" FontWeight="Bold" Foreground="#E0E0E0" Margin="0,20,0,10"/>

                        <CheckBox x:Name="chkAlwaysOnTop" Content="Always on top" Foreground="#E0E0E0" Margin="10,5" FontSize="13"/>
                        <CheckBox x:Name="chkShowGridLines" Content="Show grid lines on map" Foreground="#E0E0E0" Margin="10,5" FontSize="13"/>

                        <TextBlock Text="Data Management" FontSize="16" FontWeight="Bold" Foreground="#E0E0E0" Margin="0,20,0,10"/>

                        <StackPanel Orientation="Horizontal" Margin="10,5">
                            <Button x:Name="btnRefresh" Content="  Refresh Exchange Data  " Background="#0F3460" Foreground="#E0E0E0" Padding="10,5" BorderBrush="#00CC66" Margin="0,0,10,0" FontSize="13" Cursor="Hand"/>
                            <Button x:Name="btnResetDefaults" Content="  Reset to Defaults  " Background="#0F3460" Foreground="#E0E0E0" Padding="10,5" BorderBrush="#FFD700" Margin="0,0,10,0" FontSize="13" Cursor="Hand"/>
                        </StackPanel>

                        <TextBlock Text="Market Data API Keys (Optional)" FontSize="16" FontWeight="Bold" Foreground="#E0E0E0" Margin="0,20,0,10"/>
                        <TextBlock Text="Enter free API keys to unlock stock indices and individual stock quotes in the Market Data panel." Foreground="#888888" FontSize="11" Margin="10,0,0,8" TextWrapping="Wrap"/>

                        <StackPanel Margin="10,5">
                            <TextBlock Text="Twelve Data API Key (free: twelvedata.com/pricing — 800 req/day, unlocks Indices tab)" Foreground="#AAAAAA" FontSize="11" Margin="0,0,0,3"/>
                            <Grid>
                                <Grid.ColumnDefinitions>
                                    <ColumnDefinition Width="*"/>
                                    <ColumnDefinition Width="Auto"/>
                                </Grid.ColumnDefinitions>
                                <TextBox Grid.Column="0" x:Name="txtTwelveDataKey" Background="#16213E" Foreground="#E0E0E0" BorderBrush="#0F3460" Padding="5,3" FontSize="12" FontFamily="Consolas"/>
                                <TextBlock Grid.Column="1" x:Name="txtTwelveDataStatus" Text="" Foreground="#888888" FontSize="14" VerticalAlignment="Center" Margin="8,0,0,0"/>
                            </Grid>
                        </StackPanel>

                        <StackPanel Margin="10,10,10,5">
                            <TextBlock Text="Alpha Vantage API Key (free: alphavantage.co/support/#api-key — 25 req/day, unlocks Stocks tab)" Foreground="#AAAAAA" FontSize="11" Margin="0,0,0,3"/>
                            <Grid>
                                <Grid.ColumnDefinitions>
                                    <ColumnDefinition Width="*"/>
                                    <ColumnDefinition Width="Auto"/>
                                </Grid.ColumnDefinitions>
                                <TextBox Grid.Column="0" x:Name="txtAlphaVantageKey" Background="#16213E" Foreground="#E0E0E0" BorderBrush="#0F3460" Padding="5,3" FontSize="12" FontFamily="Consolas"/>
                                <TextBlock Grid.Column="1" x:Name="txtAlphaVantageStatus" Text="" Foreground="#888888" FontSize="14" VerticalAlignment="Center" Margin="8,0,0,0"/>
                            </Grid>
                        </StackPanel>

                        <Button x:Name="btnSaveApiKeys" Content="  Save API Keys  " Background="#0F3460" Foreground="#E0E0E0" Padding="10,5" BorderBrush="#00CC66" Margin="10,10,10,0" FontSize="13" Cursor="Hand" HorizontalAlignment="Left"/>

                        <StackPanel Margin="10,15,10,5">
                            <TextBlock Text="Key Storage Backend" Foreground="#AAAAAA" FontSize="11" Margin="0,0,0,3"/>
                            <ComboBox x:Name="cmbSecretsBackend" Width="220" HorizontalAlignment="Left" Style="{StaticResource DarkComboBoxStyle}" ItemContainerStyle="{StaticResource DarkComboBoxItemStyle}"/>
                            <TextBlock x:Name="txtSecretsBackendNote" Text="" Foreground="#666666" FontSize="9" Margin="0,3,0,0" TextWrapping="Wrap"/>
                        </StackPanel>

                        <TextBlock x:Name="txtLastUpdated" Text="Last updated: --" Foreground="#888888" Margin="10,15,0,0" FontSize="12"/>
                    </StackPanel>
                </ScrollViewer>
            </TabItem>
        </TabControl>

        <!-- Status Bar -->
        <Border Grid.Row="2" Background="#16213E" BorderBrush="#0F3460" BorderThickness="0,1,0,0">
            <Grid Margin="10,0">
                <Grid.ColumnDefinitions>
                    <ColumnDefinition Width="*"/>
                    <ColumnDefinition Width="Auto"/>
                </Grid.ColumnDefinitions>
                <TextBlock Grid.Column="0" x:Name="txtStatus" Text="Ready" Foreground="#AAAAAA" VerticalAlignment="Center" FontSize="11"/>
                <TextBlock Grid.Column="1" x:Name="txtNextClose" Text="" Foreground="#FFD700" VerticalAlignment="Center" FontSize="11" FontWeight="SemiBold"/>
            </Grid>
        </Border>
    </Grid>
</Window>
'@

    # ── Build Window ──────────────────────────────────────────────
    [xml]$xamlXml = $xaml
    $reader = New-Object System.Xml.XmlNodeReader $xamlXml
    $window = [Windows.Markup.XamlReader]::Load($reader)

    # Find named controls
    $txtHeaderClock = $window.FindName('txtHeaderClock')
    $canvasMap = $window.FindName('canvasMap')
    $panelWorldClocks = $window.FindName('panelWorldClocks')
    $panelCards = $window.FindName('panelCards')
    $gridExchanges = $window.FindName('gridExchanges')
    $txtFilter = $window.FindName('txtFilter')
    $chkNotify30 = $window.FindName('chkNotify30')
    $chkNotify15 = $window.FindName('chkNotify15')
    $chkNotify5 = $window.FindName('chkNotify5')
    $chkAlwaysOnTop = $window.FindName('chkAlwaysOnTop')
    $btnRefresh = $window.FindName('btnRefresh')
    $btnResetDefaults = $window.FindName('btnResetDefaults')
    $txtLastUpdated = $window.FindName('txtLastUpdated')
    $txtStatus = $window.FindName('txtStatus')
    $txtNextClose = $window.FindName('txtNextClose')
    $btnSelectAll = $window.FindName('btnSelectAll')
    $btnDeselectAll = $window.FindName('btnDeselectAll')
    $chkShowGridLines = $window.FindName('chkShowGridLines')

    # Map toolbar controls
    $cmbMapStyle = $window.FindName('cmbMapStyle')
    $cmbMapProjection = $window.FindName('cmbMapProjection')
    $chkTerminator = $window.FindName('chkTerminator')
    $chkTimeZoneBands = $window.FindName('chkTimeZoneBands')
    $chkPoliticalBounds = $window.FindName('chkPoliticalBounds')
    $chkEarthquakes = $window.FindName('chkEarthquakes')
    $chkVolcanoes = $window.FindName('chkVolcanoes')
    $chkSubmarineCables = $window.FindName('chkSubmarineCables')
    $chkPowerPlants = $window.FindName('chkPowerPlants')
    $chkTZBoundaries = $window.FindName('chkTZBoundaries')
    $chkConflictZones = $window.FindName('chkConflictZones')
    $btnOverlays = $window.FindName('btnOverlays')
    $popOverlays = $window.FindName('popOverlays')
    $btnZoomIn = $window.FindName('btnZoomIn')
    $btnZoomOut = $window.FindName('btnZoomOut')
    $btnZoomReset = $window.FindName('btnZoomReset')
    $btnGlobeSpin = $window.FindName('btnGlobeSpin')
    $btnExit = $window.FindName('btnExit')
    $flyoutPanel = $window.FindName('flyoutPanel')
    $flyoutTitle = $window.FindName('flyoutTitle')
    $flyoutContent = $window.FindName('flyoutContent')
    $btnCloseFlyout = $window.FindName('btnCloseFlyout')

    # Market Data Panel controls
    $marketDataPanel = $window.FindName('marketDataPanel')
    $marketDataTitle = $window.FindName('marketDataTitle')
    $marketTabBar = $window.FindName('marketTabBar')
    $marketDataContent = $window.FindName('marketDataContent')
    $btnTabNews = $window.FindName('btnTabNews')
    $btnTabFx = $window.FindName('btnTabFx')
    $btnTabCrypto = $window.FindName('btnTabCrypto')
    $btnTabIndices = $window.FindName('btnTabIndices')
    $btnTabCommodities = $window.FindName('btnTabCommodities')
    $btnTabStocks = $window.FindName('btnTabStocks')

    # API key controls
    $txtTwelveDataKey = $window.FindName('txtTwelveDataKey')
    $txtAlphaVantageKey = $window.FindName('txtAlphaVantageKey')
    $txtTwelveDataStatus = $window.FindName('txtTwelveDataStatus')
    $txtAlphaVantageStatus = $window.FindName('txtAlphaVantageStatus')
    $btnSaveApiKeys = $window.FindName('btnSaveApiKeys')
    $cmbSecretsBackend = $window.FindName('cmbSecretsBackend')
    $txtSecretsBackendNote = $window.FindName('txtSecretsBackendNote')

    # Populate secrets backend ComboBox
    $secretsBackendOptions = @(
        @{ Label = 'Windows Credential Manager'; Value = 'CredentialManager'; Note = 'Default. Uses Windows Credential Manager (DPAPI-encrypted, per-user).' },
        @{ Label = 'SecretManagement Module'; Value = 'SecretManagement'; Note = 'Requires Microsoft.PowerShell.SecretManagement module and a registered vault.' },
        @{ Label = 'Encrypted XML (DPAPI)'; Value = 'CliXml'; Note = 'Export-Clixml with DPAPI. Stored in %APPDATA%\CountdownDashboard\api-keys.clixml.' }
    )
    foreach ($opt in $secretsBackendOptions) {
        $item = New-Object System.Windows.Controls.ComboBoxItem
        $item.Content = $opt.Label; $item.Tag = $opt.Value
        $cmbSecretsBackend.Items.Add($item) | Out-Null
    }
    $cmbSecretsBackend.SelectedIndex = 0
    $txtSecretsBackendNote.Text = $secretsBackendOptions[0].Note

    $cmbSecretsBackend.Add_SelectionChanged({
            $selected = $cmbSecretsBackend.SelectedItem
            if ($selected -and $selected.Tag) {
                $script:secretsBackend = $selected.Tag
                $note = ($secretsBackendOptions | Where-Object { $_.Value -eq $selected.Tag }).Note
                $txtSecretsBackendNote.Text = $note
            }
        }.GetNewClosure())

    $script:currentFlyoutCode = $null
    $script:activeMarketTab = 'News'
    $script:fxBaseCurrency = 'USD'
    $script:cryptoCurrency = 'usd'
    $script:commodityBaseCurrency = 'USD'
    $script:secretsBackend = 'CredentialManager'

    # ── Map State ─────────────────────────────────────────────────

    $script:mapMarkers = @{}
    $script:mapClusters = @{}
    $script:expandedCluster = $null
    $script:mapZoom = 1.0
    $script:mapPanX = 0.0
    $script:mapPanY = 0.0
    $script:mapDragging = $false
    $script:mapDragStart = $null
    $script:mapPanStart = $null
    $script:mapProjection = 'Flat'  # 'Flat' or 'Globe'
    $script:globeCenterLat = 20.0
    $script:globeCenterLon = 0.0
    $script:globeSpinning = $false
    $script:earthquakeData = $null
    $script:volcanoData = $null
    $script:liveFeedLastFetch = [DateTime]::MinValue
    $script:submarineCableData = $null
    $script:powerPlantData = $null
    $script:tzBoundaryData = $null
    $script:conflictZoneData = $null

    # Populate Map Style ComboBox
    $mapStyleOptions = @(
        @{ Label = 'Earth at Night'; Value = 'NightSatellite' },
        @{ Label = 'Blue Marble'; Value = 'BlueMarble' },
        @{ Label = 'Vector Dark'; Value = 'VectorDark' },
        @{ Label = 'Minimal'; Value = 'Minimal' }
    )
    foreach ($opt in $mapStyleOptions) {
        $item = New-Object System.Windows.Controls.ComboBoxItem
        $item.Content = $opt.Label; $item.Tag = $opt.Value
        $cmbMapStyle.Items.Add($item) | Out-Null
    }
    $cmbMapStyle.SelectedIndex = 0

    # Populate Map Projection ComboBox
    $projOptions = @(
        @{ Label = 'Flat Map'; Value = 'Flat' },
        @{ Label = 'Globe'; Value = 'Globe' }
    )
    foreach ($opt in $projOptions) {
        $item = New-Object System.Windows.Controls.ComboBoxItem
        $item.Content = $opt.Label; $item.Tag = $opt.Value
        $cmbMapProjection.Items.Add($item) | Out-Null
    }
    $cmbMapProjection.SelectedIndex = 0

    # ── Market Data Cache ─────────────────────────────────────────

    $script:newsCache = @{ Data = $null; LastFetch = [datetime]::MinValue }
    $script:fxCache = @{ Data = $null; LastFetch = [datetime]::MinValue }
    $script:cryptoCache = @{ Data = $null; LastFetch = [datetime]::MinValue }
    $script:indicesCache = @{ Data = $null; LastFetch = [datetime]::MinValue }
    $script:commoditiesCache = @{ Data = $null; LastFetch = [datetime]::MinValue }
    $script:stockQuoteCache = @{}

    # ── Market Data Fetch Functions ───────────────────────────────

    function Get-FinancialNews {
        # Check cache (5 minute TTL)
        if ($script:newsCache.Data -and ([datetime]::Now - $script:newsCache.LastFetch).TotalMinutes -lt 5) {
            return $script:newsCache.Data
        }

        $feeds = @(
            @{ Name = 'Reuters'; Url = 'https://feeds.reuters.com/reuters/businessNews' },
            @{ Name = 'BBC'; Url = 'https://feeds.bbc.co.uk/news/business/rss.xml' },
            @{ Name = 'CNBC'; Url = 'https://search.cnbc.com/rs/search/combinedcms/view.xml?partnerId=wrss01&id=100003114' }
        )

        $allItems = @()
        foreach ($feed in $feeds) {
            try {
                $response = Invoke-WebRequest -Uri $feed.Url -UseBasicParsing -TimeoutSec 10
                [xml]$rss = $response.Content
                $items = $rss.rss.channel.item | Select-Object -First 5
                foreach ($item in $items) {
                    $pubDate = $null
                    try { $pubDate = [datetime]::Parse($item.pubDate) } catch { $pubDate = [datetime]::Now }
                    $allItems += [PSCustomObject]@{
                        Title     = ($item.title -replace '<[^>]+>', '').Trim()
                        Source    = $feed.Name
                        Published = $pubDate
                        Link      = $item.link
                    }
                }
            }
            catch {
                Write-Verbose "Failed to fetch $($feed.Name) RSS: $($_.Exception.Message)"
            }
        }

        if ($allItems.Count -eq 0) { return $null }

        # Deduplicate by title similarity (exact match after lowercase + trim)
        $seen = @{}
        $deduped = @()
        foreach ($item in ($allItems | Sort-Object Published -Descending)) {
            $key = ($item.Title.ToLower().Trim() -replace '[^a-z0-9 ]', '')
            if (-not $seen.ContainsKey($key)) {
                $seen[$key] = $true
                $deduped += $item
            }
        }

        $result = $deduped | Select-Object -First 10
        $script:newsCache.Data = $result
        $script:newsCache.LastFetch = [datetime]::Now
        return $result
    }

    function Get-ForexRates {
        param([string]$BaseCurrency = 'USD')
        # Check cache (10 minute TTL, invalidate if base changed)
        if ($script:fxCache.Data -and $script:fxCache.Data.Base -eq $BaseCurrency -and ([datetime]::Now - $script:fxCache.LastFetch).TotalMinutes -lt 10) {
            return $script:fxCache.Data
        }

        $allCurrencies = @('USD', 'EUR', 'GBP', 'JPY', 'CHF', 'CAD', 'AUD', 'CNY', 'HKD', 'SGD', 'KRW', 'INR', 'BRL', 'ZAR', 'MXN')
        $currencies = ($allCurrencies | Where-Object { $_ -ne $BaseCurrency }) -join ','

        # Try Frankfurter API first
        try {
            $url = "https://api.frankfurter.dev/v1/latest?base=$BaseCurrency&symbols=$currencies"
            $response = Invoke-RestMethod -Uri $url -TimeoutSec 10
            $rates = @()
            foreach ($prop in $response.rates.PSObject.Properties) {
                $rates += [PSCustomObject]@{
                    Currency = $prop.Name
                    Rate     = [math]::Round([double]$prop.Value, 4)
                    Source   = 'Frankfurter'
                }
            }
            $result = [PSCustomObject]@{ Base = $BaseCurrency; Date = $response.date; Rates = $rates }
            $script:fxCache.Data = $result
            $script:fxCache.LastFetch = [datetime]::Now
            return $result
        }
        catch {
            Write-Verbose "Frankfurter API failed: $($_.Exception.Message)"
        }

        # Fallback: ECB XML
        try {
            $ecbResponse = Invoke-WebRequest -Uri 'https://www.ecb.europa.eu/stats/eurofxref/eurofxref-daily.xml' -UseBasicParsing -TimeoutSec 10
            [xml]$ecbXml = $ecbResponse.Content
            $ns = New-Object System.Xml.XmlNamespaceManager $ecbXml.NameTable
            $ns.AddNamespace('gesmes', 'http://www.gesmes.org/xml/2002-08-01')
            $ns.AddNamespace('ecb', 'http://www.ecb.int/vocabulary/2002-08-01/eurofxref')
            $cubeNodes = $ecbXml.SelectNodes('//ecb:Cube[@currency]', $ns)
            $eurRates = @{}
            foreach ($node in $cubeNodes) {
                $eurRates[$node.currency] = [double]$node.rate
            }
            # Convert EUR-based rates to selected base
            $rates = @()
            if ($BaseCurrency -eq 'EUR') {
                foreach ($cur in ($currencies -split ',')) {
                    if ($eurRates.ContainsKey($cur)) {
                        $rates += [PSCustomObject]@{ Currency = $cur; Rate = [math]::Round($eurRates[$cur], 4); Source = 'ECB' }
                    }
                }
            }
            elseif ($eurRates.ContainsKey($BaseCurrency)) {
                $baseRate = $eurRates[$BaseCurrency]
                foreach ($cur in ($currencies -split ',')) {
                    if ($cur -eq $BaseCurrency) { continue }
                    if ($cur -eq 'EUR') {
                        $rates += [PSCustomObject]@{ Currency = 'EUR'; Rate = [math]::Round(1 / $baseRate, 4); Source = 'ECB' }
                    }
                    elseif ($eurRates.ContainsKey($cur)) {
                        $converted = [math]::Round($eurRates[$cur] / $baseRate, 4)
                        $rates += [PSCustomObject]@{ Currency = $cur; Rate = $converted; Source = 'ECB' }
                    }
                }
            }
            $result = [PSCustomObject]@{ Base = $BaseCurrency; Date = (Get-Date).ToString('yyyy-MM-dd'); Rates = $rates }
            $script:fxCache.Data = $result
            $script:fxCache.LastFetch = [datetime]::Now
            return $result
        }
        catch {
            Write-Verbose "ECB fallback failed: $($_.Exception.Message)"
        }

        return $null
    }

    function Get-CryptoPrices {
        param([string]$VsCurrency = 'usd')
        # Check cache (5 minute TTL, invalidate if currency changed)
        if ($script:cryptoCache.Data -and $script:cryptoCache.Currency -eq $VsCurrency -and ([datetime]::Now - $script:cryptoCache.LastFetch).TotalMinutes -lt 5) {
            return $script:cryptoCache.Data
        }

        try {
            $url = "https://api.coingecko.com/api/v3/coins/markets?vs_currency=$VsCurrency&order=market_cap_desc&per_page=10&sparkline=false"
            $response = Invoke-RestMethod -Uri $url -TimeoutSec 10
            $result = @()
            foreach ($coin in $response) {
                $result += [PSCustomObject]@{
                    Symbol    = $coin.symbol.ToUpper()
                    Name      = $coin.name
                    Price     = [double]$coin.current_price
                    Change24h = [math]::Round([double]$coin.price_change_percentage_24h, 2)
                    MarketCap = [double]$coin.market_cap
                }
            }
            $script:cryptoCache.Data = $result
            $script:cryptoCache.Currency = $VsCurrency
            $script:cryptoCache.LastFetch = [datetime]::Now
            return $result
        }
        catch {
            Write-Verbose "CoinGecko API failed: $($_.Exception.Message)"
            return $null
        }
    }

    function Get-StockIndices {
        param([string]$ApiKey)
        if ([string]::IsNullOrWhiteSpace($ApiKey)) { return $null }

        # Check cache (15 minute TTL)
        if ($script:indicesCache.Data -and ([datetime]::Now - $script:indicesCache.LastFetch).TotalMinutes -lt 15) {
            return $script:indicesCache.Data
        }

        # Free tier: index symbols (SPX, DJI etc.) require paid plans.
        # Use index-tracking ETFs which are available on the free tier.
        $indices = @(
            @{ Symbol = 'SPY'; Name = 'S&P 500 (SPY)' },
            @{ Symbol = 'QQQ'; Name = 'NASDAQ 100 (QQQ)' },
            @{ Symbol = 'DIA'; Name = 'Dow Jones (DIA)' },
            @{ Symbol = 'EWU'; Name = 'FTSE 100 (EWU)' },
            @{ Symbol = 'EWG'; Name = 'DAX (EWG)' },
            @{ Symbol = 'EWJ'; Name = 'Nikkei 225 (EWJ)' },
            @{ Symbol = 'FXI'; Name = 'China Large-Cap (FXI)' },
            @{ Symbol = 'INDA'; Name = 'India (INDA)' }
        )

        # Fetch in small batches to stay within 8 credits/min free tier limit
        $result = @()
        $batches = @()
        for ($i = 0; $i -lt $indices.Count; $i += 8) {
            $batch = $indices[$i..([math]::Min($i + 7, $indices.Count - 1))]
            $batches += , @($batch)
        }

        try {
            foreach ($batch in $batches) {
                $symbolList = ($batch | ForEach-Object { $_.Symbol }) -join ','
                $url = "https://api.twelvedata.com/quote?symbol=$symbolList&apikey=$ApiKey"
                $response = Invoke-RestMethod -Uri $url -TimeoutSec 15
                foreach ($idx in $batch) {
                    # Single symbol returns data directly; multiple returns keyed object
                    $data = if ($batch.Count -eq 1) { $response } else { $response.($idx.Symbol) }
                    if ($data -and $data.close -and -not $data.code) {
                        $change = 0
                        if ($data.previous_close -and [double]$data.previous_close -ne 0) {
                            $change = [math]::Round((([double]$data.close - [double]$data.previous_close) / [double]$data.previous_close) * 100, 2)
                        }
                        $result += [PSCustomObject]@{
                            Symbol = $idx.Symbol
                            Name   = $idx.Name
                            Price  = [math]::Round([double]$data.close, 2)
                            Change = $change
                        }
                    }
                }
            }
            if ($result.Count -gt 0) {
                $script:indicesCache.Data = $result
                $script:indicesCache.LastFetch = [datetime]::Now
            }
            return $result
        }
        catch {
            Write-Verbose "Twelve Data API failed: $($_.Exception.Message)"
            return $null
        }
    }

    function Get-CommodityPrices {
        param([string]$ApiKey, [string]$BaseCurrency = 'USD')
        if ([string]::IsNullOrWhiteSpace($ApiKey)) { return $null }

        # Check cache (15 minute TTL, invalidate if currency changed)
        if ($script:commoditiesCache.Data -and $script:commoditiesCache.Currency -eq $BaseCurrency -and ([datetime]::Now - $script:commoditiesCache.LastFetch).TotalMinutes -lt 15) {
            return $script:commoditiesCache.Data
        }

        # Commodity-tracking ETFs available on free Twelve Data tier
        $commodities = @(
            @{ Symbol = 'GLD'; Name = 'Gold (GLD)'; Unit = 'oz' },
            @{ Symbol = 'SLV'; Name = 'Silver (SLV)'; Unit = 'oz' },
            @{ Symbol = 'USO'; Name = 'Crude Oil (USO)'; Unit = 'bbl' },
            @{ Symbol = 'UNG'; Name = 'Natural Gas (UNG)'; Unit = 'mmBtu' },
            @{ Symbol = 'PPLT'; Name = 'Platinum (PPLT)'; Unit = 'oz' },
            @{ Symbol = 'WEAT'; Name = 'Wheat (WEAT)'; Unit = 'bu' },
            @{ Symbol = 'DBA'; Name = 'Agriculture (DBA)'; Unit = '' }
        )

        $result = @()
        try {
            $symbolList = ($commodities | ForEach-Object { $_.Symbol }) -join ','
            $url = "https://api.twelvedata.com/quote?symbol=$symbolList&apikey=$ApiKey"
            $response = Invoke-RestMethod -Uri $url -TimeoutSec 15

            # Get FX conversion rate if not USD
            $fxRate = 1.0
            if ($BaseCurrency -ne 'USD') {
                try {
                    $fxUrl = "https://api.frankfurter.dev/v1/latest?base=USD&symbols=$BaseCurrency"
                    $fxResponse = Invoke-RestMethod -Uri $fxUrl -TimeoutSec 10
                    $fxRate = [double]$fxResponse.rates.$BaseCurrency
                }
                catch {
                    Write-Verbose "FX conversion failed, showing USD prices: $($_.Exception.Message)"
                    $fxRate = 1.0
                    $BaseCurrency = 'USD'
                }
            }

            foreach ($cmd in $commodities) {
                $data = if ($commodities.Count -eq 1) { $response } else { $response.($cmd.Symbol) }
                if ($data -and $data.close -and -not $data.code) {
                    $change = 0
                    if ($data.previous_close -and [double]$data.previous_close -ne 0) {
                        $change = [math]::Round((([double]$data.close - [double]$data.previous_close) / [double]$data.previous_close) * 100, 2)
                    }
                    $result += [PSCustomObject]@{
                        Symbol   = $cmd.Symbol
                        Name     = $cmd.Name
                        Price    = [math]::Round([double]$data.close * $fxRate, 2)
                        Change   = $change
                        Currency = $BaseCurrency
                    }
                }
            }
            if ($result.Count -gt 0) {
                $script:commoditiesCache.Data = $result
                $script:commoditiesCache.Currency = $BaseCurrency
                $script:commoditiesCache.LastFetch = [datetime]::Now
            }
            return $result
        }
        catch {
            Write-Verbose "Twelve Data commodities fetch failed: $($_.Exception.Message)"
            return $null
        }
    }

    function Get-StockQuote {
        param([string]$ApiKey, [string]$Ticker)
        if ([string]::IsNullOrWhiteSpace($ApiKey) -or [string]::IsNullOrWhiteSpace($Ticker)) { return $null }

        $Ticker = $Ticker.Trim().ToUpper()

        # Check cache (30 minute TTL per ticker)
        if ($script:stockQuoteCache.ContainsKey($Ticker)) {
            $cached = $script:stockQuoteCache[$Ticker]
            if ($cached.Data -and ([datetime]::Now - $cached.LastFetch).TotalMinutes -lt 30) {
                return $cached.Data
            }
        }

        try {
            $url = "https://www.alphavantage.co/query?function=GLOBAL_QUOTE&symbol=$Ticker&apikey=$ApiKey"
            $response = Invoke-RestMethod -Uri $url -TimeoutSec 10
            $quote = $response.'Global Quote'
            if ($quote -and $quote.'05. price') {
                $result = [PSCustomObject]@{
                    Symbol    = $quote.'01. symbol'
                    Price     = [math]::Round([double]$quote.'05. price', 2)
                    Change    = [math]::Round([double]$quote.'09. change', 2)
                    ChangePct = $quote.'10. change percent' -replace '%', ''
                    Open      = [math]::Round([double]$quote.'02. open', 2)
                    High      = [math]::Round([double]$quote.'03. high', 2)
                    Low       = [math]::Round([double]$quote.'04. low', 2)
                    Volume    = $quote.'06. volume'
                }
                $script:stockQuoteCache[$Ticker] = @{ Data = $result; LastFetch = [datetime]::Now }
                return $result
            }
            return $null
        }
        catch {
            Write-Verbose "Alpha Vantage API failed: $($_.Exception.Message)"
            return $null
        }
    }

    # ── Major Stocks Quick-Pick Data ──────────────────────────────

    $script:majorStocks = [ordered]@{
        'US'           = @(
            @{ Symbol = 'AAPL'; Name = 'Apple' }, @{ Symbol = 'MSFT'; Name = 'Microsoft' },
            @{ Symbol = 'AMZN'; Name = 'Amazon' }, @{ Symbol = 'GOOGL'; Name = 'Alphabet' },
            @{ Symbol = 'META'; Name = 'Meta' }, @{ Symbol = 'TSLA'; Name = 'Tesla' },
            @{ Symbol = 'NVDA'; Name = 'NVIDIA' }, @{ Symbol = 'JPM'; Name = 'JPMorgan' },
            @{ Symbol = 'V'; Name = 'Visa' }, @{ Symbol = 'JNJ'; Name = 'J&J' }
        )
        'UK'           = @(
            @{ Symbol = 'LLOY.L'; Name = 'Lloyds' }, @{ Symbol = 'BARC.L'; Name = 'Barclays' },
            @{ Symbol = 'BP.L'; Name = 'BP' }, @{ Symbol = 'SHEL.L'; Name = 'Shell' },
            @{ Symbol = 'HSBA.L'; Name = 'HSBC' }, @{ Symbol = 'AZN.L'; Name = 'AstraZeneca' },
            @{ Symbol = 'GSK.L'; Name = 'GSK' }, @{ Symbol = 'RIO.L'; Name = 'Rio Tinto' },
            @{ Symbol = 'VOD.L'; Name = 'Vodafone' }, @{ Symbol = 'ULVR.L'; Name = 'Unilever' }
        )
        'Europe'       = @(
            @{ Symbol = 'SAP.DE'; Name = 'SAP' }, @{ Symbol = 'SIE.DE'; Name = 'Siemens' },
            @{ Symbol = 'MC.PA'; Name = 'LVMH' }, @{ Symbol = 'OR.PA'; Name = "L'Or" + [char]0x00E9 + "al" },
            @{ Symbol = 'ASML.AS'; Name = 'ASML' }, @{ Symbol = 'NESN.SW'; Name = 'Nestl' + [char]0x00E9 }
        )
        'Asia-Pacific' = @(
            @{ Symbol = '9984.T'; Name = 'SoftBank' }, @{ Symbol = '7203.T'; Name = 'Toyota' },
            @{ Symbol = '0005.HK'; Name = 'HSBC HK' }, @{ Symbol = '0700.HK'; Name = 'Tencent' },
            @{ Symbol = 'BHP.AX'; Name = 'BHP' }, @{ Symbol = 'CBA.AX'; Name = 'CommBank' }
        )
    }
    $script:selectedStockRegion = 'US'

    # ── Stock Profile & Statistics (Twelve Data) ──────────────────

    $script:stockProfileCache = @{}
    $script:stockStatsCache = @{}

    function Get-StockProfile {
        param([string]$ApiKey, [string]$Ticker)
        if ([string]::IsNullOrWhiteSpace($ApiKey) -or [string]::IsNullOrWhiteSpace($Ticker)) { return $null }
        $Ticker = $Ticker.Trim().ToUpper()

        # Check cache (24hr TTL)
        if ($script:stockProfileCache.ContainsKey($Ticker)) {
            $cached = $script:stockProfileCache[$Ticker]
            if ($cached.Data -and ([datetime]::Now - $cached.LastFetch).TotalHours -lt 24) {
                return $cached.Data
            }
        }

        try {
            $url = "https://api.twelvedata.com/profile?symbol=$Ticker&apikey=$ApiKey"
            $response = Invoke-RestMethod -Uri $url -TimeoutSec 10
            if ($response -and $response.name -and -not $response.code) {
                $result = [PSCustomObject]@{
                    Name        = $response.name
                    Sector      = $response.sector
                    Industry    = $response.industry
                    Exchange    = $response.exchange
                    Country     = $response.country
                    Employees   = $response.employees
                    Website     = $response.website
                    CEO         = $response.CEO
                    Description = $response.description
                }
                $script:stockProfileCache[$Ticker] = @{ Data = $result; LastFetch = [datetime]::Now }
                return $result
            }
            return $null
        }
        catch {
            Write-Verbose "Twelve Data profile failed for $Ticker`: $($_.Exception.Message)"
            return $null
        }
    }

    function Get-StockStatistics {
        param([string]$ApiKey, [string]$Ticker)
        if ([string]::IsNullOrWhiteSpace($ApiKey) -or [string]::IsNullOrWhiteSpace($Ticker)) { return $null }
        $Ticker = $Ticker.Trim().ToUpper()

        # Check cache (1hr TTL)
        if ($script:stockStatsCache.ContainsKey($Ticker)) {
            $cached = $script:stockStatsCache[$Ticker]
            if ($cached.Data -and ([datetime]::Now - $cached.LastFetch).TotalHours -lt 1) {
                return $cached.Data
            }
        }

        try {
            $url = "https://api.twelvedata.com/statistics?symbol=$Ticker&apikey=$ApiKey"
            $response = Invoke-RestMethod -Uri $url -TimeoutSec 10
            if ($response -and $response.statistics -and -not $response.code) {
                $stats = $response.statistics
                $result = [PSCustomObject]@{
                    MarketCap         = if ($stats.valuations_metrics.market_capitalization) { [double]$stats.valuations_metrics.market_capitalization } else { $null }
                    TrailingPE        = if ($stats.valuations_metrics.trailing_pe) { [math]::Round([double]$stats.valuations_metrics.trailing_pe, 2) } else { $null }
                    ForwardPE         = if ($stats.valuations_metrics.forward_pe) { [math]::Round([double]$stats.valuations_metrics.forward_pe, 2) } else { $null }
                    Week52High        = if ($stats.stock_price_summary.fifty_two_week_high) { [math]::Round([double]$stats.stock_price_summary.fifty_two_week_high, 2) } else { $null }
                    Week52Low         = if ($stats.stock_price_summary.fifty_two_week_low) { [math]::Round([double]$stats.stock_price_summary.fifty_two_week_low, 2) } else { $null }
                    Beta              = if ($stats.stock_price_summary.beta) { [math]::Round([double]$stats.stock_price_summary.beta, 3) } else { $null }
                    DividendYield     = if ($stats.dividends_and_splits.forward_annual_dividend_yield) { [math]::Round([double]$stats.dividends_and_splits.forward_annual_dividend_yield * 100, 2) } else { $null }
                    DividendFreq      = $stats.dividends_and_splits.dividend_frequency
                    SharesOutstanding = if ($stats.stock_statistics.shares_outstanding) { [double]$stats.stock_statistics.shares_outstanding } else { $null }
                    ProfitMargin      = if ($stats.financials.profit_margin) { [math]::Round([double]$stats.financials.profit_margin * 100, 2) } else { $null }
                }
                $script:stockStatsCache[$Ticker] = @{ Data = $result; LastFetch = [datetime]::Now }
                return $result
            }
            return $null
        }
        catch {
            Write-Verbose "Twelve Data statistics failed for $Ticker`: $($_.Exception.Message)"
            return $null
        }
    }

    # ── Market Data Panel Rendering ───────────────────────────────

    function Format-RelativeTime {
        param([datetime]$Time)
        $diff = [datetime]::Now - $Time
        if ($diff.TotalMinutes -lt 1) { return 'just now' }
        if ($diff.TotalMinutes -lt 60) { return "$([int]$diff.TotalMinutes)m ago" }
        if ($diff.TotalHours -lt 24) { return "$([int]$diff.TotalHours)h ago" }
        return $Time.ToString('MMM dd')
    }

    function Format-LargeNumber {
        param([double]$Number)
        if ($Number -ge 1e12) { return "$([math]::Round($Number / 1e12, 2))T" }
        if ($Number -ge 1e9) { return "$([math]::Round($Number / 1e9, 2))B" }
        if ($Number -ge 1e6) { return "$([math]::Round($Number / 1e6, 2))M" }
        return $Number.ToString('N0')
    }

    function New-ThemedComboBox {
        $cmb = New-Object System.Windows.Controls.ComboBox
        $cmb.Style = $window.FindResource('DarkComboBoxStyle')
        $cmb.ItemContainerStyle = $window.FindResource('DarkComboBoxItemStyle')
        $cmb.HorizontalAlignment = 'Left'
        return $cmb
    }

    function Set-ActiveMarketTab {
        param([string]$TabName)
        $script:activeMarketTab = $TabName
        $converter = [System.Windows.Media.BrushConverter]::new()

        # Reset all tab buttons
        foreach ($btn in @($btnTabNews, $btnTabFx, $btnTabCrypto, $btnTabIndices, $btnTabCommodities, $btnTabStocks)) {
            if ($btn -and $btn.Visibility -eq 'Visible') {
                $btn.Background = $converter.ConvertFromString('#16213E')
                $btn.Foreground = $converter.ConvertFromString('#AAAAAA')
                $btn.BorderBrush = $converter.ConvertFromString('#0F3460')
            }
        }

        # Highlight active tab
        $activeBtn = switch ($TabName) {
            'News' { $btnTabNews }
            'FX' { $btnTabFx }
            'Crypto' { $btnTabCrypto }
            'Indices' { $btnTabIndices }
            'Commodities' { $btnTabCommodities }
            'Stocks' { $btnTabStocks }
        }
        if ($activeBtn) {
            $activeBtn.Background = $converter.ConvertFromString('#0F3460')
            $activeBtn.Foreground = $converter.ConvertFromString('#00CC66')
            $activeBtn.BorderBrush = $converter.ConvertFromString('#00CC66')
        }

        Update-MarketDataContent
    }

    function Update-MarketDataContent {
        $marketDataContent.Children.Clear()
        $converter = [System.Windows.Media.BrushConverter]::new()

        switch ($script:activeMarketTab) {
            'News' { Render-NewsPanel -Converter $converter }
            'FX' { Render-FxPanel -Converter $converter }
            'Crypto' { Render-CryptoPanel -Converter $converter }
            'Indices' { Render-IndicesPanel -Converter $converter }
            'Commodities' { Render-CommoditiesPanel -Converter $converter }
            'Stocks' { Render-StocksPanel -Converter $converter }
        }
    }

    function Add-MarketDataRow {
        param($Converter, [string]$Left, [string]$Right, [string]$RightColor = '#E0E0E0', [string]$LeftColor = '#AAAAAA', [double]$LeftFontSize = 11, [double]$RightFontSize = 12)
        $grid = New-Object System.Windows.Controls.Grid
        $col1 = New-Object System.Windows.Controls.ColumnDefinition; $col1.Width = [System.Windows.GridLength]::new(1, [System.Windows.GridUnitType]::Star)
        $col2 = New-Object System.Windows.Controls.ColumnDefinition; $col2.Width = [System.Windows.GridLength]::new(0, [System.Windows.GridUnitType]::Auto)
        $grid.ColumnDefinitions.Add($col1); $grid.ColumnDefinitions.Add($col2)
        $grid.Margin = [System.Windows.Thickness]::new(0, 0, 0, 4)

        $lblLeft = New-Object System.Windows.Controls.TextBlock
        $lblLeft.Text = $Left; $lblLeft.FontSize = $LeftFontSize; $lblLeft.Foreground = $Converter.ConvertFromString($LeftColor)
        $lblLeft.TextTrimming = 'CharacterEllipsis'; $lblLeft.VerticalAlignment = 'Center'
        [System.Windows.Controls.Grid]::SetColumn($lblLeft, 0)
        $grid.Children.Add($lblLeft) | Out-Null

        $lblRight = New-Object System.Windows.Controls.TextBlock
        $lblRight.Text = $Right; $lblRight.FontSize = $RightFontSize; $lblRight.Foreground = $Converter.ConvertFromString($RightColor)
        $lblRight.FontFamily = [System.Windows.Media.FontFamily]::new('Consolas'); $lblRight.VerticalAlignment = 'Center'
        $lblRight.HorizontalAlignment = 'Right'
        [System.Windows.Controls.Grid]::SetColumn($lblRight, 1)
        $grid.Children.Add($lblRight) | Out-Null

        $marketDataContent.Children.Add($grid) | Out-Null
    }

    function Add-MarketSeparator {
        param($Converter)
        $sep = New-Object System.Windows.Shapes.Rectangle
        $sep.Height = 1; $sep.Fill = $Converter.ConvertFromString('#0F3460')
        $sep.Margin = [System.Windows.Thickness]::new(0, 4, 0, 4)
        $marketDataContent.Children.Add($sep) | Out-Null
    }

    function Render-NewsPanel {
        param($Converter)
        $news = Get-FinancialNews
        if (-not $news) {
            $lbl = New-Object System.Windows.Controls.TextBlock
            $lbl.Text = 'News unavailable — check internet connection'
            $lbl.Foreground = $Converter.ConvertFromString('#888888'); $lbl.FontSize = 11; $lbl.TextWrapping = 'Wrap'
            $lbl.Margin = [System.Windows.Thickness]::new(0, 10, 0, 0)
            $marketDataContent.Children.Add($lbl) | Out-Null
            return
        }

        foreach ($item in $news) {
            $panel = New-Object System.Windows.Controls.StackPanel
            $panel.Margin = [System.Windows.Thickness]::new(0, 0, 0, 8)
            $panel.Cursor = [System.Windows.Input.Cursors]::Hand

            $title = New-Object System.Windows.Controls.TextBlock
            $title.Text = $item.Title
            $title.FontSize = 11; $title.Foreground = $Converter.ConvertFromString('#E0E0E0')
            $title.TextWrapping = 'Wrap'; $title.MaxHeight = 36
            $panel.Children.Add($title) | Out-Null

            $meta = New-Object System.Windows.Controls.TextBlock
            $timeStr = Format-RelativeTime -Time $item.Published
            $meta.Text = "$($item.Source) · $timeStr"
            $meta.FontSize = 9; $meta.Foreground = $Converter.ConvertFromString('#666666')
            $meta.Margin = [System.Windows.Thickness]::new(0, 2, 0, 0)
            $panel.Children.Add($meta) | Out-Null

            # Click to open link
            $url = $item.Link
            $panel.Tag = $url
            $panel.Add_MouseLeftButtonDown({
                    param($sender, $e)
                    try { Start-Process $sender.Tag } catch { }
                })

            $marketDataContent.Children.Add($panel) | Out-Null
        }

        # Last updated footer
        Add-MarketSeparator -Converter $Converter
        $footer = New-Object System.Windows.Controls.TextBlock
        $footer.Text = "Updated: $(Format-RelativeTime -Time $script:newsCache.LastFetch)"
        $footer.FontSize = 9; $footer.Foreground = $Converter.ConvertFromString('#555555')
        $footer.HorizontalAlignment = 'Right'
        $marketDataContent.Children.Add($footer) | Out-Null
    }

    function Render-FxPanel {
        param($Converter)

        # Base currency selector row
        $selectorGrid = New-Object System.Windows.Controls.Grid
        $sc1 = New-Object System.Windows.Controls.ColumnDefinition; $sc1.Width = [System.Windows.GridLength]::new(0, [System.Windows.GridUnitType]::Auto)
        $sc2 = New-Object System.Windows.Controls.ColumnDefinition; $sc2.Width = [System.Windows.GridLength]::new(1, [System.Windows.GridUnitType]::Star)
        $selectorGrid.ColumnDefinitions.Add($sc1); $selectorGrid.ColumnDefinitions.Add($sc2)
        $selectorGrid.Margin = [System.Windows.Thickness]::new(0, 0, 0, 8)

        $lblBase = New-Object System.Windows.Controls.TextBlock
        $lblBase.Text = 'Base Currency: '; $lblBase.FontSize = 11; $lblBase.Foreground = $Converter.ConvertFromString('#AAAAAA')
        $lblBase.VerticalAlignment = 'Center'
        [System.Windows.Controls.Grid]::SetColumn($lblBase, 0)
        $selectorGrid.Children.Add($lblBase) | Out-Null

        $cmbFxBase = New-ThemedComboBox
        $fxCurrencies = @('USD', 'EUR', 'GBP', 'JPY', 'CHF', 'CAD', 'AUD', 'CNY', 'HKD', 'SGD', 'INR', 'BRL', 'ZAR', 'MXN')
        foreach ($cur in $fxCurrencies) { $cmbFxBase.Items.Add($cur) | Out-Null }
        $cmbFxBase.SelectedItem = $script:fxBaseCurrency
        $cmbFxBase.Add_SelectionChanged({
                param($sender, $e)
                $script:fxBaseCurrency = $sender.SelectedItem.ToString()
                $script:fxCache.Data = $null  # Invalidate cache
                Update-MarketDataContent
            })
        [System.Windows.Controls.Grid]::SetColumn($cmbFxBase, 1)
        $selectorGrid.Children.Add($cmbFxBase) | Out-Null
        $marketDataContent.Children.Add($selectorGrid) | Out-Null

        $fx = Get-ForexRates -BaseCurrency $script:fxBaseCurrency
        if (-not $fx) {
            $lbl = New-Object System.Windows.Controls.TextBlock
            $lbl.Text = 'FX data unavailable — check internet connection'
            $lbl.Foreground = $Converter.ConvertFromString('#888888'); $lbl.FontSize = 11; $lbl.TextWrapping = 'Wrap'
            $lbl.Margin = [System.Windows.Thickness]::new(0, 10, 0, 0)
            $marketDataContent.Children.Add($lbl) | Out-Null
            return
        }

        # Header
        $hdr = New-Object System.Windows.Controls.TextBlock
        $hdr.Text = "Base: $($fx.Base) | Date: $($fx.Date)"
        $hdr.FontSize = 10; $hdr.Foreground = $Converter.ConvertFromString('#888888')
        $hdr.Margin = [System.Windows.Thickness]::new(0, 0, 0, 6)
        $marketDataContent.Children.Add($hdr) | Out-Null

        foreach ($rate in $fx.Rates) {
            Add-MarketDataRow -Converter $Converter -Left "$($fx.Base)/$($rate.Currency)" -Right "$($rate.Rate)"
        }

        Add-MarketSeparator -Converter $Converter
        $footer = New-Object System.Windows.Controls.TextBlock
        $footer.Text = "Source: $($fx.Rates[0].Source) | Updated: $(Format-RelativeTime -Time $script:fxCache.LastFetch)"
        $footer.FontSize = 9; $footer.Foreground = $Converter.ConvertFromString('#555555')
        $footer.TextWrapping = 'Wrap'
        $marketDataContent.Children.Add($footer) | Out-Null
    }

    function Render-CryptoPanel {
        param($Converter)

        # Currency selector row
        $selectorGrid = New-Object System.Windows.Controls.Grid
        $sc1 = New-Object System.Windows.Controls.ColumnDefinition; $sc1.Width = [System.Windows.GridLength]::new(0, [System.Windows.GridUnitType]::Auto)
        $sc2 = New-Object System.Windows.Controls.ColumnDefinition; $sc2.Width = [System.Windows.GridLength]::new(1, [System.Windows.GridUnitType]::Star)
        $selectorGrid.ColumnDefinitions.Add($sc1); $selectorGrid.ColumnDefinitions.Add($sc2)
        $selectorGrid.Margin = [System.Windows.Thickness]::new(0, 0, 0, 8)

        $lblCur = New-Object System.Windows.Controls.TextBlock
        $lblCur.Text = 'Display Currency: '; $lblCur.FontSize = 11; $lblCur.Foreground = $Converter.ConvertFromString('#AAAAAA')
        $lblCur.VerticalAlignment = 'Center'
        [System.Windows.Controls.Grid]::SetColumn($lblCur, 0)
        $selectorGrid.Children.Add($lblCur) | Out-Null

        $cmbCryptoCur = New-ThemedComboBox
        $cryptoCurrencies = @(
            @{ Display = 'USD ($)'; Value = 'usd' },
            @{ Display = 'EUR (€)'; Value = 'eur' },
            @{ Display = 'GBP (£)'; Value = 'gbp' },
            @{ Display = 'JPY (¥)'; Value = 'jpy' },
            @{ Display = 'CHF'; Value = 'chf' },
            @{ Display = 'CAD (C$)'; Value = 'cad' },
            @{ Display = 'AUD (A$)'; Value = 'aud' },
            @{ Display = 'CNY (¥)'; Value = 'cny' },
            @{ Display = 'BTC'; Value = 'btc' }
        )
        $selectedIdx = 0
        for ($i = 0; $i -lt $cryptoCurrencies.Count; $i++) {
            $cmbCryptoCur.Items.Add($cryptoCurrencies[$i].Display) | Out-Null
            if ($cryptoCurrencies[$i].Value -eq $script:cryptoCurrency) { $selectedIdx = $i }
        }
        $cmbCryptoCur.SelectedIndex = $selectedIdx
        $cmbCryptoCur.Tag = $cryptoCurrencies
        $cmbCryptoCur.Add_SelectionChanged({
                param($sender, $e)
                $currencies = $sender.Tag
                $script:cryptoCurrency = $currencies[$sender.SelectedIndex].Value
                $script:cryptoCache.Data = $null  # Invalidate cache
                Update-MarketDataContent
            })
        [System.Windows.Controls.Grid]::SetColumn($cmbCryptoCur, 1)
        $selectorGrid.Children.Add($cmbCryptoCur) | Out-Null
        $marketDataContent.Children.Add($selectorGrid) | Out-Null

        $crypto = Get-CryptoPrices -VsCurrency $script:cryptoCurrency
        if (-not $crypto) {
            $lbl = New-Object System.Windows.Controls.TextBlock
            $lbl.Text = 'Crypto data unavailable — check internet connection'
            $lbl.Foreground = $Converter.ConvertFromString('#888888'); $lbl.FontSize = 11; $lbl.TextWrapping = 'Wrap'
            $lbl.Margin = [System.Windows.Thickness]::new(0, 10, 0, 0)
            $marketDataContent.Children.Add($lbl) | Out-Null
            return
        }

        # Currency symbol for display
        $curSymbol = switch ($script:cryptoCurrency) {
            'usd' { '$' }
            'eur' { [char]0x20AC }
            'gbp' { [char]0x00A3 }
            'jpy' { [char]0x00A5 }
            'cny' { [char]0x00A5 }
            'cad' { 'C$' }
            'aud' { 'A$' }
            'btc' { [char]0x20BF }
            default { '' }
        }

        foreach ($coin in $crypto) {
            $changeColor = if ($coin.Change24h -ge 0) { '#00CC66' } else { '#FF4444' }
            $arrow = if ($coin.Change24h -ge 0) { [char]0x25B2 } else { [char]0x25BC }
            $priceStr = if ($coin.Price -ge 1) { "$curSymbol$($coin.Price.ToString('N2'))" } else { "$curSymbol$($coin.Price.ToString('N6'))" }

            $grid = New-Object System.Windows.Controls.Grid
            $c1 = New-Object System.Windows.Controls.ColumnDefinition; $c1.Width = [System.Windows.GridLength]::new(50, [System.Windows.GridUnitType]::Pixel)
            $c2 = New-Object System.Windows.Controls.ColumnDefinition; $c2.Width = [System.Windows.GridLength]::new(1, [System.Windows.GridUnitType]::Star)
            $c3 = New-Object System.Windows.Controls.ColumnDefinition; $c3.Width = [System.Windows.GridLength]::new(0, [System.Windows.GridUnitType]::Auto)
            $grid.ColumnDefinitions.Add($c1); $grid.ColumnDefinitions.Add($c2); $grid.ColumnDefinitions.Add($c3)
            $grid.Margin = [System.Windows.Thickness]::new(0, 0, 0, 4)

            $lblSym = New-Object System.Windows.Controls.TextBlock
            $lblSym.Text = $coin.Symbol; $lblSym.FontSize = 11; $lblSym.FontWeight = [System.Windows.FontWeights]::Bold
            $lblSym.Foreground = $Converter.ConvertFromString('#E0E0E0'); $lblSym.VerticalAlignment = 'Center'
            [System.Windows.Controls.Grid]::SetColumn($lblSym, 0)
            $grid.Children.Add($lblSym) | Out-Null

            $lblPrice = New-Object System.Windows.Controls.TextBlock
            $lblPrice.Text = $priceStr; $lblPrice.FontSize = 11
            $lblPrice.Foreground = $Converter.ConvertFromString('#E0E0E0')
            $lblPrice.FontFamily = [System.Windows.Media.FontFamily]::new('Consolas'); $lblPrice.HorizontalAlignment = 'Right'
            $lblPrice.VerticalAlignment = 'Center'
            [System.Windows.Controls.Grid]::SetColumn($lblPrice, 1)
            $grid.Children.Add($lblPrice) | Out-Null

            $lblChange = New-Object System.Windows.Controls.TextBlock
            $lblChange.Text = " $arrow $($coin.Change24h)%"; $lblChange.FontSize = 10
            $lblChange.Foreground = $Converter.ConvertFromString($changeColor)
            $lblChange.FontFamily = [System.Windows.Media.FontFamily]::new('Consolas')
            $lblChange.VerticalAlignment = 'Center'; $lblChange.Margin = [System.Windows.Thickness]::new(6, 0, 0, 0)
            [System.Windows.Controls.Grid]::SetColumn($lblChange, 2)
            $grid.Children.Add($lblChange) | Out-Null

            $marketDataContent.Children.Add($grid) | Out-Null
        }

        Add-MarketSeparator -Converter $Converter
        $footer = New-Object System.Windows.Controls.TextBlock
        $footer.Text = "Source: CoinGecko | Updated: $(Format-RelativeTime -Time $script:cryptoCache.LastFetch)"
        $footer.FontSize = 9; $footer.Foreground = $Converter.ConvertFromString('#555555')
        $footer.TextWrapping = 'Wrap'
        $marketDataContent.Children.Add($footer) | Out-Null
    }

    function Render-IndicesPanel {
        param($Converter)
        $apiKey = $txtTwelveDataKey.Text.Trim()
        if ([string]::IsNullOrWhiteSpace($apiKey)) {
            $lbl = New-Object System.Windows.Controls.TextBlock
            $lbl.Text = 'Enter a Twelve Data API key in Settings to view index data.'
            $lbl.Foreground = $Converter.ConvertFromString('#888888'); $lbl.FontSize = 11; $lbl.TextWrapping = 'Wrap'
            $lbl.Margin = [System.Windows.Thickness]::new(0, 10, 0, 0)
            $marketDataContent.Children.Add($lbl) | Out-Null
            return
        }

        $indices = Get-StockIndices -ApiKey $apiKey
        if (-not $indices -or $indices.Count -eq 0) {
            $lbl = New-Object System.Windows.Controls.TextBlock
            $lbl.Text = 'Index data unavailable — may be rate-limited (8 req/min free tier). Try again in a minute.'
            $lbl.Foreground = $Converter.ConvertFromString('#888888'); $lbl.FontSize = 11; $lbl.TextWrapping = 'Wrap'
            $lbl.Margin = [System.Windows.Thickness]::new(0, 10, 0, 0)
            $marketDataContent.Children.Add($lbl) | Out-Null
            return
        }

        foreach ($idx in $indices) {
            $changeColor = if ($idx.Change -ge 0) { '#00CC66' } else { '#FF4444' }
            $arrow = if ($idx.Change -ge 0) { [char]0x25B2 } else { [char]0x25BC }

            $grid = New-Object System.Windows.Controls.Grid
            $c1 = New-Object System.Windows.Controls.ColumnDefinition; $c1.Width = [System.Windows.GridLength]::new(1, [System.Windows.GridUnitType]::Star)
            $c2 = New-Object System.Windows.Controls.ColumnDefinition; $c2.Width = [System.Windows.GridLength]::new(0, [System.Windows.GridUnitType]::Auto)
            $c3 = New-Object System.Windows.Controls.ColumnDefinition; $c3.Width = [System.Windows.GridLength]::new(0, [System.Windows.GridUnitType]::Auto)
            $grid.ColumnDefinitions.Add($c1); $grid.ColumnDefinitions.Add($c2); $grid.ColumnDefinitions.Add($c3)
            $grid.Margin = [System.Windows.Thickness]::new(0, 0, 0, 4)

            $lblName = New-Object System.Windows.Controls.TextBlock
            $lblName.Text = $idx.Name; $lblName.FontSize = 11
            $lblName.Foreground = $Converter.ConvertFromString('#E0E0E0'); $lblName.VerticalAlignment = 'Center'
            [System.Windows.Controls.Grid]::SetColumn($lblName, 0)
            $grid.Children.Add($lblName) | Out-Null

            $lblPrice = New-Object System.Windows.Controls.TextBlock
            $lblPrice.Text = $idx.Price.ToString('N2'); $lblPrice.FontSize = 11
            $lblPrice.Foreground = $Converter.ConvertFromString('#E0E0E0')
            $lblPrice.FontFamily = [System.Windows.Media.FontFamily]::new('Consolas'); $lblPrice.HorizontalAlignment = 'Right'
            $lblPrice.VerticalAlignment = 'Center'
            [System.Windows.Controls.Grid]::SetColumn($lblPrice, 1)
            $grid.Children.Add($lblPrice) | Out-Null

            $lblChange = New-Object System.Windows.Controls.TextBlock
            $lblChange.Text = " $arrow $($idx.Change)%"; $lblChange.FontSize = 10
            $lblChange.Foreground = $Converter.ConvertFromString($changeColor)
            $lblChange.FontFamily = [System.Windows.Media.FontFamily]::new('Consolas')
            $lblChange.VerticalAlignment = 'Center'; $lblChange.Margin = [System.Windows.Thickness]::new(6, 0, 0, 0)
            [System.Windows.Controls.Grid]::SetColumn($lblChange, 2)
            $grid.Children.Add($lblChange) | Out-Null

            $marketDataContent.Children.Add($grid) | Out-Null
        }

        Add-MarketSeparator -Converter $Converter
        $footer = New-Object System.Windows.Controls.TextBlock
        $footer.Text = "Source: Twelve Data (ETF proxies) | Updated: $(Format-RelativeTime -Time $script:indicesCache.LastFetch)"
        $footer.FontSize = 9; $footer.Foreground = $Converter.ConvertFromString('#555555')
        $footer.TextWrapping = 'Wrap'
        $marketDataContent.Children.Add($footer) | Out-Null
    }

    function Render-CommoditiesPanel {
        param($Converter)
        $apiKey = $txtTwelveDataKey.Text.Trim()
        if ([string]::IsNullOrWhiteSpace($apiKey)) {
            $lbl = New-Object System.Windows.Controls.TextBlock
            $lbl.Text = 'Enter a Twelve Data API key in Settings to view commodity data.'
            $lbl.Foreground = $Converter.ConvertFromString('#888888'); $lbl.FontSize = 11; $lbl.TextWrapping = 'Wrap'
            $lbl.Margin = [System.Windows.Thickness]::new(0, 10, 0, 0)
            $marketDataContent.Children.Add($lbl) | Out-Null
            return
        }

        # Currency selector row
        $selectorGrid = New-Object System.Windows.Controls.Grid
        $sc1 = New-Object System.Windows.Controls.ColumnDefinition; $sc1.Width = [System.Windows.GridLength]::new(0, [System.Windows.GridUnitType]::Auto)
        $sc2 = New-Object System.Windows.Controls.ColumnDefinition; $sc2.Width = [System.Windows.GridLength]::new(1, [System.Windows.GridUnitType]::Star)
        $selectorGrid.ColumnDefinitions.Add($sc1); $selectorGrid.ColumnDefinitions.Add($sc2)
        $selectorGrid.Margin = [System.Windows.Thickness]::new(0, 0, 0, 8)

        $lblCur = New-Object System.Windows.Controls.TextBlock
        $lblCur.Text = 'Display Currency: '; $lblCur.FontSize = 11; $lblCur.Foreground = $Converter.ConvertFromString('#AAAAAA')
        $lblCur.VerticalAlignment = 'Center'
        [System.Windows.Controls.Grid]::SetColumn($lblCur, 0)
        $selectorGrid.Children.Add($lblCur) | Out-Null

        $cmbCommodityCur = New-ThemedComboBox
        $commodityCurrencies = @(
            @{ Display = 'USD ($)'; Value = 'USD'; Symbol = '$' },
            @{ Display = 'EUR (€)'; Value = 'EUR'; Symbol = [string][char]0x20AC },
            @{ Display = 'GBP (£)'; Value = 'GBP'; Symbol = [string][char]0x00A3 },
            @{ Display = 'JPY (¥)'; Value = 'JPY'; Symbol = [string][char]0x00A5 },
            @{ Display = 'CHF'; Value = 'CHF'; Symbol = 'CHF ' },
            @{ Display = 'CAD (C$)'; Value = 'CAD'; Symbol = 'C$' },
            @{ Display = 'AUD (A$)'; Value = 'AUD'; Symbol = 'A$' },
            @{ Display = 'CNY (¥)'; Value = 'CNY'; Symbol = [string][char]0x00A5 }
        )
        $selectedIdx = 0
        for ($i = 0; $i -lt $commodityCurrencies.Count; $i++) {
            $cmbCommodityCur.Items.Add($commodityCurrencies[$i].Display) | Out-Null
            if ($commodityCurrencies[$i].Value -eq $script:commodityBaseCurrency) { $selectedIdx = $i }
        }
        $cmbCommodityCur.SelectedIndex = $selectedIdx
        $cmbCommodityCur.Tag = $commodityCurrencies
        $cmbCommodityCur.Add_SelectionChanged({
                param($sender, $e)
                $currencies = $sender.Tag
                $script:commodityBaseCurrency = $currencies[$sender.SelectedIndex].Value
                $script:commoditiesCache.Data = $null  # Invalidate cache
                Update-MarketDataContent
            })
        [System.Windows.Controls.Grid]::SetColumn($cmbCommodityCur, 1)
        $selectorGrid.Children.Add($cmbCommodityCur) | Out-Null
        $marketDataContent.Children.Add($selectorGrid) | Out-Null

        $commodities = Get-CommodityPrices -ApiKey $apiKey -BaseCurrency $script:commodityBaseCurrency
        if (-not $commodities -or $commodities.Count -eq 0) {
            $lbl = New-Object System.Windows.Controls.TextBlock
            $lbl.Text = 'Commodity data unavailable — check API key or try again shortly (rate limit: 8 req/min)'
            $lbl.Foreground = $Converter.ConvertFromString('#888888'); $lbl.FontSize = 11; $lbl.TextWrapping = 'Wrap'
            $lbl.Margin = [System.Windows.Thickness]::new(0, 10, 0, 0)
            $marketDataContent.Children.Add($lbl) | Out-Null
            return
        }

        # Resolve currency symbol for display
        $curSymbol = ($commodityCurrencies | Where-Object { $_.Value -eq $script:commodityBaseCurrency } | Select-Object -First 1).Symbol
        if (-not $curSymbol) { $curSymbol = '$' }

        foreach ($cmd in $commodities) {
            $changeColor = if ($cmd.Change -ge 0) { '#00CC66' } else { '#FF4444' }
            $arrow = if ($cmd.Change -ge 0) { [char]0x25B2 } else { [char]0x25BC }

            $grid = New-Object System.Windows.Controls.Grid
            $c1 = New-Object System.Windows.Controls.ColumnDefinition; $c1.Width = [System.Windows.GridLength]::new(1, [System.Windows.GridUnitType]::Star)
            $c2 = New-Object System.Windows.Controls.ColumnDefinition; $c2.Width = [System.Windows.GridLength]::new(0, [System.Windows.GridUnitType]::Auto)
            $c3 = New-Object System.Windows.Controls.ColumnDefinition; $c3.Width = [System.Windows.GridLength]::new(0, [System.Windows.GridUnitType]::Auto)
            $grid.ColumnDefinitions.Add($c1); $grid.ColumnDefinitions.Add($c2); $grid.ColumnDefinitions.Add($c3)
            $grid.Margin = [System.Windows.Thickness]::new(0, 0, 0, 4)

            $lblName = New-Object System.Windows.Controls.TextBlock
            $lblName.Text = $cmd.Name; $lblName.FontSize = 11
            $lblName.Foreground = $Converter.ConvertFromString('#E0E0E0'); $lblName.VerticalAlignment = 'Center'
            [System.Windows.Controls.Grid]::SetColumn($lblName, 0)
            $grid.Children.Add($lblName) | Out-Null

            $lblPrice = New-Object System.Windows.Controls.TextBlock
            $lblPrice.Text = "$curSymbol$($cmd.Price.ToString('N2'))"; $lblPrice.FontSize = 11
            $lblPrice.Foreground = $Converter.ConvertFromString('#E0E0E0')
            $lblPrice.FontFamily = [System.Windows.Media.FontFamily]::new('Consolas'); $lblPrice.HorizontalAlignment = 'Right'
            $lblPrice.VerticalAlignment = 'Center'
            [System.Windows.Controls.Grid]::SetColumn($lblPrice, 1)
            $grid.Children.Add($lblPrice) | Out-Null

            $lblChange = New-Object System.Windows.Controls.TextBlock
            $lblChange.Text = " $arrow $($cmd.Change)%"; $lblChange.FontSize = 10
            $lblChange.Foreground = $Converter.ConvertFromString($changeColor)
            $lblChange.FontFamily = [System.Windows.Media.FontFamily]::new('Consolas')
            $lblChange.VerticalAlignment = 'Center'; $lblChange.Margin = [System.Windows.Thickness]::new(6, 0, 0, 0)
            [System.Windows.Controls.Grid]::SetColumn($lblChange, 2)
            $grid.Children.Add($lblChange) | Out-Null

            $marketDataContent.Children.Add($grid) | Out-Null
        }

        Add-MarketSeparator -Converter $Converter
        $note = New-Object System.Windows.Controls.TextBlock
        $note.Text = "Prices shown as ETF proxies converted to $($script:commodityBaseCurrency) via Frankfurter API."
        $note.FontSize = 9; $note.Foreground = $Converter.ConvertFromString('#666666')
        $note.TextWrapping = 'Wrap'; $note.Margin = [System.Windows.Thickness]::new(0, 0, 0, 4)
        $marketDataContent.Children.Add($note) | Out-Null

        $footer = New-Object System.Windows.Controls.TextBlock
        $footer.Text = "Source: Twelve Data | Updated: $(Format-RelativeTime -Time $script:commoditiesCache.LastFetch)"
        $footer.FontSize = 9; $footer.Foreground = $Converter.ConvertFromString('#555555')
        $footer.TextWrapping = 'Wrap'
        $marketDataContent.Children.Add($footer) | Out-Null
    }

    function Render-StocksPanel {
        param($Converter)
        $avKey = $txtAlphaVantageKey.Text.Trim()
        $tdKey = $txtTwelveDataKey.Text.Trim()
        if ([string]::IsNullOrWhiteSpace($avKey)) {
            $lbl = New-Object System.Windows.Controls.TextBlock
            $lbl.Text = 'Enter an Alpha Vantage API key in Settings to look up stock quotes.'
            $lbl.Foreground = $Converter.ConvertFromString('#888888'); $lbl.FontSize = 11; $lbl.TextWrapping = 'Wrap'
            $lbl.Margin = [System.Windows.Thickness]::new(0, 10, 0, 0)
            $marketDataContent.Children.Add($lbl) | Out-Null
            return
        }

        # Search box
        $searchPanel = New-Object System.Windows.Controls.StackPanel
        $searchPanel.Orientation = 'Horizontal'; $searchPanel.Margin = [System.Windows.Thickness]::new(0, 0, 0, 8)

        $txtTicker = New-Object System.Windows.Controls.TextBox
        $txtTicker.Width = 120; $txtTicker.Background = $Converter.ConvertFromString('#16213E')
        $txtTicker.Foreground = $Converter.ConvertFromString('#E0E0E0'); $txtTicker.BorderBrush = $Converter.ConvertFromString('#0F3460')
        $txtTicker.Padding = [System.Windows.Thickness]::new(5, 3, 5, 3); $txtTicker.FontSize = 12
        $txtTicker.FontFamily = [System.Windows.Media.FontFamily]::new('Consolas')
        $txtTicker.ToolTip = 'Enter ticker symbol (e.g. AAPL, MSFT)'
        $searchPanel.Children.Add($txtTicker) | Out-Null

        $btnSearch = New-Object System.Windows.Controls.Button
        $btnSearch.Content = ' Look Up '; $btnSearch.Background = $Converter.ConvertFromString('#0F3460')
        $btnSearch.Foreground = $Converter.ConvertFromString('#E0E0E0'); $btnSearch.BorderBrush = $Converter.ConvertFromString('#00CC66')
        $btnSearch.Padding = [System.Windows.Thickness]::new(8, 3, 8, 3); $btnSearch.Margin = [System.Windows.Thickness]::new(6, 0, 0, 0)
        $btnSearch.FontSize = 11; $btnSearch.Cursor = [System.Windows.Input.Cursors]::Hand
        $searchPanel.Children.Add($btnSearch) | Out-Null

        $marketDataContent.Children.Add($searchPanel) | Out-Null

        # ── Quick-Pick Region Selector ──
        $quickPickHeader = New-Object System.Windows.Controls.TextBlock
        $quickPickHeader.Text = 'Quick Pick — Major Stocks'
        $quickPickHeader.Foreground = $Converter.ConvertFromString('#AAAAAA'); $quickPickHeader.FontSize = 10
        $quickPickHeader.Margin = [System.Windows.Thickness]::new(0, 2, 0, 4)
        $marketDataContent.Children.Add($quickPickHeader) | Out-Null

        $regionRow = New-Object System.Windows.Controls.StackPanel
        $regionRow.Orientation = 'Horizontal'; $regionRow.Margin = [System.Windows.Thickness]::new(0, 0, 0, 4)

        $stockButtonsPanel = New-Object System.Windows.Controls.WrapPanel
        $stockButtonsPanel.Margin = [System.Windows.Thickness]::new(0, 0, 0, 8)

        # Result area (created early so event handlers can reference it)
        $resultPanel = New-Object System.Windows.Controls.StackPanel
        $resultPanel.Name = 'stockResultPanel'

        # Store references in script scope so nested Click handlers can access them
        $script:stockTxtTicker = $txtTicker
        $script:stockResultPanel = $resultPanel
        $script:stockConverter = $Converter
        $script:stockButtonsPanel = $stockButtonsPanel

        # Helper: populate stock buttons for a region
        $script:stockPopulateButtons = $null
        $populateStockButtons = {
            param([string]$Region)
            $script:stockButtonsPanel.Children.Clear()
            $stocks = $script:majorStocks[$Region]
            foreach ($s in $stocks) {
                $sym = $s.Symbol
                $nam = $s.Name
                $btn = New-Object System.Windows.Controls.Button
                $btn.Content = $sym; $btn.ToolTip = $nam
                $btn.Background = $script:stockConverter.ConvertFromString('#16213E')
                $btn.Foreground = $script:stockConverter.ConvertFromString('#E0E0E0')
                $btn.BorderBrush = $script:stockConverter.ConvertFromString('#0F3460')
                $btn.Padding = [System.Windows.Thickness]::new(6, 2, 6, 2)
                $btn.Margin = [System.Windows.Thickness]::new(0, 0, 4, 4)
                $btn.FontSize = 10; $btn.FontFamily = [System.Windows.Media.FontFamily]::new('Consolas')
                $btn.Cursor = [System.Windows.Input.Cursors]::Hand
                $btn.Tag = $sym
                $btn.Add_Click({
                        $ticker = $this.Tag
                        $script:stockTxtTicker.Text = $ticker
                        $script:stockResultPanel.Children.Clear()

                        $loading = New-Object System.Windows.Controls.TextBlock
                        $loading.Text = "Looking up $ticker..."
                        $loading.Foreground = $script:stockConverter.ConvertFromString('#888888'); $loading.FontSize = 11
                        $script:stockResultPanel.Children.Add($loading) | Out-Null

                        $window.Dispatcher.BeginInvoke([System.Windows.Threading.DispatcherPriority]::Background, [Action] {
                                $t = $script:stockTxtTicker.Text.Trim().ToUpper()
                                $ak = $txtAlphaVantageKey.Text.Trim()
                                $dk = $txtTwelveDataKey.Text.Trim()
                                $quote = Get-StockQuote -ApiKey $ak -Ticker $t
                                $profile = if (-not [string]::IsNullOrWhiteSpace($dk)) { Get-StockProfile -ApiKey $dk -Ticker $t } else { $null }
                                $stats = if (-not [string]::IsNullOrWhiteSpace($dk)) { Get-StockStatistics -ApiKey $dk -Ticker $t } else { $null }
                                $script:stockResultPanel.Children.Clear()
                                if ($quote) {
                                    Render-SingleStockQuote -Quote $quote -Profile $profile -Statistics $stats -Panel $script:stockResultPanel -Converter $script:stockConverter
                                }
                                else {
                                    $err = New-Object System.Windows.Controls.TextBlock
                                    $err.Text = "No data for '$t' - check symbol or API limit (25/day)"
                                    $err.Foreground = $script:stockConverter.ConvertFromString('#FF4444'); $err.FontSize = 11; $err.TextWrapping = 'Wrap'
                                    $script:stockResultPanel.Children.Add($err) | Out-Null
                                }
                            })
                    })
                $script:stockButtonsPanel.Children.Add($btn) | Out-Null
            }
        }
        $script:stockPopulateButtons = $populateStockButtons

        # Region tab buttons
        $script:stockRegionRow = $null
        foreach ($region in $script:majorStocks.Keys) {
            $thisRegion = $region  # local copy for closure capture
            $rbtn = New-Object System.Windows.Controls.Button
            $rbtn.Content = " $thisRegion "; $rbtn.Tag = $thisRegion
            $isActive = ($thisRegion -eq $script:selectedStockRegion)
            $rbtn.Background = $Converter.ConvertFromString($(if ($isActive) { '#0F3460' } else { '#16213E' }))
            $rbtn.Foreground = $Converter.ConvertFromString($(if ($isActive) { '#00CC66' } else { '#AAAAAA' }))
            $rbtn.BorderBrush = $Converter.ConvertFromString($(if ($isActive) { '#00CC66' } else { '#0F3460' }))
            $rbtn.Padding = [System.Windows.Thickness]::new(6, 2, 6, 2)
            $rbtn.Margin = [System.Windows.Thickness]::new(0, 0, 4, 0)
            $rbtn.FontSize = 10; $rbtn.Cursor = [System.Windows.Input.Cursors]::Hand
            $rbtn.Add_Click({
                    $selectedRegion = $this.Tag
                    $script:selectedStockRegion = $selectedRegion
                    foreach ($child in $script:stockRegionRow.Children) {
                        $isActive = ($child.Tag -eq $selectedRegion)
                        $child.Background = $script:stockConverter.ConvertFromString($(if ($isActive) { '#0F3460' } else { '#16213E' }))
                        $child.Foreground = $script:stockConverter.ConvertFromString($(if ($isActive) { '#00CC66' } else { '#AAAAAA' }))
                        $child.BorderBrush = $script:stockConverter.ConvertFromString($(if ($isActive) { '#00CC66' } else { '#0F3460' }))
                    }
                    & $script:stockPopulateButtons $selectedRegion
                })
            $regionRow.Children.Add($rbtn) | Out-Null
        }
        $script:stockRegionRow = $regionRow

        $marketDataContent.Children.Add($regionRow) | Out-Null
        & $populateStockButtons $script:selectedStockRegion
        $marketDataContent.Children.Add($stockButtonsPanel) | Out-Null

        # Add result panel to layout
        $marketDataContent.Children.Add($resultPanel) | Out-Null

        # Show cached quotes
        if ($script:stockQuoteCache.Count -gt 0) {
            foreach ($key in $script:stockQuoteCache.Keys) {
                $cached = $script:stockQuoteCache[$key]
                if ($cached.Data) { Render-SingleStockQuote -Quote $cached.Data -Panel $resultPanel -Converter $Converter }
            }
        }

        # Search button click
        $btnSearch.Add_Click({
                $ticker = $script:stockTxtTicker.Text.Trim().ToUpper()
                if ([string]::IsNullOrWhiteSpace($ticker)) { return }
                $script:stockResultPanel.Children.Clear()

                $loading = New-Object System.Windows.Controls.TextBlock
                $loading.Text = "Looking up $ticker..."; $loading.Foreground = $script:stockConverter.ConvertFromString('#888888'); $loading.FontSize = 11
                $script:stockResultPanel.Children.Add($loading) | Out-Null

                $window.Dispatcher.BeginInvoke([System.Windows.Threading.DispatcherPriority]::Background, [Action] {
                        $t = $script:stockTxtTicker.Text.Trim().ToUpper()
                        $ak = $txtAlphaVantageKey.Text.Trim()
                        $dk = $txtTwelveDataKey.Text.Trim()
                        $quote = Get-StockQuote -ApiKey $ak -Ticker $t
                        $profile = if (-not [string]::IsNullOrWhiteSpace($dk)) { Get-StockProfile -ApiKey $dk -Ticker $t } else { $null }
                        $stats = if (-not [string]::IsNullOrWhiteSpace($dk)) { Get-StockStatistics -ApiKey $dk -Ticker $t } else { $null }
                        $script:stockResultPanel.Children.Clear()
                        if ($quote) {
                            Render-SingleStockQuote -Quote $quote -Profile $profile -Statistics $stats -Panel $script:stockResultPanel -Converter $script:stockConverter
                        }
                        else {
                            $err = New-Object System.Windows.Controls.TextBlock
                            $err.Text = "No data for '$t' - check symbol or API limit (25/day)"
                            $err.Foreground = $script:stockConverter.ConvertFromString('#FF4444'); $err.FontSize = 11; $err.TextWrapping = 'Wrap'
                            $script:stockResultPanel.Children.Add($err) | Out-Null
                        }
                    })
            })

        # Enter key in ticker textbox triggers search
        $script:stockTxtTicker.Add_KeyDown({
                if ($_.Key -eq [System.Windows.Input.Key]::Return) {
                    $btnSearch.RaiseEvent([System.Windows.RoutedEventArgs]::new([System.Windows.Controls.Button]::ClickEvent))
                }
            })
    }

    function Render-SingleStockQuote {
        param($Quote, $Profile, $Statistics, $Panel, $Converter)
        $changeColor = if ([double]$Quote.Change -ge 0) { '#00CC66' } else { '#FF4444' }
        $arrow = if ([double]$Quote.Change -ge 0) { [char]0x25B2 } else { [char]0x25BC }

        $border = New-Object System.Windows.Controls.Border
        $border.Background = $Converter.ConvertFromString('#16213E')
        $border.CornerRadius = [System.Windows.CornerRadius]::new(4)
        $border.Padding = [System.Windows.Thickness]::new(10, 8, 10, 8)
        $border.Margin = [System.Windows.Thickness]::new(0, 0, 0, 6)

        $stack = New-Object System.Windows.Controls.StackPanel

        # ── Header: Company name + symbol + price + change ──
        $headerText = "$($Quote.Symbol)  `$$($Quote.Price)  $arrow $($Quote.Change) ($($Quote.ChangePct)%)"
        if ($Profile -and $Profile.Name) {
            $headerText = "$($Profile.Name) ($($Quote.Symbol))  `$$($Quote.Price)  $arrow $($Quote.Change) ($($Quote.ChangePct)%)"
        }
        $header = New-Object System.Windows.Controls.TextBlock
        $header.Text = $headerText
        $header.FontSize = 12; $header.FontWeight = [System.Windows.FontWeights]::Bold
        $header.Foreground = $Converter.ConvertFromString($changeColor)
        $header.FontFamily = [System.Windows.Media.FontFamily]::new('Consolas')
        $header.TextWrapping = 'Wrap'
        $stack.Children.Add($header) | Out-Null

        # ── Profile line: sector, industry, exchange ──
        if ($Profile) {
            $profileParts = @()
            if ($Profile.Sector) { $profileParts += $Profile.Sector }
            if ($Profile.Industry) { $profileParts += $Profile.Industry }
            if ($Profile.Exchange) { $profileParts += $Profile.Exchange }
            if ($profileParts.Count -gt 0) {
                $profLine = New-Object System.Windows.Controls.TextBlock
                $profLine.Text = ($profileParts -join '  |  ')
                $profLine.FontSize = 10; $profLine.Foreground = $Converter.ConvertFromString('#7799BB')
                $profLine.FontFamily = [System.Windows.Media.FontFamily]::new('Consolas')
                $profLine.Margin = [System.Windows.Thickness]::new(0, 2, 0, 0)
                $profLine.TextWrapping = 'Wrap'
                $stack.Children.Add($profLine) | Out-Null
            }
        }

        # ── OHLCV line ──
        $details = New-Object System.Windows.Controls.TextBlock
        $details.Text = "O: $($Quote.Open)  H: $($Quote.High)  L: $($Quote.Low)  Vol: $($Quote.Volume)"
        $details.FontSize = 10; $details.Foreground = $Converter.ConvertFromString('#888888')
        $details.FontFamily = [System.Windows.Media.FontFamily]::new('Consolas')
        $details.Margin = [System.Windows.Thickness]::new(0, 3, 0, 0)
        $stack.Children.Add($details) | Out-Null

        # ── Statistics line: Market Cap, P/E, 52-week range, Beta, Dividend ──
        if ($Statistics) {
            $statParts = @()
            if ($null -ne $Statistics.MarketCap) { $statParts += "MCap: $(Format-LargeNumber $Statistics.MarketCap)" }
            if ($null -ne $Statistics.TrailingPE) { $statParts += "P/E: $($Statistics.TrailingPE)" }
            if ($null -ne $Statistics.Week52High -and $null -ne $Statistics.Week52Low) {
                $statParts += "52w: $($Statistics.Week52Low)-$($Statistics.Week52High)"
            }
            if ($null -ne $Statistics.Beta) { $statParts += [string][char]0x03B2 + ": $($Statistics.Beta)" }
            if ($null -ne $Statistics.DividendYield) { $statParts += "Div: $($Statistics.DividendYield)%" }
            if ($null -ne $Statistics.ProfitMargin) { $statParts += "Margin: $($Statistics.ProfitMargin)%" }

            if ($statParts.Count -gt 0) {
                $statLine = New-Object System.Windows.Controls.TextBlock
                $statLine.Text = ($statParts -join '  |  ')
                $statLine.FontSize = 10; $statLine.Foreground = $Converter.ConvertFromString('#AAAAAA')
                $statLine.FontFamily = [System.Windows.Media.FontFamily]::new('Consolas')
                $statLine.Margin = [System.Windows.Thickness]::new(0, 2, 0, 0)
                $statLine.TextWrapping = 'Wrap'
                $stack.Children.Add($statLine) | Out-Null
            }
        }

        $border.Child = $stack
        $Panel.Children.Add($border) | Out-Null
    }

    # ── Data Model ────────────────────────────────────────────────

    # Create observable wrapper objects
    $script:exchangeRows = [System.Collections.ObjectModel.ObservableCollection[object]]::new()
    $script:notificationsFired = @{}

    foreach ($ex in $allExchanges) {
        $isActive = $ex.IsDefault
        # Apply saved preferences if available
        if ($script:userPrefs -and $script:userPrefs.ActiveExchanges) {
            $isActive = $script:userPrefs.ActiveExchanges -contains $ex.Code
        }
        $row = [PSCustomObject]@{
            DisplayName      = $ex.Name
            Code             = $ex.Code
            Symbol           = $ex.Symbol
            Country          = $ex.Country
            City             = $ex.City
            TimeZoneId       = $ex.TimeZoneId
            OpenTime         = $ex.OpenTimeLocal
            CloseTime        = $ex.CloseTimeLocal
            LunchBreakStart  = $ex.LunchBreakStart
            LunchBreakEnd    = $ex.LunchBreakEnd
            Latitude         = $ex.Latitude
            Longitude        = $ex.Longitude
            IsActive         = $isActive
            IsDefault        = $ex.IsDefault
            StatusText       = ''
            CountdownText    = ''
            CurrentLocalTime = ''
            StatusColor      = '#666666'
        }
        $script:exchangeRows.Add($row)
    }

    $gridExchanges.ItemsSource = $script:exchangeRows

    # ── World Clock Data ──────────────────────────────────────────

    # Base world clock zones (always shown)
    $script:baseClockZones = @(
        @{ Label = 'New York'; TzId = 'Eastern Standard Time'; Abbrev = 'ET' },
        @{ Label = 'London'; TzId = 'GMT Standard Time'; Abbrev = 'GMT' },
        @{ Label = 'Frankfurt'; TzId = 'W. Europe Standard Time'; Abbrev = 'CET' },
        @{ Label = 'Tokyo'; TzId = 'Tokyo Standard Time'; Abbrev = 'JST' },
        @{ Label = 'Hong Kong'; TzId = 'China Standard Time'; Abbrev = 'HKT' },
        @{ Label = 'Sydney'; TzId = 'AUS Eastern Standard Time'; Abbrev = 'AEST' }
    )

    $script:clockLabels = @{}

    function Rebuild-WorldClocks {
        # Keep the header TextBlock (first child)
        while ($panelWorldClocks.Children.Count -gt 1) {
            $panelWorldClocks.Children.RemoveAt($panelWorldClocks.Children.Count - 1)
        }
        $script:clockLabels = @{}

        $converter = [System.Windows.Media.BrushConverter]::new()
        $coveredTzIds = @{}

        # Helper: add a clock entry
        function Add-ClockEntry {
            param([string]$Label, [string]$TzId, [string]$Abbrev, [bool]$IsDynamic = $false)

            $panel = New-Object System.Windows.Controls.StackPanel
            $panel.Margin = [System.Windows.Thickness]::new(0, 0, 0, 8)

            $lblCity = New-Object System.Windows.Controls.TextBlock
            $lblCity.Text = $Label
            $lblCity.Foreground = $converter.ConvertFromString($(if ($IsDynamic) { '#7799BB' } else { '#AAAAAA' }))
            $lblCity.FontSize = 11

            $lblTime = New-Object System.Windows.Controls.TextBlock
            $lblTime.Text = '--:--:--'
            $lblTime.Foreground = $converter.ConvertFromString($(if ($IsDynamic) { '#44BB88' } else { '#00CC66' }))
            $lblTime.FontSize = 16
            $lblTime.FontFamily = [System.Windows.Media.FontFamily]::new('Consolas')
            $lblTime.FontWeight = [System.Windows.FontWeights]::Bold

            $panel.Children.Add($lblCity) | Out-Null
            $panel.Children.Add($lblTime) | Out-Null
            $panelWorldClocks.Children.Add($panel) | Out-Null

            $script:clockLabels[$Label] = @{ TimeLabel = $lblTime; CityLabel = $lblCity; TzId = $TzId; Abbrev = $Abbrev }
        }

        # Add base clocks
        foreach ($tz in $script:baseClockZones) {
            Add-ClockEntry -Label $tz.Label -TzId $tz.TzId -Abbrev $tz.Abbrev
            $coveredTzIds[$tz.TzId] = $true
        }

        # Add separator if there will be dynamic clocks
        $dynamicExchanges = @($script:exchangeRows | Where-Object { $_.IsActive -and -not $coveredTzIds.ContainsKey($_.TimeZoneId) })
        if ($dynamicExchanges.Count -gt 0) {
            $sep = New-Object System.Windows.Shapes.Rectangle
            $sep.Height = 1
            $sep.Margin = [System.Windows.Thickness]::new(0, 4, 0, 8)
            $sep.Fill = $converter.ConvertFromString('#0F3460')
            $panelWorldClocks.Children.Add($sep) | Out-Null

            foreach ($ex in $dynamicExchanges) {
                if ($coveredTzIds.ContainsKey($ex.TimeZoneId)) { continue }
                $coveredTzIds[$ex.TimeZoneId] = $true
                # Derive abbreviation from timezone ID
                $abbrev = ($ex.TimeZoneId -replace ' Standard Time$', '' -replace ' ', '' -split '(?=[A-Z])' | ForEach-Object { $_[0] }) -join ''
                Add-ClockEntry -Label $ex.City -TzId $ex.TimeZoneId -Abbrev $abbrev -IsDynamic $true
            }
        }
    }

    # ── World Map Drawing ────────────────────────────────────────

    function Convert-LatLonToCanvas {
        param([double]$Lat, [double]$Lon, [double]$Width, [double]$Height)
        if ($script:mapProjection -eq 'Globe') {
            return Convert-LatLonToGlobe -Lat $Lat -Lon $Lon -Width $Width -Height $Height
        }
        # Equirectangular projection
        $x = (($Lon + 180) / 360) * $Width
        $y = ((90 - $Lat) / 180) * $Height
        return @{ X = $x; Y = $y; Visible = $true }
    }

    function Convert-LatLonToGlobe {
        param([double]$Lat, [double]$Lon, [double]$Width, [double]$Height)
        $R = [Math]::Min($Width, $Height) / 2 * 0.92
        $cx = $Width / 2; $cy = $Height / 2
        $latR = $Lat * [Math]::PI / 180
        $lonR = $Lon * [Math]::PI / 180
        $cLatR = $script:globeCenterLat * [Math]::PI / 180
        $cLonR = $script:globeCenterLon * [Math]::PI / 180
        $cosC = [Math]::Sin($cLatR) * [Math]::Sin($latR) + [Math]::Cos($cLatR) * [Math]::Cos($latR) * [Math]::Cos($lonR - $cLonR)
        if ($cosC -lt 0) { return @{ X = -1000; Y = -1000; Visible = $false } }
        $x = $R * [Math]::Cos($latR) * [Math]::Sin($lonR - $cLonR)
        $y = $R * ([Math]::Cos($cLatR) * [Math]::Sin($latR) - [Math]::Sin($cLatR) * [Math]::Cos($latR) * [Math]::Cos($lonR - $cLonR))
        return @{ X = $cx + $x; Y = $cy - $y; Visible = $true }
    }

    function Get-SolarTerminator {
        param([double]$Width, [double]$Height)
        $now = [DateTime]::UtcNow
        $dayOfYear = $now.DayOfYear
        $hour = $now.Hour + $now.Minute / 60.0
        # Solar declination (approximate)
        $declination = -23.44 * [Math]::Cos(2 * [Math]::PI / 365 * ($dayOfYear + 10))
        $declRad = $declination * [Math]::PI / 180
        # Sub-solar longitude
        $subSolarLon = - ($hour / 24.0 * 360 - 180)
        if ($subSolarLon -lt -180) { $subSolarLon += 360 }
        if ($subSolarLon -gt 180) { $subSolarLon -= 360 }
        $points = [System.Collections.Generic.List[object]]::new()
        for ($lon = -180; $lon -le 180; $lon += 2) {
            $lonRad = ($lon - $subSolarLon) * [Math]::PI / 180
            $latRad = [Math]::Atan( - [Math]::Cos($lonRad) / [Math]::Tan($declRad))
            $lat = $latRad * 180 / [Math]::PI
            $pos = Convert-LatLonToCanvas -Lat $lat -Lon $lon -Width $Width -Height $Height
            if ($pos.Visible) { $points.Add($pos) }
        }
        return @{ Points = $points; SubSolarLon = $subSolarLon; Declination = $declination }
    }

    function Get-ClusterPriorityColor {
        param([string[]]$Codes)
        $priority = @('Open', 'ClosingImminent', 'ClosingSoon', 'PreMarket', 'LunchBreak', 'Holiday', 'Closed')
        $bestStatus = 'Closed'
        $bestRank = 99
        foreach ($code in $Codes) {
            $row = $script:exchangeRows | Where-Object { $_.Code -eq $code } | Select-Object -First 1
            if (-not $row) { continue }
            $exObj = [PSCustomObject]@{
                TimeZoneId = $row.TimeZoneId; OpenTimeLocal = $row.OpenTime; CloseTimeLocal = $row.CloseTime
                LunchBreakStart = $row.LunchBreakStart; LunchBreakEnd = $row.LunchBreakEnd; Code = $row.Code
            }
            $si = Get-ExchangeStatus -Exchange $exObj
            $rank = $priority.IndexOf($si.Status)
            if ($rank -ge 0 -and $rank -lt $bestRank) { $bestRank = $rank; $bestStatus = $si.Status }
        }
        return Get-StatusColor -Status $bestStatus
    }

    function Add-MapMarker {
        param([PSCustomObject]$Row, [double]$CanvasWidth, [double]$CanvasHeight, [double]$OffsetX = 0, [double]$OffsetY = 0, [bool]$IsExpanded = $false)

        $pos = Convert-LatLonToCanvas -Lat $Row.Latitude -Lon $Row.Longitude -Width $CanvasWidth -Height $CanvasHeight
        if (-not $pos.Visible) { return }
        $px = $pos.X + $OffsetX
        $py = $pos.Y + $OffsetY
        $converter = [System.Windows.Media.BrushConverter]::new()

        # Outer glow circle
        $glow = New-Object System.Windows.Shapes.Ellipse
        $glow.Width = 18; $glow.Height = 18
        $glow.Fill = $converter.ConvertFromString('#3300CC66')
        $glow.IsHitTestVisible = $false
        [System.Windows.Controls.Canvas]::SetLeft($glow, $px - 9)
        [System.Windows.Controls.Canvas]::SetTop($glow, $py - 9)
        $canvasMap.Children.Add($glow) | Out-Null

        # Main marker dot
        $marker = New-Object System.Windows.Shapes.Ellipse
        $marker.Width = 10; $marker.Height = 10
        $statusColor = Get-StatusColor -Status 'Closed'
        $marker.Fill = $converter.ConvertFromString($statusColor)
        $marker.Stroke = $converter.ConvertFromString('#FFFFFF')
        $marker.StrokeThickness = 1
        $marker.Cursor = [System.Windows.Input.Cursors]::Hand
        [System.Windows.Controls.Canvas]::SetLeft($marker, $px - 5)
        [System.Windows.Controls.Canvas]::SetTop($marker, $py - 5)

        # Tooltip
        $tp = New-Object System.Windows.Controls.ToolTip
        $tp.Content = "$($Row.DisplayName) ($($Row.Symbol))"
        [System.Windows.Controls.ToolTipService]::SetToolTip($marker, $tp)

        # Click handler
        $marker.Tag = $Row.Code
        $marker.Add_MouseLeftButtonDown({
                param($sender, $e)
                Show-ExchangeFlyout -ExchangeCode $sender.Tag
                $e.Handled = $true
            })

        $canvasMap.Children.Add($marker) | Out-Null

        # Label — position alternates for expanded cluster members
        $lblOffsetX = 7; $lblOffsetY = -7
        if ($IsExpanded -and $OffsetY -gt 0) { $lblOffsetY = 8 }
        $lbl = New-Object System.Windows.Controls.TextBlock
        $lbl.Text = $Row.Symbol
        $lbl.FontSize = 9
        $lbl.Foreground = $converter.ConvertFromString('#CCCCCC')
        $lbl.IsHitTestVisible = $false
        [System.Windows.Controls.Canvas]::SetLeft($lbl, $px + $lblOffsetX)
        [System.Windows.Controls.Canvas]::SetTop($lbl, $py + $lblOffsetY)
        $canvasMap.Children.Add($lbl) | Out-Null

        $script:mapMarkers[$Row.Code] = @{ Marker = $marker; Glow = $glow; Tooltip = $tp; Label = $lbl }
    }

    function Add-ClusterMarker {
        param([hashtable]$Cluster, [double]$CanvasWidth, [double]$CanvasHeight)
        $codes = $Cluster.Codes
        $centerLat = ($Cluster.Rows | Measure-Object -Property Latitude -Average).Average
        $centerLon = ($Cluster.Rows | Measure-Object -Property Longitude -Average).Average
        $pos = Convert-LatLonToCanvas -Lat $centerLat -Lon $centerLon -Width $CanvasWidth -Height $CanvasHeight
        if (-not $pos.Visible) { return }
        $converter = [System.Windows.Media.BrushConverter]::new()
        $clusterColor = Get-ClusterPriorityColor -Codes $codes

        # Outer glow
        $glow = New-Object System.Windows.Shapes.Ellipse
        $glow.Width = 26; $glow.Height = 26
        $glow.Fill = $converter.ConvertFromString(($clusterColor -replace '^#', '#33'))
        $glow.IsHitTestVisible = $false
        [System.Windows.Controls.Canvas]::SetLeft($glow, $pos.X - 13)
        [System.Windows.Controls.Canvas]::SetTop($glow, $pos.Y - 13)
        $canvasMap.Children.Add($glow) | Out-Null

        # Main circle (larger than single marker)
        $dot = New-Object System.Windows.Shapes.Ellipse
        $dot.Width = 18; $dot.Height = 18
        $dot.Fill = $converter.ConvertFromString($clusterColor)
        $dot.Stroke = $converter.ConvertFromString('#FFFFFF')
        $dot.StrokeThickness = 1.5
        $dot.Cursor = [System.Windows.Input.Cursors]::Hand
        [System.Windows.Controls.Canvas]::SetLeft($dot, $pos.X - 9)
        [System.Windows.Controls.Canvas]::SetTop($dot, $pos.Y - 9)
        $canvasMap.Children.Add($dot) | Out-Null

        # Count badge
        $badge = New-Object System.Windows.Controls.TextBlock
        $badge.Text = "$($codes.Count)"
        $badge.FontSize = 10
        $badge.FontWeight = [System.Windows.FontWeights]::Bold
        $badge.Foreground = $converter.ConvertFromString('#FFFFFF')
        $badge.IsHitTestVisible = $false
        $badge.TextAlignment = 'Center'
        [System.Windows.Controls.Canvas]::SetLeft($badge, $pos.X - 4)
        [System.Windows.Controls.Canvas]::SetTop($badge, $pos.Y - 7)
        $canvasMap.Children.Add($badge) | Out-Null

        # Cluster label (comma-separated symbols)
        $symbols = ($Cluster.Rows | ForEach-Object { $_.Symbol }) -join ', '
        $lbl = New-Object System.Windows.Controls.TextBlock
        $lbl.Text = $symbols
        $lbl.FontSize = 9
        $lbl.Foreground = $converter.ConvertFromString('#CCCCCC')
        $lbl.IsHitTestVisible = $false
        [System.Windows.Controls.Canvas]::SetLeft($lbl, $pos.X + 12)
        [System.Windows.Controls.Canvas]::SetTop($lbl, $pos.Y - 7)
        $canvasMap.Children.Add($lbl) | Out-Null

        # Tooltip
        $names = ($Cluster.Rows | ForEach-Object { "$($_.DisplayName) ($($_.Symbol))" }) -join "`n"
        $tp = New-Object System.Windows.Controls.ToolTip
        $tp.Content = $names
        [System.Windows.Controls.ToolTipService]::SetToolTip($dot, $tp)

        # Store cluster id on dot for click handler
        $clusterKey = ($codes | Sort-Object) -join '|'
        $dot.Tag = $clusterKey
        $dot.Add_MouseLeftButtonDown({
                param($sender, $e)
                $key = $sender.Tag
                if ($script:expandedCluster -eq $key) {
                    # Collapse: re-render map
                    $script:expandedCluster = $null
                    Initialize-WorldMap
                }
                else {
                    # Expand this cluster
                    $script:expandedCluster = $key
                    Initialize-WorldMap
                }
                $e.Handled = $true
            })

        # Store cluster elements for dynamic color updates
        $script:mapClusters[$clusterKey] = @{ Dot = $dot; Glow = $glow; Badge = $badge; Label = $lbl; Tooltip = $tp; Codes = $codes }
    }

    function Draw-SolarTerminator {
        param([double]$Width, [double]$Height)
        $converter = [System.Windows.Media.BrushConverter]::new()

        if ($script:mapProjection -eq 'Globe') {
            # Globe mode: project terminator curve onto the visible hemisphere
            $now = [DateTime]::UtcNow
            $dayOfYear = $now.DayOfYear
            $hour = $now.Hour + $now.Minute / 60.0
            $declination = -23.44 * [Math]::Cos(2 * [Math]::PI / 365 * ($dayOfYear + 10))
            $declRad = $declination * [Math]::PI / 180
            $subSolarLon = - ($hour / 24.0 * 360 - 180)
            if ($subSolarLon -lt -180) { $subSolarLon += 360 }
            if ($subSolarLon -gt 180) { $subSolarLon -= 360 }

            # Sample terminator curve
            $terminatorPoints = [System.Collections.Generic.List[object]]::new()
            for ($lon = -180; $lon -le 180; $lon += 2) {
                $lonRad = ($lon - $subSolarLon) * [Math]::PI / 180
                $latRad = [Math]::Atan( - [Math]::Cos($lonRad) / [Math]::Tan($declRad))
                $lat = $latRad * 180 / [Math]::PI
                $p = Convert-LatLonToGlobe -Lat $lat -Lon $lon -Width $Width -Height $Height
                if ($p.Visible) {
                    $terminatorPoints.Add(@{ X = $p.X; Y = $p.Y; Lat = $lat; Lon = $lon })
                }
            }

            if ($terminatorPoints.Count -ge 2) {
                # Draw the terminator curve as a polyline
                $polyLine = New-Object System.Windows.Shapes.Polyline
                $polyLine.Stroke = $converter.ConvertFromString('#66FF8800')
                $polyLine.StrokeThickness = 1.5
                $polyLine.IsHitTestVisible = $false
                $pts = New-Object System.Windows.Media.PointCollection
                foreach ($pt in $terminatorPoints) {
                    $pts.Add([System.Windows.Point]::new($pt.X, $pt.Y))
                }
                $polyLine.Points = $pts
                $canvasMap.Children.Add($polyLine) | Out-Null

                # Build night-side shading: fill the hemisphere on the night side
                # Sample globe limb (visible edge) on the night side to close the polygon
                $R = [Math]::Min($Width, $Height) / 2 * 0.92
                $cx = $Width / 2; $cy = $Height / 2
                $sunNorth = $declination -gt 0

                $nightPoly = New-Object System.Windows.Shapes.Polygon
                $nightPoly.Fill = $converter.ConvertFromString('#40000000')
                $nightPoly.IsHitTestVisible = $false
                $nightPoints = New-Object System.Windows.Media.PointCollection

                # Add terminator points
                foreach ($pt in $terminatorPoints) {
                    $nightPoints.Add([System.Windows.Point]::new($pt.X, $pt.Y))
                }

                # Close along the globe limb on the night side
                # Determine which side of the terminator is night by checking a polar point
                $nightPoleLat = if ($sunNorth) { -90 } else { 90 }
                $lastPt = $terminatorPoints[$terminatorPoints.Count - 1]
                $firstPt = $terminatorPoints[0]

                # Sample the globe edge arc from last terminator point back to first
                # Use the limb (visible circle) with angular sweep on the night side
                $lastAngle = [Math]::Atan2($lastPt.Y - $cy, $lastPt.X - $cx)
                $firstAngle = [Math]::Atan2($firstPt.Y - $cy, $firstPt.X - $cx)

                # Check which arc sweep direction covers the night side
                $nightPoleP = Convert-LatLonToGlobe -Lat $nightPoleLat -Lon 0 -Width $Width -Height $Height
                if ($nightPoleP.Visible) {
                    $nightPoleAngle = [Math]::Atan2($nightPoleP.Y - $cy, $nightPoleP.X - $cx)
                }
                else {
                    # Night pole is on far side — night covers most of visible hemisphere
                    $nightPoleAngle = [Math]::Atan2( - ($firstPt.Y - $cy), - ($firstPt.X - $cx))
                }

                # Sweep from lastAngle toward nightPoleAngle and back to firstAngle
                # Normalize angles
                $sweep = $firstAngle - $lastAngle
                if ($sweep -gt [Math]::PI) { $sweep -= 2 * [Math]::PI }
                if ($sweep -lt - [Math]::PI) { $sweep += 2 * [Math]::PI }

                # Choose the sweep direction that includes the night pole
                $midAngle = $lastAngle + $sweep / 2
                $testAngleDiff = $nightPoleAngle - $midAngle
                if ($testAngleDiff -gt [Math]::PI) { $testAngleDiff -= 2 * [Math]::PI }
                if ($testAngleDiff -lt - [Math]::PI) { $testAngleDiff += 2 * [Math]::PI }

                if ([Math]::Abs($testAngleDiff) -gt [Math]::PI / 2) {
                    # Night pole is on the other arc — reverse sweep
                    $sweep = if ($sweep -gt 0) { $sweep - 2 * [Math]::PI } else { $sweep + 2 * [Math]::PI }
                }

                $steps = [Math]::Max(20, [int]([Math]::Abs($sweep) / [Math]::PI * 40))
                for ($i = 1; $i -le $steps; $i++) {
                    $angle = $lastAngle + ($sweep * $i / $steps)
                    $nightPoints.Add([System.Windows.Point]::new($cx + $R * [Math]::Cos($angle), $cy + $R * [Math]::Sin($angle)))
                }

                $nightPoly.Points = $nightPoints
                $canvasMap.Children.Add($nightPoly) | Out-Null
            }
            return
        }

        # Flat mode: original implementation
        $solar = Get-SolarTerminator -Width $Width -Height $Height
        if ($solar.Points.Count -lt 2) { return }

        $poly = New-Object System.Windows.Shapes.Polygon
        $poly.Fill = $converter.ConvertFromString('#40000000')
        $poly.IsHitTestVisible = $false
        $pointCollection = New-Object System.Windows.Media.PointCollection

        $sunNorth = $solar.Declination -gt 0

        if ($sunNorth) {
            foreach ($pt in $solar.Points) {
                $pointCollection.Add([System.Windows.Point]::new($pt.X, $pt.Y))
            }
            $pointCollection.Add([System.Windows.Point]::new($Width, $Height))
            $pointCollection.Add([System.Windows.Point]::new(0, $Height))
        }
        else {
            foreach ($pt in $solar.Points) {
                $pointCollection.Add([System.Windows.Point]::new($pt.X, $pt.Y))
            }
            $pointCollection.Add([System.Windows.Point]::new($Width, 0))
            $pointCollection.Add([System.Windows.Point]::new(0, 0))
        }
        $poly.Points = $pointCollection
        $canvasMap.Children.Add($poly) | Out-Null
    }

    function Draw-TimeZoneBands {
        param([double]$Width, [double]$Height)
        $converter = [System.Windows.Media.BrushConverter]::new()
        $bandColors = @('#10FF6600', '#100066FF')

        if ($script:mapProjection -eq 'Globe') {
            # Globe mode: draw TZ bands as curved polygon strips
            for ($utcOffset = -12; $utcOffset -le 12; $utcOffset++) {
                $lonLeft = $utcOffset * 15 - 7.5
                $lonRight = $utcOffset * 15 + 7.5
                $colorIdx = if ($utcOffset % 2 -eq 0) { 0 } else { 1 }

                # Build left edge (top to bottom) and right edge (bottom to top)
                $leftEdge = [System.Collections.Generic.List[object]]::new()
                $rightEdge = [System.Collections.Generic.List[object]]::new()
                for ($lat = 85; $lat -ge -85; $lat -= 5) {
                    $pL = Convert-LatLonToGlobe -Lat $lat -Lon $lonLeft -Width $Width -Height $Height
                    $pR = Convert-LatLonToGlobe -Lat $lat -Lon $lonRight -Width $Width -Height $Height
                    if ($pL.Visible) { $leftEdge.Add([System.Windows.Point]::new($pL.X, $pL.Y)) }
                    if ($pR.Visible) { $rightEdge.Add([System.Windows.Point]::new($pR.X, $pR.Y)) }
                }
                if ($leftEdge.Count -lt 3 -or $rightEdge.Count -lt 3) { continue }

                $poly = New-Object System.Windows.Shapes.Polygon
                $poly.Fill = $converter.ConvertFromString($bandColors[$colorIdx])
                $poly.IsHitTestVisible = $false
                $pts = New-Object System.Windows.Media.PointCollection
                foreach ($pt in $leftEdge) { $pts.Add($pt) }
                $rightEdge.Reverse()
                foreach ($pt in $rightEdge) { $pts.Add($pt) }
                $poly.Points = $pts
                $canvasMap.Children.Add($poly) | Out-Null

                # Label at top of band if visible
                $midLon = $utcOffset * 15
                $pMid = Convert-LatLonToGlobe -Lat 80 -Lon $midLon -Width $Width -Height $Height
                if ($pMid.Visible) {
                    $lbl = New-Object System.Windows.Controls.TextBlock
                    $lbl.Text = if ($utcOffset -ge 0) { "+$utcOffset" } else { "$utcOffset" }
                    $lbl.FontSize = 7
                    $lbl.Foreground = $converter.ConvertFromString('#55FFFFFF')
                    $lbl.IsHitTestVisible = $false
                    [System.Windows.Controls.Canvas]::SetLeft($lbl, $pMid.X - 5)
                    [System.Windows.Controls.Canvas]::SetTop($lbl, $pMid.Y - 10)
                    $canvasMap.Children.Add($lbl) | Out-Null
                }
            }
            return
        }

        # Flat mode: original implementation
        for ($utcOffset = -12; $utcOffset -le 12; $utcOffset++) {
            $lonLeft = $utcOffset * 15 - 7.5
            $lonRight = $utcOffset * 15 + 7.5
            $xLeft = (($lonLeft + 180) / 360) * $Width
            $xRight = (($lonRight + 180) / 360) * $Width
            $bandWidth = $xRight - $xLeft
            if ($bandWidth -le 0) { continue }
            $colorIdx = if ($utcOffset % 2 -eq 0) { 0 } else { 1 }
            $rect = New-Object System.Windows.Shapes.Rectangle
            $rect.Width = $bandWidth; $rect.Height = $Height
            $rect.Fill = $converter.ConvertFromString($bandColors[$colorIdx])
            $rect.IsHitTestVisible = $false
            [System.Windows.Controls.Canvas]::SetLeft($rect, $xLeft)
            [System.Windows.Controls.Canvas]::SetTop($rect, 0)
            $canvasMap.Children.Add($rect) | Out-Null
            # UTC offset label at top
            $lbl = New-Object System.Windows.Controls.TextBlock
            $lbl.Text = if ($utcOffset -ge 0) { "+$utcOffset" } else { "$utcOffset" }
            $lbl.FontSize = 7
            $lbl.Foreground = $converter.ConvertFromString('#55FFFFFF')
            $lbl.IsHitTestVisible = $false
            [System.Windows.Controls.Canvas]::SetLeft($lbl, $xLeft + 2)
            [System.Windows.Controls.Canvas]::SetTop($lbl, 1)
            $canvasMap.Children.Add($lbl) | Out-Null
        }
    }

    function Draw-PoliticalBoundaries {
        param([double]$Width, [double]$Height)
        $converter = [System.Windows.Media.BrushConverter]::new()

        # Load boundary data
        $dataPath = Join-Path $PSScriptRoot '..\Data\overlays\political-boundaries.json'
        if (-not (Test-Path $dataPath)) { return }
        $data = Get-Content $dataPath -Raw | ConvertFrom-Json

        $strokeBrush = $converter.ConvertFromString('#33AAAAAA')
        $da = New-Object System.Windows.Media.DoubleCollection; $da.Add(4.0); $da.Add(2.0)

        foreach ($boundary in $data.boundaries) {
            $coords = $boundary.Coords
            if ($coords.Count -lt 2) { continue }

            if ($script:mapProjection -eq 'Globe') {
                $polyLine = New-Object System.Windows.Shapes.Polyline
                $polyLine.Stroke = $strokeBrush
                $polyLine.StrokeThickness = 0.5
                $polyLine.StrokeDashArray = $da
                $polyLine.IsHitTestVisible = $false
                $pts = New-Object System.Windows.Media.PointCollection
                $prevVisible = $false
                foreach ($c in $coords) {
                    $lat = [double]$c[0]; $lon = [double]$c[1]
                    $p = Convert-LatLonToGlobe -Lat $lat -Lon $lon -Width $Width -Height $Height
                    if ($p.Visible) {
                        $pts.Add([System.Windows.Point]::new($p.X, $p.Y))
                        $prevVisible = $true
                    }
                    elseif ($prevVisible -and $pts.Count -ge 2) {
                        # Segment break — draw what we have and start new
                        $polyLine.Points = $pts
                        $canvasMap.Children.Add($polyLine) | Out-Null
                        $polyLine = New-Object System.Windows.Shapes.Polyline
                        $polyLine.Stroke = $strokeBrush
                        $polyLine.StrokeThickness = 0.5
                        $polyLine.StrokeDashArray = $da
                        $polyLine.IsHitTestVisible = $false
                        $pts = New-Object System.Windows.Media.PointCollection
                        $prevVisible = $false
                    }
                }
                if ($pts.Count -ge 2) {
                    $polyLine.Points = $pts
                    $canvasMap.Children.Add($polyLine) | Out-Null
                }
            }
            else {
                $polyLine = New-Object System.Windows.Shapes.Polyline
                $polyLine.Stroke = $strokeBrush
                $polyLine.StrokeThickness = 0.5
                $polyLine.StrokeDashArray = $da
                $polyLine.IsHitTestVisible = $false
                $pts = New-Object System.Windows.Media.PointCollection
                foreach ($c in $coords) {
                    $lat = [double]$c[0]; $lon = [double]$c[1]
                    $x = (($lon + 180) / 360) * $Width
                    $y = ((90 - $lat) / 180) * $Height
                    $pts.Add([System.Windows.Point]::new($x, $y))
                }
                $polyLine.Points = $pts
                $canvasMap.Children.Add($polyLine) | Out-Null
            }
        }
    }

    # ── Live Data Feed Functions ──

    function Fetch-LiveFeedData {
        # Only fetch if more than 5 minutes since last fetch
        if (([DateTime]::UtcNow - $script:liveFeedLastFetch).TotalMinutes -lt 5) { return }
        $script:liveFeedLastFetch = [DateTime]::UtcNow

        # Earthquake data from USGS
        try {
            $eqUrl = 'https://earthquake.usgs.gov/earthquakes/feed/v1.0/summary/2.5_day.geojson'
            $response = Invoke-RestMethod -Uri $eqUrl -TimeoutSec 10 -ErrorAction Stop
            $script:earthquakeData = @()
            foreach ($f in $response.features) {
                $script:earthquakeData += @{
                    Lon   = [double]$f.geometry.coordinates[0]
                    Lat   = [double]$f.geometry.coordinates[1]
                    Depth = [double]$f.geometry.coordinates[2]
                    Mag   = [double]$f.properties.mag
                    Place = [string]$f.properties.place
                    Time  = [DateTimeOffset]::FromUnixTimeMilliseconds([long]$f.properties.time).UtcDateTime
                }
            }
        }
        catch {
            Write-Verbose "Earthquake feed fetch failed: $($_.Exception.Message)"
        }

        # Volcano data from Smithsonian GVP
        try {
            $volUrl = 'https://volcano.si.edu/news/WeeklyVolcanoCAP.xml'
            $response = Invoke-RestMethod -Uri $volUrl -TimeoutSec 10 -ErrorAction Stop
            $script:volcanoData = @()
            foreach ($entry in $response) {
                # CAP format: look for circle elements in area
                $areas = $entry | Select-Xml -XPath '//area' -ErrorAction SilentlyContinue
                if (-not $areas) {
                    # Try parsing as simple RSS/Atom with geo coordinates
                    $circle = $null
                    if ($entry.circle) { $circle = $entry.circle }
                    elseif ($entry.area -and $entry.area.circle) { $circle = $entry.area.circle }
                    if ($circle -and $circle -match '([-\d.]+),([-\d.]+)') {
                        $script:volcanoData += @{
                            Lat        = [double]$Matches[1]
                            Lon        = [double]$Matches[2]
                            Name       = [string]$entry.headline
                            AlertLevel = 'Warning'
                        }
                    }
                }
            }
        }
        catch {
            Write-Verbose "Volcano feed fetch failed: $($_.Exception.Message)"
            # Use a small static set of notable active volcanoes as fallback
            if (-not $script:volcanoData) {
                $script:volcanoData = @(
                    @{ Lat = 63.63; Lon = -19.62; Name = 'Katla'; AlertLevel = 'Advisory' },
                    @{ Lat = -8.34; Lon = 115.51; Name = 'Agung'; AlertLevel = 'Advisory' },
                    @{ Lat = 46.20; Lon = 122.18; Name = 'Mt St Helens'; AlertLevel = 'Normal' },
                    @{ Lat = 37.75; Lon = 14.99; Name = 'Etna'; AlertLevel = 'Advisory' },
                    @{ Lat = 19.42; Lon = -155.29; Name = 'Kilauea'; AlertLevel = 'Watch' },
                    @{ Lat = -1.47; Lon = 29.23; Name = 'Nyiragongo'; AlertLevel = 'Warning' },
                    @{ Lat = 35.36; Lon = 138.73; Name = 'Fuji'; AlertLevel = 'Normal' },
                    @{ Lat = -38.69; Lon = 176.07; Name = 'Ruapehu'; AlertLevel = 'Advisory' }
                )
            }
        }
    }

    function Draw-Earthquakes {
        param([double]$Width, [double]$Height)
        if (-not $script:earthquakeData -or $script:earthquakeData.Count -eq 0) { return }
        $converter = [System.Windows.Media.BrushConverter]::new()

        foreach ($eq in $script:earthquakeData) {
            # Size based on magnitude (2.5-9+)
            $radius = [Math]::Max(3, [Math]::Min(16, ($eq.Mag - 1) * 3))

            # Color by depth: shallow (<70km) = red, intermediate (70-300) = orange, deep (>300) = yellow
            $color = if ($eq.Depth -lt 70) { '#CCFF4444' }
            elseif ($eq.Depth -lt 300) { '#CCFF8844' }
            else { '#CCFFCC44' }

            if ($script:mapProjection -eq 'Globe') {
                $p = Convert-LatLonToGlobe -Lat $eq.Lat -Lon $eq.Lon -Width $Width -Height $Height
                if (-not $p.Visible) { continue }
                $x = $p.X; $y = $p.Y
            }
            else {
                $x = (($eq.Lon + 180) / 360) * $Width
                $y = ((90 - $eq.Lat) / 180) * $Height
            }

            # Outer glow ring
            $glow = New-Object System.Windows.Shapes.Ellipse
            $glow.Width = $radius * 3; $glow.Height = $radius * 3
            $glow.Fill = $converter.ConvertFromString(($color -replace '#CC', '#22'))
            $glow.IsHitTestVisible = $false
            [System.Windows.Controls.Canvas]::SetLeft($glow, $x - $radius * 1.5)
            [System.Windows.Controls.Canvas]::SetTop($glow, $y - $radius * 1.5)
            $canvasMap.Children.Add($glow) | Out-Null

            # Inner circle
            $dot = New-Object System.Windows.Shapes.Ellipse
            $dot.Width = $radius; $dot.Height = $radius
            $dot.Fill = $converter.ConvertFromString($color)
            $dot.Stroke = $converter.ConvertFromString('#88FFFFFF')
            $dot.StrokeThickness = 0.5
            $dot.IsHitTestVisible = $false
            $dot.ToolTip = "M$($eq.Mag) — $($eq.Place)`nDepth: $($eq.Depth)km`n$($eq.Time.ToString('yyyy-MM-dd HH:mm')) UTC"
            [System.Windows.Controls.Canvas]::SetLeft($dot, $x - $radius / 2)
            [System.Windows.Controls.Canvas]::SetTop($dot, $y - $radius / 2)
            $canvasMap.Children.Add($dot) | Out-Null
        }
    }

    function Draw-Volcanoes {
        param([double]$Width, [double]$Height)
        if (-not $script:volcanoData -or $script:volcanoData.Count -eq 0) { return }
        $converter = [System.Windows.Media.BrushConverter]::new()

        foreach ($vol in $script:volcanoData) {
            # Color by alert level
            $color = switch ($vol.AlertLevel) {
                'Warning' { '#CCFF2222' }
                'Watch' { '#CCFF8800' }
                'Advisory' { '#CCFFCC00' }
                default { '#CC44AA44' }
            }

            if ($script:mapProjection -eq 'Globe') {
                $p = Convert-LatLonToGlobe -Lat $vol.Lat -Lon $vol.Lon -Width $Width -Height $Height
                if (-not $p.Visible) { continue }
                $x = $p.X; $y = $p.Y
            }
            else {
                $x = (($vol.Lon + 180) / 360) * $Width
                $y = ((90 - $vol.Lat) / 180) * $Height
            }

            # Triangle marker
            $tri = New-Object System.Windows.Shapes.Polygon
            $tri.Fill = $converter.ConvertFromString($color)
            $tri.Stroke = $converter.ConvertFromString('#88FFFFFF')
            $tri.StrokeThickness = 0.5
            $tri.IsHitTestVisible = $false
            $tri.ToolTip = "$($vol.Name)`nAlert: $($vol.AlertLevel)"
            $s = 6 # half-size
            $pts = New-Object System.Windows.Media.PointCollection
            $pts.Add([System.Windows.Point]::new($x, $y - $s * 1.3))       # top
            $pts.Add([System.Windows.Point]::new($x - $s, $y + $s * 0.7))  # bottom-left
            $pts.Add([System.Windows.Point]::new($x + $s, $y + $s * 0.7))  # bottom-right
            $tri.Points = $pts
            $canvasMap.Children.Add($tri) | Out-Null
        }
    }

    function Draw-SubmarineCables {
        param([double]$Width, [double]$Height)
        $converter = [System.Windows.Media.BrushConverter]::new()

        # Load cable data on first use
        if (-not $script:submarineCableData) {
            $dataPath = Join-Path $PSScriptRoot '..\Data\overlays\submarine-cables.json'
            if (-not (Test-Path $dataPath)) { return }
            $script:submarineCableData = Get-Content $dataPath -Raw | ConvertFrom-Json
        }
        if (-not $script:submarineCableData.cables) { return }

        foreach ($cable in $script:submarineCableData.cables) {
            $cableColor = if ($cable.Color) { $cable.Color } else { '#00BCD4' }
            # Reduce alpha for subtlety
            try { $strokeBrush = $converter.ConvertFromString($cableColor) } catch { $strokeBrush = $converter.ConvertFromString('#00BCD4') }
            $strokeBrush = $strokeBrush.Clone()
            $strokeBrush.Opacity = 0.45

            foreach ($segment in $cable.Segments) {
                if ($segment.Count -lt 2) { continue }

                if ($script:mapProjection -eq 'Globe') {
                    $polyLine = New-Object System.Windows.Shapes.Polyline
                    $polyLine.Stroke = $strokeBrush
                    $polyLine.StrokeThickness = 1.0
                    $polyLine.IsHitTestVisible = $false
                    $polyLine.ToolTip = $cable.Name
                    $pts = New-Object System.Windows.Media.PointCollection
                    $prevVisible = $false
                    foreach ($c in $segment) {
                        $lat = [double]$c[0]; $lon = [double]$c[1]
                        $p = Convert-LatLonToGlobe -Lat $lat -Lon $lon -Width $Width -Height $Height
                        if ($p.Visible) {
                            $pts.Add([System.Windows.Point]::new($p.X, $p.Y))
                            $prevVisible = $true
                        }
                        elseif ($prevVisible -and $pts.Count -ge 2) {
                            $polyLine.Points = $pts
                            $canvasMap.Children.Add($polyLine) | Out-Null
                            $polyLine = New-Object System.Windows.Shapes.Polyline
                            $polyLine.Stroke = $strokeBrush
                            $polyLine.StrokeThickness = 1.0
                            $polyLine.IsHitTestVisible = $false
                            $polyLine.ToolTip = $cable.Name
                            $pts = New-Object System.Windows.Media.PointCollection
                            $prevVisible = $false
                        }
                    }
                    if ($pts.Count -ge 2) {
                        $polyLine.Points = $pts
                        $canvasMap.Children.Add($polyLine) | Out-Null
                    }
                }
                else {
                    $polyLine = New-Object System.Windows.Shapes.Polyline
                    $polyLine.Stroke = $strokeBrush
                    $polyLine.StrokeThickness = 1.0
                    $polyLine.IsHitTestVisible = $false
                    $polyLine.ToolTip = $cable.Name
                    $pts = New-Object System.Windows.Media.PointCollection
                    foreach ($c in $segment) {
                        $lat = [double]$c[0]; $lon = [double]$c[1]
                        $x = (($lon + 180) / 360) * $Width
                        $y = ((90 - $lat) / 180) * $Height
                        $pts.Add([System.Windows.Point]::new($x, $y))
                    }
                    $polyLine.Points = $pts
                    $canvasMap.Children.Add($polyLine) | Out-Null
                }
            }
        }
    }

    function Draw-TZBoundaries {
        param([double]$Width, [double]$Height)
        $converter = [System.Windows.Media.BrushConverter]::new()

        # Load TZ boundary data on first use
        if (-not $script:tzBoundaryData) {
            $dataPath = Join-Path $PSScriptRoot '..\Data\overlays\tz-boundaries.json'
            if (-not (Test-Path $dataPath)) { return }
            $script:tzBoundaryData = Get-Content $dataPath -Raw | ConvertFrom-Json
        }
        if (-not $script:tzBoundaryData.zones) { return }

        # UTC-offset → hue mapping: spread -12..+14 across a colour wheel
        $da = New-Object System.Windows.Media.DoubleCollection; $da.Add(3.0); $da.Add(1.5)

        foreach ($zone in $script:tzBoundaryData.zones) {
            # Parse offset hours for colour
            $hue = 180  # default cyan
            if ($zone.UtcOffset -match '^([+-])(\d{2}):(\d{2})$') {
                $hrs = [int]$Matches[2] + [int]$Matches[3] / 60
                if ($Matches[1] -eq '-') { $hrs = - $hrs }
                $hue = (($hrs + 12) / 26) * 360  # -12..+14 → 0..360
            }
            # HSL→RGB (saturation 0.7, lightness 0.55)
            $s = 0.7; $l = 0.55
            $c2 = (1 - [Math]::Abs(2 * $l - 1)) * $s
            $hp = $hue / 60
            $x2 = $c2 * (1 - [Math]::Abs($hp % 2 - 1))
            $m = $l - $c2 / 2
            switch ([Math]::Floor($hp) % 6) {
                0 { $r = $c2; $g = $x2; $b = 0 }
                1 { $r = $x2; $g = $c2; $b = 0 }
                2 { $r = 0; $g = $c2; $b = $x2 }
                3 { $r = 0; $g = $x2; $b = $c2 }
                4 { $r = $x2; $g = 0; $b = $c2 }
                5 { $r = $c2; $g = 0; $b = $x2 }
            }
            $hexR = [int](($r + $m) * 255); $hexG = [int](($g + $m) * 255); $hexB = [int](($b + $m) * 255)
            $colorHex = '#{0:X2}{1:X2}{2:X2}' -f $hexR, $hexG, $hexB
            try { $strokeBrush = $converter.ConvertFromString($colorHex) } catch { $strokeBrush = $converter.ConvertFromString('#55BBDD') }
            $strokeBrush = $strokeBrush.Clone()
            $strokeBrush.Opacity = 0.5
            $tip = "$($zone.TzId)  (UTC $($zone.UtcOffset))"

            foreach ($ring in $zone.Rings) {
                if ($ring.Count -lt 3) { continue }

                if ($script:mapProjection -eq 'Globe') {
                    $polyLine = New-Object System.Windows.Shapes.Polyline
                    $polyLine.Stroke = $strokeBrush
                    $polyLine.StrokeThickness = 0.7
                    $polyLine.StrokeDashArray = $da
                    $polyLine.IsHitTestVisible = $false
                    $polyLine.ToolTip = $tip
                    $pts = New-Object System.Windows.Media.PointCollection
                    $prevVisible = $false
                    foreach ($c in $ring) {
                        $lat = [double]$c[0]; $lon = [double]$c[1]
                        $p = Convert-LatLonToGlobe -Lat $lat -Lon $lon -Width $Width -Height $Height
                        if ($p.Visible) {
                            $pts.Add([System.Windows.Point]::new($p.X, $p.Y))
                            $prevVisible = $true
                        }
                        elseif ($prevVisible -and $pts.Count -ge 2) {
                            $polyLine.Points = $pts
                            $canvasMap.Children.Add($polyLine) | Out-Null
                            $polyLine = New-Object System.Windows.Shapes.Polyline
                            $polyLine.Stroke = $strokeBrush
                            $polyLine.StrokeThickness = 0.7
                            $polyLine.StrokeDashArray = $da
                            $polyLine.IsHitTestVisible = $false
                            $polyLine.ToolTip = $tip
                            $pts = New-Object System.Windows.Media.PointCollection
                            $prevVisible = $false
                        }
                    }
                    if ($pts.Count -ge 2) {
                        $polyLine.Points = $pts
                        $canvasMap.Children.Add($polyLine) | Out-Null
                    }
                }
                else {
                    $polyLine = New-Object System.Windows.Shapes.Polyline
                    $polyLine.Stroke = $strokeBrush
                    $polyLine.StrokeThickness = 0.7
                    $polyLine.StrokeDashArray = $da
                    $polyLine.IsHitTestVisible = $false
                    $polyLine.ToolTip = $tip
                    $pts = New-Object System.Windows.Media.PointCollection
                    foreach ($c in $ring) {
                        $lat = [double]$c[0]; $lon = [double]$c[1]
                        $x = (($lon + 180) / 360) * $Width
                        $y = ((90 - $lat) / 180) * $Height
                        $pts.Add([System.Windows.Point]::new($x, $y))
                    }
                    $polyLine.Points = $pts
                    $canvasMap.Children.Add($polyLine) | Out-Null
                }
            }
        }
    }

    function Draw-PowerPlants {
        param([double]$Width, [double]$Height)
        $converter = [System.Windows.Media.BrushConverter]::new()

        # Load power plant data on first use
        if (-not $script:powerPlantData) {
            $dataPath = Join-Path $PSScriptRoot '..\Data\overlays\power-plants.json'
            if (-not (Test-Path $dataPath)) { return }
            $script:powerPlantData = Get-Content $dataPath -Raw | ConvertFrom-Json
        }
        if (-not $script:powerPlantData.plants) { return }

        foreach ($plant in $script:powerPlantData.plants) {
            $pos = Convert-LatLonToCanvas -Lat $plant.Lat -Lon $plant.Lon -Width $Width -Height $Height
            if (-not $pos.Visible) { continue }

            # Size based on capacity (MW)
            $mw = [double]$plant.MW
            $size = if ($mw -ge 5000) { 5.0 } elseif ($mw -ge 2000) { 3.5 } else { 2.5 }

            $dot = New-Object System.Windows.Shapes.Ellipse
            $dot.Width = $size; $dot.Height = $size
            try { $dot.Fill = $converter.ConvertFromString($plant.Color) } catch { $dot.Fill = $converter.ConvertFromString('#AAAAAA') }
            $dot.Opacity = 0.7
            $dot.IsHitTestVisible = $false
            $dot.ToolTip = "$($plant.Name) ($($plant.Country))`n$($plant.Fuel) — $($plant.MW) MW"
            [System.Windows.Controls.Canvas]::SetLeft($dot, $pos.X - $size / 2)
            [System.Windows.Controls.Canvas]::SetTop($dot, $pos.Y - $size / 2)
            $canvasMap.Children.Add($dot) | Out-Null
        }
    }

    function Draw-ConflictZones {
        param([double]$Width, [double]$Height)
        $converter = [System.Windows.Media.BrushConverter]::new()

        # Load conflict zone data on first use
        if (-not $script:conflictZoneData) {
            $dataPath = Join-Path $PSScriptRoot '..\Data\overlays\conflict-zones.json'
            if (-not (Test-Path $dataPath)) { return }
            $script:conflictZoneData = Get-Content $dataPath -Raw | ConvertFrom-Json
        }
        if (-not $script:conflictZoneData.zones) { return }

        foreach ($zone in $script:conflictZoneData.zones) {
            $pos = Convert-LatLonToCanvas -Lat $zone.Lat -Lon $zone.Lon -Width $Width -Height $Height
            if (-not $pos.Visible) { continue }

            # Size and intensity based on event count
            $events = [int]$zone.Events
            $size = if ($events -ge 500) { 8.0 } elseif ($events -ge 100) { 5.5 } elseif ($events -ge 20) { 3.5 } else { 2.0 }
            $alpha = if ($events -ge 500) { 180 } elseif ($events -ge 100) { 140 } elseif ($events -ge 20) { 100 } else { 60 }

            $dot = New-Object System.Windows.Shapes.Ellipse
            $dot.Width = $size; $dot.Height = $size
            $dot.Fill = New-Object System.Windows.Media.SolidColorBrush(
                [System.Windows.Media.Color]::FromArgb($alpha, 255, 50, 50))
            $dot.IsHitTestVisible = $false
            $dot.ToolTip = "$($zone.Country)`nEvents: $($zone.Events), Deaths: $($zone.Deaths)`nMost recent: $($zone.LastYear)"
            [System.Windows.Controls.Canvas]::SetLeft($dot, $pos.X - $size / 2)
            [System.Windows.Controls.Canvas]::SetTop($dot, $pos.Y - $size / 2)
            $canvasMap.Children.Add($dot) | Out-Null
        }
    }

    function Draw-OverlayLegend {
        param(
            [double]$Width,
            [double]$Height,
            [bool]$ShowConflict,
            [bool]$ShowPower,
            [bool]$ShowVolcano
        )
        if (-not $ShowConflict -and -not $ShowPower -and -not $ShowVolcano) { return }

        $converter = [System.Windows.Media.BrushConverter]::new()
        $panel = New-Object System.Windows.Controls.StackPanel
        $panel.Orientation = 'Vertical'

        $border = New-Object System.Windows.Controls.Border
        $border.Background = $converter.ConvertFromString('#CC1A1A2E')
        $border.BorderBrush = $converter.ConvertFromString('#440F3460')
        $border.BorderThickness = [System.Windows.Thickness]::new(1)
        $border.CornerRadius = [System.Windows.CornerRadius]::new(4)
        $border.Padding = [System.Windows.Thickness]::new(8, 6, 8, 6)
        $border.Child = $panel
        $border.IsHitTestVisible = $false

        # ── Conflict Zones section ──
        if ($ShowConflict) {
            $hdr = New-Object System.Windows.Controls.TextBlock
            $hdr.Text = 'Conflict Zones'
            $hdr.Foreground = $converter.ConvertFromString('#FF5555')
            $hdr.FontSize = 9; $hdr.FontWeight = 'Bold'
            $hdr.Margin = [System.Windows.Thickness]::new(0, 0, 0, 2)
            $panel.Children.Add($hdr) | Out-Null

            $conflictItems = @(
                @{ Size = 8.0; Alpha = 180; Label = '500+ events' },
                @{ Size = 5.5; Alpha = 140; Label = '100-499' },
                @{ Size = 3.5; Alpha = 100; Label = '20-99' },
                @{ Size = 2.0; Alpha = 60; Label = '< 20' }
            )
            foreach ($item in $conflictItems) {
                $row = New-Object System.Windows.Controls.StackPanel
                $row.Orientation = 'Horizontal'
                $row.Margin = [System.Windows.Thickness]::new(4, 1, 0, 1)
                $swatch = New-Object System.Windows.Shapes.Ellipse
                $swatch.Width = $item.Size; $swatch.Height = $item.Size
                $swatch.Fill = New-Object System.Windows.Media.SolidColorBrush(
                    [System.Windows.Media.Color]::FromArgb($item.Alpha, 255, 50, 50))
                $swatch.Margin = [System.Windows.Thickness]::new(0, 0, 5, 0)
                $swatch.VerticalAlignment = 'Center'
                $lbl = New-Object System.Windows.Controls.TextBlock
                $lbl.Text = $item.Label
                $lbl.Foreground = $converter.ConvertFromString('#CCCCCC')
                $lbl.FontSize = 8; $lbl.VerticalAlignment = 'Center'
                $row.Children.Add($swatch) | Out-Null
                $row.Children.Add($lbl) | Out-Null
                $panel.Children.Add($row) | Out-Null
            }
        }

        # ── Power Plants section ──
        if ($ShowPower) {
            $hdr = New-Object System.Windows.Controls.TextBlock
            $hdr.Text = 'Power Plants'
            $hdr.Foreground = $converter.ConvertFromString('#AA8855')
            $hdr.FontSize = 9; $hdr.FontWeight = 'Bold'
            $hdr.Margin = [System.Windows.Thickness]::new(0, $(if ($ShowConflict) { 5 } else { 0 }), 0, 2)
            $panel.Children.Add($hdr) | Out-Null

            $fuelItems = @(
                @{ Color = '#FFD700'; Label = 'Nuclear' },
                @{ Color = '#4A4A4A'; Label = 'Coal' },
                @{ Color = '#FF8C00'; Label = 'Gas' },
                @{ Color = '#8B4513'; Label = 'Oil' },
                @{ Color = '#1E90FF'; Label = 'Hydro' },
                @{ Color = '#32CD32'; Label = 'Wind' },
                @{ Color = '#FFD93D'; Label = 'Solar' },
                @{ Color = '#DC143C'; Label = 'Geothermal' }
            )
            foreach ($item in $fuelItems) {
                $row = New-Object System.Windows.Controls.StackPanel
                $row.Orientation = 'Horizontal'
                $row.Margin = [System.Windows.Thickness]::new(4, 1, 0, 1)
                $swatch = New-Object System.Windows.Shapes.Ellipse
                $swatch.Width = 6; $swatch.Height = 6
                $swatch.Fill = $converter.ConvertFromString($item.Color)
                $swatch.Opacity = 0.7
                $swatch.Margin = [System.Windows.Thickness]::new(0, 0, 5, 0)
                $swatch.VerticalAlignment = 'Center'
                $lbl = New-Object System.Windows.Controls.TextBlock
                $lbl.Text = $item.Label
                $lbl.Foreground = $converter.ConvertFromString('#CCCCCC')
                $lbl.FontSize = 8; $lbl.VerticalAlignment = 'Center'
                $row.Children.Add($swatch) | Out-Null
                $row.Children.Add($lbl) | Out-Null
                $panel.Children.Add($row) | Out-Null
            }
        }

        # ── Volcanoes section ──
        if ($ShowVolcano) {
            $hdr = New-Object System.Windows.Controls.TextBlock
            $hdr.Text = 'Volcanoes'
            $hdr.Foreground = $converter.ConvertFromString('#FF8800')
            $hdr.FontSize = 9; $hdr.FontWeight = 'Bold'
            $hdr.Margin = [System.Windows.Thickness]::new(0, $(if ($ShowConflict -or $ShowPower) { 5 } else { 0 }), 0, 2)
            $panel.Children.Add($hdr) | Out-Null

            $volcanoItems = @(
                @{ Color = '#FF2222'; Label = 'Warning' },
                @{ Color = '#FF8800'; Label = 'Watch' },
                @{ Color = '#FFCC00'; Label = 'Advisory' },
                @{ Color = '#44AA44'; Label = 'Normal' }
            )
            foreach ($item in $volcanoItems) {
                $row = New-Object System.Windows.Controls.StackPanel
                $row.Orientation = 'Horizontal'
                $row.Margin = [System.Windows.Thickness]::new(4, 1, 0, 1)
                $tri = New-Object System.Windows.Shapes.Polygon
                $tri.Fill = $converter.ConvertFromString($item.Color)
                $tri.Stroke = $converter.ConvertFromString('#88FFFFFF')
                $tri.StrokeThickness = 0.5
                $pts = New-Object System.Windows.Media.PointCollection
                $pts.Add([System.Windows.Point]::new(4, 0))
                $pts.Add([System.Windows.Point]::new(0, 8))
                $pts.Add([System.Windows.Point]::new(8, 8))
                $tri.Points = $pts
                $tri.Margin = [System.Windows.Thickness]::new(0, 0, 5, 0)
                $tri.VerticalAlignment = 'Center'
                $lbl = New-Object System.Windows.Controls.TextBlock
                $lbl.Text = $item.Label
                $lbl.Foreground = $converter.ConvertFromString('#CCCCCC')
                $lbl.FontSize = 8; $lbl.VerticalAlignment = 'Center'
                $row.Children.Add($tri) | Out-Null
                $row.Children.Add($lbl) | Out-Null
                $panel.Children.Add($row) | Out-Null
            }
        }

        # Position at bottom-right of map canvas (avoids market data panel on the left)
        $border.Measure([System.Windows.Size]::new([double]::PositiveInfinity, [double]::PositiveInfinity))
        $legendWidth = $border.DesiredSize.Width
        $legendHeight = $border.DesiredSize.Height
        [System.Windows.Controls.Canvas]::SetLeft($border, $Width - $legendWidth - 10)
        [System.Windows.Controls.Canvas]::SetTop($border, $Height - $legendHeight - 10)
        $canvasMap.Children.Add($border) | Out-Null
    }

    function Draw-GlobeOutline {
        param([double]$Width, [double]$Height)
        $converter = [System.Windows.Media.BrushConverter]::new()
        $R = [Math]::Min($Width, $Height) / 2 * 0.92
        $cx = $Width / 2; $cy = $Height / 2

        # Star field background (seeded for consistency)
        $rng = New-Object System.Random(42)
        for ($i = 0; $i -lt 200; $i++) {
            $sx = $rng.NextDouble() * $Width
            $sy = $rng.NextDouble() * $Height
            # Skip stars inside the globe circle
            $dx = $sx - $cx; $dy = $sy - $cy
            if ([Math]::Sqrt($dx * $dx + $dy * $dy) -lt ($R + 4)) { continue }
            $starSize = 0.5 + $rng.NextDouble() * 1.5
            $brightness = [int](80 + $rng.NextDouble() * 175)
            $star = New-Object System.Windows.Shapes.Ellipse
            $star.Width = $starSize; $star.Height = $starSize
            $hexBright = $brightness.ToString('X2')
            $star.Fill = $converter.ConvertFromString("#$hexBright$hexBright$hexBright")
            $star.IsHitTestVisible = $false
            [System.Windows.Controls.Canvas]::SetLeft($star, $sx)
            [System.Windows.Controls.Canvas]::SetTop($star, $sy)
            $canvasMap.Children.Add($star) | Out-Null
        }

        # Atmosphere glow — radial gradient ring behind globe
        $glowR = $R * 1.08
        $glowEllipse = New-Object System.Windows.Shapes.Ellipse
        $glowEllipse.Width = $glowR * 2; $glowEllipse.Height = $glowR * 2
        $glowBrush = New-Object System.Windows.Media.RadialGradientBrush
        $glowBrush.GradientStops.Add((New-Object System.Windows.Media.GradientStop([System.Windows.Media.Color]::FromArgb(0, 60, 120, 255), 0.85)))
        $glowBrush.GradientStops.Add((New-Object System.Windows.Media.GradientStop([System.Windows.Media.Color]::FromArgb(40, 60, 120, 255), 0.92)))
        $glowBrush.GradientStops.Add((New-Object System.Windows.Media.GradientStop([System.Windows.Media.Color]::FromArgb(0, 60, 120, 255), 1.0)))
        $glowEllipse.Fill = $glowBrush
        $glowEllipse.IsHitTestVisible = $false
        [System.Windows.Controls.Canvas]::SetLeft($glowEllipse, $cx - $glowR)
        [System.Windows.Controls.Canvas]::SetTop($glowEllipse, $cy - $glowR)
        $canvasMap.Children.Add($glowEllipse) | Out-Null

        # Globe circle
        $circle = New-Object System.Windows.Shapes.Ellipse
        $circle.Width = $R * 2; $circle.Height = $R * 2
        $circle.Stroke = $converter.ConvertFromString('#333355')
        $circle.StrokeThickness = 1
        $circle.Fill = $converter.ConvertFromString('#0D1B2A')
        $circle.IsHitTestVisible = $false
        [System.Windows.Controls.Canvas]::SetLeft($circle, $cx - $R)
        [System.Windows.Controls.Canvas]::SetTop($circle, $cy - $R)
        $canvasMap.Children.Add($circle) | Out-Null
        # Graticule lines
        $gratBrush = $converter.ConvertFromString('#1AFFFFFF')
        # Latitude circles
        for ($lat = -60; $lat -le 60; $lat += 30) {
            $polyLine = New-Object System.Windows.Shapes.Polyline
            $polyLine.Stroke = $gratBrush; $polyLine.StrokeThickness = 0.5
            $polyLine.IsHitTestVisible = $false
            $pts = New-Object System.Windows.Media.PointCollection
            for ($lon = -180; $lon -le 180; $lon += 5) {
                $p = Convert-LatLonToGlobe -Lat $lat -Lon $lon -Width $Width -Height $Height
                if ($p.Visible) { $pts.Add([System.Windows.Point]::new($p.X, $p.Y)) }
            }
            $polyLine.Points = $pts
            $canvasMap.Children.Add($polyLine) | Out-Null
        }
        # Longitude arcs
        for ($lon = -180; $lon -lt 180; $lon += 30) {
            $polyLine = New-Object System.Windows.Shapes.Polyline
            $polyLine.Stroke = $gratBrush; $polyLine.StrokeThickness = 0.5
            $polyLine.IsHitTestVisible = $false
            $pts = New-Object System.Windows.Media.PointCollection
            for ($lat = -90; $lat -le 90; $lat += 5) {
                $p = Convert-LatLonToGlobe -Lat $lat -Lon $lon -Width $Width -Height $Height
                if ($p.Visible) { $pts.Add([System.Windows.Point]::new($p.X, $p.Y)) }
            }
            $polyLine.Points = $pts
            $canvasMap.Children.Add($polyLine) | Out-Null
        }
    }

    function Initialize-WorldMap {
        $canvasMap.Children.Clear()
        $script:mapMarkers = @{}
        $script:mapClusters = @{}

        $converter = [System.Windows.Media.BrushConverter]::new()
        $w = $canvasMap.ActualWidth
        $h = $canvasMap.ActualHeight

        if ($w -lt 50 -or $h -lt 50) { return }

        # Determine map style from ComboBox
        $selectedStyle = 'NightSatellite'
        if ($cmbMapStyle.SelectedItem -and $cmbMapStyle.SelectedItem.Tag) {
            $selectedStyle = $cmbMapStyle.SelectedItem.Tag
        }

        # Globe mode: draw globe background and graticule
        if ($script:mapProjection -eq 'Globe') {
            $mapBg = New-Object System.Windows.Shapes.Rectangle
            $mapBg.Width = $w; $mapBg.Height = $h
            $mapBg.Fill = $converter.ConvertFromString('#0A0A1A')
            $mapBg.IsHitTestVisible = $false
            $canvasMap.Children.Add($mapBg) | Out-Null
            Draw-GlobeOutline -Width $w -Height $h
            # Draw continent polygons on globe
            Draw-GlobeContinents -Width $w -Height $h
        }
        else {
            # Flat map backgrounds
            $useImage = $false
            $imageToLoad = $null
            switch ($selectedStyle) {
                'NightSatellite' { $imageToLoad = $mapImagePath }
                'BlueMarble' {
                    $bmPath = Join-Path $appDataDir 'bluemarble.jpg'
                    if (-not (Test-Path $bmPath)) {
                        try {
                            $bmUrl = 'https://eoimages.gsfc.nasa.gov/images/imagerecords/73000/73909/world.topo.bathy.200412.3x5400x2700.jpg'
                            Invoke-WebRequest -Uri $bmUrl -OutFile $bmPath -UseBasicParsing -TimeoutSec 60
                        }
                        catch { Write-Verbose "Blue Marble download failed: $_" }
                    }
                    if (Test-Path $bmPath) { $imageToLoad = $bmPath }
                }
            }
            if ($imageToLoad -and (Test-Path $imageToLoad) -and $selectedStyle -ne 'VectorDark' -and $selectedStyle -ne 'Minimal') {
                try {
                    $uri = New-Object System.Uri($imageToLoad, [System.UriKind]::Absolute)
                    $bmp = New-Object System.Windows.Media.Imaging.BitmapImage
                    $bmp.BeginInit()
                    $bmp.UriSource = $uri
                    $bmp.CacheOption = [System.Windows.Media.Imaging.BitmapCacheOption]::OnLoad
                    $bmp.EndInit()
                    $img = New-Object System.Windows.Controls.Image
                    $img.Source = $bmp; $img.Width = $w; $img.Height = $h
                    $img.Stretch = [System.Windows.Media.Stretch]::Fill
                    $img.IsHitTestVisible = $false
                    [System.Windows.Controls.Canvas]::SetLeft($img, 0)
                    [System.Windows.Controls.Canvas]::SetTop($img, 0)
                    $canvasMap.Children.Add($img) | Out-Null
                    $useImage = $true
                }
                catch { Write-Verbose "Failed to load map image: $($_.Exception.Message)" }
            }

            if (-not $useImage) {
                $bgColor = if ($selectedStyle -eq 'Minimal') { '#111122' } else { '#0D1B2A' }
                $mapBg = New-Object System.Windows.Shapes.Rectangle
                $mapBg.Width = $w; $mapBg.Height = $h
                $mapBg.Fill = $converter.ConvertFromString($bgColor)
                $canvasMap.Children.Add($mapBg) | Out-Null
            }

            # Grid lines
            $showGrid = (-not $useImage) -or ($chkShowGridLines -and $chkShowGridLines.IsChecked)
            if ($showGrid) {
                $gridBrush = $converter.ConvertFromString($(if ($useImage) { '#33FFFFFF' } else { '#152238' }))
                for ($i = 1; $i -lt 6; $i++) {
                    $line = New-Object System.Windows.Shapes.Line
                    $line.X1 = 0; $line.X2 = $w; $line.Y1 = ($h / 6) * $i; $line.Y2 = $line.Y1
                    $line.Stroke = $gridBrush; $line.StrokeThickness = 0.5
                    $line.IsHitTestVisible = $false
                    $canvasMap.Children.Add($line) | Out-Null
                }
                for ($i = 1; $i -lt 12; $i++) {
                    $line = New-Object System.Windows.Shapes.Line
                    $line.X1 = ($w / 12) * $i; $line.X2 = $line.X1; $line.Y1 = 0; $line.Y2 = $h
                    $line.Stroke = $gridBrush; $line.StrokeThickness = 0.5
                    $line.IsHitTestVisible = $false
                    $canvasMap.Children.Add($line) | Out-Null
                }
                $eqLine = New-Object System.Windows.Shapes.Line
                $eqLine.X1 = 0; $eqLine.X2 = $w; $eqLine.Y1 = $h / 2; $eqLine.Y2 = $h / 2
                $eqLine.Stroke = $converter.ConvertFromString($(if ($useImage) { '#44FFFFFF' } else { '#1E3A5F' }))
                $eqLine.StrokeThickness = 1; $eqLine.IsHitTestVisible = $false
                $da = New-Object System.Windows.Media.DoubleCollection; $da.Add(4.0); $da.Add(4.0)
                $eqLine.StrokeDashArray = $da
                $canvasMap.Children.Add($eqLine) | Out-Null
            }

            # Draw continent outlines (when no satellite image)
            if (-not $useImage) {
                $fillColor = switch ($selectedStyle) {
                    'Minimal' { '#1A1A33' }
                    default { '#1B3A2A' }
                }
                $strokeColor = switch ($selectedStyle) {
                    'Minimal' { '#333366' }
                    default { '#2D6B45' }
                }
                $landFill = $converter.ConvertFromString($fillColor)
                $landStroke = $converter.ConvertFromString($strokeColor)
                Draw-FlatContinents -Width $w -Height $h -Fill $landFill -Stroke $landStroke
            }
        } # end Flat vs Globe

        # ── Overlays (before markers) ──

        # Political boundaries
        if ($chkPoliticalBounds -and $chkPoliticalBounds.IsChecked) {
            Draw-PoliticalBoundaries -Width $w -Height $h
        }

        # Timezone bands
        if ($chkTimeZoneBands -and $chkTimeZoneBands.IsChecked) {
            Draw-TimeZoneBands -Width $w -Height $h
        }

        # Solar terminator (day/night overlay)
        if ($chkTerminator -and $chkTerminator.IsChecked) {
            Draw-SolarTerminator -Width $w -Height $h
        }

        # Live feed overlays (fetch data if stale, then render)
        $needLiveFetch = ($chkEarthquakes -and $chkEarthquakes.IsChecked) -or ($chkVolcanoes -and $chkVolcanoes.IsChecked)
        if ($needLiveFetch) { Fetch-LiveFeedData }

        if ($chkEarthquakes -and $chkEarthquakes.IsChecked) {
            Draw-Earthquakes -Width $w -Height $h
        }
        if ($chkVolcanoes -and $chkVolcanoes.IsChecked) {
            Draw-Volcanoes -Width $w -Height $h
        }

        # Submarine cables
        if ($chkSubmarineCables -and $chkSubmarineCables.IsChecked) {
            Draw-SubmarineCables -Width $w -Height $h
        }

        # Timezone boundaries
        if ($chkTZBoundaries -and $chkTZBoundaries.IsChecked) {
            Draw-TZBoundaries -Width $w -Height $h
        }

        # Power plants
        if ($chkPowerPlants -and $chkPowerPlants.IsChecked) {
            Draw-PowerPlants -Width $w -Height $h
        }

        # Conflict zones
        if ($chkConflictZones -and $chkConflictZones.IsChecked) {
            Draw-ConflictZones -Width $w -Height $h
        }

        # Overlay legend
        Draw-OverlayLegend -Width $w -Height $h `
            -ShowConflict ($chkConflictZones -and $chkConflictZones.IsChecked) `
            -ShowPower ($chkPowerPlants -and $chkPowerPlants.IsChecked) `
            -ShowVolcano ($chkVolcanoes -and $chkVolcanoes.IsChecked)

        # ── Place exchange markers with clustering ──

        $activeRows = @($script:exchangeRows | Where-Object { $_.IsActive })
        $clusterThreshold = 20  # pixel distance to cluster

        # Compute pixel positions for all active rows
        $positions = @{}
        foreach ($row in $activeRows) {
            $pos = Convert-LatLonToCanvas -Lat $row.Latitude -Lon $row.Longitude -Width $w -Height $h
            $positions[$row.Code] = @{ Row = $row; X = $pos.X; Y = $pos.Y; Visible = $pos.Visible; Clustered = $false }
        }

        # Find clusters (groups of markers within threshold distance)
        $clusters = [System.Collections.Generic.List[hashtable]]::new()
        $clustered = @{}
        foreach ($code1 in $positions.Keys) {
            if ($clustered.ContainsKey($code1)) { continue }
            $p1 = $positions[$code1]
            if (-not $p1.Visible) { continue }
            $group = @($code1)
            foreach ($code2 in $positions.Keys) {
                if ($code2 -eq $code1 -or $clustered.ContainsKey($code2)) { continue }
                $p2 = $positions[$code2]
                if (-not $p2.Visible) { continue }
                $dist = [Math]::Sqrt(($p1.X - $p2.X) * ($p1.X - $p2.X) + ($p1.Y - $p2.Y) * ($p1.Y - $p2.Y))
                if ($dist -lt $clusterThreshold) {
                    $group += $code2
                    $clustered[$code2] = $true
                }
            }
            if ($group.Count -gt 1) {
                $clustered[$code1] = $true
                $clusterRows = @($group | ForEach-Object { $positions[$_].Row })
                $clusters.Add(@{ Codes = $group; Rows = $clusterRows })
            }
        }

        # Render non-clustered markers
        foreach ($code in $positions.Keys) {
            if ($clustered.ContainsKey($code)) { continue }
            $p = $positions[$code]
            if (-not $p.Visible) { continue }
            Add-MapMarker -Row $p.Row -CanvasWidth $w -CanvasHeight $h
        }

        # Render clusters
        foreach ($cluster in $clusters) {
            $clusterKey = ($cluster.Codes | Sort-Object) -join '|'
            if ($script:expandedCluster -eq $clusterKey) {
                # Expanded: show individual markers in a fan layout
                $count = $cluster.Codes.Count
                $fanRadius = 25
                for ($i = 0; $i -lt $count; $i++) {
                    $angle = (2 * [Math]::PI / $count) * $i - [Math]::PI / 2
                    $ox = [Math]::Cos($angle) * $fanRadius
                    $oy = [Math]::Sin($angle) * $fanRadius
                    $row = $cluster.Rows[$i]
                    Add-MapMarker -Row $row -CanvasWidth $w -CanvasHeight $h -OffsetX $ox -OffsetY $oy -IsExpanded $true
                }
                # Draw a subtle connecting line from center to each expanded marker
                $centerLat = ($cluster.Rows | Measure-Object -Property Latitude -Average).Average
                $centerLon = ($cluster.Rows | Measure-Object -Property Longitude -Average).Average
                $centerPos = Convert-LatLonToCanvas -Lat $centerLat -Lon $centerLon -Width $w -Height $h
                if ($centerPos.Visible) {
                    for ($i = 0; $i -lt $count; $i++) {
                        $angle = (2 * [Math]::PI / $count) * $i - [Math]::PI / 2
                        $ox = [Math]::Cos($angle) * $fanRadius
                        $oy = [Math]::Sin($angle) * $fanRadius
                        $cLine = New-Object System.Windows.Shapes.Line
                        $cLine.X1 = $centerPos.X; $cLine.Y1 = $centerPos.Y
                        $cLine.X2 = $centerPos.X + $ox; $cLine.Y2 = $centerPos.Y + $oy
                        $cLine.Stroke = $converter.ConvertFromString('#44FFFFFF')
                        $cLine.StrokeThickness = 0.5; $cLine.IsHitTestVisible = $false
                        $canvasMap.Children.Add($cLine) | Out-Null
                    }
                }
            }
            else {
                # Collapsed: show cluster dot
                Add-ClusterMarker -Cluster $cluster -CanvasWidth $w -CanvasHeight $h
            }
        }
    }

    # Continent polygon data — loads Natural Earth-sourced JSON if available, falls back to inline data
    # Source: https://www.naturalearthdata.com/ (Public Domain) — generated by tools/Convert-NaturalEarthToPS.ps1
    function Get-ContinentData {
        # Try to load high-resolution Natural Earth data from JSON
        $jsonPath = Join-Path $PSScriptRoot '..\Data\continent-coords.json'
        if (Test-Path $jsonPath) {
            try {
                $jsonData = Get-Content $jsonPath -Raw | ConvertFrom-Json
                $regions = if ($jsonData.regions) { $jsonData.regions } else { $jsonData }
                $result = foreach ($r in $regions) {
                    $coords = foreach ($c in $r.Coords) {
                        , @([double]$c[0], [double]$c[1])
                    }
                    @{ Name = $r.Name; Coords = $coords }
                }
                return $result
            }
            catch {
                Write-Verbose "Failed to load continent data from JSON, using fallback: $_"
            }
        }

        # Fallback: original hand-crafted polygon data (~383 points, 12 regions)
        return @(
            @{ Name = 'NorthAmerica'; Coords = @(
                    @(49, -125), @(55, -130), @(60, -140), @(64, -140), @(68, -165), @(72, -160), @(72, -140),
                    @(70, -130), @(75, -120), @(78, -100), @(78, -75), @(75, -60), @(70, -55), @(65, -60),
                    @(60, -65), @(55, -60), @(52, -56), @(47, -53), @(45, -60), @(43, -66), @(42, -70),
                    @(40, -74), @(35, -75), @(30, -82), @(25, -80), @(25, -82), @(20, -87), @(15, -83),
                    @(10, -84), @(8, -77), @(10, -74), @(12, -72), @(15, -75), @(18, -88), @(20, -90),
                    @(22, -97), @(25, -100), @(28, -96), @(30, -94), @(30, -90), @(29, -89), @(30, -88),
                    @(30, -85), @(32, -81), @(35, -77), @(37, -76), @(39, -74), @(41, -72), @(42, -71),
                    @(43, -70), @(45, -67), @(47, -68), @(48, -65), @(49, -67), @(48, -69), @(46, -72),
                    @(46, -76), @(44, -79), @(42, -83), @(43, -87), @(45, -85), @(46, -84), @(48, -88),
                    @(49, -95), @(49, -105), @(49, -115), @(49, -125)
                )
            },
            @{ Name = 'SouthAmerica'; Coords = @(
                    @(12, -72), @(10, -74), @(8, -77), @(5, -77), @(2, -80), @(-2, -80), @(-5, -81),
                    @(-6, -77), @(-4, -70), @(-2, -50), @(-3, -42), @(-5, -35), @(-8, -35), @(-13, -39),
                    @(-18, -40), @(-23, -42), @(-28, -49), @(-33, -52), @(-38, -57), @(-42, -62),
                    @(-46, -66), @(-50, -70), @(-52, -70), @(-55, -68), @(-56, -70), @(-55, -73),
                    @(-50, -75), @(-46, -76), @(-42, -73), @(-38, -73), @(-35, -72), @(-30, -71),
                    @(-27, -70), @(-22, -70), @(-18, -70), @(-15, -75), @(-10, -78), @(-5, -80),
                    @(-2, -80), @(2, -78), @(5, -77), @(8, -72), @(10, -72), @(12, -72)
                )
            },
            @{ Name = 'Europe'; Coords = @(
                    @(36, -6), @(37, -9), @(39, -9), @(42, -9), @(43, -8), @(44, -2), @(46, -2),
                    @(47, 2), @(49, 0), @(51, 2), @(53, 5), @(54, 8), @(56, 8), @(55, 12), @(57, 10),
                    @(58, 12), @(60, 11), @(63, 10), @(65, 12), @(68, 15), @(70, 20), @(71, 26),
                    @(70, 28), @(69, 31), @(67, 26), @(60, 28), @(58, 23), @(55, 22), @(54, 18),
                    @(54, 14), @(53, 14), @(52, 10), @(50, 6), @(48, 8), @(47, 7), @(46, 14),
                    @(45, 14), @(44, 13), @(42, 15), @(40, 18), @(38, 24), @(36, 28), @(35, 25),
                    @(38, 20), @(39, 20), @(41, 18), @(40, 15), @(38, 13), @(37, 15), @(36, 14),
                    @(36, 12), @(38, 10), @(39, 3), @(38, 0), @(37, -2), @(36, -5), @(36, -6)
                )
            },
            @{ Name = 'Africa'; Coords = @(
                    @(37, -6), @(36, -5), @(36, 0), @(35, 10), @(33, 10), @(30, 10), @(32, 32),
                    @(30, 33), @(22, 36), @(15, 42), @(12, 44), @(12, 50), @(5, 42), @(0, 42),
                    @(-5, 39), @(-10, 40), @(-15, 40), @(-18, 36), @(-22, 35), @(-25, 33),
                    @(-30, 31), @(-34, 26), @(-34, 18), @(-30, 16), @(-25, 15), @(-18, 12),
                    @(-12, 14), @(-6, 12), @(0, 10), @(4, 10), @(5, 1), @(4, -2), @(5, -5),
                    @(4, -7), @(5, -10), @(10, -15), @(15, -17), @(20, -17), @(22, -16),
                    @(25, -14), @(28, -10), @(30, -10), @(32, -5), @(35, -2), @(37, -6)
                )
            },
            @{ Name = 'Asia'; Coords = @(
                    @(70, 28), @(72, 40), @(73, 55), @(75, 70), @(77, 100), @(75, 110),
                    @(73, 140), @(70, 170), @(67, 170), @(65, 165), @(62, 160), @(60, 150),
                    @(57, 140), @(55, 135), @(52, 140), @(48, 135), @(44, 132), @(40, 130),
                    @(38, 128), @(36, 127), @(35, 129), @(33, 126), @(30, 122), @(25, 120),
                    @(22, 114), @(22, 108), @(18, 108), @(12, 105), @(8, 105), @(5, 103),
                    @(2, 104), @(-2, 106), @(-7, 106), @(-8, 115), @(-6, 120), @(-7, 130),
                    @(-8, 140), @(-5, 142), @(-6, 148), @(-10, 148), @(-8, 140), @(-8, 130),
                    @(-8, 120), @(-8, 115), @(-10, 110), @(-8, 106), @(-5, 100), @(-2, 98),
                    @(5, 98), @(8, 100), @(12, 100), @(13, 93), @(8, 80), @(5, 80),
                    @(8, 77), @(15, 75), @(20, 73), @(25, 68), @(28, 65), @(25, 60),
                    @(28, 51), @(30, 48), @(32, 36), @(35, 36), @(38, 35), @(40, 28),
                    @(42, 28), @(45, 30), @(47, 35), @(50, 40), @(53, 40), @(55, 40),
                    @(58, 50), @(60, 55), @(60, 60), @(58, 70), @(55, 72), @(52, 55),
                    @(50, 50), @(50, 45), @(52, 40), @(55, 40), @(58, 50), @(60, 55),
                    @(60, 28), @(67, 26), @(69, 31), @(70, 28)
                )
            },
            @{ Name = 'India'; Coords = @(
                    @(28, 68), @(30, 74), @(30, 78), @(28, 85), @(26, 89), @(22, 90),
                    @(22, 88), @(18, 84), @(15, 80), @(10, 80), @(8, 77), @(10, 76),
                    @(13, 75), @(15, 74), @(20, 73), @(22, 69), @(25, 68), @(28, 68)
                )
            },
            @{ Name = 'Australia'; Coords = @(
                    @(-12, 130), @(-12, 136), @(-14, 136), @(-14, 141), @(-18, 146),
                    @(-24, 152), @(-28, 154), @(-33, 152), @(-37, 150), @(-39, 146),
                    @(-39, 144), @(-37, 140), @(-35, 137), @(-35, 135), @(-32, 132),
                    @(-32, 128), @(-34, 122), @(-34, 116), @(-32, 115), @(-28, 114),
                    @(-24, 114), @(-22, 114), @(-18, 122), @(-15, 129), @(-12, 130)
                )
            },
            @{ Name = 'Japan'; Coords = @(
                    @(45, 142), @(43, 145), @(40, 140), @(38, 140), @(36, 140),
                    @(34, 135), @(33, 131), @(34, 130), @(35, 133), @(36, 136),
                    @(37, 137), @(39, 140), @(41, 140), @(43, 143), @(45, 142)
                )
            },
            @{ Name = 'UK'; Coords = @(
                    @(50, -6), @(51, -5), @(52, -4), @(53, -3), @(54, -3), @(55, -2),
                    @(56, -3), @(57, -2), @(58, -3), @(59, -3), @(58, -5), @(57, -6),
                    @(56, -5), @(55, -5), @(54, -5), @(53, -4), @(52, -5), @(51, -5), @(50, -6)
                )
            },
            @{ Name = 'Greenland'; Coords = @(
                    @(60, -45), @(63, -42), @(68, -30), @(72, -22), @(76, -20), @(78, -18),
                    @(80, -25), @(82, -35), @(82, -50), @(80, -60), @(78, -68), @(76, -70),
                    @(72, -55), @(68, -52), @(65, -53), @(62, -50), @(60, -45)
                )
            },
            @{ Name = 'Iceland'; Coords = @(
                    @(64, -22), @(65, -18), @(66, -16), @(66, -14), @(65, -14),
                    @(64, -16), @(63, -18), @(63, -22), @(64, -22)
                )
            },
            @{ Name = 'NewZealand'; Coords = @(
                    @(-35, 174), @(-37, 176), @(-39, 177), @(-41, 176), @(-42, 174),
                    @(-44, 170), @(-46, 167), @(-47, 167), @(-46, 169), @(-44, 172),
                    @(-42, 173), @(-40, 176), @(-38, 176), @(-36, 175), @(-35, 174)
                )
            }
        )
    }

    function Draw-FlatContinents {
        param([double]$Width, [double]$Height, $Fill, $Stroke)
        foreach ($cont in Get-ContinentData) {
            $poly = New-Object System.Windows.Shapes.Polygon
            $poly.Fill = $Fill; $poly.Stroke = $Stroke
            $poly.StrokeThickness = 1; $poly.Opacity = 0.85
            $poly.IsHitTestVisible = $false
            $points = New-Object System.Windows.Media.PointCollection
            foreach ($c in $cont.Coords) {
                $px = (($c[1] + 180) / 360) * $Width
                $py = ((90 - $c[0]) / 180) * $Height
                $points.Add([System.Windows.Point]::new($px, $py))
            }
            $poly.Points = $points
            $canvasMap.Children.Add($poly) | Out-Null
        }
    }

    function Draw-GlobeContinents {
        param([double]$Width, [double]$Height)
        $converter = [System.Windows.Media.BrushConverter]::new()
        $landFill = $converter.ConvertFromString('#1B3A2A')
        $landStroke = $converter.ConvertFromString('#2D6B45')
        foreach ($cont in Get-ContinentData) {
            $poly = New-Object System.Windows.Shapes.Polygon
            $poly.Fill = $landFill; $poly.Stroke = $landStroke
            $poly.StrokeThickness = 1; $poly.Opacity = 0.85
            $poly.IsHitTestVisible = $false
            $points = New-Object System.Windows.Media.PointCollection
            $anyVisible = $false
            foreach ($c in $cont.Coords) {
                $p = Convert-LatLonToGlobe -Lat $c[0] -Lon $c[1] -Width $Width -Height $Height
                if ($p.Visible) {
                    $points.Add([System.Windows.Point]::new($p.X, $p.Y))
                    $anyVisible = $true
                }
            }
            if ($anyVisible -and $points.Count -ge 3) {
                $poly.Points = $points
                $canvasMap.Children.Add($poly) | Out-Null
            }
        }
    }

    # ── Exchange Detail Flyout ────────────────────────────────────

    function Show-ExchangeFlyout {
        param([string]$ExchangeCode)

        # Toggle off if same exchange clicked
        if ($script:currentFlyoutCode -eq $ExchangeCode -and $flyoutPanel.Visibility -eq 'Visible') {
            $flyoutPanel.Visibility = 'Collapsed'
            $script:currentFlyoutCode = $null
            $marketDataPanel.Visibility = 'Visible'
            return
        }

        $script:currentFlyoutCode = $ExchangeCode
        $marketDataPanel.Visibility = 'Collapsed'
        $converter = [System.Windows.Media.BrushConverter]::new()

        # Find the exchange row
        $row = $script:exchangeRows | Where-Object { $_.Code -eq $ExchangeCode } | Select-Object -First 1
        if (-not $row) { return }

        # Get status info
        $exObj = [PSCustomObject]@{
            TimeZoneId      = $row.TimeZoneId
            OpenTimeLocal   = $row.OpenTime
            CloseTimeLocal  = $row.CloseTime
            LunchBreakStart = $row.LunchBreakStart
            LunchBreakEnd   = $row.LunchBreakEnd
            Code            = $row.Code
        }
        $statusInfo = Get-ExchangeStatus -Exchange $exObj
        $statusColor = Get-StatusColor -Status $statusInfo.Status

        # Set header
        $flyoutTitle.Text = "$($row.Symbol) - $($row.DisplayName)"

        # Clear and rebuild content
        $flyoutContent.Children.Clear()

        # Helper: add a detail row
        function Add-FlyoutRow {
            param([string]$Label, [string]$Value, [string]$ValueColor = '#E0E0E0', [bool]$IsLink = $false, [string]$LinkUrl = '')

            $panel = New-Object System.Windows.Controls.StackPanel
            $panel.Margin = [System.Windows.Thickness]::new(0, 0, 0, 6)

            $lblKey = New-Object System.Windows.Controls.TextBlock
            $lblKey.Text = $Label
            $lblKey.FontSize = 10
            $lblKey.Foreground = $converter.ConvertFromString('#888888')
            $panel.Children.Add($lblKey) | Out-Null

            if ($IsLink -and $LinkUrl) {
                $lblVal = New-Object System.Windows.Controls.TextBlock
                $lblVal.FontSize = 12
                $lblVal.Cursor = [System.Windows.Input.Cursors]::Hand
                $lblVal.Foreground = $converter.ConvertFromString('#4488FF')
                $lblVal.TextDecorations = [System.Windows.TextDecorations]::Underline
                $lblVal.Text = $Value
                $url = $LinkUrl
                $lblVal.Add_MouseLeftButtonDown({
                        try { Start-Process $url } catch { }
                    }.GetNewClosure())
                $panel.Children.Add($lblVal) | Out-Null
            }
            else {
                $lblVal = New-Object System.Windows.Controls.TextBlock
                $lblVal.Text = $Value
                $lblVal.FontSize = 12
                $lblVal.Foreground = $converter.ConvertFromString($ValueColor)
                $lblVal.TextWrapping = 'Wrap'
                $panel.Children.Add($lblVal) | Out-Null
            }

            $flyoutContent.Children.Add($panel) | Out-Null
        }

        # Status and countdown
        Add-FlyoutRow -Label 'Status' -Value $statusInfo.Text -ValueColor $statusColor
        Add-FlyoutRow -Label 'Countdown' -Value $row.CountdownText -ValueColor $statusColor
        Add-FlyoutRow -Label 'Local Time' -Value $row.CurrentLocalTime
        Add-FlyoutRow -Label 'Hours' -Value "$($row.OpenTime) – $($row.CloseTime)"
        Add-FlyoutRow -Label 'Location' -Value "$($row.City), $($row.Country)"

        # Divider
        $divider = New-Object System.Windows.Shapes.Rectangle
        $divider.Height = 1
        $divider.Fill = $converter.ConvertFromString('#0F3460')
        $divider.Margin = [System.Windows.Thickness]::new(0, 4, 0, 8)
        $flyoutContent.Children.Add($divider) | Out-Null

        # Exchange details from JSON
        if ($script:exchangeDetails.ContainsKey($ExchangeCode)) {
            $detail = $script:exchangeDetails[$ExchangeCode]
            Add-FlyoutRow -Label 'Primary Index' -Value $detail.PrimaryIndex
            Add-FlyoutRow -Label 'Market Cap' -Value $detail.MarketCapUSD -ValueColor '#00CC66'
            Add-FlyoutRow -Label 'Listed Companies' -Value "$($detail.ListedCompanies)"
            Add-FlyoutRow -Label 'Currency' -Value "$($detail.CurrencySymbol) $($detail.Currency)"
            Add-FlyoutRow -Label 'Founded' -Value "$($detail.Founded)"
            Add-FlyoutRow -Label 'Regulator' -Value $detail.Regulator
            Add-FlyoutRow -Label 'Website' -Value $detail.Website -IsLink $true -LinkUrl $detail.Website

            # Description divider + text
            $div2 = New-Object System.Windows.Shapes.Rectangle
            $div2.Height = 1
            $div2.Fill = $converter.ConvertFromString('#0F3460')
            $div2.Margin = [System.Windows.Thickness]::new(0, 4, 0, 8)
            $flyoutContent.Children.Add($div2) | Out-Null

            $descBlock = New-Object System.Windows.Controls.TextBlock
            $descBlock.Text = $detail.Description
            $descBlock.FontSize = 11
            $descBlock.Foreground = $converter.ConvertFromString('#AAAAAA')
            $descBlock.TextWrapping = 'Wrap'
            $flyoutContent.Children.Add($descBlock) | Out-Null
        }
        else {
            Add-FlyoutRow -Label 'Details' -Value 'No additional data available' -ValueColor '#888888'
        }

        $flyoutPanel.Visibility = 'Visible'
    }

    # ── Countdown Card Creation ──────────────────────────────────

    $script:cardElements = @{}

    function New-CountdownCard {
        param([PSCustomObject]$Row)

        $converter = [System.Windows.Media.BrushConverter]::new()

        $card = New-Object System.Windows.Controls.Border
        $card.Width = 220
        $card.Margin = [System.Windows.Thickness]::new(5)
        $card.Padding = [System.Windows.Thickness]::new(10)
        $card.CornerRadius = [System.Windows.CornerRadius]::new(6)
        $card.Background = $converter.ConvertFromString('#16213E')
        $card.BorderBrush = $converter.ConvertFromString('#0F3460')
        $card.BorderThickness = [System.Windows.Thickness]::new(1)

        $stack = New-Object System.Windows.Controls.StackPanel

        # Exchange name
        $lblName = New-Object System.Windows.Controls.TextBlock
        $lblName.Text = "$($Row.Symbol) - $($Row.City)"
        $lblName.FontSize = 13
        $lblName.FontWeight = [System.Windows.FontWeights]::Bold
        $lblName.Foreground = $converter.ConvertFromString('#E0E0E0')
        $lblName.Margin = [System.Windows.Thickness]::new(0, 0, 0, 4)
        $stack.Children.Add($lblName) | Out-Null

        # Full name
        $lblFull = New-Object System.Windows.Controls.TextBlock
        $lblFull.Text = $Row.DisplayName
        $lblFull.FontSize = 10
        $lblFull.Foreground = $converter.ConvertFromString('#888888')
        $lblFull.TextTrimming = 'CharacterEllipsis'
        $lblFull.Margin = [System.Windows.Thickness]::new(0, 0, 0, 6)
        $stack.Children.Add($lblFull) | Out-Null

        # Countdown text
        $lblCountdown = New-Object System.Windows.Controls.TextBlock
        $lblCountdown.Text = '--:--:--'
        $lblCountdown.FontSize = 24
        $lblCountdown.FontFamily = [System.Windows.Media.FontFamily]::new('Consolas')
        $lblCountdown.FontWeight = [System.Windows.FontWeights]::Bold
        $lblCountdown.Foreground = $converter.ConvertFromString('#00CC66')
        $lblCountdown.HorizontalAlignment = 'Center'
        $lblCountdown.Margin = [System.Windows.Thickness]::new(0, 2, 0, 4)
        $stack.Children.Add($lblCountdown) | Out-Null

        # Status text
        $lblStatus = New-Object System.Windows.Controls.TextBlock
        $lblStatus.Text = 'Loading...'
        $lblStatus.FontSize = 11
        $lblStatus.Foreground = $converter.ConvertFromString('#AAAAAA')
        $lblStatus.HorizontalAlignment = 'Center'
        $stack.Children.Add($lblStatus) | Out-Null

        # Hours line
        $lblHours = New-Object System.Windows.Controls.TextBlock
        $lblHours.Text = "Hours: $($Row.OpenTime) - $($Row.CloseTime)"
        $lblHours.FontSize = 10
        $lblHours.Foreground = $converter.ConvertFromString('#666666')
        $lblHours.HorizontalAlignment = 'Center'
        $lblHours.Margin = [System.Windows.Thickness]::new(0, 4, 0, 0)
        $stack.Children.Add($lblHours) | Out-Null

        $card.Child = $stack
        $card.Cursor = [System.Windows.Input.Cursors]::Hand

        # Click handler for flyout — use Tag to avoid closure scope issues
        $card.Tag = $Row.Code
        $card.Add_MouseLeftButtonDown({
                param($sender, $e)
                Show-ExchangeFlyout -ExchangeCode $sender.Tag
            })

        $script:cardElements[$Row.Code] = @{
            Card      = $card
            Countdown = $lblCountdown
            Status    = $lblStatus
            Border    = $card
        }

        return $card
    }

    function Rebuild-Cards {
        $panelCards.Children.Clear()
        $script:cardElements = @{}
        foreach ($row in $script:exchangeRows) {
            if ($row.IsActive) {
                $card = New-CountdownCard -Row $row
                $panelCards.Children.Add($card) | Out-Null
            }
        }
    }

    # ── Notification System ───────────────────────────────────────

    $script:notifyIcon = New-Object System.Windows.Forms.NotifyIcon
    $script:notifyIcon.Icon = [System.Drawing.SystemIcons]::Information
    $script:notifyIcon.Text = 'Stock Exchange Countdown'
    $script:notifyIcon.Visible = $true

    # Context menu
    $contextMenu = New-Object System.Windows.Forms.ContextMenuStrip
    $showItem = New-Object System.Windows.Forms.ToolStripMenuItem
    $showItem.Text = 'Show Dashboard'
    $showItem.Add_Click({ $window.WindowState = 'Normal'; $window.Activate() })
    $contextMenu.Items.Add($showItem) | Out-Null

    $separator = New-Object System.Windows.Forms.ToolStripSeparator
    $contextMenu.Items.Add($separator) | Out-Null

    $exitItem = New-Object System.Windows.Forms.ToolStripMenuItem
    $exitItem.Text = 'Exit'
    $exitItem.Add_Click({
            $script:forceExit = $true
            $script:notifyIcon.Visible = $false
            $script:notifyIcon.Dispose()
            $window.Close()
        })
    $contextMenu.Items.Add($exitItem) | Out-Null
    $script:notifyIcon.ContextMenuStrip = $contextMenu

    $script:notifyIcon.Add_MouseDoubleClick({
            $window.WindowState = 'Normal'
            $window.Activate()
        })

    # Exit button in header bar
    $btnExit.Add_Click({
            $result = [System.Windows.MessageBox]::Show('Exit the dashboard?', 'Confirm Exit', 'YesNo', 'Question')
            if ($result -eq 'Yes') {
                $script:forceExit = $true
                $script:notifyIcon.Visible = $false
                $script:notifyIcon.Dispose()
                $window.Close()
            }
        })

    function Send-Notification {
        param([string]$Title, [string]$Message)
        try {
            $script:notifyIcon.ShowBalloonTip(5000, $Title, $Message, [System.Windows.Forms.ToolTipIcon]::Information)
        }
        catch {
            Write-Verbose "Notification failed: $($_.Exception.Message)"
        }
    }

    function Check-Notifications {
        param([PSCustomObject]$Row, [hashtable]$StatusInfo)

        if ($StatusInfo.Status -notin @('Open', 'ClosingSoon', 'ClosingImminent')) { return }
        if (-not $StatusInfo.TimeRemaining) { return }

        $key30 = "$($Row.Code)_30_$((Get-Date).ToString('yyyy-MM-dd'))"
        $key15 = "$($Row.Code)_15_$((Get-Date).ToString('yyyy-MM-dd'))"
        $key5 = "$($Row.Code)_5_$((Get-Date).ToString('yyyy-MM-dd'))"

        $mins = $StatusInfo.TimeRemaining.TotalMinutes

        if ($chkNotify30.IsChecked -and $mins -le 30 -and $mins -gt 15 -and -not $script:notificationsFired.ContainsKey($key30)) {
            Send-Notification -Title "$($Row.Symbol) Closing Soon" -Message "$($Row.DisplayName) closes in 30 minutes"
            $script:notificationsFired[$key30] = $true
        }
        if ($chkNotify15.IsChecked -and $mins -le 15 -and $mins -gt 5 -and -not $script:notificationsFired.ContainsKey($key15)) {
            Send-Notification -Title "$($Row.Symbol) Closing" -Message "$($Row.DisplayName) closes in 15 minutes"
            $script:notificationsFired[$key15] = $true
        }
        if ($chkNotify5.IsChecked -and $mins -le 5 -and -not $script:notificationsFired.ContainsKey($key5)) {
            Send-Notification -Title "$($Row.Symbol) CLOSING IMMINENT" -Message "$($Row.DisplayName) closes in 5 minutes!"
            $script:notificationsFired[$key5] = $true
        }
    }

    # ── Timer Tick Handler ────────────────────────────────────────

    function Update-AllDisplays {
        $converter = [System.Windows.Media.BrushConverter]::new()

        # Update header UTC clock
        $txtHeaderClock.Text = "UTC: $([DateTime]::UtcNow.ToString('HH:mm:ss'))"

        # Update world clocks
        foreach ($key in $script:clockLabels.Keys) {
            $info = $script:clockLabels[$key]
            try {
                $tz = [System.TimeZoneInfo]::FindSystemTimeZoneById($info.TzId)
                $localTime = [System.TimeZoneInfo]::ConvertTimeFromUtc([DateTime]::UtcNow, $tz)
                $info.TimeLabel.Text = "$($localTime.ToString('HH:mm:ss')) $($info.Abbrev)"
            }
            catch {
                $info.TimeLabel.Text = 'Error'
            }
        }

        # Track next close for status bar
        $nextCloseExchange = $null
        $nextCloseRemaining = $null

        foreach ($row in $script:exchangeRows) {
            # Build a temporary exchange object for the status function
            $exObj = [PSCustomObject]@{
                TimeZoneId      = $row.TimeZoneId
                OpenTimeLocal   = $row.OpenTime
                CloseTimeLocal  = $row.CloseTime
                LunchBreakStart = $row.LunchBreakStart
                LunchBreakEnd   = $row.LunchBreakEnd
                Code            = $row.Code
            }

            $statusInfo = Get-ExchangeStatus -Exchange $exObj
            $color = Get-StatusColor -Status $statusInfo.Status

            # Update DataGrid row properties
            $row.StatusText = $statusInfo.Text
            $row.StatusColor = $color

            if ($statusInfo.TimeRemaining) {
                $row.CountdownText = "{0:D2}:{1:D2}:{2:D2}" -f [int]$statusInfo.TimeRemaining.TotalHours, $statusInfo.TimeRemaining.Minutes, $statusInfo.TimeRemaining.Seconds
            }
            else {
                $row.CountdownText = '--:--:--'
            }

            # Update local time
            try {
                $localNow = Get-ExchangeLocalTime -Exchange $exObj
                $row.CurrentLocalTime = $localNow.ToString('HH:mm:ss')
            }
            catch {
                $row.CurrentLocalTime = '--:--:--'
            }

            # Track next close
            if ($row.IsActive -and $statusInfo.CountdownTarget -eq 'Close' -and $statusInfo.TimeRemaining) {
                if (-not $nextCloseRemaining -or $statusInfo.TimeRemaining -lt $nextCloseRemaining) {
                    $nextCloseExchange = $row.Symbol
                    $nextCloseRemaining = $statusInfo.TimeRemaining
                }
            }

            # Update card if active
            if ($row.IsActive -and $script:cardElements.ContainsKey($row.Code)) {
                $ce = $script:cardElements[$row.Code]
                $ce.Countdown.Text = $row.CountdownText
                $ce.Countdown.Foreground = $converter.ConvertFromString($color)
                $ce.Status.Text = $statusInfo.Text
                $ce.Status.Foreground = $converter.ConvertFromString($color)
                $ce.Border.BorderBrush = $converter.ConvertFromString($color)
            }

            # Update map marker
            if ($script:mapMarkers -and $script:mapMarkers.ContainsKey($row.Code)) {
                $mm = $script:mapMarkers[$row.Code]
                $mm.Marker.Fill = $converter.ConvertFromString($color)
                $mm.Glow.Fill = $converter.ConvertFromString(($color -replace '^#', '#33'))
                $mm.Tooltip.Content = "$($row.DisplayName) ($($row.Symbol))`n$($statusInfo.Text)"
            }

            # Check notifications
            if ($row.IsActive) {
                Check-Notifications -Row $row -StatusInfo $statusInfo
            }
        }

        # Update cluster markers
        if ($script:mapClusters) {
            foreach ($key in @($script:mapClusters.Keys)) {
                $cl = $script:mapClusters[$key]
                $clusterColor = Get-ClusterPriorityColor -Codes $cl.Codes
                $cl.Dot.Fill = $converter.ConvertFromString($clusterColor)
                $cl.Glow.Fill = $converter.ConvertFromString(($clusterColor -replace '^#', '#33'))
                # Update cluster tooltip with current status
                $names = @()
                foreach ($code in $cl.Codes) {
                    $r = $script:exchangeRows | Where-Object { $_.Code -eq $code } | Select-Object -First 1
                    if ($r) { $names += "$($r.DisplayName) ($($r.Symbol)) — $($r.StatusText)" }
                }
                $cl.Tooltip.Content = $names -join "`n"
            }
        }

        # Update status bar
        if ($nextCloseExchange -and $nextCloseRemaining) {
            $ncs = "{0:D2}:{1:D2}:{2:D2}" -f [int]$nextCloseRemaining.TotalHours, $nextCloseRemaining.Minutes, $nextCloseRemaining.Seconds
            $txtNextClose.Text = "Next close: $nextCloseExchange in $ncs"
            $window.Title = "PSExchangeClock — Next: $nextCloseExchange $ncs"
        }
        else {
            $txtNextClose.Text = 'All markets closed'
            $window.Title = 'PSExchangeClock — Global Market Dashboard'
        }

        $txtStatus.Text = "Active: $(@($script:exchangeRows | Where-Object { $_.IsActive }).Count) exchanges | Updated: $([DateTime]::Now.ToString('HH:mm:ss'))"

        # Refresh DataGrid display
        $gridExchanges.Items.Refresh()
    }

    # ── Event Handlers ────────────────────────────────────────────

    # Flyout close button — show market data panel when flyout closes
    $btnCloseFlyout.Add_Click({
            $flyoutPanel.Visibility = 'Collapsed'
            $script:currentFlyoutCode = $null
            $marketDataPanel.Visibility = 'Visible'
        })

    # Market Data tab button handlers
    $btnTabNews.Add_Click({ Set-ActiveMarketTab -TabName 'News' })
    $btnTabFx.Add_Click({ Set-ActiveMarketTab -TabName 'FX' })
    $btnTabCrypto.Add_Click({ Set-ActiveMarketTab -TabName 'Crypto' })
    $btnTabIndices.Add_Click({ Set-ActiveMarketTab -TabName 'Indices' })
    $btnTabCommodities.Add_Click({ Set-ActiveMarketTab -TabName 'Commodities' })
    $btnTabStocks.Add_Click({ Set-ActiveMarketTab -TabName 'Stocks' })

    # API key save button
    $btnSaveApiKeys.Add_Click({
            # Store API keys in the secure backend instead of JSON
            $tdKeyVal = $txtTwelveDataKey.Text.Trim()
            $avKeyVal = $txtAlphaVantageKey.Text.Trim()
            if (-not [string]::IsNullOrWhiteSpace($tdKeyVal)) {
                Set-SecureApiKey -ServiceName 'TwelveData' -ApiKey $tdKeyVal
            }
            if (-not [string]::IsNullOrWhiteSpace($avKeyVal)) {
                Set-SecureApiKey -ServiceName 'AlphaVantage' -ApiKey $avKeyVal
            }
            Save-UserPreferences
            Update-ApiKeyTabVisibility
            $txtStatus.Text = "API keys saved to $($script:secretsBackend)"
            # Refresh market data content if on an API-dependent tab
            if ($script:activeMarketTab -in @('Indices', 'Commodities', 'Stocks')) {
                Update-MarketDataContent
            }
        })

    # Helper: update tab visibility based on API keys
    function Update-ApiKeyTabVisibility {
        $converter = [System.Windows.Media.BrushConverter]::new()
        if (-not [string]::IsNullOrWhiteSpace($txtTwelveDataKey.Text.Trim())) {
            $btnTabIndices.Visibility = 'Visible'
            $btnTabCommodities.Visibility = 'Visible'
            $txtTwelveDataStatus.Text = [char]0x2714  # checkmark
            $txtTwelveDataStatus.Foreground = $converter.ConvertFromString('#00CC66')
        }
        else {
            $btnTabIndices.Visibility = 'Collapsed'
            $btnTabCommodities.Visibility = 'Collapsed'
            $txtTwelveDataStatus.Text = ''
            if ($script:activeMarketTab -in @('Indices', 'Commodities')) { Set-ActiveMarketTab -TabName 'News' }
        }
        if (-not [string]::IsNullOrWhiteSpace($txtAlphaVantageKey.Text.Trim())) {
            $btnTabStocks.Visibility = 'Visible'
            $txtAlphaVantageStatus.Text = [char]0x2714
            $txtAlphaVantageStatus.Foreground = $converter.ConvertFromString('#00CC66')
        }
        else {
            $btnTabStocks.Visibility = 'Collapsed'
            $txtAlphaVantageStatus.Text = ''
            if ($script:activeMarketTab -eq 'Stocks') { Set-ActiveMarketTab -TabName 'News' }
        }
    }

    # Select All / Deselect All buttons
    $btnSelectAll.Add_Click({
            $currentSource = $gridExchanges.ItemsSource
            foreach ($row in $currentSource) {
                $row.IsActive = $true
            }
            $gridExchanges.Items.Refresh()
            Rebuild-Cards
            Rebuild-WorldClocks
            Initialize-WorldMap
            Save-UserPreferences
            $txtStatus.Text = "All visible exchanges selected"
        })

    $btnDeselectAll.Add_Click({
            $currentSource = $gridExchanges.ItemsSource
            foreach ($row in $currentSource) {
                $row.IsActive = $false
            }
            $gridExchanges.Items.Refresh()
            Rebuild-Cards
            Rebuild-WorldClocks
            Initialize-WorldMap
            Save-UserPreferences
            $txtStatus.Text = "All visible exchanges deselected"
        })

    # DataGrid cell edit — instant sync when Active toggled
    $gridExchanges.Add_CellEditEnding({
            param($cellSender, $cellArgs)
            if ($cellArgs.Column.Header -eq 'Active') {
                $window.Dispatcher.BeginInvoke([System.Windows.Threading.DispatcherPriority]::Background, [Action] {
                        Rebuild-Cards
                        Rebuild-WorldClocks
                        Initialize-WorldMap
                        Save-UserPreferences
                    })
            }
        })

    # Grid lines toggle
    $chkShowGridLines.Add_Checked({ Initialize-WorldMap })
    $chkShowGridLines.Add_Unchecked({ Initialize-WorldMap })

    # ── Map Toolbar Event Handlers ────────────────────────────────

    # Map style ComboBox
    $cmbMapStyle.Add_SelectionChanged({ Initialize-WorldMap })

    # Map projection ComboBox
    $cmbMapProjection.Add_SelectionChanged({
            $sel = $cmbMapProjection.SelectedItem
            if ($sel -and $sel.Tag) {
                $script:mapProjection = $sel.Tag
                $script:expandedCluster = $null
                Initialize-WorldMap
            }
        })

    # Overlay toggles
    $chkTerminator.Add_Checked({ Initialize-WorldMap })
    $chkTerminator.Add_Unchecked({ Initialize-WorldMap })
    $chkTimeZoneBands.Add_Checked({ Initialize-WorldMap })
    $chkTimeZoneBands.Add_Unchecked({ Initialize-WorldMap })
    $chkPoliticalBounds.Add_Checked({ Initialize-WorldMap })
    $chkPoliticalBounds.Add_Unchecked({ Initialize-WorldMap })
    $chkEarthquakes.Add_Checked({ Initialize-WorldMap })
    $chkEarthquakes.Add_Unchecked({ Initialize-WorldMap })
    $chkVolcanoes.Add_Checked({ Initialize-WorldMap })
    $chkVolcanoes.Add_Unchecked({ Initialize-WorldMap })
    $chkSubmarineCables.Add_Checked({ Initialize-WorldMap })
    $chkSubmarineCables.Add_Unchecked({ Initialize-WorldMap })
    $chkPowerPlants.Add_Checked({ Initialize-WorldMap })
    $chkPowerPlants.Add_Unchecked({ Initialize-WorldMap })
    $chkTZBoundaries.Add_Checked({ Initialize-WorldMap })
    $chkTZBoundaries.Add_Unchecked({ Initialize-WorldMap })
    $chkConflictZones.Add_Checked({ Initialize-WorldMap })
    $chkConflictZones.Add_Unchecked({ Initialize-WorldMap })

    # Overlays flyout popup toggle
    $btnOverlays.Add_Checked({
            $popOverlays.PlacementTarget = $btnOverlays
            $popOverlays.IsOpen = $true
        })
    $btnOverlays.Add_Unchecked({
            $popOverlays.IsOpen = $false
        })
    $popOverlays.Add_Closed({
            $btnOverlays.IsChecked = $false
        })

    # Zoom controls
    $btnZoomIn.Add_Click({
            $script:mapZoom = [Math]::Min($script:mapZoom * 1.4, 8.0)
            Apply-MapZoom
        })

    $btnZoomOut.Add_Click({
            $script:mapZoom = [Math]::Max($script:mapZoom / 1.4, 1.0)
            if ($script:mapZoom -le 1.01) {
                $script:mapZoom = 1.0; $script:mapPanX = 0; $script:mapPanY = 0
            }
            Apply-MapZoom
        })

    $btnZoomReset.Add_Click({
            $script:mapZoom = 1.0; $script:mapPanX = 0; $script:mapPanY = 0
            Apply-MapZoom
        })

    # Globe auto-rotation toggle
    $btnGlobeSpin.Add_Checked({
            $script:globeSpinning = $true
        })
    $btnGlobeSpin.Add_Unchecked({
            $script:globeSpinning = $false
        })

    # Mouse wheel zoom on canvas
    $canvasMap.Add_MouseWheel({
            param($sender, $e)
            $delta = $e.Delta
            $oldZoom = $script:mapZoom
            if ($delta -gt 0) {
                $script:mapZoom = [Math]::Min($script:mapZoom * 1.15, 8.0)
            }
            else {
                $script:mapZoom = [Math]::Max($script:mapZoom / 1.15, 1.0)
            }
            if ($script:mapZoom -le 1.01) {
                $script:mapZoom = 1.0; $script:mapPanX = 0; $script:mapPanY = 0
            }
            # Zoom toward cursor position
            if ($script:mapZoom -gt 1.0 -and $oldZoom -gt 0) {
                $mousePos = $e.GetPosition($canvasMap)
                $scale = $script:mapZoom / $oldZoom
                $script:mapPanX = $mousePos.X - $scale * ($mousePos.X - $script:mapPanX)
                $script:mapPanY = $mousePos.Y - $scale * ($mousePos.Y - $script:mapPanY)
            }
            Apply-MapZoom
            $e.Handled = $true
        })

    # Mouse drag to pan (right button for flat, left button for globe rotation)
    $canvasMap.Add_MouseRightButtonDown({
            param($sender, $e)
            if ($script:mapZoom -gt 1.0) {
                $script:mapDragging = $true
                $script:mapDragStart = $e.GetPosition($canvasMap)
                $script:mapPanStart = @{ X = $script:mapPanX; Y = $script:mapPanY }
                $canvasMap.CaptureMouse()
                $e.Handled = $true
            }
        })

    $canvasMap.Add_MouseMove({
            param($sender, $e)
            if ($script:mapDragging -and $script:mapDragStart) {
                $pos = $e.GetPosition($canvasMap)
                $script:mapPanX = $script:mapPanStart.X + ($pos.X - $script:mapDragStart.X)
                $script:mapPanY = $script:mapPanStart.Y + ($pos.Y - $script:mapDragStart.Y)
                Apply-MapZoom
            }
            if ($script:globeDragging -and $script:mapProjection -eq 'Globe') {
                $pos = $e.GetPosition($canvasMap)
                $dx = $pos.X - $script:globeDragStart.X
                $dy = $pos.Y - $script:globeDragStart.Y
                $script:globeCenterLon = $script:globeDragStartLon - $dx * 0.3
                $script:globeCenterLat = [Math]::Max(-80, [Math]::Min(80, $script:globeDragStartLat + $dy * 0.3))
                Initialize-WorldMap
                $script:globeDragStart = $pos
                $script:globeDragStartLon = $script:globeCenterLon
                $script:globeDragStartLat = $script:globeCenterLat
            }
        })

    $canvasMap.Add_MouseRightButtonUp({
            param($sender, $e)
            if ($script:mapDragging) {
                $script:mapDragging = $false
                $canvasMap.ReleaseMouseCapture()
                $e.Handled = $true
            }
        })

    # Globe rotation via left mouse drag
    $script:globeDragging = $false
    $script:globeDragStart = $null
    $script:globeDragStartLon = 0
    $script:globeDragStartLat = 0

    $canvasMap.Add_MouseLeftButtonDown({
            param($sender, $e)
            if ($script:mapProjection -eq 'Globe') {
                $script:globeDragging = $true
                $script:globeDragStart = $e.GetPosition($canvasMap)
                $script:globeDragStartLon = $script:globeCenterLon
                $script:globeDragStartLat = $script:globeCenterLat
                $canvasMap.CaptureMouse()
            }
        })

    $canvasMap.Add_MouseLeftButtonUp({
            param($sender, $e)
            if ($script:globeDragging) {
                $script:globeDragging = $false
                $canvasMap.ReleaseMouseCapture()
            }
        })

    function Apply-MapZoom {
        if (-not $canvasMap.RenderTransform -or $canvasMap.RenderTransform -isnot [System.Windows.Media.TransformGroup]) {
            $tg = New-Object System.Windows.Media.TransformGroup
            $tg.Children.Add((New-Object System.Windows.Media.ScaleTransform))
            $tg.Children.Add((New-Object System.Windows.Media.TranslateTransform))
            $canvasMap.RenderTransform = $tg
        }
        $tg = $canvasMap.RenderTransform
        $st = $tg.Children[0]
        $tt = $tg.Children[1]
        $st.ScaleX = $script:mapZoom
        $st.ScaleY = $script:mapZoom
        if ($script:mapZoom -le 1.01) {
            $tt.X = 0; $tt.Y = 0
        }
        else {
            # Clamp pan to prevent scrolling beyond map edges
            $cw = $canvasMap.ActualWidth; $ch = $canvasMap.ActualHeight
            $maxPanX = $cw * ($script:mapZoom - 1); $maxPanY = $ch * ($script:mapZoom - 1)
            $script:mapPanX = [Math]::Max(-$maxPanX, [Math]::Min(0, $script:mapPanX))
            $script:mapPanY = [Math]::Max(-$maxPanY, [Math]::Min(0, $script:mapPanY))
            $tt.X = $script:mapPanX
            $tt.Y = $script:mapPanY
        }
    }

    # Filter textbox
    $txtFilter.Add_TextChanged({
            $filter = $txtFilter.Text.Trim()
            if ([string]::IsNullOrEmpty($filter)) {
                $gridExchanges.ItemsSource = $script:exchangeRows
            }
            else {
                $filtered = [System.Collections.ObjectModel.ObservableCollection[object]]::new()
                foreach ($row in $script:exchangeRows) {
                    if ($row.DisplayName -match [regex]::Escape($filter) -or
                        $row.Country -match [regex]::Escape($filter) -or
                        $row.Symbol -match [regex]::Escape($filter) -or
                        $row.Code -match [regex]::Escape($filter) -or
                        $row.City -match [regex]::Escape($filter)) {
                        $filtered.Add($row)
                    }
                }
                $gridExchanges.ItemsSource = $filtered
            }
        })

    # Always on top
    $chkAlwaysOnTop.Add_Checked({ $window.Topmost = $true })
    $chkAlwaysOnTop.Add_Unchecked({ $window.Topmost = $false })

    # Refresh button
    $btnRefresh.Add_Click({
            $txtStatus.Text = 'Refreshing exchange data...'
            try {
                if (Test-Path $scraperPath) {
                    $refreshed = & $scraperPath -ForceRefresh
                    if ($refreshed) {
                        $txtStatus.Text = "Refreshed $($refreshed.Count) exchanges"
                        $txtLastUpdated.Text = "Last updated: $([DateTime]::Now.ToString('yyyy-MM-dd HH:mm:ss'))"
                    }
                }
            }
            catch {
                $txtStatus.Text = "Refresh failed: $($_.Exception.Message)"
            }
        })

    # Reset defaults
    $btnResetDefaults.Add_Click({
            foreach ($row in $script:exchangeRows) {
                $row.IsActive = $row.IsDefault
            }
            $gridExchanges.Items.Refresh()
            Rebuild-Cards
            Rebuild-WorldClocks
            Initialize-WorldMap
            Save-UserPreferences
            $txtStatus.Text = 'Reset to default exchanges'
        })

    # Canvas resize handler to redraw map
    $canvasMap.Add_SizeChanged({
            if ($canvasMap.ActualWidth -gt 50 -and $canvasMap.ActualHeight -gt 50) {
                Initialize-WorldMap
            }
        })

    # Window loaded
    $window.Add_Loaded({
            Rebuild-WorldClocks
            Rebuild-Cards

            # Apply user preferences to UI controls
            if ($script:userPrefs) {
                if ($script:userPrefs.NotifyThresholds) {
                    $chkNotify30.IsChecked = $script:userPrefs.NotifyThresholds -contains 30
                    $chkNotify15.IsChecked = $script:userPrefs.NotifyThresholds -contains 15
                    $chkNotify5.IsChecked = $script:userPrefs.NotifyThresholds -contains 5
                }
                if ($script:userPrefs.AlwaysOnTop) {
                    $chkAlwaysOnTop.IsChecked = $true
                    $window.Topmost = $true
                }
                # Load API keys from secure storage
                $script:secretsBackend = if ($script:userPrefs.SecretsBackend) { $script:userPrefs.SecretsBackend } else { 'CredentialManager' }
                # Sync ComboBox to loaded backend
                for ($i = 0; $i -lt $cmbSecretsBackend.Items.Count; $i++) {
                    if ($cmbSecretsBackend.Items[$i].Tag -eq $script:secretsBackend) {
                        $cmbSecretsBackend.SelectedIndex = $i; break
                    }
                }
                $tdKey = Get-SecureApiKey -ServiceName 'TwelveData' -Backend $script:secretsBackend
                $avKey = Get-SecureApiKey -ServiceName 'AlphaVantage' -Backend $script:secretsBackend
                if ($tdKey) { $txtTwelveDataKey.Text = $tdKey }
                if ($avKey) { $txtAlphaVantageKey.Text = $avKey }

                # Auto-migrate plaintext keys from old JSON format
                if ($script:userPrefs.TwelveDataApiKey -and -not $tdKey) {
                    Set-SecureApiKey -ServiceName 'TwelveData' -ApiKey $script:userPrefs.TwelveDataApiKey -Backend $script:secretsBackend
                    $txtTwelveDataKey.Text = $script:userPrefs.TwelveDataApiKey
                }
                if ($script:userPrefs.AlphaVantageApiKey -and -not $avKey) {
                    Set-SecureApiKey -ServiceName 'AlphaVantage' -ApiKey $script:userPrefs.AlphaVantageApiKey -Backend $script:secretsBackend
                    $txtAlphaVantageKey.Text = $script:userPrefs.AlphaVantageApiKey
                }
                # Remove migrated keys from JSON
                if ($script:userPrefs.TwelveDataApiKey -or $script:userPrefs.AlphaVantageApiKey) {
                    Save-UserPreferences
                }

                # Load commodity base currency preference
                if ($script:userPrefs.CommodityBaseCurrency) {
                    $script:commodityBaseCurrency = $script:userPrefs.CommodityBaseCurrency
                }
            }

            # Set up API key tab visibility
            Update-ApiKeyTabVisibility

            # Set last updated from cache
            $cachePath = Join-Path $dataDir 'exchange-data.json'
            if (Test-Path $cachePath) {
                try {
                    $cache = Get-Content $cachePath -Raw | ConvertFrom-Json
                    if ($cache.LastUpdated) {
                        $txtLastUpdated.Text = "Last updated: $($cache.LastUpdated)"
                    }
                }
                catch { }
            }

            # Initial update
            Update-AllDisplays

            # Initialize market data panel with News tab
            Set-ActiveMarketTab -TabName 'News'
        })

    # Window closing — minimize to tray instead of closing (unless force exit)
    $script:forceExit = $false

    $window.Add_Closing({
            param($closeSender, $closeArgs)
            if ($script:forceExit) {
                $closeArgs.Cancel = $false
            }
            else {
                $closeArgs.Cancel = $true
                $window.WindowState = 'Minimized'
                $window.ShowInTaskbar = $false
                $script:notifyIcon.ShowBalloonTip(2000, 'Stock Exchange Countdown', 'Minimized to system tray. Double-click icon to restore.', [System.Windows.Forms.ToolTipIcon]::Info)
            }
        })

    # Restore from tray on state change
    $window.Add_StateChanged({
            if ($window.WindowState -ne 'Minimized') {
                $window.ShowInTaskbar = $true
            }
        })

    # ── Start Timers ──────────────────────────────────────────────

    # Main 1-second countdown timer
    $timer = New-Object System.Windows.Threading.DispatcherTimer
    $timer.Interval = [TimeSpan]::FromSeconds(1)
    $timer.Add_Tick({ Update-AllDisplays })
    $timer.Start()

    # Market data refresh timer (60 seconds)
    $marketDataTimer = New-Object System.Windows.Threading.DispatcherTimer
    $marketDataTimer.Interval = [TimeSpan]::FromSeconds(60)
    $marketDataTimer.Add_Tick({
            # Only refresh if market data panel is visible and on the Dashboard tab
            if ($marketDataPanel.Visibility -eq 'Visible') {
                Update-MarketDataContent
            }
        })
    $marketDataTimer.Start()

    # Map overlay refresh timer (5 minutes — re-renders solar terminator + time line)
    $mapOverlayTimer = New-Object System.Windows.Threading.DispatcherTimer
    $mapOverlayTimer.Interval = [TimeSpan]::FromMinutes(5)
    $mapOverlayTimer.Add_Tick({
            if ($canvasMap.ActualWidth -gt 50 -and $canvasMap.ActualHeight -gt 50) {
                Initialize-WorldMap
            }
        })
    $mapOverlayTimer.Start()

    # Globe spin timer (100ms — smooth auto-rotation when enabled)
    $globeSpinTimer = New-Object System.Windows.Threading.DispatcherTimer
    $globeSpinTimer.Interval = [TimeSpan]::FromMilliseconds(100)
    $globeSpinTimer.Add_Tick({
            if ($script:globeSpinning -and -not $script:globeDragging -and $script:mapProjection -eq 'Globe') {
                $script:globeCenterLon = ($script:globeCenterLon + 0.5) % 360
                if ($canvasMap.ActualWidth -gt 50 -and $canvasMap.ActualHeight -gt 50) {
                    Initialize-WorldMap
                }
            }
        })
    $globeSpinTimer.Start()

    # ── Show Window ───────────────────────────────────────────────

    # Use ShowDialog so the script can be re-run in the same session
    $app = [System.Windows.Application]::Current
    if (-not $app) {
        $app = New-Object System.Windows.Application
        $app.ShutdownMode = [System.Windows.ShutdownMode]::OnExplicitShutdown
    }
    $window.ShowDialog()

    # Cleanup
    $timer.Stop()
    $marketDataTimer.Stop()
    $mapOverlayTimer.Stop()
    $globeSpinTimer.Stop()
    if ($script:notifyIcon) {
        $script:notifyIcon.Visible = $false
        $script:notifyIcon.Dispose()
    }
} # end function New-StockExchangeCountdownDashboard
