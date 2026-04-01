<#
.SYNOPSIS
    Stock Exchange Countdown Dashboard
.DESCRIPTION
    WPF dashboard showing live countdown timers to stock exchange closing times worldwide.
    Features: world map with markers, world clocks, color-coded status, sortable exchange list,
    configurable exchange selection, and desktop notifications before close.
.EXAMPLE
    .\New-StockExchangeCountdownDashboard.ps1
#>
[CmdletBinding()]
param()

# ── Assemblies ────────────────────────────────────────────────
Add-Type -AssemblyName PresentationFramework
Add-Type -AssemblyName PresentationCore
Add-Type -AssemblyName WindowsBase
Add-Type -AssemblyName System.Windows.Forms

# ── Load Exchange Data ────────────────────────────────────────
$scriptDir = $PSScriptRoot
$scraperPath = Join-Path $scriptDir 'Get-StockExchangeData.ps1'
$holidaysPath = Join-Path $scriptDir 'holidays.json'
$detailsPath = Join-Path $scriptDir 'exchange-details.json'
$prefsPath = Join-Path $scriptDir 'user-preferences.json'
$mapImagePath = Join-Path $scriptDir 'worldmap.jpg'

if (Test-Path $scraperPath) {
    $allExchanges = & $scraperPath
}
else {
    Write-Warning "Scraper not found at $scraperPath. Using inline fallback."
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
        ActiveExchanges  = $activeExchanges
        NotifyThresholds = @(30, 15, 5)
        AlwaysOnTop      = $false
    }
    try {
        if ($chkNotify30) { $prefs.NotifyThresholds = @() }
        if ($chkNotify30 -and $chkNotify30.IsChecked) { $prefs.NotifyThresholds += 30 }
        if ($chkNotify15 -and $chkNotify15.IsChecked) { $prefs.NotifyThresholds += 15 }
        if ($chkNotify5 -and $chkNotify5.IsChecked) { $prefs.NotifyThresholds += 5 }
        if ($chkAlwaysOnTop) { $prefs.AlwaysOnTop = [bool]$chkAlwaysOnTop.IsChecked }
    }
    catch { }
    $prefs | ConvertTo-Json -Depth 3 | Set-Content -Path $prefsPath -Encoding UTF8
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
        Title="Stock Exchange Countdown Dashboard"
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
                </Grid.ColumnDefinitions>
                <TextBlock Grid.Column="0" Text="&#x1F4C8; Stock Exchange Countdown Dashboard" FontSize="20" FontWeight="Bold" Foreground="#E0E0E0" VerticalAlignment="Center"/>
                <TextBlock Grid.Column="1" x:Name="txtHeaderClock" Text="UTC: --:--:--" FontSize="14" Foreground="#00CC66" VerticalAlignment="Center" FontFamily="Consolas"/>
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
                            <Canvas x:Name="canvasMap" Background="Transparent" ClipToBounds="True"/>
                            <!-- Detail Flyout Panel (hidden by default) -->
                            <Border x:Name="flyoutPanel" Background="#1A1A2E" BorderBrush="#0F3460" BorderThickness="1" CornerRadius="6"
                                    Width="280" HorizontalAlignment="Right" VerticalAlignment="Stretch" Margin="0,5,5,5"
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
$flyoutPanel = $window.FindName('flyoutPanel')
$flyoutTitle = $window.FindName('flyoutTitle')
$flyoutContent = $window.FindName('flyoutContent')
$btnCloseFlyout = $window.FindName('btnCloseFlyout')

$script:currentFlyoutCode = $null

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

$script:mapMarkers = @{}

function Initialize-WorldMap {
    $canvasMap.Children.Clear()

    $converter = [System.Windows.Media.BrushConverter]::new()
    $w = $canvasMap.ActualWidth
    $h = $canvasMap.ActualHeight

    if ($w -lt 50 -or $h -lt 50) { return }

    # Try to use NASA Earth at Night image
    $useImage = $false
    if (Test-Path $mapImagePath) {
        try {
            $uri = New-Object System.Uri($mapImagePath, [System.UriKind]::Absolute)
            $bmp = New-Object System.Windows.Media.Imaging.BitmapImage
            $bmp.BeginInit()
            $bmp.UriSource = $uri
            $bmp.CacheOption = [System.Windows.Media.Imaging.BitmapCacheOption]::OnLoad
            $bmp.EndInit()

            $img = New-Object System.Windows.Controls.Image
            $img.Source = $bmp
            $img.Width = $w
            $img.Height = $h
            $img.Stretch = [System.Windows.Media.Stretch]::Fill
            $img.IsHitTestVisible = $false
            [System.Windows.Controls.Canvas]::SetLeft($img, 0)
            [System.Windows.Controls.Canvas]::SetTop($img, 0)
            $canvasMap.Children.Add($img) | Out-Null
            $useImage = $true
        }
        catch {
            Write-Verbose "Failed to load map image, using polygon fallback: $($_.Exception.Message)"
        }
    }

    if (-not $useImage) {
        # Ocean background
        $mapBg = New-Object System.Windows.Shapes.Rectangle
        $mapBg.Width = $w; $mapBg.Height = $h
        $mapBg.Fill = $converter.ConvertFromString('#0D1B2A')
        $canvasMap.Children.Add($mapBg) | Out-Null
    }

    # Grid lines (optional, controlled by checkbox)
    $showGrid = (-not $useImage) -or ($chkShowGridLines -and $chkShowGridLines.IsChecked)
    if ($showGrid) {
        $gridBrush = $converter.ConvertFromString($(if ($useImage) { '#33FFFFFF' } else { '#152238' }))
        for ($i = 1; $i -lt 6; $i++) {
            $line = New-Object System.Windows.Shapes.Line
            $line.X1 = 0; $line.X2 = $w; $line.Y1 = ($h / 6) * $i; $line.Y2 = $line.Y1
            $line.Stroke = $gridBrush; $line.StrokeThickness = 0.5
            $canvasMap.Children.Add($line) | Out-Null
        }
        for ($i = 1; $i -lt 12; $i++) {
            $line = New-Object System.Windows.Shapes.Line
            $line.X1 = ($w / 12) * $i; $line.X2 = $line.X1; $line.Y1 = 0; $line.Y2 = $h
            $line.Stroke = $gridBrush; $line.StrokeThickness = 0.5
            $canvasMap.Children.Add($line) | Out-Null
        }

        # Equator
        $eqLine = New-Object System.Windows.Shapes.Line
        $eqLine.X1 = 0; $eqLine.X2 = $w; $eqLine.Y1 = $h / 2; $eqLine.Y2 = $h / 2
        $eqLine.Stroke = $converter.ConvertFromString($(if ($useImage) { '#44FFFFFF' } else { '#1E3A5F' })); $eqLine.StrokeThickness = 1
        $da = New-Object System.Windows.Media.DoubleCollection; $da.Add(4.0); $da.Add(4.0)
        $eqLine.StrokeDashArray = $da
        $canvasMap.Children.Add($eqLine) | Out-Null
    }

    # ── Draw continent outlines (only when no image) ──
    if (-not $useImage) {
        $landFill = $converter.ConvertFromString('#1B3A2A')
        $landStroke = $converter.ConvertFromString('#2D6B45')

        # Helper: lat/lon pairs → WPF Polygon on canvas
        function Add-ContinentPoly {
            param([double[][]]$Coords)
            $poly = New-Object System.Windows.Shapes.Polygon
            $poly.Fill = $landFill
            $poly.Stroke = $landStroke
            $poly.StrokeThickness = 1
            $poly.Opacity = 0.85
            $poly.IsHitTestVisible = $false
            $points = New-Object System.Windows.Media.PointCollection
            foreach ($c in $Coords) {
                $px = (($c[1] + 180) / 360) * $w
                $py = ((90 - $c[0]) / 180) * $h
                $points.Add([System.Windows.Point]::new($px, $py))
            }
            $poly.Points = $points
            $canvasMap.Children.Add($poly) | Out-Null
        }

        # North America
        Add-ContinentPoly -Coords @(
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

        # South America
        Add-ContinentPoly -Coords @(
            @(12, -72), @(10, -74), @(8, -77), @(5, -77), @(2, -80), @(-2, -80), @(-5, -81),
            @(-6, -77), @(-4, -70), @(-2, -50), @(-3, -42), @(-5, -35), @(-8, -35), @(-13, -39),
            @(-18, -40), @(-23, -42), @(-28, -49), @(-33, -52), @(-38, -57), @(-42, -62),
            @(-46, -66), @(-50, -70), @(-52, -70), @(-55, -68), @(-56, -70), @(-55, -73),
            @(-50, -75), @(-46, -76), @(-42, -73), @(-38, -73), @(-35, -72), @(-30, -71),
            @(-27, -70), @(-22, -70), @(-18, -70), @(-15, -75), @(-10, -78), @(-5, -80),
            @(-2, -80), @(2, -78), @(5, -77), @(8, -72), @(10, -72), @(12, -72)
        )

        # Europe
        Add-ContinentPoly -Coords @(
            @(36, -6), @(37, -9), @(39, -9), @(42, -9), @(43, -8), @(44, -2), @(46, -2),
            @(47, 2), @(49, 0), @(51, 2), @(53, 5), @(54, 8), @(56, 8), @(55, 12), @(57, 10),
            @(58, 12), @(60, 11), @(63, 10), @(65, 12), @(68, 15), @(70, 20), @(71, 26),
            @(70, 28), @(69, 31), @(67, 26), @(60, 28), @(58, 23), @(55, 22), @(54, 18),
            @(54, 14), @(53, 14), @(52, 10), @(50, 6), @(48, 8), @(47, 7), @(46, 14),
            @(45, 14), @(44, 13), @(42, 15), @(40, 18), @(38, 24), @(36, 28), @(35, 25),
            @(38, 20), @(39, 20), @(41, 18), @(40, 15), @(38, 13), @(37, 15), @(36, 14),
            @(36, 12), @(38, 10), @(39, 3), @(38, 0), @(37, -2), @(36, -5), @(36, -6)
        )

        # Africa
        Add-ContinentPoly -Coords @(
            @(37, -6), @(36, -5), @(36, 0), @(35, 10), @(33, 10), @(30, 10), @(32, 32),
            @(30, 33), @(22, 36), @(15, 42), @(12, 44), @(12, 50), @(5, 42), @(0, 42),
            @(-5, 39), @(-10, 40), @(-15, 40), @(-18, 36), @(-22, 35), @(-25, 33),
            @(-30, 31), @(-34, 26), @(-34, 18), @(-30, 16), @(-25, 15), @(-18, 12),
            @(-12, 14), @(-6, 12), @(0, 10), @(4, 10), @(5, 1), @(4, -2), @(5, -5),
            @(4, -7), @(5, -10), @(10, -15), @(15, -17), @(20, -17), @(22, -16),
            @(25, -14), @(28, -10), @(30, -10), @(32, -5), @(35, -2), @(37, -6)
        )

        # Asia (simplified main body)
        Add-ContinentPoly -Coords @(
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

        # India (sub-continent)
        Add-ContinentPoly -Coords @(
            @(28, 68), @(30, 74), @(30, 78), @(28, 85), @(26, 89), @(22, 90),
            @(22, 88), @(18, 84), @(15, 80), @(10, 80), @(8, 77), @(10, 76),
            @(13, 75), @(15, 74), @(20, 73), @(22, 69), @(25, 68), @(28, 68)
        )

        # Australia
        Add-ContinentPoly -Coords @(
            @(-12, 130), @(-12, 136), @(-14, 136), @(-14, 141), @(-18, 146),
            @(-24, 152), @(-28, 154), @(-33, 152), @(-37, 150), @(-39, 146),
            @(-39, 144), @(-37, 140), @(-35, 137), @(-35, 135), @(-32, 132),
            @(-32, 128), @(-34, 122), @(-34, 116), @(-32, 115), @(-28, 114),
            @(-24, 114), @(-22, 114), @(-18, 122), @(-15, 129), @(-12, 130)
        )

        # Japan (simplified)
        Add-ContinentPoly -Coords @(
            @(45, 142), @(43, 145), @(40, 140), @(38, 140), @(36, 140),
            @(34, 135), @(33, 131), @(34, 130), @(35, 133), @(36, 136),
            @(37, 137), @(39, 140), @(41, 140), @(43, 143), @(45, 142)
        )

        # UK / Ireland
        Add-ContinentPoly -Coords @(
            @(50, -6), @(51, -5), @(52, -4), @(53, -3), @(54, -3), @(55, -2),
            @(56, -3), @(57, -2), @(58, -3), @(59, -3), @(58, -5), @(57, -6),
            @(56, -5), @(55, -5), @(54, -5), @(53, -4), @(52, -5), @(51, -5), @(50, -6)
        )

        # Greenland
        Add-ContinentPoly -Coords @(
            @(60, -45), @(63, -42), @(68, -30), @(72, -22), @(76, -20), @(78, -18),
            @(80, -25), @(82, -35), @(82, -50), @(80, -60), @(78, -68), @(76, -70),
            @(72, -55), @(68, -52), @(65, -53), @(62, -50), @(60, -45)
        )

        # Iceland
        Add-ContinentPoly -Coords @(
            @(64, -22), @(65, -18), @(66, -16), @(66, -14), @(65, -14),
            @(64, -16), @(63, -18), @(63, -22), @(64, -22)
        )

        # New Zealand
        Add-ContinentPoly -Coords @(
            @(-35, 174), @(-37, 176), @(-39, 177), @(-41, 176), @(-42, 174),
            @(-44, 170), @(-46, 167), @(-47, 167), @(-46, 169), @(-44, 172),
            @(-42, 173), @(-40, 176), @(-38, 176), @(-36, 175), @(-35, 174)
        )

    } # end if (-not $useImage) polygon fallback

    # Place exchange markers
    $script:mapMarkers = @{}
    foreach ($row in $script:exchangeRows) {
        if (-not $row.IsActive) { continue }
        Add-MapMarker -Row $row -CanvasWidth $w -CanvasHeight $h
    }
}

function Convert-LatLonToCanvas {
    param([double]$Lat, [double]$Lon, [double]$Width, [double]$Height)
    # Equirectangular projection
    $x = (($Lon + 180) / 360) * $Width
    $y = ((90 - $Lat) / 180) * $Height
    return @{ X = $x; Y = $y }
}

function Add-MapMarker {
    param([PSCustomObject]$Row, [double]$CanvasWidth, [double]$CanvasHeight)

    $pos = Convert-LatLonToCanvas -Lat $Row.Latitude -Lon $Row.Longitude -Width $CanvasWidth -Height $CanvasHeight
    $converter = [System.Windows.Media.BrushConverter]::new()

    # Outer glow circle
    $glow = New-Object System.Windows.Shapes.Ellipse
    $glow.Width = 18; $glow.Height = 18
    $glow.Fill = $converter.ConvertFromString('#3300CC66')
    $glow.IsHitTestVisible = $false
    [System.Windows.Controls.Canvas]::SetLeft($glow, $pos.X - 9)
    [System.Windows.Controls.Canvas]::SetTop($glow, $pos.Y - 9)
    $canvasMap.Children.Add($glow) | Out-Null

    # Main marker dot
    $marker = New-Object System.Windows.Shapes.Ellipse
    $marker.Width = 10; $marker.Height = 10
    $statusColor = Get-StatusColor -Status 'Closed'
    $marker.Fill = $converter.ConvertFromString($statusColor)
    $marker.Stroke = $converter.ConvertFromString('#FFFFFF')
    $marker.StrokeThickness = 1
    $marker.Cursor = [System.Windows.Input.Cursors]::Hand
    [System.Windows.Controls.Canvas]::SetLeft($marker, $pos.X - 5)
    [System.Windows.Controls.Canvas]::SetTop($marker, $pos.Y - 5)

    # Tooltip
    $tp = New-Object System.Windows.Controls.ToolTip
    $tp.Content = "$($Row.DisplayName) ($($Row.Symbol))"
    [System.Windows.Controls.ToolTipService]::SetToolTip($marker, $tp)

    # Click handler for flyout — use Tag to avoid closure scope issues
    $marker.Tag = $Row.Code
    $marker.Add_MouseLeftButtonDown({
            param($sender, $e)
            Show-ExchangeFlyout -ExchangeCode $sender.Tag
        })

    $canvasMap.Children.Add($marker) | Out-Null

    # Label
    $lbl = New-Object System.Windows.Controls.TextBlock
    $lbl.Text = $Row.Symbol
    $lbl.FontSize = 9
    $lbl.Foreground = $converter.ConvertFromString('#CCCCCC')
    $lbl.IsHitTestVisible = $false
    [System.Windows.Controls.Canvas]::SetLeft($lbl, $pos.X + 7)
    [System.Windows.Controls.Canvas]::SetTop($lbl, $pos.Y - 7)
    $canvasMap.Children.Add($lbl) | Out-Null

    $script:mapMarkers[$Row.Code] = @{ Marker = $marker; Glow = $glow; Tooltip = $tp; Label = $lbl }
}

# ── Exchange Detail Flyout ────────────────────────────────────

function Show-ExchangeFlyout {
    param([string]$ExchangeCode)

    # Toggle off if same exchange clicked
    if ($script:currentFlyoutCode -eq $ExchangeCode -and $flyoutPanel.Visibility -eq 'Visible') {
        $flyoutPanel.Visibility = 'Collapsed'
        $script:currentFlyoutCode = $null
        return
    }

    $script:currentFlyoutCode = $ExchangeCode
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
        if ($script:mapMarkers.ContainsKey($row.Code)) {
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

    # Update status bar
    if ($nextCloseExchange -and $nextCloseRemaining) {
        $ncs = "{0:D2}:{1:D2}:{2:D2}" -f [int]$nextCloseRemaining.TotalHours, $nextCloseRemaining.Minutes, $nextCloseRemaining.Seconds
        $txtNextClose.Text = "Next close: $nextCloseExchange in $ncs"
        $window.Title = "Stock Exchange Countdown - Next: $nextCloseExchange $ncs"
    }
    else {
        $txtNextClose.Text = 'All markets closed'
        $window.Title = 'Stock Exchange Countdown Dashboard'
    }

    $txtStatus.Text = "Active: $(@($script:exchangeRows | Where-Object { $_.IsActive }).Count) exchanges | Updated: $([DateTime]::Now.ToString('HH:mm:ss'))"

    # Refresh DataGrid display
    $gridExchanges.Items.Refresh()
}

# ── Event Handlers ────────────────────────────────────────────

# Flyout close button
$btnCloseFlyout.Add_Click({
        $flyoutPanel.Visibility = 'Collapsed'
        $script:currentFlyoutCode = $null
    })

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
        }

        # Set last updated from cache
        $cachePath = Join-Path $scriptDir 'exchange-data.json'
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

# ── Start Timer ───────────────────────────────────────────────

$timer = New-Object System.Windows.Threading.DispatcherTimer
$timer.Interval = [TimeSpan]::FromSeconds(1)
$timer.Add_Tick({ Update-AllDisplays })
$timer.Start()

# ── Show Window ───────────────────────────────────────────────

$app = New-Object System.Windows.Application
$app.Run($window)

# Cleanup
$timer.Stop()
if ($script:notifyIcon) {
    $script:notifyIcon.Visible = $false
    $script:notifyIcon.Dispose()
}
