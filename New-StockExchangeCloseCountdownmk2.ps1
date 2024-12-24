<# 
.SYNOPSIS
    New-StockExchangeCountdown
    This script will display a countdown to the next stock exchange close time.
#>

# Ensure Python and pip are installed
if (-not (Get-Command python -ErrorAction SilentlyContinue)) {
    Write-Error "Python is not installed. Please install Python to proceed."
    exit 1
}

if (-not (Get-Command pip -ErrorAction SilentlyContinue)) {
    Write-Error "pip is not installed. Please install pip to proceed."
    exit 1
}

# Install yfinance and trading_calendars if not already installed
$yfinanceInstalled = & python -m pip show yfinance
if (-not $yfinanceInstalled) {
    Write-Output "Installing yfinance..."
    & python -m pip install yfinance
}

$tradingCalendarsInstalled = & python -m pip show trading_calendars
if (-not $tradingCalendarsInstalled) {
    Write-Output "Installing trading_calendars..."
    & python -m pip install trading_calendars
}

# Python script to get closing times
$script = @"
import yfinance as yf
from datetime import datetime, timedelta
import pytz
import trading_calendars as tc

def get_closing_time(ticker):
    stock = yf.Ticker(ticker)
    exchange = stock.info['exchange']
    timezone = {
        'NMS': 'US/Eastern',  # NASDAQ
        'NYQ': 'US/Eastern',  # NYSE
        'LSE': 'Europe/London',  # London Stock Exchange
        'HKG': 'Asia/Hong_Kong'  # Hong Kong Stock Exchange
    }.get(exchange, 'UTC')
    
    now = datetime.now(pytz.timezone(timezone))
    if exchange in ['NMS', 'NYQ']:
        calendar = tc.get_calendar('XNYS')
    elif exchange == 'LSE':
        calendar = tc.get_calendar('XLON')
    elif exchange == 'HKG':
        calendar = tc.get_calendar('XHKG')
    else:
        calendar = None

    if calendar:
        schedule = calendar.schedule.loc[now.date():now.date()]
        if not schedule.empty:
            close_time = schedule.iloc[0]['market_close'].astimezone(pytz.timezone(timezone))
        else:
            close_time = now.replace(hour=16, minute=0, second=0, microsecond=0)
    else:
        close_time = now.replace(hour=16, minute=0, second=0, microsecond=0)
    
    if now > close_time:
        close_time += timedelta(days=1)
    
    return exchange, close_time

ny_exchange, ny_close = get_closing_time("AAPL")
london_exchange, london_close = get_closing_time("LLOY.L")
hk_exchange, hk_close = get_closing_time("0005.HK")

print(f"{ny_exchange},{ny_close}")
print(f"{london_exchange},{london_close}")
print(f"{hk_exchange},{hk_close}")
"@

# Run the Python script and capture the output
$pythonOutput = & python -c $script

# Parse the output and create custom objects
$closingTimes = @()
foreach ($line in $pythonOutput -split "`n") {
    $parts = $line -split ","
    $closingTimes += [PSCustomObject]@{
        ExchangeName = $parts[0]
        ClosingTime  = [datetime]::ParseExact($parts[1], "yyyy-MM-dd HH:mm:ss%K", $null).ToUniversalTime()
    }
}

# Load the necessary assemblies
Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

# Create a new form
$form = New-Object System.Windows.Forms.Form
$form.Text = "Stock Exchange Close Countdown"
$form.Size = New-Object System.Drawing.Size(400, 200)
$form.StartPosition = "CenterScreen"

# Create labels for each exchange
$labels = @()
$yOffset = 20
foreach ($closingTime in $closingTimes) {
    $label = New-Object System.Windows.Forms.Label
    $label.Text = "Exchange: $($closingTime.ExchangeName), Closing Time: $($closingTime.ClosingTime.ToString('dd/MM/yyyy HH:mm:ss'))"
    $label.AutoSize = $true
    $label.Location = New-Object System.Drawing.Point(10, $yOffset)
    $form.Controls.Add($label)
    $labels += $label
    $yOffset += 20
}

# Create a timer to update the countdown every second
$timer = New-Object System.Windows.Forms.Timer
$timer.Interval = 1000
$timer.Add_Tick({
    Update-Countdown
})

# Create a function to update the countdown
function Update-Countdown {
    $currentTime = (Get-Date).ToUniversalTime()
    for ($i = 0; $i -lt $closingTimes.Count; $i++) {
        $closingTime = $closingTimes[$i]
        $label = $labels[$i]
        $timeRemaining = $closingTime.ClosingTime - $currentTime
        if ($timeRemaining.TotalSeconds -le 0) {
            $label.Text = "Exchange: $($closingTime.ExchangeName), Closing Time: $($closingTime.ClosingTime.ToString('dd/MM/yyyy HH:mm:ss')), Status: Closed"
            $label.ForeColor = "Black"
            $label.Font = New-Object System.Drawing.Font($label.Font, [System.Drawing.FontStyle]::Bold)
        } else {
            $timeRemainingString = "{0:D2}:{1:D2}:{2:D2}" -f $timeRemaining.Hours, $timeRemaining.Minutes, $timeRemaining.Seconds
            $label.Text = "Exchange: $($closingTime.ExchangeName), Closing Time: $($closingTime.ClosingTime.ToString('dd/MM/yyyy HH:mm:ss')), Time Remaining: $timeRemainingString"
            if ($timeRemaining -le (New-TimeSpan -Minutes 30)) {
                $label.ForeColor = "Red"
            } elseif ($timeRemaining -le (New-TimeSpan -Hours 1)) {
                $label.ForeColor = "Yellow"
            } else {
                $label.ForeColor = "Black"
                $label.Font = New-Object System.Drawing.Font($label.Font, [System.Drawing.FontStyle]::Regular)
            }
        }
    }
}

# Start the timer
$timer.Start()

# Show the form
$form.ShowDialog() | Out-Null

# Stop the timer when the form is closed
$timer.Stop()
$timer.Dispose()
$form.Dispose()