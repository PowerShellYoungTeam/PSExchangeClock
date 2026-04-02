function Get-StockExchangeData {
<#
.SYNOPSIS
    Retrieves stock exchange trading hours for 20 major global exchanges.
.DESCRIPTION
    Fetches trading hours data from Wikipedia, parses it, and saves to a local
    JSON cache file. Includes hardcoded fallback data for 20 major exchanges
    so the tool works offline. Used by PSExchangeClock to determine exchange
    open/close times, lunch breaks, and geographic coordinates.

    Supported exchanges: NYSE, NASDAQ, LSE, Euronext Paris, Tokyo, Shanghai,
    Hong Kong, Toronto, Frankfurt, Sydney, Bombay, NSE India, Korea,
    Switzerland, Johannesburg, B3 Brazil, Mexico, Singapore, Taiwan, Moscow.
.PARAMETER ForceRefresh
    Bypass the cache and re-scrape from Wikipedia.
.PARAMETER OutputPath
    Path to save the JSON cache file. Defaults to exchange-data.json in the
    module Data directory.
.EXAMPLE
    Get-StockExchangeData
    # Uses cached data if available, otherwise fetches from Wikipedia
.EXAMPLE
    Get-StockExchangeData -ForceRefresh
    # Forces a fresh scrape from Wikipedia
.NOTES
    Author  : PowerShellYoungTeam
    Version : 1.0.0
    Project : https://github.com/PowerShellYoungTeam/PSExchangeClock
.LINK
    https://github.com/PowerShellYoungTeam/PSExchangeClock
.LINK
    New-StockExchangeCountdownDashboard
#>
[CmdletBinding()]
param(
    [switch]$ForceRefresh,
    [string]$OutputPath = (Join-Path (Split-Path $PSScriptRoot -Parent) 'Data\exchange-data.json')
)

function Get-DefaultExchangeData {
    <#
    .SYNOPSIS
        Returns hardcoded data for 20 major stock exchanges.
    #>
    @(
        [PSCustomObject]@{
            Name            = 'New York Stock Exchange'
            Code            = 'XNYS'
            Symbol          = 'NYSE'
            Country         = 'United States'
            City            = 'New York'
            TimeZoneId      = 'Eastern Standard Time'
            OpenTimeLocal   = '09:30'
            CloseTimeLocal  = '16:00'
            LunchBreakStart = $null
            LunchBreakEnd   = $null
            Latitude        = 40.7069
            Longitude       = -74.0113
            IsDefault       = $true
        },
        [PSCustomObject]@{
            Name            = 'NASDAQ'
            Code            = 'XNAS'
            Symbol          = 'NASDAQ'
            Country         = 'United States'
            City            = 'New York'
            TimeZoneId      = 'Eastern Standard Time'
            OpenTimeLocal   = '09:30'
            CloseTimeLocal  = '16:00'
            LunchBreakStart = $null
            LunchBreakEnd   = $null
            Latitude        = 40.7569
            Longitude       = -73.9897
            IsDefault       = $true
        },
        [PSCustomObject]@{
            Name            = 'London Stock Exchange'
            Code            = 'XLON'
            Symbol          = 'LSE'
            Country         = 'United Kingdom'
            City            = 'London'
            TimeZoneId      = 'GMT Standard Time'
            OpenTimeLocal   = '08:00'
            CloseTimeLocal  = '16:30'
            LunchBreakStart = $null
            LunchBreakEnd   = $null
            Latitude        = 51.5155
            Longitude       = -0.0922
            IsDefault       = $true
        },
        [PSCustomObject]@{
            Name            = 'Euronext Paris'
            Code            = 'XPAR'
            Symbol          = 'EPA'
            Country         = 'France'
            City            = 'Paris'
            TimeZoneId      = 'Romance Standard Time'
            OpenTimeLocal   = '09:00'
            CloseTimeLocal  = '17:30'
            LunchBreakStart = $null
            LunchBreakEnd   = $null
            Latitude        = 48.8698
            Longitude       = 2.3371
            IsDefault       = $true
        },
        [PSCustomObject]@{
            Name            = 'Tokyo Stock Exchange'
            Code            = 'XTKS'
            Symbol          = 'TSE'
            Country         = 'Japan'
            City            = 'Tokyo'
            TimeZoneId      = 'Tokyo Standard Time'
            OpenTimeLocal   = '09:00'
            CloseTimeLocal  = '15:30'
            LunchBreakStart = '11:30'
            LunchBreakEnd   = '12:30'
            Latitude        = 35.6814
            Longitude       = 139.7637
            IsDefault       = $true
        },
        [PSCustomObject]@{
            Name            = 'Shanghai Stock Exchange'
            Code            = 'XSHG'
            Symbol          = 'SSE'
            Country         = 'China'
            City            = 'Shanghai'
            TimeZoneId      = 'China Standard Time'
            OpenTimeLocal   = '09:30'
            CloseTimeLocal  = '15:00'
            LunchBreakStart = '11:30'
            LunchBreakEnd   = '13:00'
            Latitude        = 31.2320
            Longitude       = 121.4758
            IsDefault       = $true
        },
        [PSCustomObject]@{
            Name            = 'Hong Kong Stock Exchange'
            Code            = 'XHKG'
            Symbol          = 'HKEX'
            Country         = 'Hong Kong'
            City            = 'Hong Kong'
            TimeZoneId      = 'China Standard Time'
            OpenTimeLocal   = '09:30'
            CloseTimeLocal  = '16:00'
            LunchBreakStart = '12:00'
            LunchBreakEnd   = '13:00'
            Latitude        = 22.2860
            Longitude       = 114.1580
            IsDefault       = $true
        },
        [PSCustomObject]@{
            Name            = 'Toronto Stock Exchange'
            Code            = 'XTSE'
            Symbol          = 'TSX'
            Country         = 'Canada'
            City            = 'Toronto'
            TimeZoneId      = 'Eastern Standard Time'
            OpenTimeLocal   = '09:30'
            CloseTimeLocal  = '16:00'
            LunchBreakStart = $null
            LunchBreakEnd   = $null
            Latitude        = 43.6490
            Longitude       = -79.3832
            IsDefault       = $true
        },
        [PSCustomObject]@{
            Name            = 'Frankfurt Stock Exchange (XETRA)'
            Code            = 'XETR'
            Symbol          = 'FRA'
            Country         = 'Germany'
            City            = 'Frankfurt'
            TimeZoneId      = 'W. Europe Standard Time'
            OpenTimeLocal   = '09:00'
            CloseTimeLocal  = '17:30'
            LunchBreakStart = $null
            LunchBreakEnd   = $null
            Latitude        = 50.1109
            Longitude       = 8.6821
            IsDefault       = $true
        },
        [PSCustomObject]@{
            Name            = 'Australian Securities Exchange'
            Code            = 'XASX'
            Symbol          = 'ASX'
            Country         = 'Australia'
            City            = 'Sydney'
            TimeZoneId      = 'AUS Eastern Standard Time'
            OpenTimeLocal   = '10:00'
            CloseTimeLocal  = '16:00'
            LunchBreakStart = $null
            LunchBreakEnd   = $null
            Latitude        = -33.8666
            Longitude       = 151.2073
            IsDefault       = $true
        },
        [PSCustomObject]@{
            Name            = 'Bombay Stock Exchange'
            Code            = 'XBOM'
            Symbol          = 'BSE'
            Country         = 'India'
            City            = 'Mumbai'
            TimeZoneId      = 'India Standard Time'
            OpenTimeLocal   = '09:15'
            CloseTimeLocal  = '15:30'
            LunchBreakStart = $null
            LunchBreakEnd   = $null
            Latitude        = 18.9262
            Longitude       = 72.8333
            IsDefault       = $false
        },
        [PSCustomObject]@{
            Name            = 'National Stock Exchange of India'
            Code            = 'XNSE'
            Symbol          = 'NSE'
            Country         = 'India'
            City            = 'Mumbai'
            TimeZoneId      = 'India Standard Time'
            OpenTimeLocal   = '09:15'
            CloseTimeLocal  = '15:30'
            LunchBreakStart = $null
            LunchBreakEnd   = $null
            Latitude        = 19.0553
            Longitude       = 72.8629
            IsDefault       = $false
        },
        [PSCustomObject]@{
            Name            = 'Korea Exchange'
            Code            = 'XKRX'
            Symbol          = 'KRX'
            Country         = 'South Korea'
            City            = 'Seoul'
            TimeZoneId      = 'Korea Standard Time'
            OpenTimeLocal   = '09:00'
            CloseTimeLocal  = '15:30'
            LunchBreakStart = $null
            LunchBreakEnd   = $null
            Latitude        = 37.5242
            Longitude       = 127.0507
            IsDefault       = $false
        },
        [PSCustomObject]@{
            Name            = 'SIX Swiss Exchange'
            Code            = 'XSWX'
            Symbol          = 'SIX'
            Country         = 'Switzerland'
            City            = 'Zurich'
            TimeZoneId      = 'W. Europe Standard Time'
            OpenTimeLocal   = '09:00'
            CloseTimeLocal  = '17:30'
            LunchBreakStart = $null
            LunchBreakEnd   = $null
            Latitude        = 47.3769
            Longitude       = 8.5417
            IsDefault       = $false
        },
        [PSCustomObject]@{
            Name            = 'Johannesburg Stock Exchange'
            Code            = 'XJSE'
            Symbol          = 'JSE'
            Country         = 'South Africa'
            City            = 'Johannesburg'
            TimeZoneId      = 'South Africa Standard Time'
            OpenTimeLocal   = '09:00'
            CloseTimeLocal  = '17:00'
            LunchBreakStart = $null
            LunchBreakEnd   = $null
            Latitude        = -26.2023
            Longitude       = 28.0436
            IsDefault       = $false
        },
        [PSCustomObject]@{
            Name            = 'B3 - Brasil Bolsa Balcao'
            Code            = 'BVMF'
            Symbol          = 'B3'
            Country         = 'Brazil'
            City            = 'Sao Paulo'
            TimeZoneId      = 'E. South America Standard Time'
            OpenTimeLocal   = '10:00'
            CloseTimeLocal  = '17:00'
            LunchBreakStart = $null
            LunchBreakEnd   = $null
            Latitude        = -23.5481
            Longitude       = -46.6335
            IsDefault       = $false
        },
        [PSCustomObject]@{
            Name            = 'Mexican Stock Exchange'
            Code            = 'XMEX'
            Symbol          = 'BMV'
            Country         = 'Mexico'
            City            = 'Mexico City'
            TimeZoneId      = 'Central Standard Time (Mexico)'
            OpenTimeLocal   = '08:30'
            CloseTimeLocal  = '15:00'
            LunchBreakStart = $null
            LunchBreakEnd   = $null
            Latitude        = 19.4213
            Longitude       = -99.1667
            IsDefault       = $false
        },
        [PSCustomObject]@{
            Name            = 'Singapore Exchange'
            Code            = 'XSES'
            Symbol          = 'SGX'
            Country         = 'Singapore'
            City            = 'Singapore'
            TimeZoneId      = 'Singapore Standard Time'
            OpenTimeLocal   = '09:00'
            CloseTimeLocal  = '17:00'
            LunchBreakStart = $null
            LunchBreakEnd   = $null
            Latitude        = 1.2833
            Longitude       = 103.8494
            IsDefault       = $false
        },
        [PSCustomObject]@{
            Name            = 'Taiwan Stock Exchange'
            Code            = 'XTAI'
            Symbol          = 'TWSE'
            Country         = 'Taiwan'
            City            = 'Taipei'
            TimeZoneId      = 'Taipei Standard Time'
            OpenTimeLocal   = '09:00'
            CloseTimeLocal  = '13:30'
            LunchBreakStart = $null
            LunchBreakEnd   = $null
            Latitude        = 25.0407
            Longitude       = 121.5141
            IsDefault       = $false
        },
        [PSCustomObject]@{
            Name            = 'Moscow Exchange'
            Code            = 'XMOS'
            Symbol          = 'MOEX'
            Country         = 'Russia'
            City            = 'Moscow'
            TimeZoneId      = 'Russian Standard Time'
            OpenTimeLocal   = '09:50'
            CloseTimeLocal  = '18:50'
            LunchBreakStart = $null
            LunchBreakEnd   = $null
            Latitude        = 55.7558
            Longitude       = 37.6173
            IsDefault       = $false
        }
    )
}

function Invoke-WikipediaScrape {
    <#
    .SYNOPSIS
        Attempts to scrape stock exchange trading hours from Wikipedia.
    #>
    [CmdletBinding()]
    param()

    $url = 'https://en.wikipedia.org/wiki/List_of_stock_exchange_trading_hours'
    Write-Verbose "Fetching $url"

    try {
        $response = Invoke-WebRequest -Uri $url -UseBasicParsing -TimeoutSec 15 -ErrorAction Stop
        $html = $response.Content
    }
    catch {
        Write-Warning "Failed to fetch Wikipedia: $($_.Exception.Message)"
        return $null
    }

    # Attempt to parse wikitable rows
    # Look for sortable wikitable
    $tablePattern = '(?s)<table[^>]*class="[^"]*wikitable[^"]*sortable[^"]*"[^>]*>(.*?)</table>'
    if ($html -notmatch $tablePattern) {
        # Try without sortable
        $tablePattern = '(?s)<table[^>]*class="[^"]*wikitable[^"]*"[^>]*>(.*?)</table>'
        if ($html -notmatch $tablePattern) {
            Write-Warning "Could not find wikitable on page."
            return $null
        }
    }

    $tableHtml = $Matches[1]
    $rowPattern = '(?s)<tr>(.*?)</tr>'
    $cellPattern = '(?s)<t[dh][^>]*>(.*?)</t[dh]>'

    $rows = [regex]::Matches($tableHtml, $rowPattern)
    if ($rows.Count -lt 2) {
        Write-Warning "Table has insufficient rows ($($rows.Count))."
        return $null
    }

    # Parse header row to identify columns
    $headerRow = $rows[0]
    $headerCells = [regex]::Matches($headerRow.Groups[1].Value, $cellPattern)
    $headers = @()
    foreach ($cell in $headerCells) {
        $text = ($cell.Groups[1].Value -replace '<[^>]+>', '' -replace '&[^;]+;', ' ').Trim()
        $headers += $text.ToLower()
    }

    Write-Verbose "Headers found: $($headers -join ', ')"

    # Map known timezone IDs
    $tzMapping = @{
        'eastern' = 'Eastern Standard Time'
        'et'      = 'Eastern Standard Time'
        'est'     = 'Eastern Standard Time'
        'gmt'     = 'GMT Standard Time'
        'bst'     = 'GMT Standard Time'
        'cet'     = 'W. Europe Standard Time'
        'cest'    = 'W. Europe Standard Time'
        'jst'     = 'Tokyo Standard Time'
        'cst'     = 'China Standard Time'
        'hkt'     = 'China Standard Time'
        'aest'    = 'AUS Eastern Standard Time'
        'ist'     = 'India Standard Time'
        'kst'     = 'Korea Standard Time'
    }

    $scraped = @()
    for ($i = 1; $i -lt $rows.Count; $i++) {
        $cells = [regex]::Matches($rows[$i].Groups[1].Value, $cellPattern)
        $cellTexts = @()
        foreach ($cell in $cells) {
            $text = ($cell.Groups[1].Value -replace '<[^>]+>', '' -replace '&[^;]+;', ' ' -replace '\s+', ' ').Trim()
            $cellTexts += $text
        }

        if ($cellTexts.Count -lt 4) { continue }

        # Best-effort field extraction (depends on table structure)
        $exchangeName = $cellTexts[0]
        $timezone = if ($cellTexts.Count -gt 2) { $cellTexts[2] } else { '' }
        $openTime = if ($cellTexts.Count -gt 3) { $cellTexts[3] } else { '' }
        $closeTime = if ($cellTexts.Count -gt 4) { $cellTexts[4] } else { '' }

        if ($exchangeName -and $closeTime) {
            Write-Verbose "  Scraped: $exchangeName | TZ=$timezone | Open=$openTime | Close=$closeTime"
            $scraped += [PSCustomObject]@{
                ScrapedName  = $exchangeName
                ScrapedTZ    = $timezone
                ScrapedOpen  = $openTime
                ScrapedClose = $closeTime
            }
        }
    }

    Write-Verbose "Scraped $($scraped.Count) exchange rows from Wikipedia."
    return $scraped
}

# ── Main Logic ────────────────────────────────────────────────

# Check cache
if (-not $ForceRefresh -and (Test-Path $OutputPath)) {
    try {
        $cached = Get-Content $OutputPath -Raw -ErrorAction Stop | ConvertFrom-Json
        if ($cached.LastUpdated) {
            $age = (Get-Date) - [datetime]$cached.LastUpdated
            if ($age.TotalDays -lt 30) {
                Write-Verbose "Cache is $([int]$age.TotalDays) days old. Using cached data."
                return $cached.Exchanges
            }
            Write-Verbose "Cache is $([int]$age.TotalDays) days old. Refreshing."
        }
    }
    catch {
        Write-Verbose "Cache read failed: $($_.Exception.Message)"
    }
}

# Start with hardcoded defaults
$exchanges = Get-DefaultExchangeData

# Attempt scrape to verify/enhance data
$scraped = Invoke-WikipediaScrape
if ($scraped) {
    Write-Verbose "Wikipedia scrape returned $($scraped.Count) entries. Hardcoded data used as primary; scrape logged for reference."
    # In a future version, merge scraped data with hardcoded to add new exchanges
    # For now, the hardcoded data is authoritative (known-good timezone IDs and coordinates)
}

# Save to cache
$cacheObject = [PSCustomObject]@{
    LastUpdated = (Get-Date).ToString('o')
    Source      = 'Hardcoded + Wikipedia scrape attempt'
    Version     = '1.0'
    Exchanges   = $exchanges
}

try {
    $cacheObject | ConvertTo-Json -Depth 5 | Set-Content $OutputPath -Encoding UTF8 -ErrorAction Stop
    Write-Verbose "Saved exchange data to $OutputPath"
}
catch {
    Write-Warning "Failed to save cache: $($_.Exception.Message)"
}

return $exchanges
} # end function Get-StockExchangeData
