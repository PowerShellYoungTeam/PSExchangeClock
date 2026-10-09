function Get-IndexDefinition {
    <#
    .SYNOPSIS
        Returns the indices shown on the dashboard's Indices tab.
    .DESCRIPTION
        EtfProxy entries are US-listed ETFs on the Twelve Data free (Basic) tier,
        fetched together in one batched request. Index entries are true index
        symbols fetched individually, so a plan restriction on one index cannot
        affect the ETF batch.

        True indices need a paid Twelve Data plan (RequiredPlan), so they are off
        by default (EnabledByDefault = $false) and are only requested when the
        user lists their symbol in EnabledIndices in user-preferences.json.

        FTSE 100: Twelve Data lists symbol 'FTSE' ("FTSE 100", GBP, mic_code XLON)
        in its /indices catalogue with access level 'Grow'. The bare ticker
        'FTSE' is also an Amundi MSCI UK ETF on Euronext Paris, so the request
        is pinned with MicCode XLON.

        OMXC25: Twelve Data lists symbol 'OMXC25' ("OMX Copenhagen 25 Index", DKK)
        in its /indices catalogue with access level 'Pro' - a Basic key receives
        HTTP 403 (upgrade required).
    #>
    [CmdletBinding()]
    [OutputType([PSCustomObject[]])]
    param()

    $etfs = @(
        @{ Symbol = 'SPY'; Name = 'S&P 500 (SPY)' },
        @{ Symbol = 'QQQ'; Name = 'NASDAQ 100 (QQQ)' },
        @{ Symbol = 'DIA'; Name = 'Dow Jones (DIA)' },
        @{ Symbol = 'EWU'; Name = 'FTSE 100 (EWU)' },
        @{ Symbol = 'EWG'; Name = 'DAX (EWG)' },
        @{ Symbol = 'EWJ'; Name = 'Nikkei 225 (EWJ)' },
        @{ Symbol = 'FXI'; Name = 'China Large-Cap (FXI)' },
        @{ Symbol = 'INDA'; Name = 'India (INDA)' }
    )

    foreach ($etf in $etfs) {
        [PSCustomObject]@{
            Symbol       = $etf.Symbol
            Name         = $etf.Name
            Kind         = 'EtfProxy'
            Provider     = 'TwelveData'
            ExchangeCode = $null
            Country      = $null
            Currency     = 'USD'
            MicCode      = $null
            RequiredPlan = 'Basic'
            EnabledByDefault = $true
            DisabledReason   = $null
        }
    }

    [PSCustomObject]@{
        Symbol       = 'FTSE'
        Name         = 'FTSE 100'
        Kind         = 'Index'
        Provider     = 'TwelveData'
        ExchangeCode = 'XLON'
        Country      = 'United Kingdom'
        Currency     = 'GBP'
        MicCode      = 'XLON'
        RequiredPlan = 'Grow'
        EnabledByDefault = $false
        DisabledReason   = 'Turned off: needs a Twelve Data Grow plan or higher (free/Basic keys get HTTP 403). To enable, add "FTSE" to "EnabledIndices" in user-preferences.json.'
    }

    [PSCustomObject]@{
        Symbol       = 'OMXC25'
        Name         = 'OMX Copenhagen 25 (OMXC25)'
        Kind         = 'Index'
        Provider     = 'TwelveData'
        ExchangeCode = 'XCSE'
        Country      = 'Denmark'
        Currency     = 'DKK'
        MicCode      = $null
        RequiredPlan = 'Pro'
        EnabledByDefault = $false
        DisabledReason   = 'Turned off: needs a Twelve Data Pro plan or higher (free/Basic keys get HTTP 403). To enable, add "EnabledIndices": ["OMXC25"] to user-preferences.json.'
    }
}

function Test-IndexEnabled {
    <#
    .SYNOPSIS
        Decides whether an index is fetched and shown, from user preferences.
    .DESCRIPTION
        An index listed in the EnabledIndices preference is enabled; otherwise
        its EnabledByDefault flag applies. Existing preference files without
        EnabledIndices keep the defaults.
    #>
    [CmdletBinding()]
    [OutputType([bool])]
    param(
        [Parameter(Mandatory)][PSCustomObject]$Definition,
        $UserPreferences
    )

    if ($UserPreferences -and $UserPreferences.EnabledIndices -and (@($UserPreferences.EnabledIndices) -contains $Definition.Symbol)) {
        return $true
    }
    return [bool]$Definition.EnabledByDefault
}

function ConvertFrom-TwelveDataQuote {
    <#
    .SYNOPSIS
        Normalises a Twelve Data /quote response (success or error) into a quote object.
    .DESCRIPTION
        Error responses carry 'code' and 'message'. Codes 403 (plan upgrade
        required) and 404 (symbol not found) are permanent for a given key and
        symbol; anything else (429 rate limit, 5xx) is transient.
        QuoteTimeUtc comes from last_quote_at (last minute candle) when present,
        otherwise from timestamp (opening of the daily bar).
    #>
    [CmdletBinding()]
    [OutputType([PSCustomObject])]
    param(
        [Parameter(Mandatory)][AllowNull()]$Response,
        [string]$Symbol
    )

    $unavailable = {
        param($Code, [string]$Message)
        [PSCustomObject]@{
            Symbol           = $Symbol
            IsAvailable      = $false
            Price            = $null
            Change           = $null
            QuoteTimeUtc     = $null
            ErrorCode        = $Code
            ErrorMessage     = $Message
            IsPermanentError = ($Code -in @(403, 404))
        }
    }

    if ($null -eq $Response) {
        return & $unavailable $null 'Empty response from provider'
    }
    if ($Response.code -or $Response.status -eq 'error') {
        return & $unavailable ([int]$Response.code) "$($Response.message)"
    }

    $price = 0.0
    if (-not $Response.close -or -not [double]::TryParse("$($Response.close)", [System.Globalization.NumberStyles]::Float, [cultureinfo]::InvariantCulture, [ref]$price)) {
        return & $unavailable $null 'Quote contained no price'
    }

    $change = 0
    $prev = 0.0
    if ($Response.previous_close -and [double]::TryParse("$($Response.previous_close)", [System.Globalization.NumberStyles]::Float, [cultureinfo]::InvariantCulture, [ref]$prev) -and $prev -ne 0) {
        $change = [math]::Round((($price - $prev) / $prev) * 100, 2)
    }

    $quoteTime = $null
    $epoch = if ($Response.last_quote_at) { $Response.last_quote_at } else { $Response.timestamp }
    if ($epoch) {
        $quoteTime = [DateTimeOffset]::FromUnixTimeSeconds([long]$epoch).UtcDateTime
    }

    [PSCustomObject]@{
        Symbol           = $Symbol
        IsAvailable      = $true
        Price            = [math]::Round($price, 2)
        Change           = $change
        QuoteTimeUtc     = $quoteTime
        ErrorCode        = $null
        ErrorMessage     = $null
        IsPermanentError = $false
    }
}

function Get-TwelveDataIndexQuote {
    <#
    .SYNOPSIS
        Fetches a single symbol from the Twelve Data /quote endpoint.
    .DESCRIPTION
        Never throws: network failures and HTTP errors are returned as an
        unavailable quote object so callers can fall back to cached data.
    #>
    [CmdletBinding()]
    [OutputType([PSCustomObject])]
    param(
        [Parameter(Mandatory)][string]$Symbol,
        [Parameter(Mandatory)][string]$ApiKey,
        [string]$MicCode,
        [int]$TimeoutSec = 15
    )

    $url = "https://api.twelvedata.com/quote?symbol=$([uri]::EscapeDataString($Symbol))&apikey=$([uri]::EscapeDataString($ApiKey))"
    if ($MicCode) {
        # Disambiguates tickers shared by several instruments (ISO 10383 MIC)
        $url += "&mic_code=$([uri]::EscapeDataString($MicCode))"
    }
    try {
        $response = Invoke-RestMethod -Uri $url -TimeoutSec $TimeoutSec -ErrorAction Stop
    }
    catch {
        $code = $null
        if ($_.Exception.Response) {
            try { $code = [int]$_.Exception.Response.StatusCode }
            catch { Write-Verbose "No HTTP status code on error response: $($_.Exception.Message)" }
        }
        if ($_.ErrorDetails -and $_.ErrorDetails.Message) {
            try {
                $body = $_.ErrorDetails.Message | ConvertFrom-Json -ErrorAction Stop
                if ($body.code) { return ConvertFrom-TwelveDataQuote -Response $body -Symbol $Symbol }
            }
            catch { Write-Verbose "Error body for $Symbol was not JSON: $($_.Exception.Message)" }
        }
        Write-Verbose "Twelve Data quote for $Symbol failed: $($_.Exception.Message)"
        return [PSCustomObject]@{
            Symbol           = $Symbol
            IsAvailable      = $false
            Price            = $null
            Change           = $null
            QuoteTimeUtc     = $null
            ErrorCode        = $code
            ErrorMessage     = $_.Exception.Message
            IsPermanentError = ($code -in @(403, 404))
        }
    }

    ConvertFrom-TwelveDataQuote -Response $response -Symbol $Symbol
}

function Get-CachedIndexQuote {
    <#
    .SYNOPSIS
        Refreshes one index quote through a session cache with fallback.
    .DESCRIPTION
        - Within TtlMinutes of the last attempt, the cached entry is returned unchanged.
        - A successful fetch replaces LastGood.
        - A failed fetch keeps the previous LastGood (shown as stale or last close)
          and records LastError.
        - A permanent error (403/404) blocks further requests for that symbol and
          API key for the rest of the session, so a Basic-plan key does not spend
          a request credit on OMXC25 every refresh.
    .OUTPUTS
        The cache entry hashtable: LastGood, LastFetchUtc, LastAttemptUtc,
        LastError, ErrorCode, Blocked.
    #>
    [CmdletBinding()]
    [OutputType([hashtable])]
    param(
        [Parameter(Mandatory)][PSCustomObject]$Definition,
        [Parameter(Mandatory)][string]$ApiKey,
        [Parameter(Mandatory)][hashtable]$Cache,
        [datetime]$NowUtc = [DateTime]::UtcNow,
        [int]$TtlMinutes = 15
    )

    # Keyed per API key so entering a different key in Settings retries a blocked symbol
    $key = "$($Definition.Symbol)|$($ApiKey.GetHashCode())"
    if (-not $Cache.ContainsKey($key)) {
        $Cache[$key] = @{
            LastGood       = $null
            LastFetchUtc   = $null
            LastAttemptUtc = $null
            LastError      = $null
            ErrorCode      = $null
            Blocked        = $false
        }
    }
    $entry = $Cache[$key]

    if ($entry.Blocked) { return $entry }
    if ($entry.LastAttemptUtc -and ($NowUtc - $entry.LastAttemptUtc).TotalMinutes -lt $TtlMinutes) { return $entry }

    $entry.LastAttemptUtc = $NowUtc
    $quote = Get-TwelveDataIndexQuote -Symbol $Definition.Symbol -ApiKey $ApiKey -MicCode $Definition.MicCode
    if ($quote.IsAvailable) {
        $entry.LastGood = $quote
        $entry.LastFetchUtc = $NowUtc
        $entry.LastError = $null
        $entry.ErrorCode = $null
    }
    else {
        $entry.LastError = $quote.ErrorMessage
        $entry.ErrorCode = $quote.ErrorCode
        if ($quote.IsPermanentError) { $entry.Blocked = $true }
    }
    return $entry
}

function Get-IndexQuoteStatus {
    <#
    .SYNOPSIS
        Classifies an index quote as Intraday, Stale, LastClose or Unavailable.
    .DESCRIPTION
        Uses the dashboard's own status for the index's home exchange (hours,
        weekends, holidays.json) rather than the provider's is_market_open flag,
        because Twelve Data catalogues OMXC25 under the Stockholm MIC (XSTO),
        whose hours and holidays differ from Copenhagen's.

        - Unavailable: no quote at all.
        - Intraday: market in session and quote no older than StaleAfterMinutes.
          Twelve Data does not report whether a quote is real-time or delayed, so
          intraday values are never labelled "live".
        - Stale: market in session but quote older than StaleAfterMinutes, or
          market closed and quote older than LastCloseMaxAgeDays.
        - LastClose: market not in session and quote recent enough to be the
          latest session's close.
    #>
    [CmdletBinding()]
    [OutputType([hashtable])]
    param(
        $QuoteTimeUtc,
        [Parameter(Mandatory)][string]$MarketStatus,
        [datetime]$NowUtc = [DateTime]::UtcNow,
        [int]$StaleAfterMinutes = 30,
        [int]$LastCloseMaxAgeDays = 7
    )

    if ($null -eq $QuoteTimeUtc) {
        return @{ Status = 'Unavailable'; Label = 'Unavailable' }
    }

    $age = $NowUtc - [datetime]$QuoteTimeUtc
    $inSession = $MarketStatus -in @('Open', 'ClosingSoon', 'ClosingImminent')

    if ($inSession) {
        if ($age.TotalMinutes -le $StaleAfterMinutes) {
            return @{ Status = 'Intraday'; Label = 'Intraday (may be delayed)' }
        }
        return @{ Status = 'Stale'; Label = 'Stale' }
    }

    if ($age.TotalDays -le $LastCloseMaxAgeDays) {
        return @{ Status = 'LastClose'; Label = 'Last close' }
    }
    return @{ Status = 'Stale'; Label = 'Stale' }
}
