function Format-CountdownSpan {
    <#
    .SYNOPSIS
        Formats a TimeSpan as HH:MM:SS, truncating (not rounding) whole hours.
    #>
    [CmdletBinding()]
    [OutputType([string])]
    param([Parameter(Mandatory)][TimeSpan]$TimeSpan)

    '{0:D2}:{1:D2}:{2:D2}' -f [int][math]::Floor($TimeSpan.TotalHours), $TimeSpan.Minutes, $TimeSpan.Seconds
}

function Get-ExchangeSessionStatus {
    <#
    .SYNOPSIS
        Computes the trading status of an exchange at a given UTC instant.
    .DESCRIPTION
        Pure status calculation used by the dashboard. The exchange's local wall
        clock is derived from its Windows timezone ID, so DST is handled by the
        operating system's timezone rules rather than fixed UTC offsets.
        Countdowns are measured in UTC, so they stay correct on DST transition days.

        An exchange may optionally define OpeningAuctionStart and ClosingAuctionStart
        (HH:mm). From OpeningAuctionStart until OpenTimeLocal, and from
        ClosingAuctionStart until CloseTimeLocal, the status text reports the
        auction (call phase) rather than pre-market or continuous trading. The
        status values themselves are unchanged (PreMarket / ClosingSoon /
        ClosingImminent), so colours and notifications behave as before.
    .PARAMETER Exchange
        Exchange object with Code, TimeZoneId, OpenTimeLocal, CloseTimeLocal and
        optional LunchBreakStart, LunchBreakEnd, OpeningAuctionStart and
        ClosingAuctionStart.
    .PARAMETER UtcNow
        The instant to evaluate. Defaults to the current UTC time.
    .PARAMETER Holidays
        Hashtable of exchange code to an array of yyyy-MM-dd closure dates.
        Missing or empty data means only weekends and session hours apply.
    #>
    [CmdletBinding()]
    [OutputType([hashtable])]
    param(
        [Parameter(Mandatory)][PSCustomObject]$Exchange,
        [datetime]$UtcNow = [DateTime]::UtcNow,
        [hashtable]$Holidays = @{}
    )

    if ($UtcNow.Kind -eq [DateTimeKind]::Local) {
        $UtcNow = $UtcNow.ToUniversalTime()
    }
    else {
        $UtcNow = [DateTime]::SpecifyKind($UtcNow, [DateTimeKind]::Utc)
    }

    $tz = $null
    try {
        $tz = [System.TimeZoneInfo]::FindSystemTimeZoneById($Exchange.TimeZoneId)
        $localNow = [System.TimeZoneInfo]::ConvertTimeFromUtc($UtcNow, $tz)
    }
    catch {
        $localNow = $UtcNow
    }

    # Converts an exchange-local wall-clock time to UTC for countdown arithmetic
    $toUtc = {
        param([datetime]$Local)
        if ($tz) {
            try {
                return [System.TimeZoneInfo]::ConvertTimeToUtc([DateTime]::SpecifyKind($Local, [DateTimeKind]::Unspecified), $tz)
            }
            catch {
                # Wall-clock time inside a DST gap; fall back to treating it as UTC
                Write-Verbose "Cannot convert $Local in $($Exchange.TimeZoneId) to UTC: $($_.Exception.Message)"
            }
        }
        return [DateTime]::SpecifyKind($Local, [DateTimeKind]::Utc)
    }

    $dayOfWeek = $localNow.DayOfWeek

    # Weekend check
    if ($dayOfWeek -eq 'Saturday' -or $dayOfWeek -eq 'Sunday') {
        return @{ Status = 'Closed'; Text = 'Weekend'; TimeRemaining = $null; CountdownTarget = $null }
    }

    # Holiday check
    $todayStr = $localNow.ToString('yyyy-MM-dd')
    if ($Holidays -and $Holidays.ContainsKey($Exchange.Code) -and $Holidays[$Exchange.Code] -contains $todayStr) {
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

    # Opening auction (optional)
    $openAuctionStart = $null
    if ($Exchange.OpeningAuctionStart) {
        $oaParts = $Exchange.OpeningAuctionStart -split ':'
        $openAuctionStart = $localNow.Date.AddHours([int]$oaParts[0]).AddMinutes([int]$oaParts[1])
    }

    # Closing auction (optional)
    $auctionStart = $null
    if ($Exchange.ClosingAuctionStart) {
        $caParts = $Exchange.ClosingAuctionStart -split ':'
        $auctionStart = $localNow.Date.AddHours([int]$caParts[0]).AddMinutes([int]$caParts[1])
    }

    # Pre-market
    if ($localNow -lt $openTime) {
        $remaining = (& $toUtc $openTime) - $UtcNow
        $prefix = if ($openAuctionStart -and $localNow -ge $openAuctionStart) { 'Opening auction - opens in' } else { 'Opens in' }
        return @{ Status = 'PreMarket'; Text = "$prefix $(Format-CountdownSpan -TimeSpan $remaining)"; TimeRemaining = $remaining; CountdownTarget = 'Open' }
    }

    # Post-close
    if ($localNow -ge $closeTime) {
        return @{ Status = 'Closed'; Text = 'Closed'; TimeRemaining = $null; CountdownTarget = $null }
    }

    # Lunch break
    if ($lunchStart -and $lunchEnd -and $localNow -ge $lunchStart -and $localNow -lt $lunchEnd) {
        $remaining = (& $toUtc $lunchEnd) - $UtcNow
        return @{ Status = 'LunchBreak'; Text = "Lunch - Reopens in $(Format-CountdownSpan -TimeSpan $remaining)"; TimeRemaining = $remaining; CountdownTarget = 'Reopen' }
    }

    # Open - countdown to close
    $remaining = (& $toUtc $closeTime) - $UtcNow
    $countdownStr = Format-CountdownSpan -TimeSpan $remaining
    $text = "Closes in $countdownStr"

    if ($auctionStart -and $localNow -ge $auctionStart) {
        # Closing call: no continuous matching, uncross at the end of the period
        $text = "Closing auction - ends in $countdownStr"
    }

    if ($remaining.TotalMinutes -le 5) {
        return @{ Status = 'ClosingImminent'; Text = $text; TimeRemaining = $remaining; CountdownTarget = 'Close' }
    }
    elseif ($remaining.TotalMinutes -le 30) {
        return @{ Status = 'ClosingSoon'; Text = $text; TimeRemaining = $remaining; CountdownTarget = 'Close' }
    }
    else {
        return @{ Status = 'Open'; Text = $text; TimeRemaining = $remaining; CountdownTarget = 'Close' }
    }
}

function Test-ExchangeActive {
    <#
    .SYNOPSIS
        Decides whether an exchange is active from saved user preferences.
    .DESCRIPTION
        Saved ActiveExchanges take precedence; exchanges absent from the saved
        list stay inactive, so new exchanges never alter an existing selection.
        With no saved list, the exchange's IsDefault flag applies.
    #>
    [CmdletBinding()]
    [OutputType([bool])]
    param(
        [Parameter(Mandatory)][PSCustomObject]$Exchange,
        $UserPreferences
    )

    if ($UserPreferences -and $UserPreferences.ActiveExchanges) {
        return [bool](@($UserPreferences.ActiveExchanges) -contains $Exchange.Code)
    }
    return [bool]$Exchange.IsDefault
}
