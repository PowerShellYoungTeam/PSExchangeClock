#requires -Version 7.0
#requires -Modules @{ ModuleName = 'Pester'; ModuleVersion = '5.0.0' }

# Deterministic tests: every status is evaluated at a fixed UTC instant and all
# network calls are mocked. Run with: Invoke-Pester -Path ./Tests

BeforeAll {
    $script:moduleRoot = Join-Path (Split-Path $PSScriptRoot -Parent) 'PSExchangeClock'
    $script:dataDir = Join-Path $script:moduleRoot 'Data'

    . (Join-Path $script:moduleRoot 'Private/Get-ExchangeSessionStatus.ps1')
    . (Join-Path $script:moduleRoot 'Private/IndexQuotes.ps1')
    Import-Module (Join-Path $script:moduleRoot 'PSExchangeClock.psd1') -Force

    # Load holidays exactly as the dashboard does
    $script:holidays = @{}
    $holidayJson = Get-Content (Join-Path $script:dataDir 'holidays.json') -Raw | ConvertFrom-Json
    foreach ($prop in $holidayJson.PSObject.Properties) {
        $script:holidays[$prop.Name] = @($prop.Value)
    }

    # Built-in exchange list via the offline path (Wikipedia scrape mocked to fail)
    Mock -ModuleName PSExchangeClock Invoke-WebRequest { throw 'offline' }
    $script:allExchanges = @(Get-StockExchangeData -ForceRefresh -OutputPath (Join-Path $TestDrive 'exchange-data.json') -WarningAction SilentlyContinue)
    $script:xcse = $script:allExchanges | Where-Object { $_.Code -eq 'XCSE' }

    function Get-Utc([string]$Iso) {
        [DateTime]::SpecifyKind([DateTime]::ParseExact($Iso, 'yyyy-MM-ddTHH:mm:ss', [cultureinfo]::InvariantCulture), [DateTimeKind]::Utc)
    }
    function Get-XcseStatus([string]$UtcIso, [hashtable]$Holidays = $script:holidays) {
        Get-ExchangeSessionStatus -Exchange $script:xcse -UtcNow (Get-Utc $UtcIso) -Holidays $Holidays
    }
}

Describe 'Nasdaq Copenhagen exchange definition' {
    It 'is a built-in exchange with the expected attributes' {
        $script:xcse | Should -Not -BeNullOrEmpty
        $script:xcse.Name | Should -Be 'Nasdaq Copenhagen'
        $script:xcse.Country | Should -Be 'Denmark'
        $script:xcse.City | Should -Be 'Copenhagen'
        $script:xcse.TimeZoneId | Should -Be 'Romance Standard Time'
        $script:xcse.OpenTimeLocal | Should -Be '09:00'
        $script:xcse.CloseTimeLocal | Should -Be '17:00'
        $script:xcse.OpeningAuctionStart | Should -Be '08:45'
        $script:xcse.ClosingAuctionStart | Should -Be '16:55'
        $script:xcse.Latitude | Should -Be 55.6761
        $script:xcse.Longitude | Should -Be 12.5683
        $script:xcse.IsDefault | Should -BeFalse
    }

    It 'uses a Windows timezone ID that maps to IANA Europe/Copenhagen' {
        $windowsId = $null
        [System.TimeZoneInfo]::TryConvertIanaIdToWindowsId('Europe/Copenhagen', [ref]$windowsId) | Should -BeTrue
        $windowsId | Should -Be $script:xcse.TimeZoneId
    }

    It 'keeps all 20 previously supported exchanges' {
        $expected = 'XNYS', 'XNAS', 'XLON', 'XPAR', 'XTKS', 'XSHG', 'XHKG', 'XTSE', 'XETR', 'XASX',
        'XBOM', 'XNSE', 'XKRX', 'XSWX', 'XJSE', 'BVMF', 'XMEX', 'XSES', 'XTAI', 'XMOS'
        foreach ($code in $expected) {
            $script:allExchanges.Code | Should -Contain $code
        }
        $script:allExchanges.Count | Should -Be 21
    }

    It 'has exchange details for the flyout, in DKK and linked to OMXC25' {
        $details = Get-Content (Join-Path $script:dataDir 'exchange-details.json') -Raw | ConvertFrom-Json
        $details.XCSE | Should -Not -BeNullOrEmpty
        $details.XCSE.Currency | Should -Be 'DKK'
        $details.XCSE.PrimaryIndex | Should -Match 'OMXC25'
    }

    It 'is present in the shipped exchange-data.json fallback' {
        $shipped = Get-Content (Join-Path $script:dataDir 'exchange-data.json') -Raw | ConvertFrom-Json
        ($shipped.Exchanges | Where-Object { $_.Code -eq 'XCSE' }).TimeZoneId | Should -Be 'Romance Standard Time'
    }
}

Describe 'Nasdaq Copenhagen session status' {
    # 2026-10-08 is a Thursday; Copenhagen is on CEST (UTC+2)
    Context 'Normal trading day' {
        It 'is PreMarket before 09:00 with a countdown to the open' {
            $s = Get-XcseStatus '2026-10-08T06:30:00'      # 08:30 local
            $s.Status | Should -Be 'PreMarket'
            $s.Text | Should -Be 'Opens in 00:30:00'
            $s.CountdownTarget | Should -Be 'Open'
        }

        It 'reports the opening auction, not pre-market, from 08:45' {
            (Get-XcseStatus '2026-10-08T06:44:59').Text | Should -Be 'Opens in 00:15:01'
            $s = Get-XcseStatus '2026-10-08T06:45:00'      # 08:45 local
            $s.Status | Should -Be 'PreMarket'
            $s.Text | Should -Be 'Opening auction - opens in 00:15:00'
            (Get-XcseStatus '2026-10-08T06:59:59').Text | Should -Be 'Opening auction - opens in 00:00:01'
            (Get-XcseStatus '2026-10-08T07:00:00').Status | Should -Be 'Open'
        }

        It 'counts down more than an hour without rounding the hours up' {
            $s = Get-XcseStatus '2026-10-08T00:24:00'      # 02:24 local, 6h36m to open
            $s.Text | Should -Be 'Opens in 06:36:00'
        }

        It 'is Open during continuous trading' {
            $s = Get-XcseStatus '2026-10-08T10:00:00'      # 12:00 local
            $s.Status | Should -Be 'Open'
            $s.Text | Should -Be 'Closes in 05:00:00'
            $s.CountdownTarget | Should -Be 'Close'
        }

        It 'is ClosingSoon within 30 minutes of the close' {
            $s = Get-XcseStatus '2026-10-08T14:40:00'      # 16:40 local
            $s.Status | Should -Be 'ClosingSoon'
            $s.Text | Should -Be 'Closes in 00:20:00'
        }

        It 'reports continuous trading one second before the closing auction' {
            $s = Get-XcseStatus '2026-10-08T14:54:59'      # 16:54:59 local
            $s.Text | Should -Be 'Closes in 00:05:01'
        }

        It 'reports the closing auction, not continuous trading, from 16:55' {
            $s = Get-XcseStatus '2026-10-08T14:55:00'      # 16:55 local
            $s.Status | Should -Be 'ClosingImminent'
            $s.Text | Should -Be 'Closing auction - ends in 00:05:00'
            $s = Get-XcseStatus '2026-10-08T14:57:30'      # 16:57:30 local
            $s.Text | Should -Be 'Closing auction - ends in 00:02:30'
        }

        It 'is Closed at and after 17:00' {
            (Get-XcseStatus '2026-10-08T15:00:00').Status | Should -Be 'Closed'
            $s = Get-XcseStatus '2026-10-08T19:00:00'
            $s.Status | Should -Be 'Closed'
            $s.TimeRemaining | Should -BeNullOrEmpty
            $s.CountdownTarget | Should -BeNullOrEmpty
        }
    }

    Context 'Weekends' {
        It 'is closed on Saturday' {
            $s = Get-XcseStatus '2026-10-10T10:00:00'
            $s.Status | Should -Be 'Closed'
            $s.Text | Should -Be 'Weekend'
        }

        It 'is closed on Sunday' {
            (Get-XcseStatus '2026-10-11T10:00:00').Text | Should -Be 'Weekend'
        }
    }

    Context 'Exchange holidays (official Nasdaq schedule)' {
        It 'ships the official 2025-2027 Copenhagen closures' {
            $expected = @(
                '2025-01-01', '2025-04-17', '2025-04-18', '2025-04-21', '2025-05-29', '2025-05-30',
                '2025-06-05', '2025-06-09', '2025-12-24', '2025-12-25', '2025-12-26', '2025-12-31',
                '2026-01-01', '2026-04-02', '2026-04-03', '2026-04-06', '2026-05-14', '2026-05-15',
                '2026-05-25', '2026-06-05', '2026-12-24', '2026-12-25', '2026-12-31',
                '2027-01-01', '2027-03-25', '2027-03-26', '2027-03-29', '2027-05-06', '2027-05-07',
                '2027-05-17', '2027-12-24', '2027-12-31'
            )
            $script:holidays['XCSE'] | Should -Be $expected
        }

        It 'lists only weekdays' {
            foreach ($d in $script:holidays['XCSE']) {
                ([datetime]::ParseExact($d, 'yyyy-MM-dd', [cultureinfo]::InvariantCulture)).DayOfWeek | Should -BeIn @('Monday', 'Tuesday', 'Wednesday', 'Thursday', 'Friday')
            }
        }

        It 'is a Holiday on Christmas Eve (an exchange closure, not a Danish public holiday)' {
            $s = Get-XcseStatus '2026-12-24T10:00:00'      # Thursday, CET
            $s.Status | Should -Be 'Holiday'
            $s.TimeRemaining | Should -BeNullOrEmpty
        }

        It 'is a Holiday on Constitution Day' {
            (Get-XcseStatus '2026-06-05T10:00:00').Status | Should -Be 'Holiday'
        }

        It 'is a Holiday on the day after Ascension' {
            (Get-XcseStatus '2026-05-15T10:00:00').Status | Should -Be 'Holiday'
        }

        It 'trades normally on the next listed trading day' {
            (Get-XcseStatus '2026-12-28T10:00:00').Status | Should -Be 'Open'   # Monday, 11:00 CET
        }

        It 'does not apply another exchange''s holidays to Copenhagen' {
            # 2026-01-06 is a Stockholm/Helsinki closure but not a Copenhagen one
            (Get-XcseStatus '2026-01-06T10:00:00').Status | Should -Be 'Open'
        }
    }

    Context 'CET/CEST transitions' {
        It 'opens at 08:00 UTC on the Friday before CEST starts (CET, UTC+1)' {
            (Get-XcseStatus '2026-03-27T07:30:00').Text | Should -Be 'Opens in 00:30:00'
            (Get-XcseStatus '2026-03-27T08:00:00').Status | Should -Be 'Open'
        }

        It 'opens at 07:00 UTC on the Monday after CEST starts (UTC+2)' {
            (Get-XcseStatus '2026-03-30T06:30:00').Text | Should -Be 'Opens in 00:30:00'
            (Get-XcseStatus '2026-03-30T07:00:00').Status | Should -Be 'Open'
            (Get-XcseStatus '2026-03-30T07:30:00').Status | Should -Be 'Open'
        }

        It 'closes at 15:00 UTC on the Friday before CEST ends' {
            (Get-XcseStatus '2026-10-23T15:00:00').Status | Should -Be 'Closed'
            (Get-XcseStatus '2026-10-23T14:57:00').Text | Should -Be 'Closing auction - ends in 00:03:00'
        }

        It 'closes at 16:00 UTC on the Monday after CEST ends (CET)' {
            (Get-XcseStatus '2026-10-26T15:00:00').Text | Should -Be 'Closes in 01:00:00'
            (Get-XcseStatus '2026-10-26T15:57:00').Text | Should -Be 'Closing auction - ends in 00:03:00'
            (Get-XcseStatus '2026-10-26T16:00:00').Status | Should -Be 'Closed'
        }

        It 'treats both 2026 transition Sundays as weekends' {
            (Get-XcseStatus '2026-03-29T10:00:00').Text | Should -Be 'Weekend'
            (Get-XcseStatus '2026-10-25T10:00:00').Text | Should -Be 'Weekend'
        }

        It 'measures countdowns in real elapsed time across a weekday DST change' {
            # EU transitions are on Sundays, so use a synthetic exchange in Israel,
            # where DST started on Friday 2026-03-27 at 02:00 (UTC+2 -> UTC+3).
            $synthetic = [PSCustomObject]@{
                Code = 'TEST'; TimeZoneId = 'Israel Standard Time'; OpenTimeLocal = '09:00'; CloseTimeLocal = '17:00'
            }
            # 01:00 local (UTC+2) -> 09:00 local (UTC+3) is 7 real hours, not 8 wall-clock hours
            $s = Get-ExchangeSessionStatus -Exchange $synthetic -UtcNow (Get-Utc '2026-03-26T23:00:00')
            $s.Text | Should -Be 'Opens in 07:00:00'
        }
    }

    Context 'Computer timezone differs from Copenhagen' {
        It 'gives the same result for a UTC instant passed as a local-kind DateTime' {
            $utc = Get-Utc '2026-10-08T10:00:00'
            $a = Get-ExchangeSessionStatus -Exchange $script:xcse -UtcNow $utc -Holidays $script:holidays
            $b = Get-ExchangeSessionStatus -Exchange $script:xcse -UtcNow $utc.ToLocalTime() -Holidays $script:holidays
            $b.Text | Should -Be $a.Text
            $b.Status | Should -Be $a.Status
        }

        It 'is independent of the process timezone (<_>)' -Skip:(-not ($IsLinux -or $IsMacOS)) -ForEach @('America/New_York', 'Asia/Tokyo', 'Pacific/Auckland') {
            $privatePath = Join-Path $script:moduleRoot 'Private/Get-ExchangeSessionStatus.ps1'
            $cmd = ". '$privatePath'; " +
            '$x = [PSCustomObject]@{ Code = ''XCSE''; TimeZoneId = ''Romance Standard Time''; OpenTimeLocal = ''09:00''; CloseTimeLocal = ''17:00''; ClosingAuctionStart = ''16:55'' }; ' +
            '$u = [DateTime]::SpecifyKind([DateTime]''2026-10-08T06:30:00'', ''Utc''); ' +
            '(Get-ExchangeSessionStatus -Exchange $x -UtcNow $u).Text + ''|'' + (Get-ExchangeSessionStatus -Exchange $x -UtcNow $u.ToLocalTime()).Text + ''|'' + [TimeZoneInfo]::Local.Id'
            $oldTz = $env:TZ
            try {
                $env:TZ = $_
                $out = & (Get-Process -Id $PID).Path -NoProfile -NonInteractive -Command $cmd
            }
            finally {
                $env:TZ = $oldTz
            }
            $parts = "$out".Trim() -split '\|'
            $parts[2] | Should -Be $_
            $parts[0] | Should -Be 'Opens in 00:30:00'
            $parts[1] | Should -Be 'Opens in 00:30:00'
        }
    }

    Context 'Missing or unavailable data' {
        It 'falls back to weekday and session-hours logic when holiday data is missing' {
            (Get-XcseStatus '2026-12-24T10:00:00' -Holidays @{}).Status | Should -Be 'Open'
            (Get-XcseStatus '2026-12-24T10:00:00' -Holidays $null).Status | Should -Be 'Open'
        }

        It 'falls back when holiday data has no entry for Copenhagen' {
            (Get-XcseStatus '2026-12-24T10:00:00' -Holidays @{ XNYS = @('2026-12-24') }).Status | Should -Be 'Open'
        }

        It 'does not throw for an unresolvable timezone ID' {
            $broken = [PSCustomObject]@{ Code = 'XCSE'; TimeZoneId = 'No Such Zone'; OpenTimeLocal = '09:00'; CloseTimeLocal = '17:00' }
            { Get-ExchangeSessionStatus -Exchange $broken -UtcNow (Get-Utc '2026-10-08T10:00:00') } | Should -Not -Throw
        }

        It 'returns no countdown while closed, so notifications are not raised' {
            foreach ($t in '2026-10-08T19:00:00', '2026-10-10T10:00:00', '2026-12-24T10:00:00') {
                $s = Get-XcseStatus $t
                $s.Status | Should -Not -BeIn @('Open', 'ClosingSoon', 'ClosingImminent')
                $s.TimeRemaining | Should -BeNullOrEmpty
            }
        }
    }

    Context 'Existing exchanges are unaffected by the closing-auction field' {
        It 'keeps the plain countdown text for exchanges without ClosingAuctionStart' {
            $xetr = $script:allExchanges | Where-Object { $_.Code -eq 'XETR' }
            # 2026-10-08 15:20 UTC = 17:20 CEST, 10 minutes before the 17:30 close
            $s = Get-ExchangeSessionStatus -Exchange $xetr -UtcNow (Get-Utc '2026-10-08T15:20:00')
            $s.Text | Should -Be 'Closes in 00:10:00'
            $s.Status | Should -Be 'ClosingSoon'
        }
    }
}

Describe 'Exchange selection and persistence' {
    It 'activates Copenhagen when it is in the saved selection (after a JSON round trip)' {
        $saved = @{ ActiveExchanges = @('XNYS', 'XCSE') } | ConvertTo-Json -Depth 3 | ConvertFrom-Json
        Test-ExchangeActive -Exchange $script:xcse -UserPreferences $saved | Should -BeTrue
    }

    It 'handles a saved selection containing only Copenhagen' {
        $saved = @{ ActiveExchanges = @('XCSE') } | ConvertTo-Json -Depth 3 | ConvertFrom-Json
        Test-ExchangeActive -Exchange $script:xcse -UserPreferences $saved | Should -BeTrue
    }

    It 'leaves an existing user selection unchanged (Copenhagen not auto-enabled)' {
        $example = Get-Content (Join-Path $script:dataDir 'user-preferences.example.json') -Raw | ConvertFrom-Json
        Test-ExchangeActive -Exchange $script:xcse -UserPreferences $example | Should -BeFalse
        $xnys = $script:allExchanges | Where-Object { $_.Code -eq 'XNYS' }
        Test-ExchangeActive -Exchange $xnys -UserPreferences $example | Should -BeTrue
    }

    It 'falls back to IsDefault when no preferences are saved' {
        Test-ExchangeActive -Exchange $script:xcse -UserPreferences $null | Should -BeFalse
        $xnys = $script:allExchanges | Where-Object { $_.Code -eq 'XNYS' }
        Test-ExchangeActive -Exchange $xnys -UserPreferences $null | Should -BeTrue
    }
}

Describe 'Get-StockExchangeData cache handling' {
    BeforeAll {
        Mock -ModuleName PSExchangeClock Invoke-WebRequest { throw 'offline' }
    }

    It 'appends Copenhagen to a fresh cache that predates it, without changing cached entries' {
        $path = Join-Path $TestDrive 'old-cache.json'
        [PSCustomObject]@{
            LastUpdated = (Get-Date).AddDays(-2).ToString('o')
            Source      = 'test'
            Version     = '1.0'
            Exchanges   = @([PSCustomObject]@{ Name = 'Customised NYSE'; Code = 'XNYS'; TimeZoneId = 'Eastern Standard Time'; OpenTimeLocal = '09:30'; CloseTimeLocal = '15:00' })
        } | ConvertTo-Json -Depth 5 | Set-Content $path

        $result = @(Get-StockExchangeData -OutputPath $path)
        ($result | Where-Object { $_.Code -eq 'XNYS' }).Name | Should -Be 'Customised NYSE'
        ($result | Where-Object { $_.Code -eq 'XNYS' }).CloseTimeLocal | Should -Be '15:00'
        ($result | Where-Object { $_.Code -eq 'XCSE' }).Name | Should -Be 'Nasdaq Copenhagen'
        @($result | Where-Object { $_.Code -eq 'XNYS' }).Count | Should -Be 1
        Should -Invoke -ModuleName PSExchangeClock Invoke-WebRequest -Times 0 -Exactly -Scope It
    }

    It 'falls back to built-in data, including Copenhagen, when offline with no cache' {
        $path = Join-Path $TestDrive 'none.json'
        $result = @(Get-StockExchangeData -OutputPath $path -WarningAction SilentlyContinue)
        $result.Code | Should -Contain 'XCSE'
        Test-Path $path | Should -BeTrue
    }
}

Describe 'OMXC25 index' {
    BeforeAll {
        $script:defs = @(Get-IndexDefinition)
        $script:omx = $script:defs | Where-Object { $_.Symbol -eq 'OMXC25' }

        # Sample Twelve Data /quote payload (shape from the Twelve Data API documentation)
        $script:samplePayload = [PSCustomObject]@{
            symbol         = 'OMXC25'
            name           = 'OMX Copenhagen 25 Index'
            currency       = 'DKK'
            datetime       = '2026-10-08'
            timestamp      = 1791442800      # 2026-10-08T07:00:00Z, opening bar
            last_quote_at  = 1791453600      # 2026-10-08T10:00:00Z
            close          = '1824.92'
            previous_close = '1800.03'
            is_market_open = $true
        }
    }

    Context 'Registration' {
        It 'is registered as a true index on Nasdaq Copenhagen, Denmark, in DKK' {
            $script:omx | Should -Not -BeNullOrEmpty
            $script:omx.Name | Should -Be 'OMX Copenhagen 25 (OMXC25)'
            $script:omx.Kind | Should -Be 'Index'
            $script:omx.Provider | Should -Be 'TwelveData'
            $script:omx.ExchangeCode | Should -Be 'XCSE'
            $script:omx.Country | Should -Be 'Denmark'
            $script:omx.Currency | Should -Be 'DKK'
        }

        It 'resolves its exchange code to Nasdaq Copenhagen' {
            $ex = $script:allExchanges | Where-Object { $_.Code -eq $script:omx.ExchangeCode }
            $ex.Name | Should -Be 'Nasdaq Copenhagen'
            $ex.Country | Should -Be $script:omx.Country
        }

        It 'is turned off by default with a note explaining the Pro plan requirement' {
            $script:omx.EnabledByDefault | Should -BeFalse
            $script:omx.DisabledReason | Should -Match 'Pro plan'
            $script:omx.DisabledReason | Should -Match 'EnabledIndices'
            Test-IndexEnabled -Definition $script:omx -UserPreferences $null | Should -BeFalse
        }

        It 'stays off for existing preference files without EnabledIndices' {
            $example = Get-Content (Join-Path $script:dataDir 'user-preferences.example.json') -Raw | ConvertFrom-Json
            Test-IndexEnabled -Definition $script:omx -UserPreferences $example | Should -BeFalse
        }

        It 'is enabled when listed in EnabledIndices (after a JSON round trip)' {
            $saved = @{ EnabledIndices = @('OMXC25') } | ConvertTo-Json -Depth 3 | ConvertFrom-Json
            Test-IndexEnabled -Definition $script:omx -UserPreferences $saved | Should -BeTrue
        }

        It 'leaves the ETF proxies enabled regardless of EnabledIndices' {
            $saved = @{ EnabledIndices = @('OMXC25') } | ConvertTo-Json -Depth 3 | ConvertFrom-Json
            foreach ($etf in @($script:defs | Where-Object { $_.Kind -eq 'EtfProxy' })) {
                Test-IndexEnabled -Definition $etf -UserPreferences $saved | Should -BeTrue
                Test-IndexEnabled -Definition $etf -UserPreferences $null | Should -BeTrue
            }
        }

        It 'keeps the existing ETF proxies unchanged and in order' {
            $etfs = @($script:defs | Where-Object { $_.Kind -eq 'EtfProxy' })
            $etfs.Symbol | Should -Be @('SPY', 'QQQ', 'DIA', 'EWU', 'EWG', 'EWJ', 'FXI', 'INDA')
            $etfs[0].Name | Should -Be 'S&P 500 (SPY)'
        }
    }

    Context 'Quote parsing' {
        It 'parses a valid quote' {
            $q = ConvertFrom-TwelveDataQuote -Response $script:samplePayload -Symbol 'OMXC25'
            $q.IsAvailable | Should -BeTrue
            $q.Price | Should -Be 1824.92
            $q.Change | Should -Be 1.38
            $q.QuoteTimeUtc | Should -Be (Get-Utc '2026-10-08T10:00:00')
        }

        It 'falls back to timestamp when last_quote_at is absent' {
            $p = $script:samplePayload.PSObject.Copy()
            $p.PSObject.Properties.Remove('last_quote_at')
            (ConvertFrom-TwelveDataQuote -Response $p -Symbol 'OMXC25').QuoteTimeUtc | Should -Be (Get-Utc '2026-10-08T07:00:00')
        }

        It 'treats a quote without a price as unavailable' {
            $p = $script:samplePayload.PSObject.Copy()
            $p.close = $null
            $q = ConvertFrom-TwelveDataQuote -Response $p -Symbol 'OMXC25'
            $q.IsAvailable | Should -BeFalse
            $q.Price | Should -BeNullOrEmpty
        }

        It 'treats an empty response as unavailable' {
            (ConvertFrom-TwelveDataQuote -Response $null -Symbol 'OMXC25').IsAvailable | Should -BeFalse
        }

        It 'marks HTTP 403 (plan upgrade required) as a permanent error' {
            $q = ConvertFrom-TwelveDataQuote -Response ([PSCustomObject]@{ code = 403; message = 'upgrade'; status = 'error' }) -Symbol 'OMXC25'
            $q.IsAvailable | Should -BeFalse
            $q.ErrorCode | Should -Be 403
            $q.IsPermanentError | Should -BeTrue
        }

        It 'marks HTTP 429 (rate limit) as transient' {
            $q = ConvertFrom-TwelveDataQuote -Response ([PSCustomObject]@{ code = 429; message = 'limit'; status = 'error' }) -Symbol 'OMXC25'
            $q.IsPermanentError | Should -BeFalse
        }
    }

    Context 'Retrieval through Twelve Data (mocked)' {
        It 'requests symbol OMXC25 from the /quote endpoint and parses the result' {
            Mock Invoke-RestMethod { $script:samplePayload }
            $q = Get-TwelveDataIndexQuote -Symbol 'OMXC25' -ApiKey 'test-key'
            $q.IsAvailable | Should -BeTrue
            $q.Price | Should -Be 1824.92
            Should -Invoke Invoke-RestMethod -Times 1 -Exactly -Scope It -ParameterFilter {
                $Uri -like 'https://api.twelvedata.com/quote?symbol=OMXC25&apikey=test-key'
            }
        }

        It 'returns an unavailable quote instead of throwing on a network failure' {
            Mock Invoke-RestMethod { throw [System.Net.Http.HttpRequestException]::new('No such host is known') }
            $q = Get-TwelveDataIndexQuote -Symbol 'OMXC25' -ApiKey 'test-key' -ErrorAction Stop
            $q.IsAvailable | Should -BeFalse
            $q.IsPermanentError | Should -BeFalse
            $q.ErrorMessage | Should -Match 'No such host'
        }
    }

    Context 'Cache, fallback and back-off' {
        It 'keeps the last good quote when a later refresh fails' {
            $cache = @{}
            Mock Get-TwelveDataIndexQuote { ConvertFrom-TwelveDataQuote -Response $script:samplePayload -Symbol 'OMXC25' }
            $e = Get-CachedIndexQuote -Definition $script:omx -ApiKey 'k' -Cache $cache -NowUtc (Get-Utc '2026-10-08T10:00:00')
            $e.LastGood.Price | Should -Be 1824.92

            Mock Get-TwelveDataIndexQuote { [PSCustomObject]@{ IsAvailable = $false; ErrorCode = $null; ErrorMessage = 'timeout'; IsPermanentError = $false } }
            $e = Get-CachedIndexQuote -Definition $script:omx -ApiKey 'k' -Cache $cache -NowUtc (Get-Utc '2026-10-08T10:20:00')
            $e.LastGood.Price | Should -Be 1824.92
            $e.LastError | Should -Be 'timeout'
            $e.Blocked | Should -BeFalse
        }

        It 'does not call the provider again within the 15-minute TTL' {
            $cache = @{}
            Mock Get-TwelveDataIndexQuote { ConvertFrom-TwelveDataQuote -Response $script:samplePayload -Symbol 'OMXC25' }
            Get-CachedIndexQuote -Definition $script:omx -ApiKey 'k' -Cache $cache -NowUtc (Get-Utc '2026-10-08T10:00:00') | Out-Null
            Get-CachedIndexQuote -Definition $script:omx -ApiKey 'k' -Cache $cache -NowUtc (Get-Utc '2026-10-08T10:10:00') | Out-Null
            Should -Invoke Get-TwelveDataIndexQuote -Times 1 -Exactly -Scope It
            Get-CachedIndexQuote -Definition $script:omx -ApiKey 'k' -Cache $cache -NowUtc (Get-Utc '2026-10-08T10:16:00') | Out-Null
            Should -Invoke Get-TwelveDataIndexQuote -Times 2 -Exactly -Scope It
        }

        It 'stops requesting after HTTP 403 for that key, but retries with a different key' {
            $cache = @{}
            Mock Get-TwelveDataIndexQuote { ConvertFrom-TwelveDataQuote -Response ([PSCustomObject]@{ code = 403; message = 'Pro plan'; status = 'error' }) -Symbol 'OMXC25' }
            $e = Get-CachedIndexQuote -Definition $script:omx -ApiKey 'basic' -Cache $cache -NowUtc (Get-Utc '2026-10-08T10:00:00')
            $e.Blocked | Should -BeTrue
            $e.ErrorCode | Should -Be 403
            Get-CachedIndexQuote -Definition $script:omx -ApiKey 'basic' -Cache $cache -NowUtc (Get-Utc '2026-10-08T12:00:00') | Out-Null
            Should -Invoke Get-TwelveDataIndexQuote -Times 1 -Exactly -Scope It
            Get-CachedIndexQuote -Definition $script:omx -ApiKey 'pro' -Cache $cache -NowUtc (Get-Utc '2026-10-08T12:00:00') | Out-Null
            Should -Invoke Get-TwelveDataIndexQuote -Times 2 -Exactly -Scope It
        }

        It 'retries after the TTL following a rate-limit error' {
            $cache = @{}
            Mock Get-TwelveDataIndexQuote { ConvertFrom-TwelveDataQuote -Response ([PSCustomObject]@{ code = 429; message = 'limit'; status = 'error' }) -Symbol 'OMXC25' }
            $e = Get-CachedIndexQuote -Definition $script:omx -ApiKey 'k' -Cache $cache -NowUtc (Get-Utc '2026-10-08T10:00:00')
            $e.Blocked | Should -BeFalse
            $e.LastGood | Should -BeNullOrEmpty
            Get-CachedIndexQuote -Definition $script:omx -ApiKey 'k' -Cache $cache -NowUtc (Get-Utc '2026-10-08T10:15:00') | Out-Null
            Should -Invoke Get-TwelveDataIndexQuote -Times 2 -Exactly -Scope It
        }
    }

    Context 'Quote status (live/delayed/stale/last close) against Copenhagen hours' {
        BeforeAll {
            function Get-OmxStatus([string]$NowIso, $QuoteIso) {
                $now = Get-Utc $NowIso
                $market = (Get-ExchangeSessionStatus -Exchange $script:xcse -UtcNow $now -Holidays $script:holidays).Status
                $qt = if ($QuoteIso) { Get-Utc $QuoteIso } else { $null }
                Get-IndexQuoteStatus -QuoteTimeUtc $qt -MarketStatus $market -NowUtc $now
            }
        }

        It 'is Intraday (never "live") for a fresh quote during the session' {
            $s = Get-OmxStatus '2026-10-08T10:10:00' '2026-10-08T10:00:00'
            $s.Status | Should -Be 'Intraday'
            $s.Label | Should -Not -Match 'live'
        }

        It 'is Intraday during the closing auction with a recent quote' {
            (Get-OmxStatus '2026-10-08T14:57:00' '2026-10-08T14:54:00').Status | Should -Be 'Intraday'
        }

        It 'is Stale for a delayed quote more than 30 minutes old during the session' {
            (Get-OmxStatus '2026-10-08T10:45:00' '2026-10-08T10:00:00').Status | Should -Be 'Stale'
        }

        It 'is Stale during the session when only yesterday''s close is available' {
            (Get-OmxStatus '2026-10-08T10:00:00' '2026-10-07T14:59:00').Status | Should -Be 'Stale'
        }

        It 'is LastClose after 17:00' {
            (Get-OmxStatus '2026-10-08T18:00:00' '2026-10-08T14:59:00').Status | Should -Be 'LastClose'
        }

        It 'is LastClose before the open' {
            (Get-OmxStatus '2026-10-09T06:00:00' '2026-10-08T14:59:00').Status | Should -Be 'LastClose'
        }

        It 'is LastClose across the Easter closure' {
            # Last 2026 session before Easter: Wed 1 Apr; Tue 7 Apr pre-open
            (Get-OmxStatus '2026-04-07T06:30:00' '2026-04-01T14:59:00').Status | Should -Be 'LastClose'
        }

        It 'is LastClose on a Copenhagen holiday even though Stockholm trades' {
            # 2026-05-15: Copenhagen closed (day after Ascension); Twelve Data files OMXC25 under XSTO
            (Get-OmxStatus '2026-05-15T10:00:00' '2026-05-13T14:59:00').Status | Should -Be 'LastClose'
        }

        It 'is Stale for a quote over seven days old while closed' {
            (Get-OmxStatus '2026-10-10T10:00:00' '2026-09-30T14:59:00').Status | Should -Be 'Stale'
        }

        It 'is Unavailable with no quote' {
            (Get-OmxStatus '2026-10-08T10:00:00' $null).Status | Should -Be 'Unavailable'
        }
    }
}

Describe 'FTSE 100 index' {
    BeforeAll {
        $script:ftse = Get-IndexDefinition | Where-Object { $_.Symbol -eq 'FTSE' }
        $script:xlon = $script:allExchanges | Where-Object { $_.Code -eq 'XLON' }
    }

    It 'is registered as a true index on the London Stock Exchange, in GBP' {
        $script:ftse | Should -Not -BeNullOrEmpty
        $script:ftse.Name | Should -Be 'FTSE 100'
        $script:ftse.Kind | Should -Be 'Index'
        $script:ftse.ExchangeCode | Should -Be 'XLON'
        $script:ftse.Country | Should -Be 'United Kingdom'
        $script:ftse.Currency | Should -Be 'GBP'
        $script:xlon.Name | Should -Be 'London Stock Exchange'
    }

    It 'is pinned to the XLON MIC because the bare ticker FTSE is ambiguous' {
        $script:ftse.MicCode | Should -Be 'XLON'
    }

    It 'is turned off by default with a note explaining the Grow plan requirement' {
        $script:ftse.RequiredPlan | Should -Be 'Grow'
        $script:ftse.EnabledByDefault | Should -BeFalse
        $script:ftse.DisabledReason | Should -Match 'Grow plan'
        $script:ftse.DisabledReason | Should -Match '"FTSE"'
        Test-IndexEnabled -Definition $script:ftse -UserPreferences $null | Should -BeFalse
    }

    It 'is enabled independently of OMXC25' {
        $omx = Get-IndexDefinition | Where-Object { $_.Symbol -eq 'OMXC25' }
        $saved = @{ EnabledIndices = @('FTSE') } | ConvertTo-Json -Depth 3 | ConvertFrom-Json
        Test-IndexEnabled -Definition $script:ftse -UserPreferences $saved | Should -BeTrue
        Test-IndexEnabled -Definition $omx -UserPreferences $saved | Should -BeFalse
    }

    It 'sends mic_code=XLON with the quote request' {
        Mock Invoke-RestMethod { [PSCustomObject]@{ symbol = 'FTSE'; close = '9120.50'; previous_close = '9100.00'; last_quote_at = 1791453600 } }
        $q = Get-TwelveDataIndexQuote -Symbol 'FTSE' -ApiKey 'test-key' -MicCode 'XLON'
        $q.IsAvailable | Should -BeTrue
        $q.Price | Should -Be 9120.5
        Should -Invoke Invoke-RestMethod -Times 1 -Exactly -Scope It -ParameterFilter {
            $Uri -eq 'https://api.twelvedata.com/quote?symbol=FTSE&apikey=test-key&mic_code=XLON'
        }
    }

    It 'passes the MIC through the cache layer' {
        Mock Get-TwelveDataIndexQuote { [PSCustomObject]@{ IsAvailable = $false; ErrorCode = 403; ErrorMessage = 'Grow'; IsPermanentError = $true } }
        $e = Get-CachedIndexQuote -Definition $script:ftse -ApiKey 'basic' -Cache @{} -NowUtc (Get-Utc '2026-10-08T10:00:00')
        $e.Blocked | Should -BeTrue
        Should -Invoke Get-TwelveDataIndexQuote -Times 1 -Exactly -Scope It -ParameterFilter { $Symbol -eq 'FTSE' -and $MicCode -eq 'XLON' }
    }

    It 'judges quote status against London hours' {
        # 2026-10-08 is a Thursday; London is on BST (UTC+1); LSE closes 16:30 local = 15:30 UTC
        $open = (Get-ExchangeSessionStatus -Exchange $script:xlon -UtcNow (Get-Utc '2026-10-08T15:00:00')).Status
        Get-IndexQuoteStatus -QuoteTimeUtc (Get-Utc '2026-10-08T14:55:00') -MarketStatus $open -NowUtc (Get-Utc '2026-10-08T15:00:00') |
            ForEach-Object { $_.Status } | Should -Be 'Intraday'
        $closed = (Get-ExchangeSessionStatus -Exchange $script:xlon -UtcNow (Get-Utc '2026-10-08T15:45:00')).Status
        $closed | Should -Be 'Closed'
        (Get-IndexQuoteStatus -QuoteTimeUtc (Get-Utc '2026-10-08T15:29:00') -MarketStatus $closed -NowUtc (Get-Utc '2026-10-08T15:45:00')).Status | Should -Be 'LastClose'
    }
}
