@{
    # Module manifest for PSExchangeClock
    RootModule        = 'PSExchangeClock.psm1'
    ModuleVersion     = '1.0.0'
    GUID              = 'a3e4f8b2-7c1d-4e5a-9f2b-8d6c3a1e7f04'
    Author            = 'PowerShellYoungTeam'
    CompanyName       = 'PowerShellYoungTeam'
    Copyright         = '(c) 2026 PowerShellYoungTeam. All rights reserved.'
    Description       = 'A real-time WPF dashboard for monitoring global stock exchange closing times, market data, forex, crypto, commodities, indices, and financial news — covering 20 major exchanges worldwide.'
    PowerShellVersion = '5.1'
    CLRVersion        = '4.0'

    # Functions and aliases exported from this module
    FunctionsToExport = @(
        'New-StockExchangeCountdownDashboard',
        'Get-StockExchangeData'
    )
    AliasesToExport   = @(
        'Start-PSExchangeClock'
    )
    CmdletsToExport   = @()
    VariablesToExport  = @()

    # Private data for PSGallery publishing
    PrivateData       = @{
        PSData = @{
            Tags         = @('Stock', 'Exchange', 'Dashboard', 'WPF', 'Finance', 'Market', 'Trading', 'Countdown', 'Forex', 'Crypto', 'Commodities')
            LicenseUri   = 'https://github.com/PowerShellYoungTeam/PSExchangeClock/blob/main/LICENSE'
            ProjectUri   = 'https://github.com/PowerShellYoungTeam/PSExchangeClock'
            ReleaseNotes = 'v1.0.0 - Initial release as a PowerShell module. Rebranded from Countdown-Gui-s.'
        }
    }
}
