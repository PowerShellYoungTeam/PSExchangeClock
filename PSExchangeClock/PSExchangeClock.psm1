# PSExchangeClock module loader
$PSModuleRoot = $PSScriptRoot

# Dot-source all public functions
$publicFunctions = Get-ChildItem -Path "$PSModuleRoot\Public\*.ps1" -ErrorAction SilentlyContinue
foreach ($function in $publicFunctions) {
    try {
        . $function.FullName
    }
    catch {
        Write-Error "Failed to import function $($function.BaseName): $_"
    }
}

# Register alias
Set-Alias -Name 'Start-PSExchangeClock' -Value 'New-StockExchangeCountdownDashboard'

# Export public functions and alias
Export-ModuleMember -Function $publicFunctions.BaseName -Alias 'Start-PSExchangeClock'
