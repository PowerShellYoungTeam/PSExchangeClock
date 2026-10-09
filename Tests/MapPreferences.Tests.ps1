#requires -Version 7.0
#requires -Modules @{ ModuleName = 'Pester'; ModuleVersion = '5.0.0' }

BeforeAll {
    $script:moduleRoot = Join-Path (Split-Path $PSScriptRoot -Parent) 'PSExchangeClock'
    . (Join-Path $script:moduleRoot 'Private/Get-MapPreferenceState.ps1')

    # Mirrors the option values and toggle names in New-StockExchangeCountdownDashboard
    $script:styles = @('NightSatellite', 'BlueMarble', 'VectorDark', 'Minimal')
    $script:projections = @('Flat', 'Globe')
    $script:overlays = @('DayNight', 'TimeZoneBands', 'PoliticalBorders', 'TimeZoneBoundaries', 'Earthquakes',
        'Volcanoes', 'NaturalEvents', 'ConflictZones', 'SubmarineCables', 'PowerPlants')

    function Resolve-Map($Prefs) {
        Get-MapPreferenceState -UserPreferences $Prefs -ValidStyles $script:styles -ValidProjections $script:projections -OverlayNames $script:overlays
    }
    function ConvertTo-SavedPrefs([hashtable]$Prefs) {
        # Same serialisation as Save-UserPreferences
        $Prefs | ConvertTo-Json -Depth 3 | ConvertFrom-Json
    }
}

Describe 'Map preference restore' {
    It 'restores style, projection and overlays after a save/load round trip' {
        $saved = ConvertTo-SavedPrefs @{ MapStyle = 'BlueMarble'; MapProjection = 'Globe'; MapOverlays = @('Earthquakes', 'PowerPlants') }
        $s = Resolve-Map $saved
        $s.Style | Should -Be 'BlueMarble'
        $s.Projection | Should -Be 'Globe'
        $s.Overlays | Should -Be @('Earthquakes', 'PowerPlants')
    }

    It 'restores a single saved overlay' {
        $s = Resolve-Map (ConvertTo-SavedPrefs @{ MapOverlays = @('Volcanoes') })
        $s.Overlays | Should -Be @('Volcanoes')
    }

    It 'treats an empty saved overlay list as all overlays off' {
        $s = Resolve-Map (ConvertTo-SavedPrefs @{ MapOverlays = @() })
        # An empty list is "all off", not "unset"
        $null -eq $s.Overlays | Should -BeFalse
        @($s.Overlays).Count | Should -Be 0
    }

    It 'keeps the built-in defaults when nothing is saved' {
        $s = Resolve-Map (ConvertTo-SavedPrefs @{ ActiveExchanges = @('XNYS') })
        $s.Style | Should -BeNullOrEmpty
        $s.Projection | Should -BeNullOrEmpty
        $null -eq $s.Overlays | Should -BeTrue
    }

    It 'keeps the built-in defaults when there is no preferences file' {
        $s = Resolve-Map $null
        $s.Style | Should -BeNullOrEmpty
        $null -eq $s.Overlays | Should -BeTrue
    }

    It 'ignores unknown style and projection values' {
        $s = Resolve-Map (ConvertTo-SavedPrefs @{ MapStyle = 'Sepia'; MapProjection = 'Mercator' })
        $s.Style | Should -BeNullOrEmpty
        $s.Projection | Should -BeNullOrEmpty
    }

    It 'ignores overlay names with no toggle, such as TimeLine from older example files' {
        $s = Resolve-Map (ConvertTo-SavedPrefs @{ MapOverlays = @('DayNight', 'TimeLine') })
        $s.Overlays | Should -Be @('DayNight')
    }

    It 'loads the shipped example preferences' {
        $example = Get-Content (Join-Path $script:moduleRoot 'Data/user-preferences.example.json') -Raw | ConvertFrom-Json
        $s = Resolve-Map $example
        $s.Style | Should -Be 'NightSatellite'
        $s.Projection | Should -Be 'Flat'
        $s.Overlays | Should -Be @('DayNight')
    }
}
