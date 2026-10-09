function Get-MapPreferenceState {
    <#
    .SYNOPSIS
        Resolves the saved map style, projection and overlays from user preferences.
    .DESCRIPTION
        Unknown or missing values resolve to $null so the caller keeps its own
        defaults. Overlays resolve to $null only when MapOverlays is absent; an
        empty saved list means "all overlays off". Unknown overlay names (for
        example 'TimeLine' from older example files, which has no toggle) are
        ignored.
    .OUTPUTS
        Hashtable with Style, Projection and Overlays.
    #>
    [CmdletBinding()]
    [OutputType([hashtable])]
    param(
        $UserPreferences,
        [string[]]$ValidStyles = @(),
        [string[]]$ValidProjections = @(),
        [string[]]$OverlayNames = @()
    )

    $state = @{ Style = $null; Projection = $null; Overlays = $null }
    if (-not $UserPreferences) { return $state }

    if ($UserPreferences.MapStyle -and $ValidStyles -contains $UserPreferences.MapStyle) {
        $state.Style = [string]$UserPreferences.MapStyle
    }
    if ($UserPreferences.MapProjection -and $ValidProjections -contains $UserPreferences.MapProjection) {
        $state.Projection = [string]$UserPreferences.MapProjection
    }
    if ($UserPreferences.PSObject.Properties['MapOverlays']) {
        $saved = @($UserPreferences.MapOverlays | Where-Object { $_ })
        $state.Overlays = @($OverlayNames | Where-Object { $saved -contains $_ })
    }
    return $state
}
