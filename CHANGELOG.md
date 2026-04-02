# Changelog

All notable changes to PSExchangeClock will be documented in this file.

The format is based on [Keep a Changelog](https://keepachangelog.com/en/1.1.0/),
and this project adheres to [Semantic Versioning](https://semver.org/spec/v2.0.0.html).

## [1.1.0] - 2026-07-15

### Fixed

- Overlapping exchange markers (XNYS/XNAS in New York, XBOM/XNSE in Mumbai) now render as cluster markers with a count badge; click to expand into a radial fan layout

### Added

- **Zoom & pan** — mouse-wheel zoom (1×–8×), right-drag to pan, reset button
- **Solar terminator** — semi-transparent day/night overlay updated every 5 minutes
- **Current time line** — vertical UTC time indicator on the flat map
- **Map style selector** — Night Satellite (default), Blue Marble, Vector Dark, and Minimal styles
- **Projection selector** — switch between Flat (equirectangular) and Globe (orthographic) views
- **Globe view** — orthographic projection with left-drag rotation, graticule lines, and continent outlines
- **Timezone bands overlay** — alternating colour strips with UTC offset labels
- **Trading sessions overlay** — Asia, Europe, and Americas session bands
- Map toolbar with controls for style, projection, overlays, and zoom
- 5-minute overlay refresh timer for terminator and time line

### Changed

- `Initialize-WorldMap` rewritten to support multiple styles, projections, overlays, and clustering
- `Convert-LatLonToCanvas` updated with globe-mode routing and visibility flag
- `Add-MapMarker` updated with offset and expanded-fan support

## [1.0.0] - 2026-04-02

### Added

- Rebranded from **Countdown-Gui-s** to **PSExchangeClock** and restructured as a proper PowerShell module
- Module manifest (`.psd1`) and loader (`.psm1`) for `Install-Module` support
- `Start-PSExchangeClock` alias for `New-StockExchangeCountdownDashboard`
- User preferences now stored in `%APPDATA%\PSExchangeClock\` (module directory may be read-only)
- Comprehensive `README.md` with full feature documentation, API setup guide, and exchange table
- `about_PSExchangeClock` help topic (`Get-Help about_PSExchangeClock`)
- Expanded comment-based help for all exported functions (`.SYNOPSIS`, `.DESCRIPTION`, `.EXAMPLE`, `.NOTES`, `.LINK`)
- `CHANGELOG.md` (this file)
- MIT License
- GitHub Actions CI/CD:
  - `lint.yml` — PSScriptAnalyzer on push and pull requests
  - `publish.yml` — Automatic publish to PowerShell Gallery on version tags (`v*`)

### Removed

- Legacy scripts: `Countdown-LSEClose.ps1`, `Countdown-Shutdown.Ps1`, `New-StockExchangeCloseCountdown.ps1`, `New-StockExchangeCloseCountdownmk2.ps1`

### Changed

- Window title: "Stock Exchange Countdown Dashboard" → "PSExchangeClock — Global Market Dashboard"
- Data files moved to `PSExchangeClock/Data/` subdirectory
- Exchange data cache and map image now stored alongside user preferences in AppData

### Features (carried forward from pre-module development)

- Live countdown timers with color transitions for 20 global stock exchanges
- Interactive world map with clickable markers (NASA Earth at Night imagery)
- World clocks across multiple time zones
- Market data sidebar: financial news, forex, crypto, indices, commodities, stock quotes
- Desktop notifications at configurable thresholds before exchange close
- Secure API key management (Credential Manager / SecretManagement / CliXml)
- Lunch break tracking (Tokyo, Shanghai, Hong Kong)
- Holiday support via `holidays.json`
- System tray integration, always-on-top toggle
- Sortable/filterable exchange DataGrid
- User preference persistence and multi-currency support
