# PSExchangeClock

[![PowerShell Gallery Version](https://img.shields.io/powershellgallery/v/PSExchangeClock)](https://www.powershellgallery.com/packages/PSExchangeClock)
[![PowerShell Gallery Downloads](https://img.shields.io/powershellgallery/dt/PSExchangeClock)](https://www.powershellgallery.com/packages/PSExchangeClock)
[![License: MIT](https://img.shields.io/badge/License-MIT-yellow.svg)](LICENSE)
[![Lint](https://github.com/PowerShellYoungTeam/PSExchangeClock/actions/workflows/lint.yml/badge.svg)](https://github.com/PowerShellYoungTeam/PSExchangeClock/actions/workflows/lint.yml)

A real-time WPF dashboard for monitoring global stock exchange closing times, live market data, forex, cryptocurrency, commodities, indices, and financial news — built entirely in PowerShell.

<!-- SCREENSHOT: Full dashboard with Dashboard tab active showing world map, countdown cards, world clocks, and market data sidebar. Capture at a resolution of ~1100x750. -->

## Features

- **Live countdown timers** with color-coded status (green → yellow → red) for 20 major stock exchanges worldwide
- **Interactive world map** with clickable exchange markers — NASA Earth at Night satellite imagery or vector continent outlines, flat and globe projections, zoom & pan
- **Map overlays** toggled from the toolbar:
  - *Geo* — day/night terminator, timezone bands, political borders, timezone boundaries
  - *Live data* — earthquakes (USGS), volcanoes (Smithsonian GVP), conflict zones (UCDP/PRIO)
  - *Infrastructure* — submarine cables, power plants colour-coded by fuel type
  - Auto-displayed legend for conflict zones, power plants, and volcanoes
- **World clocks** showing current times across six base time zones plus dynamic zones for active exchanges
- **Market data sidebar** with tabbed content:
  - Breaking financial news (Reuters, BBC, CNBC RSS feeds)
  - Live forex rates (Frankfurter API / ECB XML fallback)
  - Cryptocurrency prices (CoinGecko API — top 10 coins)
  - Stock indices (S&P 500, NASDAQ Composite, DAX, FTSE 100, Nikkei 225 via ETF proxies)
  - Commodity prices (Gold, Silver, Oil, Natural Gas, Wheat, Platinum)
  - Individual stock quotes with company profiles and statistics
- **Desktop notifications** at configurable thresholds (30, 15, 5 min) before exchange close
- **Secure API key management** — Windows Credential Manager, SecretManagement module, or CliXml encrypted storage
- **Lunch break tracking** for Tokyo, Shanghai, and Hong Kong exchanges
- **Holiday support** via editable `holidays.json` for market closure overrides
- **System tray integration** with minimize-to-tray and always-on-top toggle
- **Professional dark-themed UI** with sortable/filterable exchange DataGrid
- **User preference persistence** — active exchanges, notification settings, preferred currencies
- **Multi-currency support** for forex, crypto, and commodity pricing
- **Offline fallback** — hardcoded exchange data works without internet; market data requires connectivity

## Quick Start

```powershell
Install-Module PSExchangeClock
Start-PSExchangeClock
```

Or use the full function name:

```powershell
New-StockExchangeCountdownDashboard
```

## Requirements

- **Windows PowerShell 5.1** or **PowerShell 7+** on Windows (WPF requires Windows)
- **.NET Framework 4.5+** (for WPF assemblies)
- **Internet connection** for market data APIs (exchange hours work offline)

### Optional API Keys

API keys unlock additional market data tabs. Both services offer free tiers:

| Service | Free Tier | Used For |
|---------|-----------|----------|
| [Twelve Data](https://twelvedata.com/) | 800 req/day, 8 req/min | Stock indices, commodities, company profiles & statistics |
| [Alpha Vantage](https://www.alphavantage.co/) | 25 req/day | Individual stock quotes (open/close/high/low/volume) |

Enter your keys in the **Settings** tab of the dashboard, or run `New-StockExchangeCountdownDashboard` — the tool works without keys (news, forex, and crypto are free and unlimited).

## Installation

### From PowerShell Gallery (recommended)

```powershell
Install-Module PSExchangeClock -Scope CurrentUser
```

### From Source

```powershell
git clone https://github.com/PowerShellYoungTeam/PSExchangeClock.git
Import-Module .\PSExchangeClock\PSExchangeClock
```

## Configuration

User preferences are stored in `%APPDATA%\PSExchangeClock\user-preferences.json` and are created automatically on first run. You can also change settings via the **Settings** tab in the dashboard.

### Preferences

| Setting | Description | Default |
|---------|-------------|---------|
| `ActiveExchanges` | Array of exchange codes to display | All 20 exchanges |
| `NotifyThresholds` | Minutes before close to show notifications | `[30, 15, 5]` |
| `AlwaysOnTop` | Keep dashboard window above other windows | `false` |
| `CommodityBaseCurrency` | Currency for commodity price display | `EUR` |
| `SecretsBackend` | Where API keys are stored | `CredentialManager` |

### Holiday Overrides

Edit the `holidays.json` file in the module's `Data` directory to add market holidays. Each exchange code maps to an array of ISO 8601 dates:

```json
{
    "XNYS": ["2026-01-01", "2026-01-19", "2026-02-16"],
    "XLON": ["2026-01-01", "2026-04-03"]
}
```

## Supported Exchanges

| Symbol | Exchange | City | Timezone |
|--------|----------|------|----------|
| NYSE | New York Stock Exchange | New York | Eastern |
| NASDAQ | NASDAQ | New York | Eastern |
| LSE | London Stock Exchange | London | GMT/BST |
| EPA | Euronext Paris | Paris | CET/CEST |
| TSE | Tokyo Stock Exchange | Tokyo | JST |
| SSE | Shanghai Stock Exchange | Shanghai | CST |
| HKEX | Hong Kong Stock Exchange | Hong Kong | HKT |
| TSX | Toronto Stock Exchange | Toronto | Eastern |
| FRA | Frankfurt Stock Exchange | Frankfurt | CET/CEST |
| ASX | Australian Securities Exchange | Sydney | AEST/AEDT |
| BSE | Bombay Stock Exchange | Mumbai | IST |
| NSE | National Stock Exchange | Mumbai | IST |
| KRX | Korea Exchange | Seoul | KST |
| SIX | SIX Swiss Exchange | Zurich | CET/CEST |
| JSE | Johannesburg Stock Exchange | Johannesburg | SAST |
| B3 | B3 (Brasil Bolsa Balcao) | Sao Paulo | BRT |
| BMV | Bolsa Mexicana de Valores | Mexico City | CST |
| SGX | Singapore Exchange | Singapore | SGT |
| TWSE | Taiwan Stock Exchange | Taipei | CST |
| MOEX | Moscow Exchange | Moscow | MSK |

## Data Sources

| Data | Source | Rate Limits |
|------|--------|-------------|
| Exchange hours | Wikipedia scrape / hardcoded fallback | N/A |
| Financial news | Reuters, BBC, CNBC RSS feeds | Unlimited |
| Forex rates | [Frankfurter API](https://www.frankfurter.app/) / ECB XML | Unlimited |
| Cryptocurrency | [CoinGecko API](https://www.coingecko.com/) | Unlimited |
| Indices & Commodities | [Twelve Data API](https://twelvedata.com/) | 800 req/day |
| Stock quotes | [Alpha Vantage API](https://www.alphavantage.co/) | 25 req/day |
| World map image | [NASA GSFC Earth at Night](https://earthobservatory.nasa.gov/) | One-time download |
| Earthquakes | [USGS Earthquake Hazards](https://earthquake.usgs.gov/) | Unlimited |
| Volcanoes | [Smithsonian GVP](https://volcano.si.edu/) | Unlimited |
| Conflict zones | [UCDP/PRIO](https://ucdp.uu.se/) | Bundled data |
| Power plants | [WRI Global Power Plant Database](https://datasets.wri.org/datasets/global-power-plant-database) | Bundled data |
| Submarine cables | [TeleGeography](https://www.submarinecablemap.com/) | Bundled data |

## Refreshing Exchange Data

The module includes hardcoded fallback data for all 20 exchanges. To pull the latest trading hours from Wikipedia:

```powershell
Get-StockExchangeData -ForceRefresh
```

## Contributing

Contributions are welcome! Please open an issue or submit a pull request.

1. Fork the repository
2. Create a feature branch (`git checkout -b feature/my-feature`)
3. Commit your changes (`git commit -am 'Add my feature'`)
4. Push to the branch (`git push origin feature/my-feature`)
5. Open a Pull Request

## License

This project is licensed under the MIT License — see the [LICENSE](LICENSE) file for details.
