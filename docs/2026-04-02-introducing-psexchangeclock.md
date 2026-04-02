---
layout: post
title: "Introducing PSExchangeClock — A Global Market Dashboard Built in PowerShell"
date: 2026-04-02
categories: [PowerShell, Projects]
tags: [PowerShell, WPF, Finance, Trading, Open-Source]
---

What started as a quick countdown timer for the London Stock Exchange close has turned into a full-blown global financial markets dashboard — and it's all built in PowerShell.

Today I'm releasing **PSExchangeClock** on the [PowerShell Gallery](https://www.powershellgallery.com/packages/PSExchangeClock).

## The Origin Story

Back in 2021, I was supporting traders and needed a simple way to see when the LSE was closing. So I threw together a wee WPF timer in PowerShell — about 80 lines, hardcoded date, job done.

Then I needed NASDAQ and NYSE. Then Frankfurt and Tokyo. Then I thought "wouldn't a world map be cool?" and before I knew it, I had a 3,500-line WPF dashboard covering 20 exchanges across every major timezone.

It was time to give it a proper name and ship it as a module.

## What It Does

PSExchangeClock is a real-time WPF dashboard that monitors global stock exchange closing times alongside live market data.

<!-- SCREENSHOT: Full dashboard view showing the Dashboard tab with world map, countdown cards visible, world clocks on the side, and the market data panel. Aim for ~1100x750 window size. -->

### Countdown Timers

The core feature — live countdown timers for 20 major exchanges worldwide. Each card shows the exchange status with color transitions:

- **Green** — market open, plenty of time
- **Yellow** — closing within 30 minutes
- **Red** — closing within 15 minutes
- **Purple** — market holiday

Desktop notifications pop up at 30, 15, and 5 minutes before close (configurable).

<!-- SCREENSHOT: Close-up of 3-4 countdown cards showing different states — one green (open), one yellow (closing soon), one red (closing imminent), and one grey (closed). -->

### Interactive World Map

Click any exchange marker on the NASA Earth at Night satellite map to see a flyout panel with full exchange details — primary index, market cap, listed companies, currency, regulator, description, and live status.

<!-- SCREENSHOT: World map tab with at least 2-3 visible exchange markers and one flyout panel open showing exchange details. -->

### Map Overlays

The map toolbar includes a toggleable overlay system with three categories:

- **Geo layers** — day/night terminator, timezone bands, political borders, and timezone boundaries
- **Live data** — earthquakes from USGS, volcano alerts from the Smithsonian GVP, and armed conflict zones from UCDP/PRIO
- **Infrastructure** — submarine cable routes and major power plants colour-coded by fuel type (nuclear, coal, gas, hydro, wind, solar, geothermal, oil)

When conflict zones, power plants, or volcanoes are active, an overlay legend appears at the bottom-left of the map showing what the markers and colours mean.

<!-- SCREENSHOT: World map with overlays enabled — show power plants or conflict zones with the legend visible in the bottom-left corner. -->

### Market Data Sidebar

Six tabbed panels of live data, no terminal required:

- **News** — Financial headlines from Reuters, BBC, and CNBC RSS feeds
- **Forex** — Live currency rates via Frankfurter API
- **Crypto** — Top 10 cryptocurrency prices from CoinGecko
- **Indices** — S&P 500, NASDAQ, DAX, FTSE 100, Nikkei 225 (via ETF proxies)
- **Commodities** — Gold, Silver, Oil, Natural Gas, Wheat, Platinum
- **Stocks** — Individual quote lookup with company profiles and statistics

<!-- SCREENSHOT: Market data sidebar showing one of the tabs active (Forex or Crypto recommended for visual appeal). -->

### All Exchanges Tab

A sortable, filterable DataGrid showing all 20 exchanges with live status indicators. Select/deselect exchanges to customise your dashboard.

<!-- SCREENSHOT: The All Exchanges tab showing the DataGrid with status indicators, a mix of Open/Closed statuses visible. -->

### Settings

Configure everything from the UI — notification thresholds, API keys (securely stored via Windows Credential Manager), display preferences, and secrets backend.

<!-- SCREENSHOT: Settings tab showing the notification checkboxes, API key fields (with keys masked/empty), and secrets backend dropdown. -->

## Install It

```powershell
Install-Module PSExchangeClock -Scope CurrentUser
Start-PSExchangeClock
```

That's it. Two commands. The dashboard launches, downloads the NASA map image on first run, and you're monitoring 20 exchanges.

For the full feature set (indices, commodities, stock quotes), grab free API keys from [Twelve Data](https://twelvedata.com/) and [Alpha Vantage](https://www.alphavantage.co/) and enter them in the Settings tab. News, forex, and crypto work without any keys.

## Technical Bits

- **~4,500 lines of PowerShell + WPF XAML** — no compiled code, no C#, just scripts
- **5 API integrations** — Frankfurter, CoinGecko, Twelve Data, Alpha Vantage, Wikipedia
- **3 secure storage backends** — Windows Credential Manager, SecretManagement module, CliXml
- **Offline fallback** — hardcoded exchange data for all 20 exchanges
- **CI/CD** — PSScriptAnalyzer linting on every PR, automatic PSGallery publish on version tags

## Links

- **GitHub**: [PowerShellYoungTeam/PSExchangeClock](https://github.com/PowerShellYoungTeam/PSExchangeClock)
- **PowerShell Gallery**: [PSExchangeClock](https://www.powershellgallery.com/packages/PSExchangeClock)

Give it a try, star the repo if you find it useful, and PRs are always welcome.
