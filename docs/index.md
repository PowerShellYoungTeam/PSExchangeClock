---
layout: default
title: PSExchangeClock
---

# PSExchangeClock

**A real-time WPF dashboard for monitoring 20 global stock exchange closing times, live market data, forex, cryptocurrency, commodities, indices, and financial news — built entirely in PowerShell.**

[![PowerShell Gallery Version](https://img.shields.io/powershellgallery/v/PSExchangeClock)](https://www.powershellgallery.com/packages/PSExchangeClock)
[![PowerShell Gallery Downloads](https://img.shields.io/powershellgallery/dt/PSExchangeClock)](https://www.powershellgallery.com/packages/PSExchangeClock)
[![License: MIT](https://img.shields.io/badge/License-MIT-yellow.svg)](https://github.com/PowerShellYoungTeam/PSExchangeClock/blob/main/LICENSE)

## Quick Start

```powershell
Install-Module PSExchangeClock -Scope CurrentUser
Start-PSExchangeClock
```

## What it does

- **Live countdown timers** with color-coded status (green → yellow → red → purple for holidays) for 20 major exchanges worldwide
- **Interactive world map** — NASA Earth at Night imagery, flat and globe projections, zoom & pan, clickable exchange markers
- **Map overlays** — day/night terminator, timezone bands and boundaries, political borders, earthquakes, volcanoes, natural events, conflict zones, submarine cables, and power plants
- **Market data sidebar** — catalog-driven financial news with source/category filters and manual refresh, live forex, crypto, indices, commodities, and individual stock quotes
- **World clocks**, desktop notifications, holiday support, and secure API key storage (Windows Credential Manager, SecretManagement, or CliXml)

## The origin story

What started in 2021 as an 80-line countdown timer for the London Stock Exchange close grew — one exchange, one overlay, one feed at a time — into a full global markets dashboard, still just PowerShell and WPF, no compiled code.

## Links

- [Source on GitHub](https://github.com/PowerShellYoungTeam/PSExchangeClock)
- [PSExchangeClock on the PowerShell Gallery](https://www.powershellgallery.com/packages/PSExchangeClock)
- [Changelog](https://github.com/PowerShellYoungTeam/PSExchangeClock/blob/main/CHANGELOG.md)

Star the repo if you find it useful — PRs are always welcome.
