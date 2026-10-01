# bambu-sheets

*(Live Bambu Lab filament catalog in Google Sheets)*

The [Bambu Lab US store](https://us.store.bambulab.com/collections/bambu-lab-3d-printer-filament) has 40+ filaments across dozens of colors, but there's no easy way to compare specs, check stock across variants, or track prices over time.

Enter bambu-sheets: a Google Apps Script that pulls live data from the Shopify API and builds a structured, always-current filament catalog directly in Google Sheets.

## Features

- Live prices and stock status fetched from the Shopify storefront API.
- Per-color stock breakdown with collapsible row groups — PLA Basic's 30+ colors don't dominate the sheet.
- Static spec data (nozzle/bed/dry temps, tensile strength, HDT, AMS compatibility, etc.) baked in for all 43 filaments.
- Product thumbnail images via `=IMAGE()` formula.
- Fast refresh mode (prices & stock only) and full rebuild mode.
- Expand/collapse all color groups from the menu.

### General Flow

- On first run, `Authorize & Setup` creates the sheet structure and grants OAuth scopes.
- `Refresh All` fetches every product from Shopify, writes summary rows + collapsed color groups, and applies stock/price styling.
- `Refresh Prices & Stock` is the fast path — it scans existing rows and updates only prices and stock, preserving the sheet structure.

### Sheet Layout

Each filament gets a **summary row** (always visible) showing specs, price, and aggregate stock status. Below it, a **collapsed row group** contains one row per color variant with individual stock status. Click the `+` on the left margin to expand.

## Installation

```
npm install -g @google/clasp
npm install
```

## First Deployment

### Setup Project / Clasp

- [Create a new Google Spreadsheet](https://sheets.new)
- Open **Extensions → Apps Script**
- Get the `Script ID` from **Project Settings** (gear icon)
- Add `Script ID` to `.clasp.json`
- Run `clasp login` in your terminal

### Enable the Sheets Advanced Service

- In the Apps Script editor, go to **Services** (+ icon)
- Add **Google Sheets API** (v4)
- This is already configured in `appsscript.json` — you just need to authorize it

### Deploy

```
npm run build
npm run deploy
```

### First Run

- Reload the spreadsheet — a **🧵 Bambu Filaments** menu will appear
- Run **Authorize & Setup** once to grant OAuth and create headers
- Run **Refresh All** to populate the sheet

## Development

Run the following to build, watch, and auto-deploy on changes:

```
npm run watch
```

Other commands in `package.json`.

### Project Structure

```
src/
  lib/
    types.ts    — Shopify API types, PriceInfo, StockInfo, RowMarker
    config.ts   — store URLs, image width, color constants
    columns.ts  — column index map (A–Z) and header labels
    shopify.ts  — fetchAllProducts, fetchProductByHandle, getPriceInfo, getStockInfo
    sheet.ts    — writeSummaryRow, writeDetailRows, ensureHeaderRow, styleStockCell
    groups.ts   — groupAndCollapse, setAllGroupsCollapsed (Sheets API v4)
  main.ts       — GAS entry points: onOpen, refreshAll, refreshPricesAndStock, etc.
```

### Updating Filament Data

Static specs (temperatures, mechanical properties) live in `STATIC_DATA` in `src/main.ts`. When Bambu adds a new filament, add an entry there keyed by the product's Shopify handle (visible in the product URL).
