// ---------------------------------------------------------------------------
// GAS entry points
// ---------------------------------------------------------------------------
// All functions at the top level of any .ts file are compiled into the global
// GAS namespace (module: "none"), so they're accessible from menus / triggers.
// ---------------------------------------------------------------------------

// ---------------------------------------------------------------------------
// Menu
// ---------------------------------------------------------------------------

function onOpen(): void {
    SpreadsheetApp.getUi()
        .createMenu('🧵 Bambu Filaments')
        .addItem('Refresh all (full rebuild)', 'refreshAll')
        .addItem('Refresh prices & stock only', 'refreshPricesAndStock')
        .addSeparator()
        .addItem('Expand all color groups', 'expandAllGroups')
        .addItem('Collapse all color groups', 'collapseAllGroups')
        .addSeparator()
        .addItem('Authorize & setup sheet', 'authorizeAndSetup')
        .addToUi();
}

// ---------------------------------------------------------------------------
// Setup
// ---------------------------------------------------------------------------

/**
 * One-time setup: grants OAuth scopes, writes headers, hides helper columns.
 * Run this manually the first time, or after clearing the sheet.
 */
function authorizeAndSetup(): void {
    const sheet = getOrCreateSheet();
    ensureHeaderRow(sheet);
    hideMarkerColumn(sheet);
    hideImageUrlColumn(sheet);
    SpreadsheetApp.getUi().alert('Setup complete. Run "Refresh all" to populate data.');
}

// ---------------------------------------------------------------------------
// Full refresh (rebuild summary + detail rows + groups)
// ---------------------------------------------------------------------------

/**
 * Fetches all products from Shopify, clears the sheet, and rebuilds every
 * summary row + grouped detail rows from scratch.
 *
 * This is the slow path — use it after the product catalog changes.
 */
function refreshAll(): void {
    const ui    = SpreadsheetApp.getUi();
    const sheet = getOrCreateSheet();

    const products = fetchAllProducts();
    if (!products) {
        ui.alert('Error fetching products. Check the Apps Script logs.');
        return;
    }

    // Clear existing data (keep header)
    const lastRow = sheet.getLastRow();
    if (lastRow > 1) {
        sheet.deleteRows(2, lastRow - 1);
    }

    // Remove stale groups
    clearAllGroups(sheet);

    // Reload header in case columns changed
    ensureHeaderRow(sheet);

    let currentRow = 2;

    for (const product of products) {
        // Write summary row
        writeSummaryRow(sheet, currentRow, product, getStaticData(product.handle));
        currentRow++;

        // Write detail rows (one per color variant)
        if (product.variants.length > 1) {
            const detailCount = writeDetailRows(sheet, currentRow, product);
            if (detailCount > 0) {
                groupAndCollapse(sheet, currentRow, currentRow + detailCount - 1);
                currentRow += detailCount;
            }
        } else if (product.variants.length === 1) {
            // Single variant — still write it so stock status is visible
            writeDetailRows(sheet, currentRow, product);
            currentRow++;
        }

        // Avoid hitting Sheets API rate limits
        if (currentRow % 20 === 0) {
            SpreadsheetApp.flush();
        }
    }

    hideMarkerColumn(sheet);
    hideImageUrlColumn(sheet);
    SpreadsheetApp.flush();

    ui.alert(`Done! Loaded ${products.length} filaments.`);
}

// ---------------------------------------------------------------------------
// Fast refresh (prices + stock only, no row rebuild)
// ---------------------------------------------------------------------------

/**
 * Scans existing summary rows, re-fetches each product by handle, and updates
 * only the price and stock columns — preserving row structure and groups.
 *
 * Much faster than refreshAll() for routine stock checks.
 */
function refreshPricesAndStock(): void {
    const ui    = SpreadsheetApp.getUi();
    const sheet = getOrCreateSheet();
    const lastRow = sheet.getLastRow();
    if (lastRow < 2) {
        ui.alert('Sheet is empty. Run "Refresh all" first.');
        return;
    }

    // Read all ROW_MARKER and PRODUCT_URL values in one batch
    const markers     = sheet.getRange(2, Col.ROW_MARKER,  lastRow - 1, 1).getValues() as string[][];
    const productUrls = sheet.getRange(2, Col.PRODUCT_URL, lastRow - 1, 1).getValues() as string[][];

    let updated = 0;

    for (let i = 0; i < markers.length; i++) {
        if (markers[i]![0] !== RowMarker.Summary) continue;

        const row = i + 2; // 1-based + header
        const url = productUrls[i]![0] ?? '';
        const handle = handleFromUrl(url);
        if (!handle) continue;

        const product = fetchProductByHandle(handle);
        if (!product) continue;

        const price = getPriceInfo(product);
        const stock = getStockInfo(product);

        sheet.getRange(row, Col.PRICE).setValue(price.display);
        sheet.getRange(row, Col.BULK_PRICE).setValue(price.compareAt);
        sheet.getRange(row, Col.IN_STOCK).setValue(stock.summary);
        styleStockCell(
            sheet.getRange(row, Col.IN_STOCK),
            stock.allInStock,
            stock.noneInStock,
        );

        // Update detail rows too
        const detailRange = findDetailRowRange(sheet, row);
        if (detailRange && product.variants.length) {
            product.variants.forEach((v, vi) => {
                const detailRow = detailRange.start + vi;
                if (detailRow > detailRange.end) return;
                sheet.getRange(detailRow, Col.IN_STOCK).setValue(v.available ? '✓' : '✗');
                styleStockCell(
                    sheet.getRange(detailRow, Col.IN_STOCK),
                    v.available,
                    !v.available,
                );
            });
        }

        updated++;

        // Throttle: Shopify allows ~2 req/s on the public API
        Utilities.sleep(600);
    }

    SpreadsheetApp.flush();
    ui.alert(`Updated prices & stock for ${updated} filaments.`);
}

// ---------------------------------------------------------------------------
// Group expand / collapse
// ---------------------------------------------------------------------------

function expandAllGroups(): void {
    setAllGroupsCollapsed(false);
}

function collapseAllGroups(): void {
    setAllGroupsCollapsed(true);
}

// ---------------------------------------------------------------------------
// Utilities
// ---------------------------------------------------------------------------

/** Returns the "Bambu Filaments" sheet, creating it if it doesn't exist. */
function getOrCreateSheet(): GoogleAppsScript.Spreadsheet.Sheet {
    const ss   = SpreadsheetApp.getActiveSpreadsheet();
    const name = 'Bambu Filaments';
    return ss.getSheetByName(name) ?? ss.insertSheet(name);
}

/**
 * Returns static spec data for a given product handle, or undefined if we
 * don't have data for it.
 */
function getStaticData(handle: string): StaticFilamentData | undefined {
    return STATIC_DATA.find(d => d.handle === handle);
}

// ---------------------------------------------------------------------------
// Static spec data
// ---------------------------------------------------------------------------
// Temperatures, speeds, and mechanical properties don't change often and
// aren't available from the Shopify API, so they're baked in here.
// Update this array when Bambu adds or changes filaments.
// ---------------------------------------------------------------------------

const STATIC_DATA: StaticFilamentData[] = [
    {
        handle: 'bambu-pla-basic',
        material: 'PLA', fiber: 'No', ams: 'Yes', enclosure: 'No', hardened: 'No',
        nozzleTemp: '220–240', bedTemp: '35–45', dryTemp: '45', dryTime: '8',
        maxSpeed: '21', tensile: '37', bendingMod: '3.3', hdt: '52',
        colorBlack: 'Yes', colorWhite: 'Yes', colorOther: 'Yes',
    },
    {
        handle: 'bambu-pla-matte',
        material: 'PLA', fiber: 'No', ams: 'Yes', enclosure: 'No', hardened: 'No',
        nozzleTemp: '220–240', bedTemp: '35–45', dryTemp: '45', dryTime: '8',
        maxSpeed: '12', tensile: '40', bendingMod: '3.5', hdt: '52',
        colorBlack: 'Yes', colorWhite: 'Yes', colorOther: 'Yes',
    },
    {
        handle: 'bambu-pla-silk',
        material: 'PLA', fiber: 'No', ams: 'Yes', enclosure: 'No', hardened: 'No',
        nozzleTemp: '220–240', bedTemp: '35–45', dryTemp: '45', dryTime: '8',
        maxSpeed: '12', tensile: '35', bendingMod: '3.1', hdt: '52',
        colorBlack: 'Yes', colorWhite: 'Yes', colorOther: 'Yes',
    },
    {
        handle: 'pla-sparkle',
        material: 'PLA', fiber: 'No', ams: 'Yes', enclosure: 'No', hardened: 'No',
        nozzleTemp: '220–240', bedTemp: '35–45', dryTemp: '45', dryTime: '8',
        maxSpeed: '12', tensile: '35', bendingMod: '3.1', hdt: '52',
        colorBlack: 'No', colorWhite: 'No', colorOther: 'Yes',
    },
    {
        handle: 'bambu-pla-metal',
        material: 'PLA', fiber: 'No', ams: 'Yes', enclosure: 'No', hardened: 'No',
        nozzleTemp: '220–240', bedTemp: '35–45', dryTemp: '45', dryTime: '8',
        maxSpeed: '12', tensile: '35', bendingMod: '3.1', hdt: '52',
        colorBlack: 'Yes', colorWhite: 'No', colorOther: 'Yes',
    },
    {
        handle: 'bambu-pla-marble',
        material: 'PLA', fiber: 'No', ams: 'Yes', enclosure: 'No', hardened: 'No',
        nozzleTemp: '220–240', bedTemp: '35–45', dryTemp: '45', dryTime: '8',
        maxSpeed: '12',
        colorBlack: 'No', colorWhite: 'Yes', colorOther: 'Yes',
    },
    {
        handle: 'bambu-pla-aero',
        material: 'PLA', fiber: 'No', ams: 'Yes', enclosure: 'No', hardened: 'No',
        nozzleTemp: '220–240', bedTemp: '35–45', dryTemp: '45', dryTime: '8',
        maxSpeed: '21', tensile: '37', bendingMod: '3.3', hdt: '52',
        colorBlack: 'No', colorWhite: 'No', colorOther: 'Yes',
        notes: 'Lightweight / foam-core structure',
    },
    {
        handle: 'bambu-pla-basic-gradient',
        material: 'PLA', fiber: 'No', ams: 'Yes', enclosure: 'No', hardened: 'No',
        nozzleTemp: '220–240', bedTemp: '35–45', dryTemp: '45', dryTime: '8',
        maxSpeed: '21',
        colorBlack: 'No', colorWhite: 'No', colorOther: 'Yes',
    },
    {
        handle: 'pla-impact',
        material: 'PLA', fiber: 'No', ams: 'Yes', enclosure: 'No', hardened: 'No',
        nozzleTemp: '220–240', bedTemp: '35–45', dryTemp: '45', dryTime: '8',
        maxSpeed: '21', tensile: '52', bendingMod: '2.8', hdt: '55',
        colorBlack: 'Yes', colorWhite: 'Yes', colorOther: 'Yes',
        notes: 'High-impact PLA',
    },
    {
        handle: 'bambu-pla-tough',
        material: 'PLA', fiber: 'No', ams: 'Yes', enclosure: 'No', hardened: 'No',
        nozzleTemp: '220–240', bedTemp: '35–45', dryTemp: '45', dryTime: '8',
        maxSpeed: '21', tensile: '38', bendingMod: '2.4', hdt: '55',
        colorBlack: 'Yes', colorWhite: 'Yes', colorOther: 'Yes',
    },
    {
        handle: 'bambu-pla-cf',
        material: 'PLA', fiber: 'CF', ams: 'Yes', enclosure: 'No', hardened: 'Yes',
        nozzleTemp: '220–240', bedTemp: '35–45', dryTemp: '45', dryTime: '8',
        maxSpeed: '12', tensile: '68', bendingMod: '6.7', hdt: '58',
        colorBlack: 'Yes', colorWhite: 'No', colorOther: 'No',
    },
    {
        handle: 'bambu-pla-gf',
        material: 'PLA', fiber: 'GF', ams: 'Yes', enclosure: 'No', hardened: 'Yes',
        nozzleTemp: '220–240', bedTemp: '35–45', dryTemp: '45', dryTime: '8',
        maxSpeed: '12', tensile: '53', bendingMod: '5.5', hdt: '60',
        colorBlack: 'No', colorWhite: 'No', colorOther: 'Yes',
    },
    {
        handle: 'support-for-pla-s',
        material: 'PLA-S', fiber: 'No', ams: 'Yes', enclosure: 'No', hardened: 'No',
        nozzleTemp: '220–240', bedTemp: '35–45', dryTemp: '45', dryTime: '8',
        colorBlack: 'No', colorWhite: 'Yes', colorOther: 'No',
        notes: 'Breakaway support material for PLA',
    },
    {
        handle: 'support-for-pla-new-edition',
        material: 'PLA-S', fiber: 'No', ams: 'Yes', enclosure: 'No', hardened: 'No',
        nozzleTemp: '220–240', bedTemp: '35–45', dryTemp: '45', dryTime: '8',
        colorBlack: 'No', colorWhite: 'Yes', colorOther: 'No',
        notes: 'Breakaway support material for PLA (new formulation)',
    },
    {
        handle: 'bambu-petg-basic',
        material: 'PETG', fiber: 'No', ams: 'Yes', enclosure: 'No', hardened: 'No',
        nozzleTemp: '230–250', bedTemp: '70–80', dryTemp: '65', dryTime: '8',
        maxSpeed: '12', tensile: '32', bendingMod: '1.8', hdt: '72',
        colorBlack: 'Yes', colorWhite: 'Yes', colorOther: 'Yes',
    },
    {
        handle: 'bambu-petg-hf',
        material: 'PETG', fiber: 'No', ams: 'Yes', enclosure: 'No', hardened: 'No',
        nozzleTemp: '240–260', bedTemp: '70–80', dryTemp: '65', dryTime: '8',
        maxSpeed: '24', tensile: '50', bendingMod: '2.4', hdt: '72',
        colorBlack: 'Yes', colorWhite: 'Yes', colorOther: 'Yes',
        notes: 'High-flow PETG',
    },
    {
        handle: 'bambu-petg-cf',
        material: 'PETG', fiber: 'CF', ams: 'Yes', enclosure: 'No', hardened: 'Yes',
        nozzleTemp: '240–260', bedTemp: '70–85', dryTemp: '65', dryTime: '8',
        maxSpeed: '12', tensile: '74', bendingMod: '5.5', hdt: '78',
        colorBlack: 'Yes', colorWhite: 'No', colorOther: 'No',
    },
    {
        handle: 'bambu-abs',
        material: 'ABS', fiber: 'No', ams: 'Yes', enclosure: 'Yes', hardened: 'No',
        nozzleTemp: '240–270', bedTemp: '90–100', dryTemp: '65', dryTime: '8',
        maxSpeed: '12', tensile: '36', bendingMod: '2.3', hdt: '88',
        colorBlack: 'Yes', colorWhite: 'Yes', colorOther: 'Yes',
    },
    {
        handle: 'bambu-abs-gf',
        material: 'ABS', fiber: 'GF', ams: 'Yes', enclosure: 'Yes', hardened: 'Yes',
        nozzleTemp: '250–270', bedTemp: '90–110', dryTemp: '80', dryTime: '8',
        maxSpeed: '12', tensile: '56', bendingMod: '5.0', hdt: '100',
        colorBlack: 'No', colorWhite: 'No', colorOther: 'Yes',
    },
    {
        handle: 'bambu-asa',
        material: 'ASA', fiber: 'No', ams: 'Yes', enclosure: 'Yes', hardened: 'No',
        nozzleTemp: '240–270', bedTemp: '90–100', dryTemp: '65', dryTime: '8',
        maxSpeed: '12', tensile: '34', bendingMod: '2.0', hdt: '90',
        colorBlack: 'Yes', colorWhite: 'Yes', colorOther: 'Yes',
        notes: 'UV resistant',
    },
    {
        handle: 'bambu-pa6-cf',
        material: 'PA6', fiber: 'CF', ams: 'HT', enclosure: 'Yes', hardened: 'Yes',
        nozzleTemp: '260–280', bedTemp: '80–90', dryTemp: '80', dryTime: '12',
        maxSpeed: '12', tensile: '150', bendingMod: '9.0', hdt: '190',
        colorBlack: 'Yes', colorWhite: 'No', colorOther: 'No',
    },
    {
        handle: 'bambu-pa-cf',
        material: 'PA', fiber: 'CF', ams: 'HT', enclosure: 'Yes', hardened: 'Yes',
        nozzleTemp: '260–280', bedTemp: '50–65', dryTemp: '80', dryTime: '12',
        maxSpeed: '12', tensile: '135', bendingMod: '8.0', hdt: '170',
        colorBlack: 'Yes', colorWhite: 'No', colorOther: 'No',
    },
    {
        handle: 'bambu-pa-gf',
        material: 'PA', fiber: 'GF', ams: 'HT', enclosure: 'Yes', hardened: 'Yes',
        nozzleTemp: '260–280', bedTemp: '50–65', dryTemp: '80', dryTime: '12',
        maxSpeed: '12', tensile: '95', bendingMod: '6.0', hdt: '155',
        colorBlack: 'Yes', colorWhite: 'No', colorOther: 'Yes',
    },
    {
        handle: 'bambu-pc',
        material: 'PC', fiber: 'No', ams: 'HT', enclosure: 'Yes', hardened: 'No',
        nozzleTemp: '260–280', bedTemp: '90–110', dryTemp: '80', dryTime: '8',
        maxSpeed: '6', tensile: '58', bendingMod: '2.1', hdt: '126',
        colorBlack: 'Yes', colorWhite: 'Yes', colorOther: 'Yes',
    },
    {
        handle: 'pc-cf',
        material: 'PC', fiber: 'CF', ams: 'HT', enclosure: 'Yes', hardened: 'Yes',
        nozzleTemp: '270–300', bedTemp: '100–120', dryTemp: '80', dryTime: '8',
        maxSpeed: '8', tensile: '92', bendingMod: '7.4', hdt: '138',
        colorBlack: 'Yes', colorWhite: 'No', colorOther: 'No',
    },
    {
        handle: 'pc-fr',
        material: 'PC', fiber: 'No', ams: 'HT', enclosure: 'Yes', hardened: 'No',
        nozzleTemp: '260–300', bedTemp: '90–110', dryTemp: '80', dryTime: '8',
        colorBlack: 'Yes', colorWhite: 'No', colorOther: 'No',
        notes: 'Flame retardant (UL94 V0)',
    },
    {
        handle: 'bambu-tpu-95a',
        material: 'TPU', fiber: 'No', ams: 'Yes', enclosure: 'No', hardened: 'No',
        nozzleTemp: '210–230', bedTemp: '35–45', dryTemp: '65', dryTime: '8',
        maxSpeed: '4', tensile: '39', hdt: '43',
        colorBlack: 'Yes', colorWhite: 'Yes', colorOther: 'Yes',
        notes: '95A shore hardness',
    },
    {
        handle: 'bambu-tpu-90a',
        material: 'TPU', fiber: 'No', ams: 'No', enclosure: 'No', hardened: 'No',
        nozzleTemp: '220–230', bedTemp: '30–40', dryTemp: '65', dryTime: '8',
        maxSpeed: '4', tensile: '25',
        colorBlack: 'Yes', colorWhite: 'No', colorOther: 'Yes',
        notes: '90A shore hardness',
    },
    {
        handle: 'bambu-pps-cf',
        material: 'PPS', fiber: 'CF', ams: 'HT', enclosure: 'Yes', hardened: 'Yes',
        nozzleTemp: '300–320', bedTemp: '120–140', dryTemp: '120', dryTime: '24',
        tensile: '136', bendingMod: '12.5', hdt: '260',
        colorBlack: 'Yes', colorWhite: 'No', colorOther: 'No',
        notes: 'High-temp engineering filament',
    },
    {
        handle: 'pps-cf',
        material: 'PPS', fiber: 'CF', ams: 'HT', enclosure: 'Yes', hardened: 'Yes',
        nozzleTemp: '300–320', bedTemp: '120–140', dryTemp: '120', dryTime: '24',
        tensile: '136', bendingMod: '12.5', hdt: '260',
        colorBlack: 'Yes', colorWhite: 'No', colorOther: 'No',
        notes: 'High-temp engineering filament',
    },
    {
        handle: 'bambu-asa-aero',
        material: 'ASA', fiber: 'No', ams: 'Yes', enclosure: 'Yes', hardened: 'No',
        nozzleTemp: '240–270', bedTemp: '90–100', dryTemp: '65', dryTime: '8',
        maxSpeed: '21',
        colorBlack: 'No', colorWhite: 'No', colorOther: 'Yes',
        notes: 'Lightweight / foam-core structure',
    },
    {
        handle: 'support-for-abs',
        material: 'Support', fiber: 'No', ams: 'Yes', enclosure: 'Yes', hardened: 'No',
        nozzleTemp: '220–240', bedTemp: '45–55', dryTemp: '65', dryTime: '8',
        colorBlack: 'No', colorWhite: 'Yes', colorOther: 'No',
        notes: 'Water-soluble or breakaway support for ABS/ASA',
    },
    {
        handle: 'support-g',
        material: 'Support', fiber: 'No', ams: 'Yes', enclosure: 'No', hardened: 'No',
        nozzleTemp: '220–240', bedTemp: '35–45', dryTemp: '45', dryTime: '8',
        colorBlack: 'No', colorWhite: 'Yes', colorOther: 'No',
        notes: 'Water-soluble support (BVOH/PVA-based)',
    },
    {
        handle: 'bambu-hips',
        material: 'HIPS', fiber: 'No', ams: 'Yes', enclosure: 'Yes', hardened: 'No',
        nozzleTemp: '220–240', bedTemp: '90–100', dryTemp: '65', dryTime: '8',
        maxSpeed: '12', tensile: '23', bendingMod: '1.7', hdt: '76',
        colorBlack: 'Yes', colorWhite: 'Yes', colorOther: 'Yes',
        notes: 'Dissolvable in d-Limonene; support for ABS',
    },
];
