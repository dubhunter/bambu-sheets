// ---------------------------------------------------------------------------
// Sheet read / write / style helpers
// ---------------------------------------------------------------------------

type Sheet = GoogleAppsScript.Spreadsheet.Sheet;
type Range = GoogleAppsScript.Spreadsheet.Range;

// ---------------------------------------------------------------------------
// Header
// ---------------------------------------------------------------------------

/** Ensures row 1 contains the correct header labels and styling. */
function ensureHeaderRow(sheet: Sheet): void {
    const header = sheet.getRange(1, 1, 1, TOTAL_COLS);
    header.setValues([HEADERS]);
    header
        .setBackground(COLOR_HEADER_BG)
        .setFontColor(COLOR_HEADER_FG)
        .setFontWeight('bold')
        .setWrap(false);
    sheet.setFrozenRows(1);
}

// ---------------------------------------------------------------------------
// Column visibility
// ---------------------------------------------------------------------------

/** Hides the ROW_MARKER column (Z) — it's only used by script logic. */
function hideMarkerColumn(sheet: Sheet): void {
    sheet.hideColumns(Col.ROW_MARKER);
}

/** Hides IMAGE_URL column (V) — it's only needed as a data source. */
function hideImageUrlColumn(sheet: Sheet): void {
    sheet.hideColumns(Col.IMAGE_URL);
}

// ---------------------------------------------------------------------------
// Writing rows
// ---------------------------------------------------------------------------

/**
 * Writes (or overwrites) the summary row for a filament product.
 * The summary row is always visible; detail rows are grouped below it.
 *
 * @param sheet      - target sheet
 * @param row        - 1-based row index
 * @param product    - Shopify product data
 * @param staticData - static spec data keyed by product handle (may be undefined)
 */
function writeSummaryRow(
    sheet: Sheet,
    row: number,
    product: ShopifyProduct,
    staticData: StaticFilamentData | undefined,
): void {
    const price    = getPriceInfo(product);
    const stock    = getStockInfo(product);
    const imageUrl = getBestImageUrl(product);
    const productUrl = `${STORE_DISPLAY}/products/${product.handle}`;

    const imageFormula = imageUrl
        ? `=IMAGE("${imageUrl}",4,${IMAGE_ROW_HEIGHT - 4},${IMAGE_ROW_HEIGHT - 4})`
        : '';

    const rowData: (string | number | boolean)[] = new Array(TOTAL_COLS).fill('');

    rowData[Col.IMAGE       - 1] = imageFormula;
    rowData[Col.NAME        - 1] = product.title;
    rowData[Col.PRICE       - 1] = price.display;
    rowData[Col.BULK_PRICE  - 1] = price.compareAt;
    rowData[Col.PRODUCT_URL - 1] = productUrl;
    rowData[Col.IMAGE_URL   - 1] = imageUrl;
    rowData[Col.IN_STOCK    - 1] = stock.summary;
    rowData[Col.ROW_MARKER  - 1] = RowMarker.Summary;

    // Overlay static spec data if available
    if (staticData) {
        rowData[Col.MATERIAL    - 1] = staticData.material    ?? '';
        rowData[Col.FIBER       - 1] = staticData.fiber       ?? '';
        rowData[Col.AMS         - 1] = staticData.ams         ?? '';
        rowData[Col.ENCLOSURE   - 1] = staticData.enclosure   ?? '';
        rowData[Col.HARDENED    - 1] = staticData.hardened    ?? '';
        rowData[Col.NOZZLE_TEMP - 1] = staticData.nozzleTemp  ?? '';
        rowData[Col.BED_TEMP    - 1] = staticData.bedTemp     ?? '';
        rowData[Col.DRY_TEMP    - 1] = staticData.dryTemp     ?? '';
        rowData[Col.DRY_TIME    - 1] = staticData.dryTime     ?? '';
        rowData[Col.MAX_SPEED   - 1] = staticData.maxSpeed    ?? '';
        rowData[Col.TENSILE     - 1] = staticData.tensile     ?? '';
        rowData[Col.BENDING_MOD - 1] = staticData.bendingMod  ?? '';
        rowData[Col.HDT         - 1] = staticData.hdt         ?? '';
        rowData[Col.COLOR_BLACK - 1] = staticData.colorBlack  ?? '';
        rowData[Col.COLOR_WHITE - 1] = staticData.colorWhite  ?? '';
        rowData[Col.COLOR_OTHER - 1] = staticData.colorOther  ?? '';
        rowData[Col.NOTES       - 1] = staticData.notes       ?? '';
    }

    sheet.getRange(row, 1, 1, TOTAL_COLS).setValues([rowData]);

    // Style the summary row
    const summaryRange = sheet.getRange(row, 1, 1, TOTAL_COLS);
    summaryRange.setBackground(COLOR_SUMMARY_BG).setFontWeight('bold');

    // Stock cell colour
    styleStockCell(sheet.getRange(row, Col.IN_STOCK), stock.allInStock, stock.noneInStock);

    // Row height for thumbnail
    sheet.setRowHeight(row, IMAGE_ROW_HEIGHT);
}

/**
 * Writes detail rows (one per variant / color) starting at `startRow`.
 * Returns the number of rows written.
 */
function writeDetailRows(
    sheet: Sheet,
    startRow: number,
    product: ShopifyProduct,
): number {
    if (!product.variants.length) return 0;

    const productUrl = `${STORE_DISPLAY}/products/${product.handle}`;
    const imageUrl   = getBestImageUrl(product);

    const data: (string | number | boolean)[][] = product.variants.map(v => {
        const row: (string | number | boolean)[] = new Array(TOTAL_COLS).fill('');
        row[Col.COLOR_NAME  - 1] = v.title;
        row[Col.IN_STOCK    - 1] = v.available ? '✓' : '✗';
        row[Col.PRICE       - 1] = v.price ? `$${parseFloat(v.price).toFixed(2)}` : '';
        row[Col.PRODUCT_URL - 1] = productUrl;
        row[Col.IMAGE_URL   - 1] = imageUrl;
        row[Col.ROW_MARKER  - 1] = RowMarker.Detail;
        return row;
    });

    const range = sheet.getRange(startRow, 1, data.length, TOTAL_COLS);
    range.setValues(data);

    // Style each detail row
    product.variants.forEach((v, i) => {
        const detailRow = startRow + i;
        sheet.getRange(detailRow, 1, 1, TOTAL_COLS)
            .setBackground(COLOR_DETAIL_BG)
            .setFontWeight('normal');

        styleStockCell(
            sheet.getRange(detailRow, Col.IN_STOCK),
            v.available,
            !v.available,
        );
    });

    return data.length;
}

// ---------------------------------------------------------------------------
// Stock cell styling
// ---------------------------------------------------------------------------

/** Colours an in-stock cell green, out-of-stock red, partial yellow. */
function styleStockCell(range: Range, isGood: boolean, isBad: boolean): void {
    if (isGood) {
        range.setBackground(COLOR_IN_STOCK);
    } else if (isBad) {
        range.setBackground(COLOR_OUT_STOCK);
    } else {
        range.setBackground(COLOR_PART_STOCK);
    }
}

// ---------------------------------------------------------------------------
// Row scanning helpers
// ---------------------------------------------------------------------------

/**
 * Returns the 1-based row number of the first summary row matching the given
 * product name, or -1 if not found.
 */
function findRowByMarkerAndName(
    sheet: Sheet,
    name: string,
): number {
    const lastRow = sheet.getLastRow();
    if (lastRow < 2) return -1;

    const markers = sheet.getRange(2, Col.ROW_MARKER, lastRow - 1, 1).getValues() as string[][];
    const names   = sheet.getRange(2, Col.NAME, lastRow - 1, 1).getValues() as string[][];

    for (let i = 0; i < markers.length; i++) {
        if (markers[i]![0] === RowMarker.Summary && names[i]![0] === name) {
            return i + 2; // +2: 1-based + skip header
        }
    }
    return -1;
}

/**
 * Given a summary row, returns the row range of its detail rows (inclusive),
 * or null if there are none.
 */
function findDetailRowRange(
    sheet: Sheet,
    summaryRow: number,
): { start: number; end: number } | null {
    const lastRow = sheet.getLastRow();
    if (summaryRow >= lastRow) return null;

    const markers = sheet
        .getRange(summaryRow + 1, Col.ROW_MARKER, lastRow - summaryRow, 1)
        .getValues() as string[][];

    let end = summaryRow; // inclusive
    for (let i = 0; i < markers.length; i++) {
        if (markers[i]![0] === RowMarker.Detail) {
            end = summaryRow + 1 + i;
        } else {
            break;
        }
    }

    if (end === summaryRow) return null; // no detail rows
    return { start: summaryRow + 1, end };
}

// ---------------------------------------------------------------------------
// Static spec data type (sourced from the original CSV / manual data)
// ---------------------------------------------------------------------------

type StaticFilamentData = {
    handle:      string;
    material?:   string;
    fiber?:      string;
    ams?:        string;
    enclosure?:  string;
    hardened?:   string;
    nozzleTemp?: string;
    bedTemp?:    string;
    dryTemp?:    string;
    dryTime?:    string;
    maxSpeed?:   string;
    tensile?:    string;
    bendingMod?: string;
    hdt?:        string;
    colorBlack?: string;
    colorWhite?: string;
    colorOther?: string;
    notes?:      string;
};
