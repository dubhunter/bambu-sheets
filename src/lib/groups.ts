// ---------------------------------------------------------------------------
// Row grouping via Google Sheets API v4 (advanced service)
// ---------------------------------------------------------------------------
// Standard SpreadsheetApp has no API for creating or collapsing row groups.
// We use Sheets.Spreadsheets.batchUpdate (advanced service) instead.
//
// Prerequisites (already set in appsscript.json):
//   enabledAdvancedServices: [{ userSymbol: "Sheets", serviceId: "sheets", version: "v4" }]
// ---------------------------------------------------------------------------

/**
 * Creates a collapsible row group for `startRow`..`endRow` (1-based, inclusive)
 * and collapses it by default.
 *
 * Call after writing detail rows for a product.
 */
function groupAndCollapse(
    sheet: GoogleAppsScript.Spreadsheet.Sheet,
    startRow: number,
    endRow: number,
): void {
    if (endRow < startRow) return;

    const ssId    = sheet.getParent().getId();
    const sheetId = sheet.getSheetId();

    // Convert to 0-based half-open range expected by the API
    const startIndex = startRow - 1;
    const endIndex   = endRow;       // exclusive in API

    const requests: GoogleAppsScript.Sheets.Schema.Request[] = [
        {
            addDimensionGroup: {
                range: {
                    sheetId,
                    dimension:  'ROWS',
                    startIndex,
                    endIndex,
                },
            },
        },
        {
            updateDimensionGroup: {
                dimensionGroup: {
                    range: {
                        sheetId,
                        dimension:  'ROWS',
                        startIndex,
                        endIndex,
                    },
                    depth:     1,
                    collapsed: true,
                },
                fields: 'collapsed',
            },
        },
    ];

    Sheets.Spreadsheets!.batchUpdate({ requests }, ssId);
}

/**
 * Expands or collapses every row group in the active sheet.
 *
 * @param collapsed - true to collapse all, false to expand all
 */
function setAllGroupsCollapsed(collapsed: boolean): void {
    const ss    = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getActiveSheet();
    const ssId  = ss.getId();

    const metaResp = Sheets.Spreadsheets!.get(ssId, {
        ranges:  [`'${sheet.getName()}'`],
        fields:  'sheets(rowGroups)',
    });

    const sheetMeta = metaResp.sheets?.[0];
    const rowGroups = sheetMeta?.rowGroups;
    if (!rowGroups?.length) return;

    const sheetId = sheet.getSheetId();

    const requests: GoogleAppsScript.Sheets.Schema.Request[] = rowGroups.map(group => ({
        updateDimensionGroup: {
            dimensionGroup: {
                range:     group.range,
                depth:     group.depth,
                sheetId,
                collapsed,
            },
            fields: 'collapsed',
        },
    }));

    Sheets.Spreadsheets!.batchUpdate({ requests }, ssId);
}

/**
 * Deletes all existing row groups in the sheet.
 * Call before a full rebuild so stale groups don't accumulate.
 */
function clearAllGroups(sheet: GoogleAppsScript.Spreadsheet.Sheet): void {
    const ssId = sheet.getParent().getId();

    const metaResp = Sheets.Spreadsheets!.get(ssId, {
        ranges:  [`'${sheet.getName()}'`],
        fields:  'sheets(rowGroups)',
    });

    const sheetMeta = metaResp.sheets?.[0];
    const rowGroups = sheetMeta?.rowGroups;
    if (!rowGroups?.length) return;

    const sheetId = sheet.getSheetId();

    const requests: GoogleAppsScript.Sheets.Schema.Request[] = rowGroups.map(group => ({
        deleteDimensionGroup: {
            range: {
                sheetId,
                dimension:  'ROWS',
                startIndex: group.range?.startIndex,
                endIndex:   group.range?.endIndex,
            },
        },
    }));

    Sheets.Spreadsheets!.batchUpdate({ requests }, ssId);
}
