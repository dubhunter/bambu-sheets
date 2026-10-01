// ---------------------------------------------------------------------------
// Column index map (1-based, matching sheet layout)
// ---------------------------------------------------------------------------
// Columns A–X are visible; Y onward are hidden helper columns.
//
//  A  IMAGE         — =IMAGE() formula showing product thumbnail
//  B  NAME          — product title (e.g. "Bambu PLA Basic")
//  C  MATERIAL      — material type (e.g. "PLA", "ABS", "ASA")
//  D  PRICE         — current price per spool (formatted string)
//  E  BULK_PRICE    — multi-pack / bulk price, if different
//  F  FIBER         — reinforcement: No | CF | GF
//  G  AMS           — AMS compatibility: Yes | No | HT
//  H  ENCLOSURE     — requires enclosure: Yes | No | Recommended
//  I  HARDENED      — requires hardened nozzle: Yes | No
//  J  NOZZLE_TEMP   — nozzle temp range, °C (e.g. "220–240")
//  K  BED_TEMP      — bed temp range, °C
//  L  DRY_TEMP      — drying temperature, °C
//  M  DRY_TIME      — drying time, hours
//  N  MAX_SPEED     — max volumetric speed, mm³/s
//  O  TENSILE       — tensile strength, MPa
//  P  BENDING_MOD   — bending modulus, GPa
//  Q  HDT           — heat deflection temperature, °C
//  R  COLOR_BLACK   — has black color option: Yes | No
//  S  COLOR_WHITE   — has white color option: Yes | No
//  T  COLOR_OTHER   — has other color options: Yes | No
//  U  PRODUCT_URL   — link to product page
//  V  IMAGE_URL     — raw CDN image URL (used to build =IMAGE() in col A)
//  W  NOTES         — misc notes
//  X  IN_STOCK      — stock status (summary row: "28/32 in stock"; detail: "✓" / "✗")
//  Y  COLOR_NAME    — variant title / color name (detail rows only)
//  Z  ROW_MARKER    — "SUMMARY" or "DETAIL" (hidden; used by script logic)
// ---------------------------------------------------------------------------

const Col = {
    IMAGE:       1,
    NAME:        2,
    MATERIAL:    3,
    PRICE:       4,
    BULK_PRICE:  5,
    FIBER:       6,
    AMS:         7,
    ENCLOSURE:   8,
    HARDENED:    9,
    NOZZLE_TEMP: 10,
    BED_TEMP:    11,
    DRY_TEMP:    12,
    DRY_TIME:    13,
    MAX_SPEED:   14,
    TENSILE:     15,
    BENDING_MOD: 16,
    HDT:         17,
    COLOR_BLACK: 18,
    COLOR_WHITE: 19,
    COLOR_OTHER: 20,
    PRODUCT_URL: 21,
    IMAGE_URL:   22,
    NOTES:       23,
    IN_STOCK:    24,
    COLOR_NAME:  25,
    ROW_MARKER:  26,
} as const;

const TOTAL_COLS = 26;

// Human-readable header labels (index 0 = col A = Col.IMAGE)
const HEADERS: string[] = [
    'Image',
    'Name',
    'Material',
    'Price',
    'Bulk Price',
    'Fiber',
    'AMS',
    'Enclosure',
    'Hardened Nozzle',
    'Nozzle Temp (°C)',
    'Bed Temp (°C)',
    'Dry Temp (°C)',
    'Dry Time (h)',
    'Max Speed (mm³/s)',
    'Tensile Strength (MPa)',
    'Bending Modulus (GPa)',
    'HDT (°C)',
    'Black',
    'White',
    'Other Colors',
    'Product URL',
    'Image URL',
    'Notes',
    'In Stock',
    'Color',
    'RowMarker',
];
