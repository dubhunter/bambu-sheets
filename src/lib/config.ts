// ---------------------------------------------------------------------------
// Global constants
// ---------------------------------------------------------------------------

const STORE_BASE    = 'https://bambulab-us.myshopify.com';
const STORE_DISPLAY = 'https://us.store.bambulab.com';
const COLLECTION    = 'bambu-lab-3d-printer-filament';
const IMAGE_WIDTH   = 192;  // px — thumbnail size in sheet

// Fetch up to 250 products per page (Shopify maximum)
const PAGE_LIMIT = 250;

// ---------------------------------------------------------------------------
// Sheet / row styling
// ---------------------------------------------------------------------------

// Header row
const COLOR_HEADER_BG   = '#1a73e8';
const COLOR_HEADER_FG   = '#ffffff';

// Summary row (one per filament type)
const COLOR_SUMMARY_BG  = '#f8f9fa';

// Detail row (one per color variant)
const COLOR_DETAIL_BG   = '#ffffff';

// Stock status cell backgrounds
const COLOR_IN_STOCK    = '#e6f4ea';   // light green
const COLOR_OUT_STOCK   = '#fce8e6';   // light red
const COLOR_PART_STOCK  = '#fef9e7';   // light yellow

// Row height (px) for rows that contain an image thumbnail
const IMAGE_ROW_HEIGHT  = 60;
