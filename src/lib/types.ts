// ---------------------------------------------------------------------------
// Shopify public storefront API shapes
// ---------------------------------------------------------------------------

type ShopifyImage = {
    id: number;
    src: string;
    width: number;
    height: number;
    position: number;
};

type ShopifyVariant = {
    id: number;
    title: string;
    price: string;
    compare_at_price: string | null;
    available: boolean;
    option1: string | null;
    option2: string | null;
    option3: string | null;
};

type ShopifyProduct = {
    id: number;
    title: string;
    handle: string;
    body_html: string;
    tags: string;
    variants: ShopifyVariant[];
    images: ShopifyImage[];
};

type ShopifyProductsResponse = {
    products: ShopifyProduct[];
};

type ShopifyProductResponse = {
    product: ShopifyProduct;
};

// ---------------------------------------------------------------------------
// Derived data used within the script
// ---------------------------------------------------------------------------

type PriceInfo = {
    display: string;
    compareAt: string;
};

type StockInfo = {
    summary: string;        // e.g. "28 / 32 in stock"
    allInStock: boolean;
    noneInStock: boolean;
    inStockCount: number;
    total: number;
};

// Tracks whether a sheet row is a filament summary or a per-color detail row.
// Stored in the hidden ROW_MARKER column (col Z) so the script can identify
// rows after the sheet has been modified.
enum RowMarker {
    Summary = 'SUMMARY',
    Detail  = 'DETAIL',
}
