// ---------------------------------------------------------------------------
// Shopify storefront API helpers
// ---------------------------------------------------------------------------
// Uses the public (no-auth) Shopify JSON API on the underlying myshopify.com
// domain, which is accessible from Google's servers without CORS or auth.
// ---------------------------------------------------------------------------

/**
 * Fetch every product in the filament collection, paginating with
 * Shopify's cursor-based `page_info` param if needed.
 */
function fetchAllProducts(): ShopifyProduct[] | null {
    const url = `${STORE_BASE}/collections/${COLLECTION}/products.json?limit=${PAGE_LIMIT}`;
    try {
        const resp = UrlFetchApp.fetch(url, { muteHttpExceptions: true });
        if (resp.getResponseCode() !== 200) {
            console.error(`fetchAllProducts HTTP ${resp.getResponseCode()}: ${url}`);
            return null;
        }
        const data = JSON.parse(resp.getContentText()) as ShopifyProductsResponse;
        return data.products ?? [];
    } catch (e) {
        console.error('fetchAllProducts error:', e);
        return null;
    }
}

/**
 * Fetch a single product by its URL handle (e.g. "bambu-pla-basic").
 * Used as a fallback when refreshing a specific row.
 */
function fetchProductByHandle(handle: string): ShopifyProduct | null {
    const url = `${STORE_BASE}/products/${handle}.json`;
    try {
        const resp = UrlFetchApp.fetch(url, { muteHttpExceptions: true });
        if (resp.getResponseCode() !== 200) {
            console.error(`fetchProductByHandle HTTP ${resp.getResponseCode()}: ${handle}`);
            return null;
        }
        const data = JSON.parse(resp.getContentText()) as ShopifyProductResponse;
        return data.product ?? null;
    } catch (e) {
        console.error('fetchProductByHandle error:', e);
        return null;
    }
}

// ---------------------------------------------------------------------------
// Data extraction helpers
// ---------------------------------------------------------------------------

/** Returns display price and compare-at price for the cheapest variant. */
function getPriceInfo(product: ShopifyProduct): PriceInfo {
    if (!product.variants.length) return { display: '', compareAt: '' };

    // Sort ascending so we show the base price
    const sorted = [...product.variants].sort(
        (a, b) => parseFloat(a.price) - parseFloat(b.price)
    );
    const cheapest = sorted[0]!;
    const price = parseFloat(cheapest.price);
    const display = isNaN(price) ? '' : `$${price.toFixed(2)}`;

    let compareAt = '';
    if (cheapest.compare_at_price) {
        const cmp = parseFloat(cheapest.compare_at_price);
        if (!isNaN(cmp) && cmp > price) compareAt = `$${cmp.toFixed(2)}`;
    }

    return { display, compareAt };
}

/** Summarises variant stock across all color/weight combinations. */
function getStockInfo(product: ShopifyProduct): StockInfo {
    const total = product.variants.length;
    const inStockCount = product.variants.filter(v => v.available).length;

    const allInStock  = inStockCount === total;
    const noneInStock = inStockCount === 0;
    const summary     = total === 1
        ? (allInStock ? 'In stock' : 'Out of stock')
        : `${inStockCount} / ${total} in stock`;

    return { summary, allInStock, noneInStock, inStockCount, total };
}

/**
 * Returns the best thumbnail URL for the product's hero image.
 * Prefers the first image; resizes via Shopify CDN width param if possible.
 */
function getBestImageUrl(product: ShopifyProduct): string {
    const img = product.images[0];
    if (!img) return '';

    // Shopify CDN supports `_192x` size suffix before the extension.
    // src looks like: https://cdn.shopify.com/.../filename.jpg?v=1234
    const src = img.src;
    const qIdx = src.indexOf('?');
    const base = qIdx >= 0 ? src.slice(0, qIdx) : src;
    const query = qIdx >= 0 ? src.slice(qIdx) : '';

    // Insert _192x before the final extension
    const dotIdx = base.lastIndexOf('.');
    if (dotIdx < 0) return src;
    const sized = `${base.slice(0, dotIdx)}_${IMAGE_WIDTH}x${base.slice(dotIdx)}${query}`;
    return sized;
}

/**
 * Extracts the Shopify product handle from a store URL.
 * Accepts both us.store.bambulab.com and bambulab-us.myshopify.com URLs.
 * Returns null if the URL doesn't look like a product page.
 */
function handleFromUrl(url: string): string | null {
    const match = url.match(/\/products\/([^/?#]+)/);
    return match ? (match[1] ?? null) : null;
}
