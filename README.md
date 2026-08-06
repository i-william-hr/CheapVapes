# Vape Price Board

> **⚠️ Beta.** This scrapes JS-rendered e-commerce pages with pattern-matching and heuristics, not a maintained per-site integration. It can and will get things wrong — misparsed prices, missed stock/flavor data, or fields it can't find at all. **Verify anything before relying on it, especially price and stock status.** Treat the output as a starting point for manual comparison, not ground truth.

Two scripts: one crawls a list of vape storefronts and ranks listings by **price per puff**, the other turns the result into a local, browsable dashboard.

---

## What it does

### `scrape_vapes.py`
Given a list of storefront URLs:

1. **Discovers product pages** — loads each seed URL in a headless browser, waits for JS to render, and crawls out links that look like individual product pages (`/products/`, `/product/`, one hop into `/collections/*` if the homepage only links to categories).
2. **Scrapes each product page** — loads it headless (handling JS-rendered content and common age-verification / cookie-consent overlays), and extracts:
   - Title, price, currency
   - Puff count (including shorthand like `150K`)
   - Flavor (from actual variant-selector markup, not just nearby text)
   - Stock status (prefers structured JSON-LD data when the site provides it, since it's more reliable than scanning visible text)
   - Pack size (for explicit multi-packs, e.g. "3-Pack")
3. **Filters and computes**:
   - Drops out-of-stock listings
   - Drops listings with a literal "random" / "mystery" flavor
   - Drops multi-flavor-in-one-device listings (e.g. "4-in-1", "Flavor Switch") entirely
   - Drops e-liquid / nic salt bottle listings (not devices, no puff count)
   - Splits pack listings into a per-unit price
   - Computes `price_per_puff` and ranks ascending
   - Best-effort detects multi-buy promo pricing on the page ("Buy 2 Get 1 Free", "20% off 3+", "3 for €25") and records it as an informational `bulk_offer` / `bulk_price_per_puff` column — this does **not** change the main ranking, since it only pays off if you actually buy that many
4. **Outputs** `output.xlsx` with three tabs — `ranked`, `excluded` (with a reason per row), and `all_raw` (everything, unfiltered) — plus a `_discovered_urls.csv` manifest and a `_debug_snippets.csv` for troubleshooting.

### `dashboard.py`
Reads `output.xlsx` and generates a single self-contained `dashboard.html` — a card-per-listing view styled like a fuel-pump price board, with:
- A big digital price-per-1,000-puffs readout per card (green for the current cheapest three)
- Free-text search across title/flavor
- Sort by price/puff, price, or puff count
- Toggleable filter chips per storefront
- A collapsible table of everything that got excluded, and why
- Direct "View listing" links back to the source page
- If a listing has a multi-buy promo (e.g. "Buy 2 Get 1 Free"), a secondary line shows the effective bulk price per 1,000 puffs alongside the single-unit price

No server required — it's one static HTML file with the data embedded, open it in any browser.

---

## Usage

```bash
pip install playwright beautifulsoup4 pandas openpyxl
playwright install chromium

# 1. Scrape
python3 scrape_vapes.py urls.txt output.xlsx

# 2. Build the dashboard
python3 dashboard.py output.xlsx dashboard.html
```

Then open `dashboard.html` directly in a browser.

`urls.txt` — one storefront homepage or collection URL per line.

---

## Known limitations

- **Site variation**: these are different templates/platforms, not one integration. Extraction is generic (regex + common patterns + JSON-LD), so a given site's markup may not match cleanly. Check `_debug_snippets.csv` and the `excluded` tab when something looks off.
- **JSON-LD dependency**: price/stock accuracy is much better on sites that embed `Product`/`Offer` schema. Sites without it fall back to weaker text-scraping heuristics and are more likely to be wrong.
- **Flavor detection is best-effort**: it targets actual variant-selector markup, but non-standard flavor UIs may return blank rather than a guess (blank is intentional — better than a wrong guess).
- **No JS execution verification**: the discovery/scrape logic assumes standard Shopify-family behavior (age-gate button text, standard product URL patterns). Non-standard storefronts may need a per-domain override — there's a marked spot in `scrape_vapes.py` (`extract_fields()`) to add one.
- **Promotional pricing**: grabs the currently-displayed price, which may be a temporary sale/discount price, not a stable reference price.
- **Multi-buy discount detection is pattern-based and incomplete**: it catches a handful of common phrasings ("Buy X Get Y Free", "N% off X+", "X for €Y") but not every storefront's wording — a real discount can go undetected. Treat the `bulk_offer` column as "found some," not "found all."
- **Rendering method**: this crawls headless via Playwright (a background browser instance), not by driving your actual open Chrome window over the DevTools Protocol (CDP). Headless is what makes running 10+ sites in parallel practical. Attaching to a real Chrome window instead only tends to matter if a site is actively fingerprinting/blocking headless browsers — if you start seeing consistent blocks or empty pages on a specific domain that work fine when you visit manually, that's the signal to switch that domain to a CDP-attached approach; it's not needed as a default for sites like these.

If something looks wrong in the output, the most useful thing to share back is the specific product URL plus the relevant row from `excluded`/`all_raw` — that's usually enough to pin down whether it's a one-off site quirk or a real bug in the extraction logic.
