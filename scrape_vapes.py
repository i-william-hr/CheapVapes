"""
Vape ad price-per-puff scraper.

Usage:
    pip install playwright beautifulsoup4 pandas openpyxl
    playwright install chromium
    python scrape_vapes.py urls.txt output.xlsx

urls.txt = one storefront/homepage/collection URL per line.

WHAT IT DOES
------------
0. DISCOVERY PASS: loads each seed URL, waits for JS to render, and crawls
   out links that look like individual product pages (/products/, /product/,
   one hop into /collections/*  if the homepage only links to categories).
   Falls back to scraping the seed page itself if no product links are found.
   Writes a `<output>_discovered_urls.csv` manifest mapping seed -> product
   pages found, so you can see/audit what got crawled.
1. Loads each discovered product URL in a headless Chromium browser (so
   JS-rendered React/Shopify pages actually populate before we read the DOM).
2. Extracts: title, price, currency, puff count, flavor, stock status, pack size.
3. Splits "buy 2 get 20% off" / "3-pack" style bundle listings into a
   per-unit row so price-per-puff is comparable across listings.
4. Drops out-of-stock listings.
5. Drops listings whose flavor is literally "random" / "mystery" (per your spec).
6. Computes price_per_puff = unit_price / puffs, sorts ascending.
7. Writes a .xlsx with both the raw extracted fields and the computed columns,
   so you can audit or re-filter.

NOTE ON SITE VARIATION
-----------------------
These are different storefronts (mostly Shopify or Shopify-clone templates,
based on what I inspected), so selectors vary a bit site to site. The
extractor below uses a generic, resilient strategy (regex over visible text +
common CSS patterns) rather than hard-coded selectors per site, which is more
robust to template differences but you WILL want to spot-check the output
against a few pages and adjust the regexes in `extract_fields()` if a
particular site's markup doesn't match cleanly. Search for "SITE-SPECIFIC
OVERRIDES" below to add per-domain tweaks if needed.
"""

import asyncio
import csv
import json
import re
import sys
from dataclasses import dataclass, field
from urllib.parse import urlparse

import pandas as pd
from bs4 import BeautifulSoup
from playwright.async_api import async_playwright

# ---------------------------------------------------------------------------
# Config
# ---------------------------------------------------------------------------

NAV_TIMEOUT_MS = 30_000
WAIT_AFTER_LOAD_MS = 2500          # let JS/XHR settle after DOM load
CONCURRENCY = 3                    # parallel browser tabs
RANDOM_FLAVOR_TERMS = {"random", "mystery", "surprise"}

OUT_OF_STOCK_TERMS = [
    "out of stock", "sold out", "unavailable", "currently unavailable",
    "notify me when available", "restock",
]

PUFF_RE = re.compile(r"([\d,]{2,7})\s*puffs?", re.IGNORECASE)
# Shorthand like "150K", "300K puffs", "42k puffs" used in titles without
# the full number spelled out, often paired with "disposable"/"vape".
PUFF_SHORTHAND_RE = re.compile(
    r"([\d]{1,3}(?:\.\d)?)\s*[kK]\b(?!\w)", re.IGNORECASE
)
PACK_RE = re.compile(
    r"(\d+)\s*[- ]?(?:pack|pcs|pc|piece|pieces|x)\b|\bpack\s*of\s*(\d+)",
    re.IGNORECASE,
)
PRICE_RE = re.compile(r"([€$£])\s?(\d{1,4}(?:[.,]\d{2})?)")

# ---------------------------------------------------------------------------
# Data model
# ---------------------------------------------------------------------------

@dataclass
class Listing:
    url: str
    domain: str
    title: str = ""
    currency: str = ""
    price: float | None = None          # listing price as shown (may be a bundle price)
    unit_price: float | None = None     # price per single vape after bundle split
    puffs: int | None = None
    pack_size: int = 1
    flavor: str = ""
    in_stock: bool = True
    price_per_puff: float | None = None
    excluded_reason: str = ""
    raw_snippet: str = ""
    # Multi-buy promo pricing (distinct from a multi-pack SKU): e.g. "buy 2
    # get 1 free" or "save 20% on 3+". Informational only — does not affect
    # single-unit ranking, since it only pays off if you actually buy that
    # many.
    bulk_offer: str = ""
    bulk_unit_price: float | None = None
    bulk_price_per_puff: float | None = None

# ---------------------------------------------------------------------------
# Extraction
# ---------------------------------------------------------------------------

def extract_jsonld_product(soup: BeautifulSoup) -> dict:
    """
    Most Shopify-family storefronts embed machine-readable product data in
    <script type="application/ld+json">, e.g.:
        {"@type": "Product", "offers": {"price": "7.99", "priceCurrency":
         "EUR", "availability": "http://schema.org/InStock"}, ...}
    This is far more reliable than scraping visible text, because it isn't
    affected by hidden "Notify me when available" markup that sits in the
    DOM (but CSS-hidden) for every product regardless of actual stock —
    which is what was causing near-100% false "out of stock" results.
    Returns {} if nothing usable is found (caller falls back to heuristics).
    """
    for script in soup.find_all("script", type="application/ld+json"):
        try:
            data = json.loads(script.string or "")
        except Exception:
            continue

        candidates = data if isinstance(data, list) else [data]
        # unwrap @graph if present (common in combined schema blocks)
        flat = []
        for c in candidates:
            if isinstance(c, dict) and "@graph" in c:
                flat.extend(c["@graph"])
            else:
                flat.append(c)

        for item in flat:
            if not isinstance(item, dict):
                continue
            item_type = item.get("@type", "")
            if item_type != "Product" and "Product" not in str(item_type):
                continue

            offers = item.get("offers", {})
            if isinstance(offers, list):
                offers = offers[0] if offers else {}

            price = offers.get("price") or offers.get("lowPrice")
            currency = offers.get("priceCurrency", "")
            availability = str(offers.get("availability", "")).lower()

            result = {}
            if price is not None:
                try:
                    result["price"] = float(str(price).replace(",", "."))
                    result["currency"] = currency
                except ValueError:
                    pass
            if availability:
                result["in_stock"] = "outofstock" not in availability.replace(" ", "")
            if item.get("name"):
                result["title"] = item["name"]
            if result:
                return result
    return {}


def guess_stock_status(page_text: str) -> bool:
    """
    Fallback only — used when no JSON-LD availability field is found.
    Deliberately narrow: many sites keep a hidden "Notify me when
    available" block in the DOM for every product (CSS-hidden, not
    removed), so broad terms like "unavailable" alone cause false
    positives. Require a more specific, less ambiguous phrase.
    """
    low = page_text.lower()
    strict_terms = ["out of stock", "sold out", "currently unavailable"]
    return not any(term in low for term in strict_terms)


def guess_puffs(text: str, allow_shorthand: bool = False) -> int | None:
    m = PUFF_RE.search(text)
    if m:
        return int(m.group(1).replace(",", ""))
    if allow_shorthand:
        # e.g. "QQ Bang 150K Disposable Vape" -> 150,000. Restricted to
        # title text (via allow_shorthand=True caller) to avoid false
        # matches on unrelated numbers elsewhere on the page (SKUs, etc).
        m2 = PUFF_SHORTHAND_RE.search(text)
        if m2:
            return int(float(m2.group(1)) * 1000)
    return None


def guess_pack_size(title: str, page_text: str) -> int:
    """
    Looks for explicit multi-pack language. Defaults to 1 (single vape).
    NOTE: "buy 2 get 20% off" promo banners are NOT the same as a bundle
    LISTING (that's a quantity discount on checkout, not a multi-vape SKU) —
    only treat it as a bundle if the title/variant itself says e.g. "3-Pack"
    or "Pack of 5".
    """
    m = PACK_RE.search(title)
    if m:
        n = m.group(1) or m.group(2)
        if n and n.isdigit():
            return int(n)
    return 1


def guess_price(text: str, title: str = "") -> tuple[str, float] | None:
    """
    Fallback only — used when JSON-LD offers/price isn't present.
    Searches from the product title's position onward (if found) rather
    than the very start of the page text, since promo banners /
    announcement bars ("free shipping over $39") often contain a price-like
    string earlier in the DOM than the actual product price.
    """
    search_text = text
    if title:
        idx = text.find(title)
        if idx != -1:
            search_text = text[idx:]
    matches = PRICE_RE.findall(search_text) or PRICE_RE.findall(text)
    if not matches:
        return None
    currency, amount = matches[0]
    return currency, float(amount.replace(",", "."))


def guess_flavor(soup: BeautifulSoup, page_text: str) -> str:
    """
    Targets the ACTUAL variant-selector markup rather than any text near
    the word "flavor" (the old version was grabbing cookie banners /
    language switchers / nav links whenever "flavor" happened to appear
    somewhere earlier in the DOM — that was the bug). Two concrete
    patterns, both common on Shopify-family storefronts:

      1. <select> whose name/id/data-attrs mention flavor/flavour/scent,
         OR whose <option> values are clearly flavor-like short phrases
         (fruit/dessert/menthol words) rather than nav/cookie text.
      2. A swatch/button group (data-option-name="Flavor" or similar) —
         collect data-value/text of each button.

    Returns "" (blank) rather than a guess if nothing clean is found —
    a blank flavor is far less misleading than garbage text.
    """
    NOISE_WORDS = [
        "cookie", "translation", "preference", "accept", "decline",
        "login", "log in", "cart", "shipping", "warehouse", "home",
        "checkout", "trusted store", "secure", "quantity", "menu",
        "sale ends", "day", "hour",
    ]

    def is_noisy(text: str) -> bool:
        low = text.lower()
        return any(w in low for w in NOISE_WORDS) or len(text) > 40

    # Pattern 1: <select> elements
    for select in soup.find_all("select"):
        attrs_blob = " ".join([
            select.get("name", ""), select.get("id", ""),
            " ".join(select.get("class", [])),
            str(select.get("data-option-name", "")),
        ]).lower()
        if not re.search(r"flavou?r|scent|geschmack|taste", attrs_blob):
            # Also check the nearest preceding label text
            label = select.find_previous(["label", "span", "div"])
            label_text = label.get_text(strip=True) if label else ""
            if not re.search(r"flavou?r", label_text, re.IGNORECASE):
                continue
        options = [
            o.get_text(strip=True) for o in select.find_all("option")
            if o.get_text(strip=True)
            and not re.search(r"choose|select an option|default title", o.get_text(strip=True), re.IGNORECASE)
        ]
        options = [o for o in options if not is_noisy(o)]
        if options:
            return " / ".join(options[:10])

    # Pattern 2: swatch/button groups with explicit flavor data attributes
    for el in soup.find_all(attrs={"data-option-name": re.compile(r"flavou?r", re.IGNORECASE)}):
        buttons = el.find_all(["button", "label", "span", "input"])
        vals = []
        for b in buttons:
            v = b.get("data-value") or b.get("value") or b.get_text(strip=True)
            if v and not is_noisy(v):
                vals.append(v.strip())
        vals = list(dict.fromkeys(vals))  # dedupe, keep order
        if vals:
            return " / ".join(vals[:10])

    # fallback: literal random/mystery near the word "flavor" (kept from
    # original — this specific narrow check is fine, unlike the removed
    # broad DOM-proximity search)
    for term in RANDOM_FLAVOR_TERMS:
        if re.search(rf"{term}\s+flavou?r", page_text, re.IGNORECASE):
            return term
    return ""


MULTI_FLAVOR_RE = re.compile(
    r"\b(\d+)[\s-]*in[\s-]*1\b|multi[\s-]?flavou?r|flavor\s+switch|flavor\s+chamber",
    re.IGNORECASE,
)


def is_multi_flavor_device(title: str) -> bool:
    """
    Flags devices explicitly sold as multiple flavors in one unit
    ("4-in-1", "Multi-Flavor", "Flavor Switch", "8 Flavor Chamber") —
    per your request, these get excluded entirely rather than ranked.
    """
    return bool(MULTI_FLAVOR_RE.search(title))


# Bottles of e-liquid / nic salts aren't disposable devices — they have no
# "puffs" figure and shouldn't be ranked alongside devices at all, even if
# a puff count happens to be scraped from an unrelated part of the page.
ELIQUID_RE = re.compile(
    r"\be[\s-]?liquid\b|\bnic\s*salt\b|\bsalt\s*nic\b|\bfreebase\b|\bshortfill\b",
    re.IGNORECASE,
)


def is_eliquid_listing(title: str) -> bool:
    """
    Catches liquid/bottle listings (e.g. 'VOEX Eliquid') so they're
    excluded from device price-per-puff ranking entirely.
    """
    return bool(ELIQUID_RE.search(title))


BULK_FREE_RE = re.compile(
    r"buy\s*(\d+)[^.\n]{0,20}get\s*(\d+)[^.\n]{0,10}free", re.IGNORECASE
)
BULK_PERCENT_RE = re.compile(
    r"(?:buy\s*(\d+)\+?|(\d+)\+)[^.\n]{0,20}?(\d{1,2})\s?%\s*(?:off|discount|savings?)"
    r"|(\d{1,2})\s?%\s*(?:off|discount|savings?)[^.\n]{0,20}?(?:buy\s*(\d+)\+?|(\d+)\+)",
    re.IGNORECASE,
)
BULK_FIXED_PRICE_RE = re.compile(
    r"(\d+)\s*(?:for|pcs? for)\s*[€$£]\s?(\d+(?:[.,]\d{2})?)", re.IGNORECASE
)


def detect_bulk_offer(page_text: str, single_price: float) -> dict:
    """
    Best-effort detection of multi-buy promo pricing that's layered on top
    of a single-unit listing (not a separate bundle SKU) — e.g. "Buy 2 Get
    1 Free" or "Save 20% on 3+". This is different from pack_size, which is
    for an actual multi-unit product listing.

    Returns {} if nothing is found. Result is informational — it does NOT
    change the primary single-unit price_per_puff ranking, since it only
    applies if you actually buy that many.
    """
    m = BULK_FREE_RE.search(page_text)
    if m:
        buy_n, free_n = int(m.group(1)), int(m.group(2))
        total_units = buy_n + free_n
        if total_units > 0:
            unit_price = round(single_price * buy_n / total_units, 4)
            return {
                "offer": f"Buy {buy_n} Get {free_n} Free",
                "unit_price": unit_price,
            }

    m = BULK_PERCENT_RE.search(page_text)
    if m:
        groups = m.groups()
        # groups: (qty_a, qty_b, pct_a) from first alt, or (pct_b, qty_c, qty_d) from second
        if groups[2] is not None:  # first alternative matched (qty ... pct)
            qty = groups[0] or groups[1]
            pct = int(groups[2])
        else:  # second alternative matched (pct ... qty)
            pct = int(groups[3])
            qty = groups[4] or groups[5]
        unit_price = round(single_price * (1 - pct / 100), 4)
        return {
            "offer": f"{pct}% off {qty}+" if qty else f"{pct}% off multi-buy",
            "unit_price": unit_price,
        }

    m = BULK_FIXED_PRICE_RE.search(page_text)
    if m:
        qty, total = int(m.group(1)), float(m.group(2).replace(",", "."))
        if qty > 1:
            unit_price = round(total / qty, 4)
            return {
                "offer": f"{qty} for {total}",
                "unit_price": unit_price,
            }

    return {}


def extract_fields(url: str, html: str) -> Listing:
    domain = urlparse(url).netloc
    soup = BeautifulSoup(html, "html.parser")
    page_text = soup.get_text(" ", strip=True)

    title_tag = soup.find(["h1"]) or soup.find("title")
    title = title_tag.get_text(strip=True) if title_tag else ""

    listing = Listing(url=url, domain=domain, title=title, raw_snippet=page_text[:300])

    jsonld = extract_jsonld_product(soup)

    # Prefer structured data (JSON-LD) when available — it's authoritative
    # and immune to hidden-DOM text-scraping false positives. Fall back to
    # regex/text heuristics only for whatever JSON-LD didn't provide.
    if "price" in jsonld:
        listing.price = jsonld["price"]
        listing.currency = jsonld.get("currency", "")
    else:
        price_info = guess_price(page_text, listing.title)
        if price_info:
            listing.currency, listing.price = price_info

    if jsonld.get("title"):
        listing.title = jsonld["title"]

    listing.puffs = guess_puffs(listing.title, allow_shorthand=True) or guess_puffs(page_text)
    listing.pack_size = guess_pack_size(listing.title, page_text)
    listing.flavor = guess_flavor(soup, page_text)

    if "in_stock" in jsonld:
        listing.in_stock = jsonld["in_stock"]
    else:
        listing.in_stock = guess_stock_status(page_text)

    if listing.price is not None:
        offer = detect_bulk_offer(page_text, listing.price)
        if offer:
            listing.bulk_offer = offer["offer"]
            listing.bulk_unit_price = offer["unit_price"]

    # ---- SITE-SPECIFIC OVERRIDES -----------------------------------------
    # Add per-domain tweaks here if the generic extractor misses something,
    # e.g.:
    # if "dkvape.shop" in domain:
    #     listing.in_stock = "OUT OF STOCK" not in html
    # -----------------------------------------------------------------------

    return listing


def split_and_filter(listings: list[Listing]) -> list[Listing]:
    """Apply bundle-splitting, stock filter, random-flavor filter,
    multi-flavor-device filter, and e-liquid filter."""
    kept = []
    for l in listings:
        if is_eliquid_listing(l.title):
            l.excluded_reason = "e-liquid / nic salt (not a device)"
            kept.append(l)
            continue
        if l.price is None or l.puffs is None:
            l.excluded_reason = "missing price or puff count"
            kept.append(l)
            continue
        if not l.in_stock:
            l.excluded_reason = "out of stock"
            kept.append(l)
            continue
        if is_multi_flavor_device(l.title):
            l.excluded_reason = "multi-flavor device (X-in-1)"
            kept.append(l)
            continue
        if l.flavor.strip().lower() in RANDOM_FLAVOR_TERMS:
            l.excluded_reason = "random/mystery flavor"
            kept.append(l)
            continue

        l.unit_price = round(l.price / l.pack_size, 4)
        l.price_per_puff = round(l.unit_price / l.puffs, 6)
        if l.bulk_unit_price is not None:
            l.bulk_price_per_puff = round(l.bulk_unit_price / l.puffs, 6)
        kept.append(l)
    return kept


# ---------------------------------------------------------------------------
# Playwright fetching
# ---------------------------------------------------------------------------

AGE_GATE_BUTTON_TEXTS = [
    "yes", "i am 18", "i'm 18", "i am over 18", "enter", "confirm",
    "18+", "verify", "accept", "i am of legal age", "continue",
]
COOKIE_BUTTON_TEXTS = ["accept", "accept all", "agree", "ok", "got it"]


async def dismiss_overlays(page):
    """
    Vape storefronts almost universally show an age-verification modal
    ("Are you 18+?") before the real page content is usable, and often a
    cookie-consent banner too. If these aren't dismissed, price/puff text
    may never render (client-side gated) or may sit behind an overlay.
    Best-effort: try a handful of common button texts, ignore failures.
    """
    for text in AGE_GATE_BUTTON_TEXTS + COOKIE_BUTTON_TEXTS:
        try:
            btn = page.get_by_role("button", name=re.compile(text, re.IGNORECASE))
            if await btn.count() > 0:
                await btn.first.click(timeout=1500)
                await page.wait_for_timeout(400)
        except Exception:
            pass
    # Some sites use plain clickable divs/links instead of <button>
    for text in AGE_GATE_BUTTON_TEXTS:
        try:
            el = page.get_by_text(re.compile(rf"^{text}$", re.IGNORECASE))
            if await el.count() > 0:
                await el.first.click(timeout=1500)
                await page.wait_for_timeout(400)
        except Exception:
            pass


async def fetch_one(context, url: str, sem: asyncio.Semaphore) -> Listing:
    async with sem:
        page = await context.new_page()
        try:
            await page.goto(url, timeout=NAV_TIMEOUT_MS, wait_until="domcontentloaded")
            await dismiss_overlays(page)
            await page.wait_for_timeout(WAIT_AFTER_LOAD_MS)
            html = await page.content()
            listing = extract_fields(url, html)
        except Exception as e:
            listing = Listing(url=url, domain=urlparse(url).netloc)
            listing.excluded_reason = f"fetch error: {e}"
        finally:
            await page.close()
        return listing


async def run(urls: list[str]) -> list[Listing]:
    sem = asyncio.Semaphore(CONCURRENCY)
    async with async_playwright() as p:
        browser = await p.chromium.launch(headless=True)
        context = await browser.new_context(
            user_agent=(
                "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 "
                "(KHTML, like Gecko) Chrome/124.0 Safari/537.36"
            )
        )
        tasks = [fetch_one(context, u, sem) for u in urls]
        results = await asyncio.gather(*tasks)
        await browser.close()
    return results


# ---------------------------------------------------------------------------
# Output
# ---------------------------------------------------------------------------

def to_dataframe(listings: list[Listing]) -> pd.DataFrame:
    rows = []
    for l in listings:
        rows.append({
            "domain": l.domain,
            "url": l.url,
            "title": l.title,
            "currency": l.currency,
            "listing_price": l.price,
            "pack_size": l.pack_size,
            "unit_price": l.unit_price,
            "puffs": l.puffs,
            "price_per_puff": l.price_per_puff,
            "flavor": l.flavor,
            "in_stock": l.in_stock,
            "excluded_reason": l.excluded_reason,
            "bulk_offer": l.bulk_offer,
            "bulk_unit_price": l.bulk_unit_price,
            "bulk_price_per_puff": l.bulk_price_per_puff,
        })
    df = pd.DataFrame(rows)
    return df


PRODUCT_LINK_PATTERNS = [
    re.compile(r"/products/"),        # Shopify
    re.compile(r"/product/"),         # WooCommerce and clones
    re.compile(r"/shop/"),
    re.compile(r"/collections/.+/products/"),
]

NON_PRODUCT_HINTS = [
    "/cart", "/account", "/login", "/register", "/search", "/blogs",
    "/pages/", "javascript:", "#", "/policies", "/checkout",
]


def looks_like_product_url(href: str) -> bool:
    if not href or any(h in href for h in NON_PRODUCT_HINTS):
        return False
    return any(p.search(href) for p in PRODUCT_LINK_PATTERNS)


async def discover_product_urls(context, seed_url: str, sem: asyncio.Semaphore,
                                 max_products: int = 60) -> list[str]:
    """
    Loads a storefront root/collection page, waits for JS to render the
    catalog, and pulls out links that look like individual product pages.
    Also follows one level of /collections/* pages if the root itself
    doesn't expose product links directly (common when the homepage only
    links to category pages, not products).
    """
    async with sem:
        page = await context.new_page()
        found: set[str] = set()
        try:
            await page.goto(seed_url, timeout=NAV_TIMEOUT_MS, wait_until="domcontentloaded")
            await dismiss_overlays(page)
            await page.wait_for_timeout(WAIT_AFTER_LOAD_MS)
            html = await page.content()
            soup = BeautifulSoup(html, "html.parser")
            base = f"{urlparse(seed_url).scheme}://{urlparse(seed_url).netloc}"

            def collect_from(soup_obj, base_url):
                links = set()
                for a in soup_obj.find_all("a", href=True):
                    href = a["href"]
                    if href.startswith("//"):
                        href = "https:" + href
                    elif href.startswith("/"):
                        href = base_url + href
                    elif not href.startswith("http"):
                        continue
                    if looks_like_product_url(href):
                        links.add(href.split("?")[0])
                return links

            found |= collect_from(soup, base)

            # If we didn't find product links directly, try one hop into
            # collection/category pages.
            if not found:
                category_links = set()
                for a in soup.find_all("a", href=True):
                    href = a["href"]
                    if href.startswith("/"):
                        href = base + href
                    if "/collections/" in href or "/category/" in href or "/shop" in href:
                        category_links.add(href.split("?")[0])
                for cat_url in list(category_links)[:5]:
                    try:
                        await page.goto(cat_url, timeout=NAV_TIMEOUT_MS, wait_until="domcontentloaded")
                        await page.wait_for_timeout(WAIT_AFTER_LOAD_MS)
                        chtml = await page.content()
                        csoup = BeautifulSoup(chtml, "html.parser")
                        found |= collect_from(csoup, base)
                    except Exception:
                        continue
        except Exception as e:
            print(f"  [discover error] {seed_url}: {e}")
        finally:
            await page.close()
        return list(found)[:max_products]


async def discover_all(seed_urls: list[str]) -> dict[str, list[str]]:
    sem = asyncio.Semaphore(CONCURRENCY)
    async with async_playwright() as p:
        browser = await p.chromium.launch(headless=True)
        context = await browser.new_context(
            user_agent=(
                "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 "
                "(KHTML, like Gecko) Chrome/124.0 Safari/537.36"
            )
        )
        tasks = {u: discover_product_urls(context, u, sem) for u in seed_urls}
        results = {}
        for u, coro in tasks.items():
            results[u] = await coro
        await browser.close()
    return results


def main():
    if len(sys.argv) != 3:
        print("Usage: python scrape_vapes.py urls.txt output.xlsx")
        sys.exit(1)

    urls_file, out_file = sys.argv[1], sys.argv[2]
    with open(urls_file) as f:
        seed_urls = [line.strip() for line in f if line.strip() and not line.startswith("#")]
    seed_urls = list(dict.fromkeys(seed_urls))  # dedupe, keep order

    print(f"Discovering product pages across {len(seed_urls)} storefronts...")
    discovered = asyncio.run(discover_all(seed_urls))

    urls = []
    manifest_rows = []
    for seed, products in discovered.items():
        print(f"  {seed}: {len(products)} product page(s) found")
        if not products:
            # fall back to scraping the seed page itself
            urls.append(seed)
            manifest_rows.append({"seed": seed, "product_url": seed, "note": "no subpages found, using seed"})
        for p in products:
            urls.append(p)
            manifest_rows.append({"seed": seed, "product_url": p, "note": ""})

    pd.DataFrame(manifest_rows).to_csv(out_file.replace(".xlsx", "_discovered_urls.csv"), index=False)
    print(f"Total product pages to scrape: {len(urls)}")

    listings = asyncio.run(run(urls))
    listings = split_and_filter(listings)
    df = to_dataframe(listings)

    included = df[df["excluded_reason"] == ""].sort_values("price_per_puff")
    excluded = df[df["excluded_reason"] != ""]

    with pd.ExcelWriter(out_file, engine="openpyxl") as writer:
        included.to_excel(writer, sheet_name="ranked", index=False)
        excluded.to_excel(writer, sheet_name="excluded", index=False)
        df.to_excel(writer, sheet_name="all_raw", index=False)

    print(f"Done. {len(included)} listings ranked, {len(excluded)} excluded.")
    print(f"Written to {out_file}")

    # ---- DIAGNOSTICS ---------------------------------------------------
    # If everything (or nearly everything) got excluded, print WHY so you
    # don't have to open the spreadsheet to start debugging.
    if len(excluded) > 0:
        print("\nExclusion reason breakdown:")
        print(excluded["excluded_reason"].value_counts().to_string())

        sample = excluded.head(3)
        print("\nSample of excluded rows (first 3):")
        for _, row in sample.iterrows():
            print(f"\n  URL: {row['url']}")
            print(f"  title: {row['title'][:80]!r}")
            print(f"  listing_price: {row['listing_price']}  puffs: {row['puffs']}  in_stock: {row['in_stock']}")
            print(f"  reason: {row['excluded_reason']}")

    debug_file = out_file.replace(".xlsx", "_debug_snippets.csv")
    pd.DataFrame([
        {"url": l.url, "title": l.title, "raw_snippet": l.raw_snippet}
        for l in listings
    ]).to_csv(debug_file, index=False)
    print(f"\nRaw text snippets (first 300 chars of each page's text) saved to {debug_file}")
    print("If everything was excluded, open that file and check whether the")
    print("snippet actually contains price/puff text near the top of the page —")
    print("if the real content loads further down (e.g. behind a cookie banner,")
    print("age-gate, or lazy-loaded block), the fix is usually to wait longer")
    print("or dismiss a blocking overlay before reading page.content().")


if __name__ == "__main__":
    main()
