"""
Generates a local HTML dashboard from the scraper's output.xlsx.

Usage:
    pip install pandas openpyxl
    python3 dashboard.py output.xlsx dashboard.html

Opens as a plain static file in any browser (no server needed). Reads the
'ranked' and 'excluded' sheets and embeds them as JSON in the page, so
searching/sorting/filtering all happens client-side in vanilla JS.
"""

import json
import sys
import html as htmlmod
from datetime import datetime

import pandas as pd


def load_data(xlsx_path: str):
    ranked = pd.read_excel(xlsx_path, sheet_name="ranked")
    try:
        excluded = pd.read_excel(xlsx_path, sheet_name="excluded")
    except Exception:
        excluded = pd.DataFrame()

    ranked = ranked.sort_values("price_per_puff").reset_index(drop=True)

    def clean(df):
        # .astype(object) first avoids a pandas gotcha where inserting None
        # into a float-dtype column silently coerces it back to NaN — which
        # would embed non-standard `NaN` tokens into the page's JSON instead
        # of proper `null`.
        df = df.astype(object).where(pd.notnull(df), None)
        return df.to_dict(orient="records")

    return clean(ranked), clean(excluded)


TEMPLATE = """<!DOCTYPE html>
<html lang="en">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1.0">
<title>Vape Price Board</title>
<link rel="preconnect" href="https://fonts.googleapis.com">
<link rel="preconnect" href="https://fonts.gstatic.com" crossorigin>
<link href="https://fonts.googleapis.com/css2?family=JetBrains+Mono:wght@400;500;700&family=Space+Grotesk:wght@400;500;600;700&display=swap" rel="stylesheet">
<style>
  :root {
    --bg-void: #0F1214;
    --bg-panel: #181F23;
    --bg-panel-hi: #1F282D;
    --line: #2A3339;
    --line-soft: #212A2F;
    --led-amber: #FFB020;
    --led-amber-dim: #8A6018;
    --led-green: #4ADE80;
    --led-green-dim: #245C3B;
    --text-hi: #EDF3F5;
    --text-lo: #8B99A2;
    --text-faint: #566067;
    --radius: 10px;
  }

  * { box-sizing: border-box; }

  body {
    margin: 0;
    background: var(--bg-void);
    background-image:
      radial-gradient(circle at 15% 0%, rgba(255,176,32,0.05), transparent 40%),
      radial-gradient(circle at 85% 10%, rgba(74,222,128,0.04), transparent 35%);
    color: var(--text-hi);
    font-family: 'Space Grotesk', sans-serif;
    -webkit-font-smoothing: antialiased;
  }

  .digital { font-family: 'JetBrains Mono', monospace; }

  header.board {
    padding: 28px 32px 22px;
    border-bottom: 1px solid var(--line);
    position: sticky;
    top: 0;
    background: rgba(15,18,20,0.92);
    backdrop-filter: blur(6px);
    z-index: 10;
  }

  .board-top {
    display: flex;
    align-items: baseline;
    justify-content: space-between;
    flex-wrap: wrap;
    gap: 10px;
  }

  .board-title {
    font-size: 22px;
    font-weight: 700;
    letter-spacing: 0.02em;
    display: flex;
    align-items: baseline;
    gap: 10px;
  }
  .board-title .dot {
    width: 9px; height: 9px; border-radius: 50%;
    background: var(--led-green);
    box-shadow: 0 0 8px var(--led-green);
    display: inline-block;
    transform: translateY(-1px);
  }

  .board-sub {
    font-size: 12.5px;
    color: var(--text-lo);
    font-family: 'JetBrains Mono', monospace;
    letter-spacing: 0.01em;
  }

  .controls {
    margin-top: 18px;
    display: flex;
    gap: 10px;
    flex-wrap: wrap;
    align-items: center;
  }

  input[type="text"], select {
    background: var(--bg-panel);
    border: 1px solid var(--line);
    color: var(--text-hi);
    font-family: 'Space Grotesk', sans-serif;
    font-size: 13.5px;
    padding: 9px 12px;
    border-radius: 7px;
    outline: none;
  }
  input[type="text"] { min-width: 220px; }
  input[type="text"]:focus, select:focus { border-color: var(--led-amber-dim); }
  input[type="text"]::placeholder { color: var(--text-faint); }

  .chip {
    font-family: 'JetBrains Mono', monospace;
    font-size: 11.5px;
    padding: 6px 10px;
    border: 1px solid var(--line);
    border-radius: 999px;
    color: var(--text-lo);
    cursor: pointer;
    user-select: none;
    transition: border-color .15s, color .15s, background .15s;
    white-space: nowrap;
  }
  .chip:hover { color: var(--text-hi); border-color: var(--text-faint); }
  .chip.active {
    color: var(--bg-void);
    background: var(--led-amber);
    border-color: var(--led-amber);
    font-weight: 500;
  }

  .chiprow { display: flex; gap: 8px; flex-wrap: wrap; }

  main { padding: 26px 32px 60px; max-width: 1400px; margin: 0 auto; }

  .grid {
    display: grid;
    grid-template-columns: repeat(auto-fill, minmax(280px, 1fr));
    gap: 16px;
  }

  .card {
    background: var(--bg-panel);
    border: 1px solid var(--line);
    border-radius: var(--radius);
    padding: 16px;
    display: flex;
    flex-direction: column;
    gap: 10px;
    position: relative;
    transition: border-color .15s, transform .15s;
  }
  .card:hover { border-color: var(--text-faint); transform: translateY(-1px); }
  .card.best { border-color: var(--led-green-dim); }

  .ribbon {
    position: absolute;
    top: -1px; left: 14px;
    background: var(--led-green);
    color: #06210f;
    font-family: 'JetBrains Mono', monospace;
    font-size: 10px;
    font-weight: 700;
    letter-spacing: 0.06em;
    padding: 3px 8px;
    border-radius: 0 0 5px 5px;
  }

  .card-top {
    display: flex;
    justify-content: space-between;
    align-items: flex-start;
    gap: 8px;
  }

  .domain {
    font-family: 'JetBrains Mono', monospace;
    font-size: 10.5px;
    color: var(--text-lo);
    letter-spacing: 0.03em;
    text-transform: uppercase;
  }

  .rank {
    font-family: 'JetBrains Mono', monospace;
    font-size: 10.5px;
    color: var(--text-faint);
  }

  .title {
    font-size: 14px;
    line-height: 1.35;
    font-weight: 500;
    color: var(--text-hi);
    display: -webkit-box;
    -webkit-line-clamp: 3;
    -webkit-box-orient: vertical;
    overflow: hidden;
    min-height: 57px;
  }

  .meta-row {
    display: flex;
    gap: 14px;
    font-size: 12px;
    color: var(--text-lo);
    font-family: 'JetBrains Mono', monospace;
  }
  .meta-row span b { color: var(--text-hi); font-weight: 500; }

  .flavors {
    font-size: 11.5px;
    color: var(--text-faint);
    line-height: 1.5;
    max-height: 32px;
    overflow: hidden;
  }

  .readout {
    margin-top: 4px;
    background: #0B0D0E;
    border: 1px solid var(--line);
    border-radius: 7px;
    padding: 10px 12px;
    display: flex;
    justify-content: space-between;
    align-items: baseline;
  }
  .readout-label {
    font-size: 9.5px;
    color: var(--text-faint);
    text-transform: uppercase;
    letter-spacing: 0.07em;
    font-family: 'JetBrains Mono', monospace;
  }
  .readout-value {
    font-family: 'JetBrains Mono', monospace;
    font-size: 20px;
    font-weight: 700;
    color: var(--led-amber);
    text-shadow: 0 0 12px rgba(255,176,32,0.35);
  }
  .card.best .readout-value {
    color: var(--led-green);
    text-shadow: 0 0 12px rgba(74,222,128,0.4);
  }
  .readout-value small { font-size: 11px; font-weight: 500; opacity: .75; }

  .price-line {
    font-size: 12px;
    color: var(--text-lo);
    font-family: 'JetBrains Mono', monospace;
  }

  .bulk-note {
    font-size: 11px;
    color: var(--text-lo);
    display: flex;
    align-items: center;
    gap: 6px;
    flex-wrap: wrap;
  }
  .bulk-tag {
    font-family: 'JetBrains Mono', monospace;
    font-size: 10px;
    padding: 2px 7px;
    border-radius: 999px;
    border: 1px solid var(--led-amber-dim);
    color: var(--led-amber);
  }

  a.buy {
    margin-top: 4px;
    display: flex;
    align-items: center;
    justify-content: center;
    gap: 6px;
    text-decoration: none;
    color: var(--bg-void);
    background: var(--text-hi);
    font-weight: 600;
    font-size: 12.5px;
    padding: 9px;
    border-radius: 7px;
    transition: opacity .15s;
  }
  a.buy:hover { opacity: 0.85; }

  .empty-state {
    padding: 60px 20px;
    text-align: center;
    color: var(--text-faint);
    font-family: 'JetBrains Mono', monospace;
    font-size: 13px;
  }

  details.excluded-panel {
    margin-top: 40px;
    border-top: 1px solid var(--line);
    padding-top: 18px;
  }
  details.excluded-panel summary {
    cursor: pointer;
    font-family: 'JetBrains Mono', monospace;
    font-size: 12.5px;
    color: var(--text-lo);
    letter-spacing: 0.02em;
  }
  details.excluded-panel summary:hover { color: var(--text-hi); }
  .excluded-table {
    margin-top: 14px;
    width: 100%;
    border-collapse: collapse;
    font-size: 12px;
  }
  .excluded-table th, .excluded-table td {
    text-align: left;
    padding: 8px 10px;
    border-bottom: 1px solid var(--line-soft);
    color: var(--text-lo);
  }
  .excluded-table th {
    font-family: 'JetBrains Mono', monospace;
    text-transform: uppercase;
    font-size: 10px;
    letter-spacing: 0.05em;
    color: var(--text-faint);
  }
  .excluded-table td.title-cell { color: var(--text-hi); max-width: 340px; }
  .reason-tag {
    font-family: 'JetBrains Mono', monospace;
    font-size: 10.5px;
    padding: 2px 8px;
    border-radius: 999px;
    border: 1px solid var(--line);
    color: var(--text-lo);
    white-space: nowrap;
  }

  @media (max-width: 520px) {
    header.board, main { padding-left: 16px; padding-right: 16px; }
    .grid { grid-template-columns: 1fr; }
  }

  @media (prefers-reduced-motion: reduce) {
    .card { transition: none; }
  }
</style>
</head>
<body>

<header class="board">
  <div class="board-top">
    <div class="board-title"><span class="dot"></span>VAPE PRICE BOARD</div>
    <div class="board-sub">__SUBTITLE__</div>
  </div>
  <div class="controls">
    <input type="text" id="search" placeholder="Search title or flavor…">
    <select id="sort">
      <option value="ppp_asc">Sort: price / 1000 puffs (low → high)</option>
      <option value="ppp_desc">Sort: price / 1000 puffs (high → low)</option>
      <option value="price_asc">Sort: price (low → high)</option>
      <option value="puffs_desc">Sort: puffs (high → low)</option>
    </select>
    <div class="chiprow" id="domain-chips"></div>
  </div>
</header>

<main>
  <div class="grid" id="grid"></div>
  <div class="empty-state" id="empty" style="display:none;">No listings match those filters.</div>

  <details class="excluded-panel">
    <summary id="excluded-summary">▸ Show excluded listings</summary>
    <table class="excluded-table">
      <thead>
        <tr><th>Domain</th><th>Title</th><th>Reason</th><th></th></tr>
      </thead>
      <tbody id="excluded-body"></tbody>
    </table>
  </details>
</main>

<script>
const RANKED = __RANKED_JSON__;
const EXCLUDED = __EXCLUDED_JSON__;

function fmtMoney(v, currency) {
  if (v === null || v === undefined) return '—';
  const sym = currency || '';
  return sym + v.toFixed(2);
}

function pppPer1000(ppp) {
  // price_per_puff is a tiny fraction (e.g. 0.00013); scale to "per 1000
  // puffs" so it reads as a normal price, like a fuel pump's price/litre.
  return (ppp * 1000);
}

function domainOf(url) {
  try { return new URL(url).hostname.replace(/^www\\./, ''); } catch(e) { return ''; }
}

let activeDomains = new Set();
let allDomains = [...new Set(RANKED.map(r => r.domain))].sort();

function buildChips() {
  const row = document.getElementById('domain-chips');
  row.innerHTML = '';
  allDomains.forEach(d => {
    const chip = document.createElement('div');
    chip.className = 'chip active';
    chip.textContent = d;
    chip.dataset.domain = d;
    chip.addEventListener('click', () => {
      if (activeDomains.has(d)) { activeDomains.delete(d); chip.classList.remove('active'); }
      else { activeDomains.add(d); chip.classList.add('active'); }
      render();
    });
    row.appendChild(chip);
  });
  activeDomains = new Set(allDomains);
}

function cardHTML(r, rank) {
  const best = rank < 3;
  const flavor = r.flavor && r.flavor.trim() ? r.flavor : 'Flavor: not listed';
  const puffs = r.puffs ? Number(r.puffs).toLocaleString() : '—';
  const ppp1000 = pppPer1000(r.price_per_puff);
  const hasBulk = r.bulk_offer && r.bulk_price_per_puff;
  const bulkPpp1000 = hasBulk ? pppPer1000(r.bulk_price_per_puff) : null;
  return `
    <div class="card ${best ? 'best' : ''}">
      ${best ? `<div class="ribbon">BEST VALUE</div>` : ''}
      <div class="card-top">
        <div class="domain">${r.domain || ''}</div>
        <div class="rank">#${rank + 1}</div>
      </div>
      <div class="title">${r.title || ''}</div>
      <div class="meta-row">
        <span><b>${puffs}</b> puffs</span>
        <span>pack of <b>${r.pack_size || 1}</b></span>
      </div>
      <div class="flavors" title="${flavor}">${flavor}</div>
      <div class="readout">
        <div>
          <div class="readout-label">per 1000 puffs</div>
          <div class="readout-value">${r.currency || ''}${ppp1000.toFixed(3)}</div>
        </div>
        <div class="price-line">${fmtMoney(r.unit_price, r.currency)} / unit</div>
      </div>
      ${hasBulk ? `
      <div class="bulk-note">
        <span class="bulk-tag">${r.bulk_offer}</span>
        <span>→ effective ${r.currency || ''}${bulkPpp1000.toFixed(3)} / 1000 puffs</span>
      </div>` : ''}
      <a class="buy" href="${r.url}" target="_blank" rel="noopener noreferrer">View listing →</a>
    </div>
  `;
}

function render() {
  const q = document.getElementById('search').value.trim().toLowerCase();
  const sortMode = document.getElementById('sort').value;

  let rows = RANKED.filter(r => activeDomains.has(r.domain));
  if (q) {
    rows = rows.filter(r =>
      (r.title || '').toLowerCase().includes(q) ||
      (r.flavor || '').toLowerCase().includes(q)
    );
  }

  rows = rows.slice().sort((a, b) => {
    if (sortMode === 'ppp_asc') return a.price_per_puff - b.price_per_puff;
    if (sortMode === 'ppp_desc') return b.price_per_puff - a.price_per_puff;
    if (sortMode === 'price_asc') return (a.unit_price||0) - (b.unit_price||0);
    if (sortMode === 'puffs_desc') return (b.puffs||0) - (a.puffs||0);
    return 0;
  });

  const grid = document.getElementById('grid');
  const empty = document.getElementById('empty');
  if (rows.length === 0) {
    grid.innerHTML = '';
    empty.style.display = 'block';
  } else {
    empty.style.display = 'none';
    // "Best value" ribbon only makes sense on the true global cheapest,
    // so only mark top-3 when sorted ascending by price/puff.
    grid.innerHTML = rows.map((r, i) =>
      cardHTML(r, sortMode === 'ppp_asc' ? i : 99)
    ).join('');
  }
}

function buildExcluded() {
  const body = document.getElementById('excluded-body');
  const summary = document.getElementById('excluded-summary');
  summary.textContent = `▸ Show excluded listings (${EXCLUDED.length})`;
  body.innerHTML = EXCLUDED.map(r => `
    <tr>
      <td>${r.domain || ''}</td>
      <td class="title-cell">${r.title || ''}</td>
      <td><span class="reason-tag">${r.excluded_reason || ''}</span></td>
      <td><a class="buy" style="padding:5px 10px;font-size:11px;display:inline-flex;" href="${r.url}" target="_blank" rel="noopener noreferrer">view →</a></td>
    </tr>
  `).join('');
}

document.getElementById('search').addEventListener('input', render);
document.getElementById('sort').addEventListener('change', render);

buildChips();
render();
buildExcluded();
</script>

</body>
</html>
"""


def main():
    if len(sys.argv) != 3:
        print("Usage: python3 dashboard.py output.xlsx dashboard.html")
        sys.exit(1)

    xlsx_path, out_path = sys.argv[1], sys.argv[2]
    ranked, excluded = load_data(xlsx_path)

    n_domains = len({r["domain"] for r in ranked if r.get("domain")})
    subtitle = (
        f"{len(ranked)} listings · {n_domains} storefronts · "
        f"generated {datetime.now().strftime('%Y-%m-%d %H:%M')}"
    )

    html_out = (
        TEMPLATE
        .replace("__SUBTITLE__", htmlmod.escape(subtitle))
        .replace("__RANKED_JSON__", json.dumps(ranked))
        .replace("__EXCLUDED_JSON__", json.dumps(excluded))
    )

    with open(out_path, "w", encoding="utf-8") as f:
        f.write(html_out)

    print(f"Dashboard written to {out_path} — open it directly in a browser.")


if __name__ == "__main__":
    main()
