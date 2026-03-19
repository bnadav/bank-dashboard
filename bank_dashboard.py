#!/usr/bin/env python3
"""
Bank Statement Dashboard Generator — Bank Hapoalim (בנק הפועלים)

Reads credit card Excel exports, renames them to MM_YYYY_CARD.xlsx,
and generates a self-contained HTML dashboard.

Usage:
    python bank_dashboard.py                     # auto-discovers *.xlsx in cwd
    python bank_dashboard.py file1.xlsx file2.xlsx
"""

import sys
import io
import os
import re
import glob as _glob
import json
from datetime import datetime
from collections import defaultdict

sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8')

try:
    import openpyxl
except ImportError:
    print("openpyxl not found. Run: pip install openpyxl")
    sys.exit(1)

# ── Constants ─────────────────────────────────────────────────────────────────

RENAMED_RE = re.compile(r'^\d{2}_\d{4}(_\d{4})?\.xlsx$')  # MM_YYYY.xlsx or MM_YYYY_CARD.xlsx

LOCAL_ANCHOR   = 'פירוט עבור הכרטיסים בארץ'
FOREIGN_ANCHOR = "פירוט עבור הכרטיסים בחו''ל"   # exact — not the דולר/יורו sub-sections

MONTH_HE = {
    1:'ינואר', 2:'פברואר', 3:'מרץ', 4:'אפריל', 5:'מאי', 6:'יוני',
    7:'יולי', 8:'אוגוסט', 9:'ספטמבר', 10:'אוקטובר', 11:'נובמבר', 12:'דצמבר',
}

# Categories: priority order, first keyword match wins
CATEGORIES = [
    ('דלק ותחבורה',    ['פז', 'yellow', 'דרך ארץ', 'מ. התחבורה', 'חניון', 'פנגו', 'ארגמן', 'רכב']),
    ('ביטוח',          ['ביטוח', 'מגדל', 'כלל ביט', 'איילון', 'כלל ח']),
    ('תקשורת',         ['פרטנר', 'hot mobile', 'אול מובייל', 'אקספון', 'סלקום', 'תקשורת']),
    ('מזון וסופרמרקט', ['carrefour', 'קרפור', 'שופרסל', 'רמי לוי', 'ויקטורי',
                         'סופר פארם', 'סופרפארם', 'פארמ', 'המשק', 'אלונית']),
    ('קפה ומסעדות',    ['ארומה', 'גרג', 'רולדין', 'מקדונלד', 'wolt', 'פלאפל', 'פיצה',
                         'מסעדה', 'קפה', 'לבנונית', 'אינגליש קייק', 'ממתקי', 'בית הוועד',
                         'גולדה', 'אניס', 'ביג אפל', 'צוללת', 'פינת', 'נייט קוקי',
                         'ניקוליסונס', 'פיצוחי']),
    ('בריאות ורפואה',  ['ביטוח לאומי', 'הדסה', 'בית חולים', 'הוטרינרי', 'ביה"ח',
                         'טרילו', 'ערן', 'ער"ן', 'קאופמן', 'רפואה', 'בריאות']),
    ('מנויים ובידור',  ['spotify', 'youtube', 'google', 'netflix', 'apple',
                         'headstart', 'edreams', 'optionstrat', 'camscanner', 'גלובס']),
    ('קניות',          ['איקאה', 'ikea', 'דלתא', 'סטימצקי', 'סבון של פעם',
                         'טופ ', 'פיקס מן', 'שלמה א.אנגל']),
    ('חינוך וספורט',   ['אור ירוק', 'כושר', 'אולימפוס', 'קופל', 'לנהיגה', 'המדור']),
    ('העברות ותרומות', ['bit', 'העברה', 'עיגול לטובה']),
    ('דמי כרטיס',      ['דמי כרטיס', 'פועלים- דמי', 'דמי כרטיס בנק']),
    ('שונות',          []),   # catch-all
]

PALETTE = [
    '#4361EE', '#F72585', '#7209B7', '#3A86FF', '#FB8500',
    '#219EBC', '#8338EC', '#06D6A0', '#FFB703', '#EF233C',
    '#2D6A4F', '#ADB5BD',
]

# ── Parsing helpers ───────────────────────────────────────────────────────────

def categorize(merchant: str, amount: float = 0.0) -> str:
    m = merchant.lower()
    # Yellow fuel app transactions under 100 ILS get their own category
    if amount < 100 and any(kw.lower() in m for kw in ['yellow', 'פז']):
        return 'Yellow'
    for name, kws in CATEGORIES[:-1]:
        if any(kw.lower() in m for kw in kws):
            return name
    return 'שונות'


def find_section(rows: list, anchor: str) -> int:
    """Return the 0-based index of the row where col A == anchor, or -1."""
    for i, row in enumerate(rows):
        if row[0] == anchor:
            return i
    return -1


def safe_float(v) -> float:
    try:
        return float(v) if v not in (None, '') else 0.0
    except (TypeError, ValueError):
        return 0.0


def parse_section(rows: list, start: int, card: str, section: str) -> list:
    """
    Parse transactions starting at `start` (the anchor row).
    Skips +3 rows (anchor + account-header + column-header) then reads
    data rows until col A is None.

    Local  columns: [0]=card [1]=bill_dt [2]=txn_dt [3]=merchant [4]=purchase [5]=charge [12]=type
    Foreign columns: [0]=card [1]=bill_dt [2]=txn_dt [3]=merchant [4]=charge [5]=purchase [6]=currency [13]=type
    """
    txns = []
    data_start = start + 3
    for row in rows[data_start:]:
        if row[0] is None:
            break
        card_cell = str(row[0]).strip()
        if not (card_cell.isdigit() and len(card_cell) == 4):
            continue   # stray non-data row
        try:
            txn_dt  = row[2]
            merchant = str(row[3] or '').strip()
            if section == 'local':
                charge = safe_float(row[5])
                purchase = safe_float(row[4])
                currency = 'ILS'
                txn_type = str(row[12] or '').strip()
            else:
                charge = safe_float(row[4])
                purchase = safe_float(row[5])
                currency = str(row[6] or 'USD').strip()
                txn_type = str(row[13] or '').strip()
            if not isinstance(txn_dt, datetime):
                continue
            txns.append({
                'card':      card_cell,
                'date':      txn_dt,
                'merchant':  merchant,
                'charge':    charge,      # amount charged this cycle (negative = credit)
                'purchase':  purchase,    # original purchase total
                'currency':  currency,
                'type':      txn_type,
                'foreign':   section == 'foreign',
                'category':  categorize(merchant, charge),
            })
        except (IndexError, TypeError):
            continue
    return txns


def process_file(path: str) -> tuple:
    """
    Load workbook, parse transactions, rename file.
    Returns (transactions: list, new_path: str).
    """
    print(f'Processing: {os.path.basename(path)}')
    try:
        wb = openpyxl.load_workbook(path)
    except Exception as e:
        print(f'  ✗ Cannot open: {e}')
        return [], path, None, None

    if 'גיליון1' not in wb.sheetnames:
        print(f'  ✗ Sheet "גיליון1" not found, skipping.')
        return [], path, None, None

    rows = list(wb['גיליון1'].iter_rows(values_only=True))

    # Extract billing period from the first local data row
    local_i = find_section(rows, LOCAL_ANCHOR)
    month, year = None, None
    if local_i != -1:
        for row in rows[local_i + 3:]:
            if row[0] is None:
                break
            s = str(row[0]).strip()
            if s.isdigit() and len(s) == 4 and isinstance(row[1], datetime):
                month = row[1].month
                year  = row[1].year
                break

    # Parse transactions first so we know how many distinct cards are in this file
    txns = []
    if local_i != -1:
        txns += parse_section(rows, local_i, '????', 'local')

    foreign_i = find_section(rows, FOREIGN_ANCHOR)
    if foreign_i != -1:
        txns += parse_section(rows, foreign_i, '????', 'foreign')

    cards_found = sorted(set(t['card'] for t in txns))

    # Rename: multi-card file → MM_YYYY.xlsx; single-card file → MM_YYYY_CARD.xlsx
    new_path = path
    if month and year:
        if len(cards_found) > 1:
            new_name = f'{month:02d}_{year}.xlsx'
        else:
            card4 = cards_found[0] if cards_found else '????'
            new_name = f'{month:02d}_{year}_{card4}.xlsx'
        new_path = os.path.join(os.path.dirname(os.path.abspath(path)), new_name)
        if os.path.abspath(path) != new_path:
            if os.path.exists(new_path):
                print(f'  ⚠ Target {new_name} already exists — duplicate file, skipping.')
                return [], new_path, month, year
            else:
                os.rename(path, new_path)
                print(f'  ✓ Renamed → {new_name}')

    print(f'  → Cards {", ".join(cards_found)}: {len(txns)} transactions')
    return txns, new_path, month, year


def discover_files(cli_args: list) -> list:
    if cli_args:
        # Explicit files: use as-is, no filtering
        return cli_args
    # Auto-discovery: skip already-renamed files
    candidates = _glob.glob('*.xlsx')
    files   = [f for f in candidates if not RENAMED_RE.match(os.path.basename(f))]
    skipped = [f for f in candidates if RENAMED_RE.match(os.path.basename(f))]
    for f in skipped:
        print(f'Skipping already-renamed file: {f}')
    return files


# ── HTML generation ───────────────────────────────────────────────────────────

def cat_color(cat: str, cat_order: list) -> str:
    idx = next((i for i, (n, _) in enumerate(cat_order) if n == cat), 0)
    return PALETTE[idx % len(PALETTE)]


def badge(cat: str, cat_order: list) -> str:
    c = cat_color(cat, cat_order)
    return (f'<span style="display:inline-block;padding:2px 8px;border-radius:20px;'
            f'font-size:.75rem;font-weight:500;background:{c}22;color:{c};'
            f'border:1px solid {c}55;white-space:nowrap">{cat}</span>')


def ils_cell(amount: float) -> str:
    if amount < 0:
        return f'<span style="color:#16a34a;font-weight:600">+{abs(amount):,.2f}&#x20AA;</span>'
    return f'<span style="color:#dc2626;font-weight:600">{amount:,.2f}&#x20AA;</span>'


def fdate(d: datetime) -> str:
    return d.strftime('%d/%m/%Y')


def build_html(txns: list, month: int, year: int) -> str:
    charges    = [t for t in txns if t['charge'] > 0]
    credits    = [t for t in txns if t['charge'] < 0]
    total_chg  = sum(t['charge'] for t in charges)
    total_crd  = abs(sum(t['charge'] for t in credits))
    avg_chg    = total_chg / len(charges) if charges else 0.0

    cards      = sorted(set(t['card'] for t in txns))
    card_totals = {c: sum(t['charge'] for t in charges if t['card'] == c) for c in cards}

    cat_agg    = defaultdict(float)
    for t in charges:
        cat_agg[t['category']] += t['charge']
    cat_order  = sorted(cat_agg.items(), key=lambda x: -x[1])

    top10      = sorted(charges, key=lambda x: -x['charge'])[:10]
    all_sorted = sorted(txns, key=lambda x: -x['date'].timestamp())

    month_he   = MONTH_HE.get(month, str(month))

    # ── Summary cards ──
    summary_html = (
        f'<div class="scard total">'
        f'<div class="sc-label">סה"כ הוצאות</div>'
        f'<div class="sc-amount">{total_chg:,.2f}&#x20AA;</div>'
        f'<div class="sc-sub">{len(charges)} עסקאות חיוב</div>'
        f'</div>'
    )
    for c in cards:
        pct = card_totals[c] / total_chg * 100 if total_chg else 0
        col = PALETTE[cards.index(c) % len(PALETTE)]
        bar = (f'<div style="height:5px;border-radius:3px;background:#eee;margin:6px 0">'
               f'<div style="width:{pct:.1f}%;height:5px;border-radius:3px;background:{col}"></div>'
               f'</div>')
        summary_html += (
            f'<div class="scard" style="border-top:3px solid {col}">'
            f'<div class="sc-label">כרטיס {c}</div>'
            f'<div class="sc-amount">{card_totals[c]:,.2f}&#x20AA;</div>'
            f'{bar}'
            f'<div class="sc-sub">{pct:.1f}% מסך ההוצאות</div>'
            f'</div>'
        )

    # ── Top 10 rows ──
    top10_html = ''
    for i, t in enumerate(top10, 1):
        top10_html += (
            f'<tr>'
            f'<td style="font-weight:700;color:#4361EE;width:36px">#{i}</td>'
            f'<td style="white-space:nowrap">{fdate(t["date"])}</td>'
            f'<td style="font-weight:500">{"🌍 " if t["foreign"] else ""}{t["merchant"]}</td>'
            f'<td>{badge(t["category"], cat_order)}</td>'
            f'<td style="text-align:left">{ils_cell(t["charge"])}</td>'
            f'<td style="font-family:monospace;color:#888;font-size:.8rem">···· {t["card"]}</td>'
            f'</tr>'
        )

    # ── Filter buttons ──
    filter_btns = '<button class="fbtn active" data-card="all">הכל</button>'
    for c in cards:
        filter_btns += f'<button class="fbtn" data-card="{c}">כרטיס {c}</button>'

    # ── All transactions rows ──
    all_rows_html = ''
    for t in all_sorted:
        all_rows_html += (
            f'<tr data-card="{t["card"]}">'
            f'<td style="white-space:nowrap">{fdate(t["date"])}</td>'
            f'<td style="font-weight:500">{"🌍 " if t["foreign"] else ""}{t["merchant"]}</td>'
            f'<td>{badge(t["category"], cat_order)}</td>'
            f'<td style="text-align:left">{ils_cell(t["charge"])}</td>'
            f'<td style="font-family:monospace;color:#888;font-size:.8rem">···· {t["card"]}</td>'
            f'<td style="color:#aaa;font-size:.78rem">{t["type"]}</td>'
            f'</tr>'
        )

    # ── Chart data (as JS constants, injected safely) ──
    chart_labels = json.dumps([n for n, _ in cat_order], ensure_ascii=False)
    chart_values = json.dumps([round(v, 2) for _, v in cat_order])
    chart_colors = json.dumps([PALETTE[i % len(PALETTE)] for i in range(len(cat_order))])

    # ── JS block (plain string, no f-string — so {} are literal) ──
    js_block = (
        "const chartLabels = " + chart_labels + ";\n"
        "const chartValues = " + chart_values + ";\n"
        "const chartColors = " + chart_colors + ";\n"
        """
Chart.defaults.font.family = "'Segoe UI', Arial, sans-serif";
Chart.defaults.font.size = 12;

new Chart(document.getElementById('pieChart'), {
  type: 'doughnut',
  data: {
    labels: chartLabels,
    datasets: [{
      data: chartValues,
      backgroundColor: chartColors,
      borderWidth: 2,
      borderColor: '#fff'
    }]
  },
  options: {
    responsive: true,
    maintainAspectRatio: false,
    plugins: {
      legend: {
        position: 'right',
        rtl: true,
        labels: { boxWidth: 12, padding: 10, font: { size: 11 } }
      },
      tooltip: {
        callbacks: {
          label: ctx => ' ' + ctx.label + ': ' + ctx.parsed.toLocaleString('he-IL', {minimumFractionDigits:2}) + ' ₪'
        }
      }
    }
  }
});

new Chart(document.getElementById('barChart'), {
  type: 'bar',
  data: {
    labels: chartLabels,
    datasets: [{
      data: chartValues,
      backgroundColor: chartColors,
      borderRadius: 4,
      borderSkipped: false
    }]
  },
  options: {
    indexAxis: 'y',
    responsive: true,
    maintainAspectRatio: false,
    plugins: {
      legend: { display: false },
      tooltip: {
        callbacks: {
          label: ctx => ' ' + ctx.parsed.x.toLocaleString('he-IL', {minimumFractionDigits:2}) + ' ₪'
        }
      }
    },
    scales: {
      x: {
        ticks: {
          callback: v => v.toLocaleString('he-IL') + '₪'
        },
        grid: { color: '#f0f0f0' }
      },
      y: {
        ticks: { font: { size: 11 } },
        grid: { display: false }
      }
    }
  }
});

// Card filter
document.querySelectorAll('.fbtn').forEach(btn => {
  btn.addEventListener('click', () => {
    document.querySelectorAll('.fbtn').forEach(b => b.classList.remove('active'));
    btn.classList.add('active');
    applyFilters();
  });
});

function applyFilters() {
  const activeCard = document.querySelector('.fbtn.active').dataset.card;
  const search = document.getElementById('searchBox').value.toLowerCase();
  document.querySelectorAll('#allTable tbody tr').forEach(row => {
    const cardMatch = activeCard === 'all' || row.dataset.card === activeCard;
    const merchant  = row.cells[1].textContent.toLowerCase();
    const searchMatch = !search || merchant.includes(search);
    row.style.display = (cardMatch && searchMatch) ? '' : 'none';
  });
}

document.getElementById('searchBox').addEventListener('input', applyFilters);

// ── Table sorting ────────────────────────────────────────────────────────────

function parseCellValue(text) {
  // Rank: #1, #2, ...
  if (/^#[0-9]+$/.test(text)) return parseInt(text.slice(1));
  // Date: DD/MM/YYYY
  if (/^[0-9]{2}[/][0-9]{2}[/][0-9]{4}$/.test(text)) {
    const [d, m, y] = text.split('/');
    return parseInt(y + m + d);
  }
  // Number / currency: strip ₪ ‏‎+ , · spaces
  const num = parseFloat(text.replace(/[₪,+ \t·\u200f\u200e]/g, ''));
  if (!isNaN(num)) return num;
  return text;
}

function makeTableSortable(tableId) {
  const table = document.getElementById(tableId);
  if (!table) return;
  const ths = Array.from(table.querySelectorAll('thead th'));
  let curCol = -1, curAsc = true;

  ths.forEach((th, idx) => {
    // Append sort-arrow indicator
    const arrow = document.createElement('span');
    arrow.style.cssText = 'margin-right:6px;font-size:.65rem;opacity:.3;vertical-align:middle;';
    arrow.textContent = '⇅';
    th.appendChild(arrow);
    th.style.cursor = 'pointer';
    th.style.userSelect = 'none';
    th.title = 'לחץ למיון';

    th.addEventListener('click', () => {
      if (curCol === idx) curAsc = !curAsc;
      else { curCol = idx; curAsc = true; }

      // Update all arrows
      ths.forEach((h, i) => {
        const a = h.querySelector('span');
        if (!a) return;
        if (i === idx) {
          a.textContent = curAsc ? ' ▲' : ' ▼';
          a.style.opacity = '1';
          a.style.color = '#4361EE';
        } else {
          a.textContent = '⇅';
          a.style.opacity = '.3';
          a.style.color = '';
        }
      });

      const tbody = table.querySelector('tbody');
      const rows = Array.from(tbody.querySelectorAll('tr'));

      rows.sort((a, b) => {
        const aVal = parseCellValue(a.cells[idx] ? a.cells[idx].textContent.trim() : '');
        const bVal = parseCellValue(b.cells[idx] ? b.cells[idx].textContent.trim() : '');
        if (typeof aVal === 'number' && typeof bVal === 'number') {
          return curAsc ? aVal - bVal : bVal - aVal;
        }
        return curAsc
          ? String(aVal).localeCompare(String(bVal), 'he')
          : String(bVal).localeCompare(String(aVal), 'he');
      });

      // Re-insert rows preserving their display state (filter not reset)
      rows.forEach(r => tbody.appendChild(r));
    });
  });
}

makeTableSortable('top10Table');
makeTableSortable('allTable');
"""
    )

    # ── Assemble full HTML (f-string; CSS/JS braces must be doubled) ──
    return f"""<!DOCTYPE html>
<html lang="he" dir="rtl">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width,initial-scale=1">
<title>דוח הוצאות — {month_he} {year}</title>
<script src="https://cdn.jsdelivr.net/npm/chart.js@4.4.0/dist/chart.umd.min.js"></script>
<style>
*, *::before, *::after {{ box-sizing: border-box; margin: 0; padding: 0; }}
body {{ font-family: 'Segoe UI', Arial, sans-serif; background: #F0F2F5; color: #1a1a2e; direction: rtl; }}
.page {{ max-width: 1200px; margin: 0 auto; padding: 24px 20px; }}

/* Header */
.header {{ margin-bottom: 24px; }}
.header h1 {{ font-size: 1.75rem; font-weight: 700; color: #1a1a2e; }}
.header p  {{ color: #777; margin-top: 4px; font-size: .9rem; }}

/* Summary cards */
.summary-row {{ display: flex; gap: 16px; margin-bottom: 20px; flex-wrap: wrap; }}
.scard {{ background: #fff; border-radius: 12px; padding: 20px; flex: 1; min-width: 150px;
          box-shadow: 0 2px 8px rgba(0,0,0,.07); border-top: 3px solid #4361EE; }}
.scard.total {{ border-top-color: #4361EE; }}
.sc-label  {{ font-size: .78rem; color: #888; margin-bottom: 6px; font-weight: 500; text-transform: uppercase; letter-spacing: .4px; }}
.sc-amount {{ font-size: 1.5rem; font-weight: 700; color: #1a1a2e; }}
.sc-sub    {{ font-size: .75rem; color: #aaa; margin-top: 4px; }}

/* Stats bar */
.stats-bar {{ display: flex; gap: 0; background: #fff; border-radius: 12px;
              box-shadow: 0 2px 8px rgba(0,0,0,.07); margin-bottom: 20px; overflow: hidden; flex-wrap: wrap; }}
.stat      {{ flex: 1; min-width: 120px; padding: 16px 20px; text-align: center; border-left: 1px solid #f0f0f0; }}
.stat:last-child {{ border-left: none; }}
.stat-n    {{ font-size: 1.3rem; font-weight: 700; color: #4361EE; }}
.stat-l    {{ font-size: .75rem; color: #888; margin-top: 2px; }}

/* Charts */
.charts-grid {{ display: grid; grid-template-columns: 1fr 1fr; gap: 16px; margin-bottom: 20px; }}
@media (max-width: 700px) {{ .charts-grid {{ grid-template-columns: 1fr; }} }}
.chart-card {{ background: #fff; border-radius: 12px; padding: 20px;
               box-shadow: 0 2px 8px rgba(0,0,0,.07); }}
.chart-card h2 {{ font-size: .9rem; font-weight: 600; color: #555; margin-bottom: 14px; }}
.chart-wrap {{ position: relative; height: 280px; }}

/* Section cards */
.section {{ background: #fff; border-radius: 12px; padding: 20px;
            box-shadow: 0 2px 8px rgba(0,0,0,.07); margin-bottom: 20px; }}
.section h2 {{ font-size: .9rem; font-weight: 600; color: #555; margin-bottom: 16px; }}

/* Tables */
table {{ width: 100%; border-collapse: collapse; font-size: .875rem; }}
th {{ background: #F8F9FB; padding: 10px 12px; font-weight: 600; color: #555;
     border-bottom: 2px solid #EAECF0; white-space: nowrap; text-align: right; }}
td {{ padding: 9px 12px; border-bottom: 1px solid #F2F4F7; vertical-align: middle; }}
tr:last-child td {{ border-bottom: none; }}
tbody tr:hover td {{ background: #FAFBFF; }}

/* Toolbar */
.toolbar {{ display: flex; gap: 8px; align-items: center; margin-bottom: 14px; flex-wrap: wrap; }}
.fbtn {{ padding: 6px 14px; border-radius: 20px; border: 1px solid #ddd; background: #fff;
         cursor: pointer; font-size: .82rem; font-family: inherit; color: #555;
         transition: all .15s; }}
.fbtn:hover {{ border-color: #4361EE; color: #4361EE; }}
.fbtn.active {{ background: #4361EE; color: #fff; border-color: #4361EE; font-weight: 600; }}
.search-box {{ margin-right: auto; padding: 6px 14px; border: 1px solid #ddd;
               border-radius: 20px; font-size: .82rem; font-family: inherit;
               outline: none; min-width: 200px; color: #333; }}
.search-box:focus {{ border-color: #4361EE; box-shadow: 0 0 0 3px #4361EE18; }}
</style>
</head>
<body>
<div class="page">

<div class="header">
  <h1>דוח הוצאות — {month_he} {year}</h1>
  <p>מופק {datetime.today().strftime('%d/%m/%Y')} &bull; {len(txns)} עסקאות &bull; {len(cards)} כרטיסים</p>
</div>

<div class="summary-row">
{summary_html}
</div>

<div class="stats-bar">
  <div class="stat"><div class="stat-n">{len(txns)}</div><div class="stat-l">סה"כ עסקאות</div></div>
  <div class="stat"><div class="stat-n">{len(cat_order)}</div><div class="stat-l">קטגוריות</div></div>
  <div class="stat"><div class="stat-n">{total_crd:,.0f}&#x20AA;</div><div class="stat-l">זיכויים</div></div>
  <div class="stat"><div class="stat-n">{avg_chg:,.0f}&#x20AA;</div><div class="stat-l">ממוצע לעסקה</div></div>
  <div class="stat"><div class="stat-n">{len(charges)}</div><div class="stat-l">עסקאות חיוב</div></div>
</div>

<div class="charts-grid">
  <div class="chart-card">
    <h2>התפלגות לפי קטגוריה</h2>
    <div class="chart-wrap"><canvas id="pieChart"></canvas></div>
  </div>
  <div class="chart-card">
    <h2>סכום לפי קטגוריה (&#x20AA;)</h2>
    <div class="chart-wrap"><canvas id="barChart"></canvas></div>
  </div>
</div>

<div class="section">
  <h2>10 ההוצאות הגבוהות ביותר</h2>
  <table id="top10Table">
    <thead><tr>
      <th></th><th>תאריך</th><th>בית עסק</th><th>קטגוריה</th><th>סכום</th><th>כרטיס</th>
    </tr></thead>
    <tbody>{top10_html}</tbody>
  </table>
</div>

<div class="section">
  <h2>כל העסקאות</h2>
  <div class="toolbar">
    {filter_btns}
    <input id="searchBox" class="search-box" type="text" placeholder="&#x1F50D; חיפוש לפי שם עסק...">
  </div>
  <table id="allTable">
    <thead><tr>
      <th>תאריך</th><th>בית עסק</th><th>קטגוריה</th><th>סכום</th><th>כרטיס</th><th>סוג עסקה</th>
    </tr></thead>
    <tbody>{all_rows_html}</tbody>
  </table>
</div>

</div><!-- .page -->
<script>
{js_block}
</script>
</body>
</html>"""


# ── Main ──────────────────────────────────────────────────────────────────────

def main():
    cli_args = sys.argv[1:]
    files = discover_files(cli_args)

    if not files:
        print('No unprocessed .xlsx files found.')
        print('(Files matching MM_YYYY_NNNN.xlsx are skipped as already processed.)')
        return

    all_txns = []
    period_month, period_year = None, None

    for f in files:
        txns, new_path, f_month, f_year = process_file(f)
        all_txns.extend(txns)
        # Use billing period from the first file that has data
        if period_month is None and f_month:
            period_month, period_year = f_month, f_year

    if not all_txns:
        print('No transactions found in any file.')
        return

    # Fallback period
    if period_month is None:
        period_month, period_year = all_txns[0]['date'].month, all_txns[0]['date'].year

    html = build_html(all_txns, period_month, period_year)
    out_file = f'dashboard_{period_month:02d}_{period_year}.html'

    with open(out_file, 'w', encoding='utf-8') as fh:
        fh.write(html)

    print(f'\n✓ Dashboard saved: {out_file}')
    print(f'  Total transactions : {len(all_txns)}')
    print(f'  Charges            : {sum(1 for t in all_txns if t["charge"] > 0)}')
    print(f'  Credits            : {sum(1 for t in all_txns if t["charge"] < 0)}')


if __name__ == '__main__':
    main()
