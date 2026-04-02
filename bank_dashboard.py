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


def _find_col(header_row, name):
    """Return the 0-based index of `name` in `header_row`, or -1."""
    for i, cell in enumerate(header_row):
        if cell and str(cell).strip() == name:
            return i
    return -1


def parse_section(rows: list, start: int, card: str, section: str) -> list:
    """
    Parse transactions starting at `start` (the anchor row).
    Skips +3 rows (anchor + account-header + column-header) then reads
    data rows until col A is None.

    Column positions are detected from the header row to handle varying layouts.
    """
    txns = []
    header_row = rows[start + 2]
    data_start = start + 3

    # Detect column positions from header
    charge_col = _find_col(header_row, "סכום חיוב בש''ח")
    purchase_col = _find_col(header_row, 'סכום קנייה')
    type_col = _find_col(header_row, 'תאור סוג עסקת אשראי')
    currency_col = _find_col(header_row, 'מטבע מקורי') if section == 'foreign' else -1

    # Fallback to hardcoded positions if header detection fails
    if charge_col == -1 or purchase_col == -1:
        if section == 'local':
            charge_col, purchase_col = 5, 4
        else:
            charge_col, purchase_col = 4, 5
    if type_col == -1:
        type_col = 12 if section == 'local' else 13
    if section == 'foreign' and currency_col == -1:
        currency_col = 6

    for row in rows[data_start:]:
        if row[0] is None:
            break
        card_cell = str(row[0]).strip()
        if not (card_cell.isdigit() and len(card_cell) == 4):
            continue   # stray non-data row
        try:
            txn_dt  = row[2]
            merchant = str(row[3] or '').strip()
            charge = safe_float(row[charge_col])
            purchase = safe_float(row[purchase_col])
            if section == 'local':
                currency = 'ILS'
            else:
                currency = str(row[currency_col] or 'USD').strip()
            txn_type = str(row[type_col] or '').strip()
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

    # Rename: only rename files that don't already match MM_YYYY.xlsx / MM_YYYY_CARD.xlsx
    new_path = path
    basename = os.path.basename(path)
    already_named = RENAMED_RE.match(basename)
    if already_named:
        # Trust the filename for month/year — it was set by a previous run or the user
        month = int(basename[:2])
        year  = int(basename[3:7])
    if month and year and not already_named:
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
        return cli_args
    # Include all xlsx files in data/ — new exports and previously renamed
    return _glob.glob('data/*.xlsx')


# ── HTML generation ───────────────────────────────────────────────────────────

# Per-month JS template — %%SUFFIX%%, %%LABELS%%, %%VALUES%%, %%COLORS%% are replaced at runtime.
# Written as a plain string (no f-string) so JS braces are literal.
_MONTH_JS = """\
var _init_%%SUFFIX%% = false;
function initMonth_%%SUFFIX%%() {
  if (_init_%%SUFFIX%%) return;
  _init_%%SUFFIX%% = true;
  var L = %%LABELS%%, V = %%VALUES%%, C = %%COLORS%%;
  new Chart(document.getElementById('pieChart_%%SUFFIX%%'), {
    type: 'doughnut',
    data: { labels: L, datasets: [{ data: V, backgroundColor: C, borderWidth: 2, borderColor: '#fff' }] },
    options: {
      responsive: true, maintainAspectRatio: false,
      plugins: {
        legend: { position: window.matchMedia('(max-width:699px)').matches ? 'bottom' : 'right', rtl: true, labels: { boxWidth: 12, padding: 10, font: { size: 11 } } },
        tooltip: { callbacks: { label: ctx => ' ' + ctx.label + ': ' + ctx.parsed.toLocaleString('he-IL', {minimumFractionDigits:2}) + ' \u20aa' } }
      }
    }
  });
  new Chart(document.getElementById('barChart_%%SUFFIX%%'), {
    type: 'bar',
    data: { labels: L, datasets: [{ data: V, backgroundColor: C, borderRadius: 4, borderSkipped: false }] },
    options: {
      indexAxis: 'y', responsive: true, maintainAspectRatio: false,
      onClick: function(evt, elements) {
        var panel = document.getElementById('catDetail_%%SUFFIX%%');
        if (!elements.length) { panel.style.display='none'; panel.dataset.cat=''; return; }
        var idx = elements[0].index;
        var cat = L[idx], val = V[idx];
        if (panel.dataset.cat === cat && panel.style.display !== 'none') {
          panel.style.display='none'; panel.dataset.cat=''; return;
        }
        panel.dataset.cat = cat;
        document.getElementById('catDetailTitle_%%SUFFIX%%').textContent =
          cat + ' \u2014 ' + val.toLocaleString('he-IL', {minimumFractionDigits:2}) + ' \u20aa';
        var tbody = document.querySelector('#catDetailTable_%%SUFFIX%% tbody');
        tbody.innerHTML = '';
        document.querySelectorAll('#allTable_%%SUFFIX%% tbody tr').forEach(function(r) {
          if (r.dataset.category === cat) tbody.appendChild(r.cloneNode(true));
        });
        panel.style.display = 'block';
        panel.scrollIntoView({behavior:'smooth', block:'nearest'});
      },
      plugins: {
        legend: { display: false },
        tooltip: { callbacks: { label: ctx => ' ' + ctx.parsed.x.toLocaleString('he-IL', {minimumFractionDigits:2}) + ' \u20aa' } }
      },
      scales: {
        x: { ticks: { callback: v => v.toLocaleString('he-IL') + '\u20aa' }, grid: { color: '#f0f0f0' } },
        y: { ticks: { font: { size: 11 } }, grid: { display: false } }
      }
    }
  });
  document.querySelectorAll('#panel_%%SUFFIX%% .fbtn').forEach(function(btn) {
    btn.addEventListener('click', function() {
      document.querySelectorAll('#panel_%%SUFFIX%% .fbtn').forEach(function(b) { b.classList.remove('active'); });
      btn.classList.add('active');
      applyFilters_%%SUFFIX%%();
    });
  });
  function applyFilters_%%SUFFIX%%() {
    var card = document.querySelector('#panel_%%SUFFIX%% .fbtn.active').dataset.card;
    var q = document.getElementById('searchBox_%%SUFFIX%%').value.toLowerCase();
    document.querySelectorAll('#allTable_%%SUFFIX%% tbody tr').forEach(function(row) {
      var ok = (card === 'all' || row.dataset.card === card) &&
               (!q || row.cells[1].textContent.toLowerCase().includes(q));
      row.style.display = ok ? '' : 'none';
    });
  }
  document.getElementById('searchBox_%%SUFFIX%%').addEventListener('input', applyFilters_%%SUFFIX%%);
  makeTableSortable('top20Table_%%SUFFIX%%');
  makeTableSortable('allTable_%%SUFFIX%%');
  makeTableSortable('catDetailTable_%%SUFFIX%%');
}
"""

_SHARED_JS = """\
Chart.defaults.font.family = "'Segoe UI', Arial, sans-serif";
Chart.defaults.font.size = 12;

function parseCellValue(text) {
  if (/^#[0-9]+$/.test(text)) return parseInt(text.slice(1));
  if (/^[0-9]{2}[/][0-9]{2}[/][0-9]{4}$/.test(text)) {
    var p = text.split('/'); return parseInt(p[2] + p[1] + p[0]);
  }
  var n = parseFloat(text.replace(/[\u20aa,+ \\t\u00b7\u200f\u200e]/g, ''));
  return isNaN(n) ? text : n;
}

function makeTableSortable(tableId) {
  var table = document.getElementById(tableId);
  if (!table) return;
  var ths = Array.from(table.querySelectorAll('thead th'));
  var curCol = -1, curAsc = true;
  ths.forEach(function(th, idx) {
    var arrow = document.createElement('span');
    arrow.style.cssText = 'margin-right:6px;font-size:.65rem;opacity:.3;vertical-align:middle;';
    arrow.textContent = '\u21c5';
    th.appendChild(arrow);
    th.style.cursor = 'pointer'; th.style.userSelect = 'none'; th.title = '\u05dc\u05d7\u05e5 \u05dc\u05de\u05d9\u05d5\u05df';
    th.addEventListener('click', function() {
      if (curCol === idx) curAsc = !curAsc; else { curCol = idx; curAsc = true; }
      ths.forEach(function(h, i) {
        var a = h.querySelector('span'); if (!a) return;
        if (i === idx) { a.textContent = curAsc ? ' \u25b2' : ' \u25bc'; a.style.opacity='1'; a.style.color='#4361EE'; }
        else           { a.textContent = '\u21c5'; a.style.opacity='.3'; a.style.color=''; }
      });
      var tbody = table.querySelector('tbody');
      var rows = Array.from(tbody.querySelectorAll('tr:not(.subtotal-row)'));
      rows.sort(function(a, b) {
        var av = parseCellValue(a.cells[idx] ? a.cells[idx].textContent.trim() : '');
        var bv = parseCellValue(b.cells[idx] ? b.cells[idx].textContent.trim() : '');
        if (typeof av === 'number' && typeof bv === 'number') return curAsc ? av-bv : bv-av;
        return curAsc ? String(av).localeCompare(String(bv),'he') : String(bv).localeCompare(String(av),'he');
      });
      var subtotals = Array.from(tbody.querySelectorAll('tr.subtotal-row'));
      rows.forEach(function(r) { tbody.appendChild(r); });
      subtotals.forEach(function(r) { tbody.appendChild(r); });
    });
  });
}

document.querySelectorAll('.tab').forEach(function(btn) {
  btn.addEventListener('click', function() {
    document.querySelectorAll('.tab').forEach(function(b) { b.classList.remove('active'); });
    btn.classList.add('active');
    document.querySelectorAll('.month-panel').forEach(function(p) { p.classList.remove('active'); });
    var s = btn.dataset.suffix;
    document.getElementById('panel_' + s).classList.add('active');
    window['initMonth_' + s]();
  });
});
"""


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


def mtotal_cell(amount: float) -> str:
    """Muted merchant-total cell — visually distinct from per-transaction ils_cell."""
    if amount < 0:
        return f'<span style="color:#16a34a;font-size:.82rem">{abs(amount):,.2f}&#x20AA;</span>'
    return f'<span style="color:#64748b;font-size:.82rem">{amount:,.2f}&#x20AA;</span>'


def fdate(d: datetime) -> str:
    return d.strftime('%d/%m/%Y')


def build_month_panel(txns: list, month: int, year: int, suffix: str) -> tuple:
    """Returns (inner_html: str, init_js: str) for one month tab panel."""
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

    top20      = sorted(charges, key=lambda x: -x['charge'])[:20]
    all_sorted = sorted(txns, key=lambda x: -x['date'].timestamp())

    merchant_totals = defaultdict(float)
    for t in txns:
        merchant_totals[t['merchant']] += t['charge']

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

    # ── Top 20 rows with subtotals at 10, 15, 20 ──
    top_html = ''
    subtotal_breaks = [10, 15, 20]
    running_sum = 0.0
    for i, t in enumerate(top20, 1):
        running_sum += t['charge']
        top_html += (
            f'<tr>'
            f'<td style="font-weight:700;color:#4361EE;width:36px">#{i}</td>'
            f'<td style="white-space:nowrap">{fdate(t["date"])}</td>'
            f'<td style="font-weight:500">{"🌍 " if t["foreign"] else ""}{t["merchant"]}</td>'
            f'<td>{badge(t["category"], cat_order)}</td>'
            f'<td style="text-align:left">{ils_cell(t["charge"])}</td>'
            f'<td style="text-align:left">{mtotal_cell(merchant_totals[t["merchant"]])}</td>'
            f'<td style="font-family:monospace;color:#888;font-size:.8rem">···· {t["card"]}</td>'
            f'</tr>'
        )
        if i in subtotal_breaks or i == len(top20):
            pct = running_sum / total_chg * 100 if total_chg else 0
            label = f'סה"כ טופ {i}'
            top_html += (
                f'<tr class="subtotal-row">'
                f'<td colspan="4" style="text-align:right;font-weight:700">{label}</td>'
                f'<td style="text-align:left;font-weight:700">{running_sum:,.2f}&#x20AA;</td>'
                f'<td colspan="2" style="color:#64748b;font-size:.85rem">{pct:.1f}% מסך ההוצאות</td>'
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
            f'<tr data-card="{t["card"]}" data-category="{t["category"]}">'
            f'<td style="white-space:nowrap">{fdate(t["date"])}</td>'
            f'<td style="font-weight:500">{"🌍 " if t["foreign"] else ""}{t["merchant"]}</td>'
            f'<td>{badge(t["category"], cat_order)}</td>'
            f'<td style="text-align:left">{ils_cell(t["charge"])}</td>'
            f'<td style="text-align:left">{mtotal_cell(merchant_totals[t["merchant"]])}</td>'
            f'<td style="font-family:monospace;color:#888;font-size:.8rem">···· {t["card"]}</td>'
            f'</tr>'
        )

    # Build chart init JS via template substitution (no f-string → braces are literal)
    chart_labels = json.dumps([n for n, _ in cat_order], ensure_ascii=False)
    chart_values = json.dumps([round(v, 2) for _, v in cat_order])
    chart_colors = json.dumps([PALETTE[i % len(PALETTE)] for i in range(len(cat_order))])
    init_js = (
        _MONTH_JS
        .replace('%%SUFFIX%%', suffix)
        .replace('%%LABELS%%', chart_labels)
        .replace('%%VALUES%%', chart_values)
        .replace('%%COLORS%%', chart_colors)
    )

    inner_html = f"""
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
    <div class="chart-wrap"><canvas id="pieChart_{suffix}"></canvas></div>
  </div>
  <div class="chart-card">
    <h2>סכום לפי קטגוריה (&#x20AA;) &nbsp;<span style="font-size:.75rem;font-weight:400;color:#aaa">לחץ על עמודה לפירוט</span></h2>
    <div class="chart-wrap"><canvas id="barChart_{suffix}"></canvas></div>
  </div>
</div>
<div id="catDetail_{suffix}" class="section" style="display:none;border-top:3px solid #4361EE">
  <div style="display:flex;justify-content:space-between;align-items:center;margin-bottom:16px">
    <h2 id="catDetailTitle_{suffix}" style="font-size:.95rem;font-weight:600;color:#555"></h2>
    <button onclick="var p=document.getElementById('catDetail_{suffix}');p.style.display='none';p.dataset.cat='';"
            style="border:none;background:none;font-size:1.2rem;cursor:pointer;color:#aaa;padding:4px 8px;line-height:1">✕</button>
  </div>
  <div class="table-wrap"><table id="catDetailTable_{suffix}">
    <thead><tr>
      <th>תאריך</th><th>בית עסק</th><th>קטגוריה</th><th>סכום</th><th>סה"כ בבית עסק</th><th>כרטיס</th>
    </tr></thead>
    <tbody></tbody>
  </table></div>
</div>
<div class="section">
  <h2>20 ההוצאות הגבוהות ביותר</h2>
  <div class="table-wrap"><table id="top20Table_{suffix}">
    <thead><tr>
      <th></th><th>תאריך</th><th>בית עסק</th><th>קטגוריה</th><th>סכום</th><th>סה"כ בבית עסק</th><th>כרטיס</th>
    </tr></thead>
    <tbody>{top_html}</tbody>
  </table></div>
</div>
<div class="section">
  <h2>כל העסקאות</h2>
  <div class="toolbar">
    {filter_btns}
    <input id="searchBox_{suffix}" class="search-box" type="text" placeholder="&#x1F50D; חיפוש לפי שם עסק...">
  </div>
  <div class="table-wrap"><table id="allTable_{suffix}">
    <thead><tr>
      <th>תאריך</th><th>בית עסק</th><th>קטגוריה</th><th>סכום</th><th>סה"כ בבית עסק</th><th>כרטיס</th>
    </tr></thead>
    <tbody>{all_rows_html}</tbody>
  </table></div>
</div>
"""
    return inner_html, init_js


def build_combined_html(months_data: list) -> str:
    """months_data: list of (txns, month, year) sorted newest-first."""
    tab_btns    = ''
    panels_html = ''
    all_init_js = ''
    first_suffix = None

    for i, (txns, month, year) in enumerate(months_data):
        suffix   = f'{month:02d}_{year}'
        month_he = MONTH_HE.get(month, str(month))
        if i == 0:
            first_suffix = suffix
        active_tab   = ' active' if i == 0 else ''
        active_panel = ' active' if i == 0 else ''
        tab_btns += f'<button class="tab{active_tab}" data-suffix="{suffix}">{month_he} {year}</button>\n    '
        inner, init_js = build_month_panel(txns, month, year, suffix)
        panels_html += f'<div class="month-panel{active_panel}" id="panel_{suffix}">{inner}</div>\n'
        all_init_js += init_js

    first_call = f'initMonth_{first_suffix}();\n' if first_suffix else ''

    return f"""<!DOCTYPE html>
<html lang="he" dir="rtl">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width,initial-scale=1">
<title>דוח הוצאות</title>
<script src="https://cdn.jsdelivr.net/npm/chart.js@4.4.0/dist/chart.umd.min.js"></script>
<style>
*, *::before, *::after {{ box-sizing: border-box; margin: 0; padding: 0; }}
body {{ font-family: 'Segoe UI', Arial, sans-serif; background: #F0F2F5; color: #1a1a2e; direction: rtl; }}
.page {{ max-width: 1200px; margin: 0 auto; padding: 24px 20px; }}
/* Tabs */
.tabs-bar {{ background: #fff; border-radius: 12px; padding: 10px 14px; margin-bottom: 20px;
             box-shadow: 0 2px 8px rgba(0,0,0,.07); display: flex; gap: 6px; flex-wrap: wrap; align-items: center; }}
.tabs-bar span {{ font-size: .8rem; color: #aaa; font-weight: 500; margin-left: 6px; }}
.tab {{ padding: 8px 20px; border-radius: 8px; border: none; background: transparent;
        cursor: pointer; font-size: .9rem; font-family: inherit; color: #555; font-weight: 500; transition: all .15s; }}
.tab:hover {{ background: #F0F2F5; color: #1a1a2e; }}
.tab.active {{ background: #4361EE; color: #fff; font-weight: 600; }}
/* Month panels */
.month-panel {{ display: none; }}
.month-panel.active {{ display: block; }}
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
.chart-card {{ background: #fff; border-radius: 12px; padding: 20px; box-shadow: 0 2px 8px rgba(0,0,0,.07); }}
.chart-card h2 {{ font-size: .9rem; font-weight: 600; color: #555; margin-bottom: 14px; }}
.chart-wrap {{ position: relative; height: 280px; }}
/* Sections */
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
.subtotal-row td {{ background: #f0f4ff !important; border-top: 2px solid #4361EE; border-bottom: 2px solid #4361EE; }}
/* Toolbar */
.toolbar {{ display: flex; gap: 8px; align-items: center; margin-bottom: 14px; flex-wrap: wrap; }}
.fbtn {{ padding: 6px 14px; border-radius: 20px; border: 1px solid #ddd; background: #fff;
         cursor: pointer; font-size: .82rem; font-family: inherit; color: #555; transition: all .15s; }}
.fbtn:hover {{ border-color: #4361EE; color: #4361EE; }}
.fbtn.active {{ background: #4361EE; color: #fff; border-color: #4361EE; font-weight: 600; }}
.search-box {{ margin-right: auto; padding: 6px 14px; border: 1px solid #ddd;
               border-radius: 20px; font-size: .82rem; font-family: inherit;
               outline: none; min-width: 200px; color: #333; }}
.search-box:focus {{ border-color: #4361EE; box-shadow: 0 0 0 3px #4361EE18; }}
/* Responsive table wrapper */
.table-wrap {{ overflow-x: auto; -webkit-overflow-scrolling: touch; }}
/* Mobile */
@media (max-width: 600px) {{
  .page {{ padding: 12px 10px; }}
  .header h1 {{ font-size: 1.3rem; }}
  .header p {{ font-size: .8rem; }}
  .scard {{ min-width: 120px; padding: 14px; }}
  .sc-amount {{ font-size: 1.2rem; }}
  .sc-sub {{ font-size: .7rem; }}
  .stat {{ min-width: 70px; padding: 12px 8px; }}
  .stat-n {{ font-size: 1.1rem; }}
  .chart-wrap {{ height: 240px; }}
  .chart-card {{ padding: 14px; }}
  .section {{ padding: 14px; }}
  th, td {{ padding: 7px 8px; font-size: .8rem; }}
  .search-box {{ min-width: 0; width: 100%; margin-right: 0; }}
  .toolbar {{ flex-direction: column; align-items: stretch; }}
  .toolbar .fbtn {{ text-align: center; }}
}}
</style>
</head>
<body>
<div class="page">
<div class="tabs-bar">
  <span>חודש:</span>
  {tab_btns}
</div>
{panels_html}
</div>
<script>
{all_init_js}
{_SHARED_JS}
{first_call}
</script>
</body>
</html>"""


# ── Main ──────────────────────────────────────────────────────────────────────

def main():
    cli_args = sys.argv[1:]
    files = discover_files(cli_args)

    if not files:
        print('No .xlsx files found.')
        return

    # Group transactions by (year, month)
    months: dict = defaultdict(list)
    for f in files:
        txns, _, f_month, f_year = process_file(f)
        if txns and f_month and f_year:
            months[(f_year, f_month)].extend(txns)

    if not months:
        print('No transactions found in any file.')
        return

    # Sort newest-first
    months_data = [
        (txns, month, year)
        for (year, month), txns in sorted(months.items(), reverse=True)
    ]

    html = build_combined_html(months_data)
    with open('dashboard.html', 'w', encoding='utf-8') as fh:
        fh.write(html)

    print(f'\n✓ Dashboard saved: dashboard.html  ({len(months_data)} month(s))')
    for txns, month, year in months_data:
        charges = sum(1 for t in txns if t['charge'] > 0)
        credits = sum(1 for t in txns if t['charge'] < 0)
        print(f'  {MONTH_HE.get(month, month)} {year}: {len(txns)} עסקאות ({charges} חיובים, {credits} זיכויים)')


if __name__ == '__main__':
    main()
