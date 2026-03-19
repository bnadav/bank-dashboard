# Bank Statement Dashboard

Personal finance tool for processing Bank Hapoalim (בנק הפועלים) credit card Excel exports and generating a monthly HTML spending dashboard.

## What this project does

Each month, Bank Hapoalim provides one Excel file containing **both credit cards** (`excelNewBank.xlsx`, `excelNewBank(1).xlsx`, etc.). This project:

1. Renames each file to `MM_YYYY_CARD.xlsx` (e.g. `02_2026_6138.xlsx`)
2. Parses all transactions from both cards and unites them
3. Categorizes spending by keyword-matching merchant names
4. Generates a self-contained `dashboard_MM_YYYY.html` with charts, top 10, and a full transaction table

## Running

```bash
# Drop new Excel files in this folder, then:
python bank_dashboard.py

# Or pass files explicitly (works on already-renamed files too):
python bank_dashboard.py 02_2026_6138.xlsx 02_2026_8689.xlsx
```

**Dependency:** `pip install openpyxl`
**Requirement:** Internet access to load Chart.js from CDN when opening the HTML dashboard.

## File naming convention

- Raw exports from the bank: any `.xlsx` name (auto-discovered)
- After processing: `MM_YYYY.xlsx` — e.g. `03_2026.xlsx` (new format: both cards in one file)
- Legacy single-card files: `MM_YYYY_CARD.xlsx` — e.g. `02_2026_6138.xlsx` (old format)
- Files already matching `MM_YYYY.xlsx` or `MM_YYYY_NNNN.xlsx` are skipped during auto-discovery (but can be passed explicitly)
- Dashboard output: `dashboard_MM_YYYY.html`

## Excel file structure

Both raw files share the same layout (sheet `גיליון1`, Hebrew/RTL, 14 columns). Each file corresponds to one credit card on account `12-782-82779`. The two known cards are `6138` and `8689`.

The sheet has these sections (identified by the string in column A):

| Section anchor (col A) | Content |
|---|---|
| `פירוט עבור הכרטיסים בארץ` | Local (Israel) transaction detail |
| `פירוט עבור הכרטיסים בחו''ל` | Foreign transaction detail |
| `פירוט עבור הכרטיסים בחו''ל בדולר` | USD sub-section (usually empty) |
| `פירוט עבור הכרטיסים בחו''ל ביורו` | EUR sub-section (usually empty) |

After each section anchor: skip 2 rows (account-number header + column header), then read data rows until col A is `None`.

**Local transaction columns (0-indexed):**
`[0]` card digits · `[1]` billing date · `[2]` transaction date · `[3]` merchant · `[4]` original purchase total · `[5]` charge this cycle (ILS) · `[6]` reference · `[12]` transaction type

**Foreign transaction columns:**
`[0]` card digits · `[1]` billing date · `[2]` transaction date · `[3]` merchant · `[4]` charge this cycle (ILS) · `[5]` original purchase amount · `[6]` original currency · `[7]` reference · `[13]` transaction type

**Key parsing notes:**
- `charge` (not `purchase`) is always used for financial totals — for installments, `purchase` is the full original price while `charge` is only this month's slice
- Negative `charge` values are credits/refunds (displayed in green)
- Use the **billing date** (col `[1]`) to determine report month/year — NOT the transaction date (col `[2]`), which can be months earlier for installment payments
- The `FOREIGN_ANCHOR` must be matched **exactly** to avoid matching the `בדולר`/`ביורו` sub-sections

## Transaction categories

`categorize(merchant, amount)` assigns a category to each transaction. It receives both the merchant name and the charge amount, which allows amount-based rules.

**Special rule — Yellow:** PAZ / Yellow app transactions with `charge < 100 ILS` are categorized as `Yellow` (small fuel top-ups / parking). Transactions of 100 ILS or more fall through to `דלק ותחבורה` (large fuel fills). This check runs before the main keyword loop.

The remaining categories are defined in `CATEGORIES` at the top of `bank_dashboard.py` as an ordered list of `(name, keywords)` tuples. Matching is case-insensitive substring search against the merchant name; first match wins; `שונות` (other) is the catch-all.

To add or adjust a keyword category, edit the `CATEGORIES` list. Order matters — more specific entries should come before broader ones (e.g. `ביטוח לאומי` before `ביטוח`). To add a new amount-based rule (like Yellow), add it at the top of `categorize()` before the main loop.

## Code structure

```
bank_dashboard.py
├── Constants
│   ├── CATEGORIES      — ordered list of (category_name, keywords)
│   ├── PALETTE         — 12-color hex palette for charts and badges
│   └── MONTH_HE        — Hebrew month names
├── Parsing
│   ├── categorize()    — keyword match → category name
│   ├── find_section()  — locate section anchor row in rows list
│   ├── safe_float()    — None/empty-safe float conversion
│   ├── parse_section() — extract transactions from one section
│   └── process_file()  — load workbook, rename, parse both sections
├── Discovery
│   └── discover_files() — CLI args passthrough or glob with rename-filter
├── HTML generation
│   ├── cat_color()     — category → hex color from PALETTE
│   ├── badge()         — colored category pill HTML
│   ├── ils_cell()      — amount span with red/green coloring
│   └── build_html()    — assemble full dashboard HTML
└── main()
```

**HTML generation note:** `build_html` uses Python f-strings for the outer HTML; all CSS/JS literal `{` `}` must be doubled as `{{` `}}`. The Chart.js initialization block and all interactive JS (card filter, search, table sort) are built with plain string concatenation (not f-strings) to avoid this issue with nested JS object literals.

**Sortable tables:** Both the Top 10 table (`id="top10Table"`) and the full transactions table (`id="allTable"`) are made sortable via `makeTableSortable(tableId)`. Clicking a column header sorts ascending; clicking again sorts descending; a `⇅` / `▲` / `▼` indicator appears on the active column. Sort type is auto-detected from cell content: dates (`DD/MM/YYYY`) → chronological, amounts / card numbers → numeric, everything else → Hebrew-aware string. Sorting preserves the card-filter and search state (hidden rows remain hidden).

## Windows / Hebrew console

`sys.stdout` is re-wrapped with UTF-8 at module load to handle Hebrew characters in console output on Windows.
