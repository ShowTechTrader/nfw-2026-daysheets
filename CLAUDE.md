# NFW 2026 — Production Tools · Claude Code Context

This file gives you full context on the project. Read it before touching anything.

---

## What this project is

Production day sheet system for **Noosa Food & Wine Festival 2026** (June 8–16, Festival Village, Noosa Main Beach). Run by Keegan Millar / ShowTechTrader.

The core loop is:
1. Keegan edits `NFW_Production_Schedule_2026_-_Master.xlsx` (schedule + notes)
2. Runs `python generate_daysheet.py`
3. Uploads the output `daysheet.html` to GitHub → live in ~60s

**Live URL:** `https://showtechtrader.github.io/nfw-2026-daysheets/daysheet.html`  
**GitHub repo:** `https://github.com/ShowTechTrader/nfw-2026-daysheets`

---

## File structure

```
files/                          ← Keegan's working directory (where you are)
├── generate_daysheet.py        ← The generator — main thing to work on
├── NFW_Production_Schedule_2026_-_Master.xlsx   ← Source of truth
├── NFWF_Monogram_Positive.png  ← Logo (embedded as base64 in HTML)
└── site-map.pdf                ← Embedded in site map accordion in HTML

GitHub repo (upload daysheet.html here after each run):
├── daysheet.html               ← Generated output — the live page
└── site-map.pdf                ← Static, only needs uploading once
```

---

## Running the generator

```bash
python generate_daysheet.py
# or explicitly:
python generate_daysheet.py --excel "NFW_Production_Schedule_2026_-_Master.xlsx" \
                            --logo NFWF_Monogram_Positive.png \
                            --output daysheet.html
```

Dependencies (install once):
```bash
pip install openpyxl requests Pillow numpy
```

---

## Excel workbook structure

### Day schedule sheets (9 sheets: "Monday 8th" through "Tuesday 16th")

Row 1: title, Row 3: headers, Row 4: day label, Rows 5+: 15-min schedule slots

| Col | Header | Notes |
|-----|--------|-------|
| A | Time | `datetime.time` objects, carry-forward |
| B | Dept | Event / Site / Technical / Vendor / Security / Theming / Crew / Deliveries / Cleaning |
| C | Activity | Free text |
| D | Responsibility | Free text |
| E | Contact | Auto-filled by XLOOKUP from col F |
| F | Company | **Dropdown** from Data Validation Sheet col D |
| G | Number | Auto-filled by XLOOKUP from col F |
| H | Notes | Free text |

Formulas in col E and G:
```
=IF(F5="","",XLOOKUP(F5,'Data Validation Sheet'!$D:$D,'Data Validation Sheet'!$C:$C,"",0))
=IF(F5="","",XLOOKUP(F5,'Data Validation Sheet'!$D:$D,'Data Validation Sheet'!$E:$E,"",0))
```

### Data Validation Sheet

| Col | Content |
|-----|---------|
| A | Departments (Site, Event, Technical, Vendors, Security, Theming, Crew, Deliveries, Cleaning) |
| C | Contact Name |
| D | Company ← **dropdown source** for day sheets col F |
| E | Contact Number |

Col F dropdown in day sheets references `'Data Validation Sheet'!$D$2:$D$50` (up to 48 companies).

**Important:** Excel/OneDrive strips data validations on save. If dropdowns disappear, re-add them with openpyxl — see the pattern used previously (DataValidation type='list', formula1=that range, showDropDown=False).

### Day Sheet Notes sheet (second-to-last tab)

This is where Keegan enters daily editorial content. The generator reads it.

**Standing data (shared across all days):**
- Row 10: Vendor WiFi — col B = username, col D = password
- Row 11: Internal WiFi — col B = username, col D = password
- Rows 15–38: Radio assignments — col A = number, col B = name (24 radios)
- Rows 15–19: Radio channels — col D = channel no., col E = use

**Per-day sections — layout constants (CRITICAL — must stay in sync):**

```python
FIRST_DAY_ROW  = 42    # Row where Day 1 section starts
SECTION_HEIGHT = 93    # Rows per day section

OFF = {
    'ticker_active':   3,  'ticker_text':      4,  'ticker_urgency':   5,
    'overview_hdr':    7,  'overview_count':  15,
    'changes_hdr':    24,  'changes_count':   15,
    'flags_hdr':      41,  'flags_count':     15,
    'to_sort_hdr':    58,  'to_sort_count':   15,
    'before_eod_hdr': 75,  'before_eod_count':15,
}
```

Day N section starts at row: `FIRST_DAY_ROW + (N-1) * SECTION_HEIGHT`  
Bullet rows for each section: rows `S + hdr_off + 1` through `S + hdr_off + count`  
Ticker active dropdown values: "Yes" / "No"  
Ticker urgency dropdown values: "Green" / "Orange" / "Red"

---

## Generator architecture (`generate_daysheet.py`)

Key functions in order:

| Function | Purpose |
|----------|---------|
| `process_logo(path)` | PIL: removes dark bg, recolours to navy `#082038`, returns base64 |
| `fetch_weather(lat, lon, date)` | Open-Meteo API — works on Keegan's Mac, blocked in Claude sandbox |
| `read_contacts_from_excel(wb)` | Reads DV sheet col C/D/E → list of `{company, name, number}` dicts |
| `read_notes_from_excel(wb)` | Reads Day Sheet Notes → dict keyed by day number |
| `read_schedule(ws)` | Reads a day sheet → list of event dicts |
| `render_schedule(events)` | Returns HTML table string |
| `render_weather(wx, iso, idx)` | Returns weather card HTML (metvuw map loaded by JS) |
| `bullets(items)` | Returns `<ul class="bullet-list">` HTML |
| `render_panel(meta, idx, events, notes, wx)` | Assembles one full day panel |
| `generate_html(days_data, logo_b64, generated_at, contacts)` | Wraps everything into final HTML |
| `main()` | Argparse entry point |

### HTML panel section order (per day)

1. Overview / Key Notes (+ site map accordion)
2. Changes
3. Flags
4. Schedule
5. Before End of Day
6. Weather
7. To Sort
8. WiFi
9. Radio Assignments

### CSS / JS approach

Everything is embedded inline — single self-contained HTML file, no external dependencies except Google Fonts. The CSS and JS are stored as Python string constants `CSS` and `JS` in the generator.

**Python f-string caution:** avoid nested double quotes inside `f"..."` — use intermediate variables. Python <3.12 doesn't support it. This has burned us before.

---

## Contacts modal

A "Contacts" button is fixed top-right. Clicking it slides in a panel from the right showing all companies from the Data Validation Sheet with name and phone number (clickable `tel:` links). Search bar filters live.

The contacts JS is a **separate** `<script>` block at the bottom of the HTML (not part of the main `JS` string constant). Data is embedded as `var CONTACTS_DATA = [...];` built at generation time from the Excel DV sheet.

---

## Brand

- Primary blue: `#D5F0FE`
- Dark navy: `#082038`
- Mid blue: `#1A9BC6`
- Fonts: Outfit (headings), DM Sans (body) via Google Fonts
- Dept badge colours: Event `#F97316` · Site `#16A34A` · Technical `#2563EB` · Vendor `#D97706` · Security `#DC2626` · Theming `#7C3AED` · Crew `#525252` · Deliveries `#0891B2` · Cleaning `#65A30D`

---

## Days configuration

```python
DAYS_META = [
    {"day":1,"sheet":"Monday 8th",    "date":"Monday 8th June",    "short":"Mon 8",  "type":"Bump-In",      "iso":"2026-06-08"},
    {"day":2,"sheet":"Tuesday 9th",   "date":"Tuesday 9th June",   "short":"Tue 9",  "type":"Bump-In",      "iso":"2026-06-09"},
    {"day":3,"sheet":"Wednesday 10th","date":"Wednesday 10th June", "short":"Wed 10", "type":"Bump-In",      "iso":"2026-06-10"},
    {"day":4,"sheet":"Thursday 11th", "date":"Thursday 11th June",  "short":"Thu 11", "type":"Bump-In",      "iso":"2026-06-11"},
    {"day":5,"sheet":"Friday 12th",   "date":"Friday 12th June",    "short":"Fri 12", "type":"Bump-In",      "iso":"2026-06-12"},
    {"day":6,"sheet":"Saturday 13th", "date":"Saturday 13th June",  "short":"Sat 13", "type":"Festival Day", "iso":"2026-06-13"},
    {"day":7,"sheet":"Sunday 14th",   "date":"Sunday 14th June",    "short":"Sun 14", "type":"Wrap Day",     "iso":"2026-06-14"},
    {"day":8,"sheet":"Monday 15th",   "date":"Monday 15th June",    "short":"Mon 15", "type":"Bump-Out",     "iso":"2026-06-15"},
    {"day":9,"sheet":"Tuesday 16th",  "date":"Tuesday 16th June",   "short":"Tue 16", "type":"Bump-Out",     "iso":"2026-06-16"},
]
```

---

## Known issues / things to watch

- **Data validation dropdowns (CRITICAL — openpyxl strips them on every load+save)**
  
  openpyxl silently strips Excel's *extended* (x14-namespace) data validations every time it loads and resaves the workbook — even if you never touch those sheets. The Dept (col B) and Company (col F) dropdowns on all 9 day sheets are stored as x14 validations and **will be deleted** by any `openpyxl.load_workbook(…).save(…)` round-trip.
  
  **The three dropdowns that must exist on every day sheet (B5:B500, F5:F500, I5:I500):**

  | Column | Range | formula1 | Source |
  |--------|-------|----------|--------|
  | B — Dept | `B5:B500` | `'Data Validation Sheet'!$A$2:$A$10` | Dept list |
  | F — Company | `F5:F500` | `'Data Validation Sheet'!$D$2:$D$50` | Company list |
  | I — Done tick | `I5:I500` | `"✓"` | Literal value |

  **Rule: after ANY openpyxl save of the master workbook, immediately re-add B and F validations using this pattern:**

  ```python
  from openpyxl.worksheet.datavalidation import DataValidation

  DAY_SHEETS = [
      'Monday 8th','Tuesday 9th','Wednesday 10th','Thursday 11th','Friday 12th',
      'Saturday 13th','Sunday 14th','Monday 15th','Tuesday 16th'
  ]

  for sheet_name in DAY_SHEETS:
      ws = wb[sheet_name]
      dv_dept = DataValidation(
          type='list',
          formula1="'Data Validation Sheet'!$A$2:$A$10",
          showDropDown=False,   # False = show the dropdown arrow
          allow_blank=True
      )
      dv_dept.sqref = 'B5:B500'
      ws.add_data_validation(dv_dept)

      dv_company = DataValidation(
          type='list',
          formula1="'Data Validation Sheet'!$D$2:$D$50",
          showDropDown=False,
          allow_blank=True
      )
      dv_company.sqref = 'F5:F500'
      ws.add_data_validation(dv_company)

  wb.save('NFW_Production_Schedule_2026_-_Master.xlsx')
  ```

  **Why `showDropDown=False`?** openpyxl uses Excel's internal XML attribute name where `False` = show the arrow, `True` = hide it. Counter-intuitive but correct.

  **Why this recurs every run:** When Excel opens a file containing openpyxl-written standard DVs, it converts them to its own x14 extended format on save. The next openpyxl load then strips that x14 block again, wiping B and F for all 9 sheets. This is a permanent cycle — the re-add block above must be included in every script that loads and saves the master workbook. It is NOT enough to add it once.

  The I-column tick validation survives because it is a literal-value DV (`"✓"`) which Excel leaves in the standard namespace. Always verify with `list(ws.data_validations.dataValidation)` after saving if in doubt.

- **Weather fetch** fails in any sandboxed environment — works fine on Keegan's Mac with internet.
- **Day Sheet Notes offsets** — if you ever change `SECTION_HEIGHT` or any offset, you MUST rebuild the Excel sheet too. They must stay in sync. Run a verification script after any layout change.
- The `generate_daysheet.py` output mentions `day_notes.json` in its footer print — that's a stale message from an earlier version. Ignore it; the generator reads from Excel only.

---

## What's next / ideas discussed but not built

- Key Contacts card (quick-reference section on day sheets for emergency numbers)
- Tomorrow's Preview section
- Crew Catering section
- End of Day Checklist
- Nearest Services (hospital, hardware store, etc.) — static, in Standing Data
