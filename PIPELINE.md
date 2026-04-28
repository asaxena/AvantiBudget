# Avanti Fellows FY27 Org — Pipeline Architecture

## Overview

The org dashboard (`index.html`) is a fully local, file-based HTML app. It reads
employee data from `data.js` (a JS file generated from the **Payroll FY27**
sheet of a Google Sheet) and renders five tabs: Org Tree, Untagged, Exits &
Leave, Vacancies, and Centres.

There is **no server** and **no build step**. Open `index.html` directly in a
browser, or visit https://asaxena.github.io/AvantiBudget/.

---

## Source of Truth

| What | Where |
|------|-------|
| Master spreadsheet | Google Sheet **"Final Employee Master - April 2026"** |
| Sheet ID | `1jcXJAGlVRJlj0_GenWiU32m_aeX7TO9ytXt_4XR4kRA` |
| URL | https://docs.google.com/spreadsheets/d/1jcXJAGlVRJlj0_GenWiU32m_aeX7TO9ytXt_4XR4kRA/edit |
| Authoritative tab | **`Payroll FY27`** — every other tab is reference only |

**Only `Payroll FY27` is read by the pipeline.** Akshay edits this sheet
manually. Claude downloads it as xlsx, regenerates `data.js`, and commits +
pushes to https://github.com/asaxena/AvantiBudget.

---

## File Inventory

| File | Lives in | Purpose |
|------|----------|---------|
| Google Sheet "Final Employee Master - April 2026" | Drive | Single source of truth. Edit only the `Payroll FY27` tab. |
| `Final Employee Master - April 2026.xlsx` | Local working dir (parent of repo) | Throwaway xlsx export from Drive. Not committed. |
| `regenerate_data_js.py` | Local working dir | Reads the xlsx, writes `data.js`. Not committed. |
| `data.js` | Repo | Generated from `Payroll FY27`. Sets `window.__ORG_DATA__` — a flat JSON array of all staff rows. **The only file in the repo that gets regenerated.** |
| `index.html` | Repo | The dashboard. Reads `window.__ORG_DATA__`, computes derived state, renders all tabs. Edited by hand only when UI changes. |

---

## Step 1 — The Payroll FY27 Sheet

### Column Layout (exact header names map to JS field names)

| Col | Header in sheet | JS field | Notes |
|-----|-----------------|----------|-------|
| A | AF ID | `id` | `AF###` for staff, `TBH###` for vacancies |
| B | Name | `name` | |
| C | Staff Type | `type` | See type values below |
| D | Designation (New) | `desig` | |
| E | Status FY27 | `status` | Active / Exit / Long Leave / Maternity / TBH / Under Notice period / Unassigned / Vendor / Active - Role Change / Not needed |
| F | Reporting To (New) | `rm_nm` | Display name of direct manager |
| G | RM AF ID | `rm_id` | AF ID of direct manager (`BOARD` for CEO level, `UNMAPPED` for orphans) |
| H | Function | `fn` | |
| I | Subject(s) | `subj` | Comma-separated, e.g. `Physics,Chemistry` |
| J | Level(s) | `lvls` | e.g. `JEE Advanced`, `NEET` |
| K | APM | `apm` | Academic Program Manager name |
| L | APM AF ID | `apm_id` | |
| M | PM | `pm` | Program Manager name |
| N | PM AF ID | `pm_id` | |
| O | SPM | `spm` | Senior Program Manager name |
| P | SPM AF ID | `spm_id` | |
| Q | Reporting Manager (Budget) | — | Ignored — used only inside the sheet for budget reporting. |
| R | RM AFID (Budget) | — | Ignored. |
| S | Centre(s) / Cost Centre (Budget) | `ctr` | Full centre name |
| T | System(s) | `sys` | e.g. `JNV`, `EMRS`, `Meritorious` |
| U | State(s) | `state` | |
| V | Mar 2026 Payroll | `mar` | Found / Not Found / — |
| W | Apr 2026 Status | `apr` | Current / Exit date / New Joiner / Not on Apr payroll |
| X | Source | `source` | NT Master Only / Centre Master (Tagged) / Centre Master (Untagged) / Centre Master (TBH) / Centre Master (Vendor) / Teacher Planning (TBH) / April New Joiner |

> The export script looks up columns **by header name**, so column reordering
> in the sheet is fine. Renaming a header will break the script — search for
> the new name in `FIELD_MAP` before saving.

### Staff Type Values

| `type` value | Meaning |
|-------------|---------|
| `Non-Teaching Staff` | NT staff (managers, ops, etc.) |
| `Teaching Staff` | Active tagged teacher |
| `Teaching Staff (TBH)` | Open teacher vacancy |
| `Vendor` | Contract/vendor teacher |

### Row Classification Rules

These rules are used by both the export script and the JS dashboard:

| Classification | Condition |
|---------------|-----------|
| **Exit / Leave** | `status` is `Exit`, `Long Leave`, or `Maternity`; OR `apr` contains `Exit`, `Terminated`, or `maternity leave` |
| **Untagged** | `source` is `Centre Master (Untagged)` |
| **Vacancy** | `type` is `Teaching Staff (TBH)` |
| **Active org** | Everyone not in the above three groups |

### Banner / Section Rows

If anyone adds merged "section banner" rows to the sheet (e.g. `FROM NT MASTER (247 rows)`),
the export script skips them: rows where both `AF ID` and `Name` are blank, or
where `Name` doesn't start with a letter.

---

## Step 2 — Generating `data.js`

The script lives at `../regenerate_data_js.py` (parent of the repo). Akshay
edits the Google Sheet, asks Claude to regenerate, and Claude:

1. Downloads the sheet as xlsx via the Drive MCP.
2. Saves it to `../Final Employee Master - April 2026.xlsx`.
3. Runs `python3 regenerate_data_js.py`.
4. Verifies the diff against the previous `data.js` (sanity counts, no IDs disappearing unexpectedly).
5. `git commit` + `git push` from inside the repo.

### The script

```python
import json, sys
from pathlib import Path
import openpyxl

ROOT = Path(__file__).parent
XLSX = ROOT / "Final Employee Master - April 2026.xlsx"
SHEET = "Payroll FY27"
OUT = ROOT / "AvantiBudget" / "data.js"

FIELD_MAP = {
    "id":     "AF ID",
    "name":   "Name",
    "type":   "Staff Type",
    "desig":  "Designation (New)",
    "fn":     "Function",
    "subj":   "Subject(s)",
    "lvls":   "Level(s)",
    "ctr":    "Centre(s) / Cost Centre (Budget)",
    "sys":    "System(s)",
    "state":  "State(s)",
    "status": "Status FY27",
    "rm_nm":  "Reporting To (New)",
    "rm_id":  "RM AF ID",
    "apm":    "APM",
    "apm_id": "APM AF ID",
    "pm":     "PM",
    "pm_id":  "PM AF ID",
    "spm":    "SPM",
    "spm_id": "SPM AF ID",
    "mar":    "Mar 2026 Payroll",
    "apr":    "Apr 2026 Status",
    "source": "Source",
}

wb = openpyxl.load_workbook(XLSX, data_only=True)
ws = wb[SHEET]
headers = [str(c.value or '').strip() for c in ws[1]]
col = {h: i for i, h in enumerate(headers)}

missing = [v for v in FIELD_MAP.values() if v not in col]
if missing:
    sys.exit(f"Missing expected columns in '{SHEET}': {missing}")

js_to_col = {k: col[v] for k, v in FIELD_MAP.items()}
afid_ci, name_ci = col["AF ID"], col["Name"]

RAW = []
for row in ws.iter_rows(min_row=2, values_only=True):
    afid = str(row[afid_ci] or '').strip()
    name = str(row[name_ci] or '').strip()
    if not afid and not name:
        continue
    if not afid and name and not name[0].isalpha():
        continue
    rec = {}
    for js_key, ci in js_to_col.items():
        v = row[ci]
        rec[js_key] = str(v).strip() if v is not None else ''
    for k in rec:
        if rec[k].lower() in ('none', 'nan'):
            rec[k] = ''
    RAW.append(rec)

with OUT.open("w", encoding="utf-8") as f:
    f.write("/* Auto-generated from 'Payroll FY27' sheet of Final Employee Master - April 2026.xlsx — do not edit manually */\n")
    f.write("window.__ORG_DATA__ = ")
    f.write(json.dumps(RAW, ensure_ascii=False, separators=(',', ':')))
    f.write(";")

print(f"Wrote {len(RAW)} rows to {OUT}")
```

> **Why `data.js` and not `data.json`?**
> Browsers block `fetch()` on `file://` URLs (same-origin policy). A
> `<script src="data.js">` tag executes synchronously and sets a global —
> no network request needed. Using `fetch('data.json')` will fail with
> "Failed to fetch" when opened from the filesystem.

### TBH ID Sequence

TBH IDs are assigned sequentially. Check the sheet for the highest existing
`TBH###` value before adding a new one.

---

## Step 3 — How `index.html` Consumes the Data

### Loading Sequence

```html
<head>
  <script src="data.js"></script>  <!-- sets window.__ORG_DATA__ before main script runs -->
  ...
</head>
<script>
  let RAW = [];

  // ... all function definitions ...

  /* ── init ── */
  RAW = window.__ORG_DATA__ || [];
  recompute();
  buildOrg();
  buildUntagged();
  buildExits();
  buildVacancies();
  buildCentres();
</script>
```

### `recompute()` — Derived State

All derived arrays are **`let`** (not `const`) so they can be reassigned after
`RAW` is populated. `recompute()` must be called after any change to `RAW`.

```javascript
let exits=[], untagged=[], vacancies=[], org=[], byId={}, children={};

function recompute(){
  exits     = RAW.filter(p => isExit(p) || isLeave(p));
  untagged  = RAW.filter(p => isUntagged(p) && !isExit(p));
  vacancies = RAW.filter(p => isVacancy(p));
  org       = RAW.filter(p => !isExit(p) && !isLeave(p) && !isUntagged(p));

  byId = {};
  org.forEach(p => { if (p.id) byId[p.id] = p; });

  children = {};
  org.forEach(p => {
    const rid = p.rm_id;
    if (rid && rid !== 'BOARD' && rid !== 'UNMAPPED' && rid !== '' && byId[rid]) {
      if (!children[rid]) children[rid] = [];
      children[rid].push(p);
    }
  });
}
```

### Classification Functions

```javascript
function isExit(p) {
  const a = p.apr.toLowerCase(), s = p.status.toLowerCase();
  return s==='exit' || a.includes('exit') || a.includes('terminated') ||
         s==='long leave' || a==='long leave' || s==='maternity' || a==='maternity leave';
}
function isLeave(p) {
  const s = p.status.toLowerCase(), a = p.apr.toLowerCase();
  return (s==='long leave' || s==='maternity' || a==='long leave') && !isExit(p);
}
function isUntagged(p) { return p.source === 'Centre Master (Untagged)'; }
function isVacancy(p)  { return p.type === 'Teaching Staff (TBH)'; }
function isTeacher(p)  {
  return p.type==='Teaching Staff' || p.type==='Teaching Staff (TBH)' || p.type==='Vendor';
}
function isTBH(p)      { return p.id.startsWith('TBH'); }
```

### Tree Building Logic (`buildOrg`)

The hierarchy is driven entirely by the `rm_id` field on each row:

- `rm_id === 'BOARD'` (or `'AF BOARD'`) → root node (renders at top level, used for CEOs)
- `rm_id` present in `byId` → person is a child of that manager
- `rm_id` blank / `UNMAPPED` / points to someone not in `byId` (e.g. an exited manager, or `'EXIT'`) → orphaned (rendered in a collapsible "No RM assigned" footer section)

Non-teaching staff and teaching staff share the same flat `RAW` array and the
same tree. Teachers are grouped visually under their APM into a `centres-wrap`
container rather than displayed as individual person cards.

### Tab Summary

| Tab | Data source | Render function |
|-----|------------|-----------------|
| Org Tree | `org` array | `buildOrg()` |
| Untagged | `untagged` array | `buildUntagged()` |
| Exits & Leave | `exits` array | `buildExits()` |
| Vacancies | `vacancies` array (TBH slots only) | `buildVacancies()` |
| Centres | `RAW` filtered live to teaching staff | `buildCentres()` |

### Search Logic (Org Tree tab)

The search box targets two distinct node types:

1. **Person cards** (`.card[data-name]`) — NT staff nodes. The card element stores `data-name`, `data-id`, `data-ctr`, `data-desig` attributes. Matching cards get `class="hi"`, non-matching get `class="dim"`.
2. **Centre cards** (`.ctr-card[data-search]`) — teacher groupings. The `data-search` attribute contains a space-joined string of centre name + all teacher IDs + all teacher names + all subjects. Non-matching centre cards are set to `display:none`. If all centre cards under a `.centres-wrap` are hidden, the wrap itself is hidden.

---

## Updating the Dashboard

To update after editing the Google Sheet:

1. Edit the `Payroll FY27` tab in **Final Employee Master - April 2026**.
2. Tell Claude "regenerate data.js" (or similar). Claude will:
   - Download the latest sheet as xlsx.
   - Run `regenerate_data_js.py` to rewrite `data.js`.
   - Diff against the previous version (row count, type/status/source bucket counts, ID set).
   - Commit and push to `main`.
3. GitHub Pages redeploys https://asaxena.github.io/AvantiBudget/ within ~1 min.

**To add a new TBH slot:** add a row with `Staff Type = Teaching Staff (TBH)`, assign the next TBH ID, set `RM AF ID` to the APM's AF ID.

**To convert a TBH to a real hire:** update `AF ID` → real AF ID, `Name` → real name, `Staff Type` → `Teaching Staff`, update `Status FY27` / `Mar 2026 Payroll` / `Apr 2026 Status`.

---

## Common Pitfalls

| Pitfall | Fix |
|---------|-----|
| `"Error loading data.json: Failed to fetch"` | Never use `fetch()` on file://. Use `<script src="data.js">` + `window.__ORG_DATA__` instead. |
| Org tree is blank / derived arrays are empty | `exits/untagged/vacancies/org/byId/children` must be `let` (not `const`), and `recompute()` must be called after `RAW` is assigned — not at parse time when `RAW` is still `[]`. |
| Person not appearing under their manager | Check `RM AF ID` in the sheet exactly matches the manager's `AF ID` (case-sensitive, no extra spaces). `EXIT` as an `RM AF ID` value will route the person to the orphan section. |
| Centre cards always visible during search | The `ctrBlock()` function must add a `data-search` attribute to each `.ctr-card`. `doSearch()` must target `.ctr-card[data-search]`, not just `.card`. |
| Banner rows polluting RAW | Export script skips rows where both `AF ID` and `Name` are blank, or where `Name` doesn't start with a letter (merged section headers). |
| `None` / `nan` strings in data | openpyxl returns Python `None` for blank cells; `str(None)` = `"None"`. Normalised in the script with `if rec[k].lower() in ('none','nan'): rec[k] = ''`. |
| Renamed a column in the sheet and `data.js` regen fails | The script exits with `Missing expected columns in 'Payroll FY27': [...]`. Update `FIELD_MAP` in `regenerate_data_js.py` to match the new header. |
