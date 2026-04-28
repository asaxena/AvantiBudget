# Avanti Fellows FY27 Org — Pipeline Architecture

## Overview

The org dashboard (`index.html`) is a fully local, file-based HTML app. It reads
employee data from `data.js` (a JS file generated from the **Employee DB - April 27 - Frozen**
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
| Authoritative tab | **`Employee DB - April 27 - Frozen`** — every other tab is reference only |

**Only `Employee DB - April 27 - Frozen` is read by the pipeline.** Akshay edits this
sheet manually. Claude downloads it as xlsx, regenerates `data.js`, and commits +
pushes to https://github.com/asaxena/AvantiBudget.

> **Sheet name has changed before** (was `Payroll FY27` in earlier versions).
> If the script errors with `Worksheet ... does not exist`, ask which tab is
> the new authoritative source and update the `SHEET` constant in
> `regenerate_data_js.py`.

---

## File Inventory

| File | Lives in | Purpose |
|------|----------|---------|
| Google Sheet "Final Employee Master - April 2026" | Drive | Single source of truth. Edit only the `Employee DB - April 27 - Frozen` tab. |
| `Final Employee Master - April 2026.xlsx` | Local working dir (parent of repo) | Throwaway xlsx export from Drive. Not committed. |
| `regenerate_data_js.py` | Local working dir | Reads the xlsx, writes `data.js`. Not committed. |
| `data.js` | Repo | Generated from the sheet. Sets `window.__ORG_DATA__` — a flat JSON array of all staff rows. **The only file in the repo that gets regenerated.** |
| `index.html` | Repo | The dashboard. Reads `window.__ORG_DATA__`, computes derived state, renders all tabs. Edited by hand only when UI changes. |

---

## Step 1 — The `Employee DB - April 27 - Frozen` Sheet

### Column Layout (exact header names map to JS field names)

| Col | Header in sheet | JS field | Notes |
|-----|-----------------|----------|-------|
| A | AF ID | `id` | `AF###` for staff, `TBH###` for vacancies, `BRL###` / `SRL###` for vendors |
| B | Name | `name` | |
| C | Staff Type | `type` | See type values below |
| D | Centre | — | **Ignored** — singular per-row centre (legacy/duplicate of column K). |
| E | Subject | — | **Ignored** — singular per-row subject (duplicate of M). |
| F | Level | — | **Ignored** — singular per-row level (duplicate of N). |
| G | Designation (New) | `desig` | |
| H | Status FY27 | `status` | Active / Active - Role Change / Exit / Long Leave / Maternity / TBH / Under Notice period / Unassigned / Vendor |
| I | Reporting To (New) | `rm_nm` | Display name of direct manager |
| J | RM AF ID | `rm_id` | AF ID of direct manager (`BOARD` for CEO level, `UNMAPPED` for orphans, `EXIT` if their manager has left) |
| K | Centre(s) / Cost Centre (Budget) | `ctr` | Full centre name |
| L | Function | `fn` | |
| M | Subject(s) | `subj` | Comma-separated, e.g. `Physics,Chemistry` |
| N | Level(s) | `lvls` | e.g. `JEE Advanced`, `NEET` |
| O | APM | `apm` | Academic Program Manager name |
| P | APM AF ID | `apm_id` | |
| Q | PM | `pm` | Program Manager name |
| R | PM AF ID | `pm_id` | |
| S | SPM | `spm` | Senior Program Manager name |
| T | SPM AF ID | `spm_id` | |
| U | Reporting Manager (Budget) | — | Ignored — used only inside the sheet for budget reporting. |
| V | RM AFID (Budget) | — | Ignored. |
| W | System(s) | `sys` | e.g. `JNV`, `EMRS`, `Meritorious` |
| X | State(s) | `state` | |
| Y | Mar 2026 Payroll | `mar` | Found / Not Found / — |
| Z | Apr 2026 Status | `apr` | Current / Exit date / New Joiner / Not on Apr payroll |
| AA | Source | `source` | NT Master Only / Centre Master (Tagged) / Centre Master (Untagged) / Centre Master (Vendor) / April New Joiner / blank |

> The export script looks up columns **by header name**, so column reordering
> in the sheet is fine. Renaming a header will break the script — update
> `FIELD_MAP` in `regenerate_data_js.py` to match the new header.

### Staff Type Values

| `type` value | Meaning |
|-------------|---------|
| `Non-Teaching Staff` | NT staff (managers, ops, etc.) |
| `Teaching Staff` | Teacher row — could be active or a vacancy depending on `Status FY27` |
| `Vendor` | Contract/vendor teacher (Brilliant, SRL, etc.) |

> **Teacher vacancies are no longer a separate type.** A row with
> `Staff Type = Teaching Staff` and `Status FY27 = TBH` is an open teacher slot
> (typically with an AF ID like `TBH123`). The dashboard's `isVacancy` reflects
> this — see Classification Functions below.

### Row Classification Rules

These rules are used by both the export script and the JS dashboard:

| Classification | Condition |
|---------------|-----------|
| **Exit / Leave** | `status` is `Exit`, `Long Leave`, or `Maternity`; OR `apr` contains `Exit`, `Terminated`, or `maternity leave` |
| **Untagged** | `source` is `Centre Master (Untagged)` |
| **Vacancy** | `status` is `TBH` AND `type` is `Teaching Staff` |
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
SHEET = "Employee DB - April 27 - Frozen"
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
    f.write("/* Auto-generated from 'Employee DB - April 27 - Frozen' sheet — do not edit manually */\n")
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
function isVacancy(p)  { return p.status === 'TBH' && p.type === 'Teaching Staff'; }
function isTeacher(p)  {
  return p.type==='Teaching Staff' || p.type==='Teaching Staff (TBH)' || p.type==='Vendor';
}
function isTBH(p)      { return p.id.startsWith('TBH'); }
```

> The `Teaching Staff (TBH)` clause inside `isTeacher` is a backwards-compat
> leftover — no current rows carry that type, but the OR keeps the dashboard
> working if the sheet schema ever flips back.

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
| Vacancies | `vacancies` array (Teaching Staff with `Status FY27 = TBH`) | `buildVacancies()` |
| Centres | `RAW` filtered live to teaching staff | `buildCentres()` |

### Search Logic (Org Tree tab)

The search box targets two distinct node types:

1. **Person cards** (`.card[data-name]`) — NT staff nodes. The card element stores `data-name`, `data-id`, `data-ctr`, `data-desig` attributes. Matching cards get `class="hi"`, non-matching get `class="dim"`.
2. **Centre cards** (`.ctr-card[data-search]`) — teacher groupings. The `data-search` attribute contains a space-joined string of centre name + all teacher IDs + all teacher names + all subjects. Non-matching centre cards are set to `display:none`. If all centre cards under a `.centres-wrap` are hidden, the wrap itself is hidden.

---

## Updating the Dashboard

To update after editing the Google Sheet:

1. Edit the `Employee DB - April 27 - Frozen` tab in **Final Employee Master - April 2026**.
2. Tell Claude "regenerate data.js" (or just "regenerate"). Claude will:
   - Download the latest sheet as xlsx.
   - Run `regenerate_data_js.py` to rewrite `data.js`.
   - Diff against the previous version (row count, type/status/source bucket counts, ID set).
   - Commit and push to `main`.
3. GitHub Pages redeploys https://asaxena.github.io/AvantiBudget/ within ~1 min.

**To add a new TBH slot:** add a row with `Staff Type = Teaching Staff`, `Status FY27 = TBH`, the next sequential `TBH###` AF ID, and `RM AF ID` set to the APM's AF ID.

**To convert a TBH to a real hire:** update `AF ID` → real AF ID, `Name` → real name (keep `Staff Type = Teaching Staff`), set `Status FY27 = Active`, update `Mar 2026 Payroll` / `Apr 2026 Status`.

---

## Common Pitfalls

| Pitfall | Fix |
|---------|-----|
| `Worksheet '<name>' does not exist` | Sheet was renamed. Update the `SHEET` constant in `regenerate_data_js.py` and PIPELINE.md. |
| Vacancies tab shows 0 after a regen | Schema change in the sheet — the dashboard's `isVacancy` may be out of sync with the new flag (currently `status === 'TBH' && type === 'Teaching Staff'`). |
| `"Error loading data.json: Failed to fetch"` | Never use `fetch()` on file://. Use `<script src="data.js">` + `window.__ORG_DATA__` instead. |
| Org tree is blank / derived arrays are empty | `exits/untagged/vacancies/org/byId/children` must be `let` (not `const`), and `recompute()` must be called after `RAW` is assigned — not at parse time when `RAW` is still `[]`. |
| Person not appearing under their manager | Check `RM AF ID` in the sheet exactly matches the manager's `AF ID` (case-sensitive, no extra spaces). `EXIT` as an `RM AF ID` value will route the person to the orphan section. |
| Centre cards always visible during search | The `ctrBlock()` function must add a `data-search` attribute to each `.ctr-card`. `doSearch()` must target `.ctr-card[data-search]`, not just `.card`. |
| Banner rows polluting RAW | Export script skips rows where both `AF ID` and `Name` are blank, or where `Name` doesn't start with a letter (merged section headers). |
| `None` / `nan` strings in data | openpyxl returns Python `None` for blank cells; `str(None)` = `"None"`. Normalised in the script with `if rec[k].lower() in ('none','nan'): rec[k] = ''`. |
| Renamed a column in the sheet and `data.js` regen fails | The script exits with `Missing expected columns in '<sheet>': [...]`. Update `FIELD_MAP` in `regenerate_data_js.py` to match the new header. |
