# Avanti FY27 Org Reconciliation — Approach & Session Protocol

**Current version:** v1.7
**Last updated:** April 27, 2026
**Workbook (source of truth):** `avanti_workbook.xlsx`
**Visual chart:** `index.html`
**This file:** `APPROACH.md`
**Repo:** https://github.com/asaxena/AvantiBudget
**Live site:** https://asaxena.github.io/AvantiBudget/

---

## ⚠️ Read this first if context has compressed

Before doing **anything** in a new session — even if a summary block at the top of the conversation tells you what's going on — follow the **Session Start** protocol below. The summary may be incomplete or stale; the GitHub repo is the source of truth.

---

## Division of labour

| Step | Who does it |
|---|---|
| Maintain the workbook structure and content | Claude |
| Maintain APPROACH.md and changelog | Claude |
| Regenerate the HTML chart (index.html) | Claude |
| **Commit updated files to the GitHub repo** | **Akshay** |

Claude's container resets between sessions but can pull from GitHub raw URLs (no auth, no rate limits) at session start. Akshay commits to the repo at session end. GitHub Pages auto-deploys index.html.

---

## Session Start — at the beginning of every session

Claude pulls the three files directly from GitHub raw:

```bash
mkdir -p /home/claude/build
curl -s -o /home/claude/build/avanti_workbook.xlsx https://raw.githubusercontent.com/asaxena/AvantiBudget/main/avanti_workbook.xlsx
curl -s -o /home/claude/index.html               https://raw.githubusercontent.com/asaxena/AvantiBudget/main/index.html
curl -s -o /home/claude/build/APPROACH.md         https://raw.githubusercontent.com/asaxena/AvantiBudget/main/APPROACH.md
```

Then:

1. **Read APPROACH.md in full.** Methodology and conventions live here.

2. **Confirm versions.** Open the workbook with openpyxl. Check the **Version** tab — newest entry is the current version. Cross-check that APPROACH.md's "Current version" header matches.

3. **Review the workbook in tab order:**
   - **README** — orientation
   - **Version** — changelog. Note current version and recent changes
   - **Decisions** — every RM override and structural change made so far, with rationale
   - **Rules** — interpretation logic
   - **Open Questions** — pending clarifications
   - **Data Issues** — known data-quality findings
   - **Employee Master (Final)** — current state of every person with original vs final RM/centre

4. **Skim source replicas** if the user's request touches teachers, centres, or budget:
   - `Src - Centre Master`, `Src - PM-SPM-APM`, `Src - Teacher Alloc`, `Src - Personnel Budget`

5. **Then start the user's task.** Don't guess at structure from memory.

The original Google Sheets (Centre Master, PM/SPM Mapping, Teacher Recruitment, Personnel Budget) are **read-only inputs**. Never modify them. All changes layer on top in the workbook. Source URLs are in the README tab.

---

## Session End — present files for the user to commit

After applying any change, regenerating downstream views, and updating APPROACH.md:

1. Save all three files to `/mnt/user-data/outputs/`:
   - `avanti_workbook.xlsx`
   - `index.html`
   - `APPROACH.md`

2. Call `present_files` so the user can download them.

3. The user commits to https://github.com/asaxena/AvantiBudget. GitHub Pages auto-deploys to https://asaxena.github.io/AvantiBudget/ within ~1 minute.

There is no automated push — committing is a deliberate user action so the user can review changes before they go live.

---

## Source files (Google Sheets — read-only inputs)

| Source | URL | What we use |
|---|---|---|
| Centre Master | docs.google.com/spreadsheets/d/1-SoEk4KWUgBSer_owdrnTNAR5tMIMZeKrq7931nfrzs | 53 centres with category, system, owner, course, plan |
| PM/SPM Mapping | docs.google.com/spreadsheets/d/1QAbyh5DYdi0Ud3ElWNirZFDhs87gl68SV91gOZhJqmc | 12 ops people (4 SPM + 2 PM + 6 APM) and cluster assignments |
| Teacher Recruitment | docs.google.com/spreadsheets/d/16dQptFruBBNpV4abeD8I6opPLRcYJ1ozxB7oPzVWYCo | "Latest 16 Apr" tab — teacher allocations with cell-colour-encoded levels |
| Personnel Budget | docs.google.com/spreadsheets/d/1uaT_U7sbUnb_FnR56AQ_HJU0qd0TMirY8CMtyF8nZT8 | "FY27 - Consolidated Personnel -" tab — 445 rows with AF, name, cc, cat, salary, RM |

These files are not modified by us. We only read. All changes layer on top of them in the workbook.

---

## Workbook structure (16 tabs)

| # | Tab | Purpose |
|---|---|---|
| 1 | README | Purpose, source links, snapshot date |
| 2 | Version | Changelog (latest is v1.1, 2026-04-27) |
| 3 | Decisions | Every RM override / structural change with rationale & source |
| 4 | Rules | Interpretation logic in plain English |
| 5 | Open Questions | Clarifications pending |
| 6 | Data Issues | Findings to fix in source sheets |
| 7 | Employee Master (Final) | Cleaned view: every person with original vs final RM/centre/role |
| 8 | **Teacher Master** (was Teachers Audit) | Real teaching FTE only — 186 records (in budget, hire-needed, centre-reassigned, needs-budget-row, subject-gap) |
| 9 | **Non-Teaching Master** (new in v1.1) | All active non-teaching + non-teaching vacancies — 370 records, with Budget Status flag |
| 10 | Reconciliation | Full join of Recruitment ↔ Personnel Budget — 262 records (superset that feeds Teacher Master + Budget Corrections) |
| 11 | Budget Corrections | Both teaching and non-teaching budget issues — 79 items (excess TBHs, no-alloc, vendors, NT needing budget rows) |
| 12 | Vacancy List | All open positions across teaching and non-teaching (95 — salary column blanked since v1.1) |
| 13 | Src - Centre Master | Source replica |
| 14 | Src - PM-SPM-APM | Source replica |
| 15 | Src - Teacher Alloc | Source replica with cell-colour-derived levels |
| 16 | Src - Personnel Budget | Source replica (salary column blanked since v1.1 — actual salaries only in source Google Sheet) |

When updating: edit the appropriate "decision" tabs first (Decisions / Rules / Open Questions / Data Issues), then propagate to Employee Master (Final), then regenerate Teacher Master / Non-Teaching Master / Reconciliation / Budget Corrections / Vacancy List, then regenerate index.html.

---

## HTML chart (`index.html`)

8 tabs visible in the rendered chart:
1. Akshay's Tree
2. Vandana's Tree
3. Centre Master (53 centres, frozen first column + sticky header, level badges next to teacher names)
4. **Teacher Master** — 186 real teaching records
5. **Non-Teaching Master** (new in v1.1) — 370 records (320 active/exit + 50 vacancies) with Budget Status filter
6. Budget Corrections — 79 items (76 teaching + 3 non-teaching), with Side filter (Teaching / Non-Teaching)
7. Vacancy List — 95 open positions, no salary column
8. Unassigned (clarifications, no-RM employees, orphaned, exits)

Mobile: tab bar replaced with dropdown menu below 768px.

---

## Salary handling (added v1.1)

The workbook is shared broadly via the GitHub repo. Salary information is therefore **blanked across the published files**:
- `Src - Personnel Budget` tab in the workbook: FY27 Salary column emptied
- `Vacancy List` tab in the workbook: Salary column emptied
- `index.html`: Salary column dropped from the Vacancy table

Actual salary data remains in the **source Personnel Budget Google Sheet**, which is restricted to the right people. If a salary query comes up in chat, look it up in the source sheet directly — never re-add salary columns to the published files.

---

## Methodology

### Build pipeline (in `/home/claude/build/`)

The reconciliation is built up in stages from the source xlsx files. Each stage produces a JSON intermediate that subsequent stages consume.

1. `step1_centre_master.py` → `centre_master.json` — 53 centres parsed with type/system/category, plus `needs_apc` flag
2. `step2_teacher_master.py` → `teacher_master.json` — teachers from "Latest 16 Apr", with levels read from cell fill colour
3. `step3_match_af_ids.py` → `teacher_master_v2.json` — AF IDs backfilled by matching names
4. `step4_employee_master.py` → `employee_master.json` — 445 personnel rows with all RM overrides applied
5. Reconciliation logic produces `teachers_reconciled_v2.json`
6. HR pipeline produces `hr_tbhs.json`
7. Output generation produces `index.html` and the workbook

### Key data conventions

- **AF IDs** are the join key everywhere. `AFxxx` for regular employees, `AFIxxx` for interns, `TBHxxx` for budget-only TBH slots.
- **Centre names** are normalised through aliases (see Rules tab). E.g. "JNV Odisha 2" = "JNV Puri".
- **Cell colours** in the recruitment sheet encode teacher level. Map below.
- **APC is centre-specific.** Required only at the 22 centres where the recruitment sheet's APC column has a name or empty cell. Centres with "NA" in APC don't need APC.

### Reconciliation statuses (8 categories)

| Status | Meaning | Lives in tab |
|---|---|---|
| `✓ In Budget` | Alloc cc matches budget cc | Teacher Master |
| `⊙ Hire needed (slot reserved)` | Subject gap with budget TBH already at centre | Teacher Master |
| `⇄ Centre Reassigned — UPDATE budget` | Named teacher's alloc cc differs from budget cc | Teacher Master |
| `⚠ Needs budget row — ADD` | Allocated FTE but no budget row | Teacher Master |
| `⚠ Subject gap — ADD slot + hire` | Subject gap with no budget slot | Teacher Master |
| `✕ Excess TBH — REMOVE` | Budget TBH at centre with no subject gap | Budget Corrections |
| `? Budgeted, no centre alloc` | Person budgeted but not allocated | Budget Corrections |
| `— Vendor (no budget needed)` | External contractor | Budget Corrections |

Non-teaching `⚠ Needs budget row — ADD` and `? Pending clarification` flags also feed into Budget Corrections.

### Role priority order

When a person appears in multiple sources:
1. **PM/SPM/APM mapping** = current role (highest priority)
2. **Teacher recruitment allocation** = current role
3. **Personnel budget designation** = fallback role

This prevents misclassifying ops people as teachers when their AF appears in multiple sheets.

---

## Cell colour → Teacher level map

Read directly from `Latest 16 Apr` tab cell fill colours:

| Hex | Level | Short |
|---|---|---|
| FF0000FF | JEE Mains Plus | JM+ |
| FF00FFFF | JEE Advanced | Adv |
| FF00FF00 | JEE Main | JM |
| FFFF00FF | NEET | NEET |
| FFFF9900 | MBBS | MBBS |
| FF980000 | CET | CET |
| FF274E13 | Foundation | Fnd |
| FFFFFF00 | Yellow = substitute/training-team teacher; inherits centre's primary level |

---

## Centre aliases

| Variant 1 | Variant 2 | Used in workbook |
|---|---|---|
| JNV Odisha 2 - CoE | JNV Puri - CoE | JNV Puri - CoE |
| Live Class - Punjab, Himachal, Uttarakhand, Delhi | Online / Boot Camp - Punjab, Himachal, Uttarakhand, Delhi | Online / Boot Camp variant |
| Foundation Mathematics | Mathematics | Mathematics (for Foundation centres) |

---

## Change Log

| Version | Date | Summary |
|---|---|---|
| v1.0 | 2026-04-26 | Initial workbook build. 60 RM/structural decisions applied. Cleaned Employee Master with original vs final RM/centre. APPROACH.md created. Drive folder set as source of truth. |
| v1.7 | 2026-04-27 | TBH137 fixed in HTML (now under TBH136/Abhishek, not Deepak); Nitin KR→Exit; AF370 Gurvir explicit Exit row added; Manvi added as Director Girls in STEM (placeholder ID 'MANVI', AF ID needed); TBH145 Director Girls in STEM under Panchali; Ananya→Manvi. nt_exits 4→6. |
| v1.6 | 2026-04-27 | NT: TBH116→Tech Interns under Surya Teja; Subhas M→unassigned; Shivangi Goyal→TBH142 (Medak APM); Rameez replaces Gurvir (AFI125/126 inherit); Rakesh Negi→Nazneen; Toshiba→Exit; TBH118 duplicate deleted. nt_no_rm 16→13, nt_exits 3→4, nt_orphaned 2→0. |
| v1.5 | 2026-04-27 | Teaching: JNV Chandigarh moved from Deepak (AF233) SPM cluster to SPM TBH 3 (TBH136) → PM TBH 1 (TBH137) under Abhishek (AF693). Triggered by updated Master Sheet — SPM column changed to '-'. NT Master TBH137 RM already updated (AF233 → TBH136) in prior step. |
| v1.4 | 2026-04-27 | NT: TBH118 Major Grants → Ruchika; Abhinav Gupta → Surya Teja; AFI119 Preetham Gowda marked Exit; Manzar → Lorina; Saksham Long Leave; TBH028 → Ananya; new TBH133 (JP Morgan) under Jotika; TBH026/027/116 city coordinators under TBH133. Teaching: rebuilt from updated Master Sheet (Name-AFID format); Aryaman now direct SPM-level under Siva; Abhishek group labelled SPM TBH. nt_no_rm 19→15, exits 2→3. |
| v1.3 | 2026-04-27 | NT org changes: AF785 Rahul Kumar → Tushanka (AF694); TBH004 Priyanka Palshetkar → Pritam (AF287); TBH119 APM Alumni SAP → Jotika (AF514); TBH117 Fundraising Specialist CSR → Ruchika (AF748); TBH118 Sr Manager Major Grants → Prafful (AF382); AF649 Namisha Singla → Shahied (AF387). Teaching tree rebuilt from Master Sheet Center Status 26-27: APM TBH names now specific (TBH 1–6), JNV Adilabad shown as separate CoE+Nodal entries, nt_no_rm reduced from 24 → 19. |
| v1.2 | 2026-04-27 | Corrected Abhishek Kumar RM: AF693 (Head Operation - North) → Agny; reverted AF838 to Panchali Dutta. Siva Dasari (AF219) moved to direct report of Agny (was under Subramanya). Changes applied to NT Master, Employee Master, Decisions tab, and HTML. |
| v1.1 | 2026-04-27 | Restructured tabs: split Reconciliation into Teacher Master (real teaching FTE, 186) and Budget Corrections (teaching + non-teaching budget issues, 79). Added Non-Teaching Master tab (370 records: 320 active/exit + 50 vacancies, with Budget Status flag). Hid salary columns across the workbook (Personnel Budget Src tab + Vacancy List + HTML Vacancy table). Migrated source-of-truth from Google Drive to GitHub repo asaxena/AvantiBudget. HTML renamed to index.html for GitHub Pages root. |
