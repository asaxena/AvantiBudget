# Avanti FY27 Org Reconciliation — Approach & Session Protocol

**Current version:** v1.0
**Last updated:** April 26, 2026
**Workbook (source of truth):** `avanti_workbook.xlsx` — kept in Google Drive by Akshay
**Visual chart:** `org_chart.html`
**This file:** `APPROACH.md`

---

## ⚠️ Read this first if context has compressed

Before doing **anything** in a new session — even if a summary block at the top of the conversation tells you what's going on — follow the **Session Start** protocol below. The summary may be incomplete or stale; the workbook in Google Drive is the source of truth.

---

## Division of labour

| Step | Who does it |
|---|---|
| Maintain the workbook structure and content | Claude |
| Maintain APPROACH.md and changelog | Claude |
| Regenerate the HTML chart | Claude |
| **Keep the canonical workbook in Google Drive** | **Akshay** |
| **Upload latest workbook + APPROACH.md to Drive at session end** | **Akshay** |
| **Attach the latest workbook + APPROACH.md at session start** | **Akshay** |

Claude's container resets between sessions, so it cannot be the keeper of the canonical files. Akshay holds the Drive copy; Claude is the editor.

---

## Session Start — at the beginning of every session

### What Akshay does (one-time per session)
Attach the latest `avanti_workbook.xlsx` and `APPROACH.md` from Drive to the chat. Without these, Claude is working blind.

### What Claude does
1. **Read APPROACH.md in full.** Methodology and conventions live here.

2. **Confirm versions.** Open the attached workbook with openpyxl. Check the **Version** tab — top row is the current version. Cross-check that APPROACH.md's "Current version" header matches.

3. **Review the workbook in tab order:**
   - **README** — orientation
   - **Version** — changelog (NEWEST AT TOP). Note current version and recent changes
   - **Decisions** — every RM override and structural change made so far, with rationale
   - **Rules** — interpretation logic
   - **Open Questions** — pending clarifications
   - **Data Issues** — known data-quality findings
   - **Employee Master (Final)** — current state of every person with original vs final RM/centre

4. **Skim source replicas** if the user's request touches teachers, centres, or budget:
   `Src - Centre Master`, `Src - PM-SPM-APM`, `Src - Teacher Alloc`, `Src - Personnel Budget`.

5. **Only then** start the user's task. Don't guess from memory or a conversation summary alone.

If APPROACH.md or the workbook isn't attached, **stop and ask Akshay to attach them** before doing any substantive work.

Build pipeline lives at `/home/claude/build/`. Source xlsx files (re-pull from Sheets if needed): `/home/claude/teacher_alloc.xlsx` and `/home/claude/personnel.xlsx`. HTML chart at `/home/claude/org_chart.html`.

---

## Session End — before the final response of every session

### What Claude does
After applying any change (RM override, role correction, new clarification, fixing a data issue, etc.):

1. **Update the relevant workbook tab.** The workbook is the source of truth — never edit only the HTML chart.
   - New RM override → add row to `Decisions`
   - New rule applied → add row to `Rules`
   - New clarification raised → add row to `Open Questions`
   - New data-quality finding → add row to `Data Issues`
   - Person's RM/centre/role changed → update their row in `Employee Master (Final)`
   - Recruitment or budget change → regenerate `Reconciliation` and `Vacancy List`

2. **Bump the version.** Add a new row to the top of the **Version** tab:
   - Increment version (v1.1, v1.2, … v2.0 for major changes)
   - Date in YYYY-MM-DD
   - Updated by (e.g. "Akshay + Claude")
   - Summary of changes (one or two sentences)
   - Tabs touched (comma-separated)

3. **Regenerate downstream views.** If the change affects org structure or teachers, also update `org_chart.html`.

4. **Update this APPROACH.md.** New rule? Add to Rules section. Methodology change? Update relevant section. Bump "Current version" and "Last updated" at the top. Add a Changelog entry at the bottom.

5. **Push all three to outputs** via `present_files` so Akshay has them inline:
   `org_chart.html`, `avanti_workbook.xlsx`, `APPROACH.md`.

6. **Tell Akshay clearly** at the end of the response:
   > "Updated to v1.X. Please upload `avanti_workbook.xlsx` and `APPROACH.md` to the Drive folder, replacing the existing files (Drive will retain version history)."

### What Akshay does
- Download the three files from the chat.
- Upload `avanti_workbook.xlsx` and `APPROACH.md` to the Drive folder, replacing the existing copies. Drive auto-versions, so prior copies remain accessible.
- (Optional) Save `org_chart.html` somewhere convenient for sharing.

---

## Source files (Google Sheets, owned by various teams)

| Source | ID | What we use |
|---|---|---|
| Centre Master | `1-SoEk4KWUgBSer_owdrnTNAR5tMIMZeKrq7931nfrzs` | 53 centres with category, system, owner, course, plan |
| PM/SPM Mapping | `1QAbyh5DYdi0Ud3ElWNirZFDhs87gl68SV91gOZhJqmc` | 12 ops people (4 SPM + 2 PM + 6 APM) and clusters |
| Teacher Recruitment | `16dQptFruBBNpV4abeD8I6opPLRcYJ1ozxB7oPzVWYCo` | "Latest 16 Apr" tab — allocations with cell-colour-encoded levels |
| Personnel Budget | `1uaT_U7sbUnb_FnR56AQ_HJU0qd0TMirY8CMtyF8nZT8` | "FY27 - Consolidated Personnel" tab — 445 rows |

---

## Methodology

### Why a workbook + chart, not just a chart

The HTML chart is one rendering. The workbook is the source of truth: it holds decisions, rules, open questions, data issues, and the cleaned employee master. Any view (HTML, PDF, dashboard) should be regenerable from the workbook. **If the workbook and chart disagree, the workbook wins.**

### How information layers

```
SOURCE FILES (Google Sheets, owned by various teams)
     │
     ▼
SOURCE REPLICAS (Src - * tabs in workbook — snapshot)
     │
     ▼
RULES + DECISIONS (Rules tab + Decisions tab — interpretation layer)
     │
     ▼
DERIVED OUTPUTS (Employee Master Final, Reconciliation, Vacancy List)
     │
     ▼
RENDERED VIEW (org_chart.html)
```

A change at any layer triggers regeneration of everything below it. Source-file changes require re-pulling.

### Workbook tabs

| # | Tab | Purpose | Updated when |
|---|---|---|---|
| 1 | README | Orientation, source links, version reference | Every snapshot |
| 2 | Version | Append-only changelog with date/version/summary | Every session that changes anything |
| 3 | Decisions | Every RM/structural change with rationale | Any RM override, role change, structural change |
| 4 | Rules | Interpretation logic | New rule emerges or existing one changes |
| 5 | Open Questions | Pending clarifications | New question raised or existing resolved |
| 6 | Data Issues | Findings to fix in source sheets | New data-quality issue surfaces |
| 7 | Employee Master (Final) | Cleaned state of every person | Any RM/centre/role/status change |
| 8 | Reconciliation | Teacher alloc ⇄ budget reconciliation | Recruitment or budget changes; APC rule changes |
| 9 | Vacancy List | All open positions | Reconciliation regenerates; new TBHs added |
| 10-13 | Src - * | Source replicas | Re-pulled when sources update |

---

## Rules (current)

The workbook's Rules tab should mirror this list.

1. **Role priority.** When a person appears in multiple sources: (1) PM/SPM/APM mapping, (2) Teacher recruitment alloc, (3) Personnel budget designation (fallback).

2. **Centre allocation = current role.** If allocated to a centre as a teacher, that's their role — they move out of any non-teaching team.

3. **Centre aliases.** `JNV Odisha 2 = JNV Puri`. `Live Class Punjab/HP/UK/Delhi = Online/Boot Camp Punjab/HP/UK/Delhi`. `Foundation Math = Mathematics`.

4. **APC requirement.** APC required ONLY where the recruitment sheet's APC column has a name OR is explicitly empty. Centres where APC = "NA" do NOT need APC. (22 of 53 centres need APC.)

5. **Cell colour → level mapping** (Latest 16 Apr tab):
   - `FF0000FF` JEE Mains Plus (blue)
   - `FF00FFFF` JEE Advanced (cyan)
   - `FF00FF00` JEE Main (green)
   - `FFFF00FF` NEET (magenta)
   - `FFFF9900` MBBS (orange)
   - `FF980000` CET (dark red)
   - `FF274E13` Foundation (dark green)
   - `FFFFFF00` (yellow) substitute / training-team — inherits centre's primary level

6. **Contracted vendors don't need budget rows.** SRL, Brilliant, Gravity, etc. — flagged as `vendor_no_budget_needed`.

7. **"Centre Reassigned" semantics.** When alloc cc differs from budget cc, the BUDGET should be updated (allocation reflects current truth).

8. **Same-name disambiguation.** Always use AF ID as unique key. AF869 Nikita Kshatriya ≠ AF847 Nikita Shrikesh Powar.

9. **Exits.** Removed from active counts; reports reassigned. Currently exited: Aman Jakar, Ajay Kumar Dasari, Gurvir Singh, Vinod Mehto, Pulak Maity, Ghazala Firdos, Atul Kumar Jha, Gaurav Kumar, Akash Singh, Bhagyashree Harasur, Md. Nisar Mollick, Suguna Sri Vagvala, Raj Kumar.

10. **Long Leave.** Saksham Srivastava on long leave; reports moved to Savin Khandelwal under Panchali Dutta.

11. **TBH classification.**
    - Auto-flagged TBH = subject gap at centre (derived)
    - Budget TBH = TBH001-TBH116 in personnel sheet (real budget slots)
    - Match by centre. Excess budget TBHs → "Remove". Excess auto-flagged → "Add slot + hire".

---

## Reconciliation status taxonomy

| Status | Meaning | Action |
|---|---|---|
| `in_budget` | Named teacher's alloc cc matches budget cc | None |
| `hire_needed_with_budget` | Subject gap with budget TBH already there | Hire |
| `centre_reassigned` | Named teacher in budget, allocated elsewhere | Update budget cc |
| `needs_budget_row` | Allocated teacher with no budget line | Add to budget |
| `hire_needed_no_budget` | Subject gap, no budget slot | Add slot + hire |
| `tbh_excess_remove` | Budget TBH at centre with no subject gap | Delete or reallocate |
| `no_alloc` | Named in budget but not at any centre | Investigate |
| `vendor_no_budget_needed` | Contractor | None |

`Reconciliation` tab keeps all 8. The chart's `Teachers Audit` tab shows the 6 real-teacher statuses; the other 2 go to `Budget Corrections`.

---

## Worked example: applying a change

Akshay says: "Move person X from manager A to manager B."

1. Find X's AF ID in `Employee Master (Final)`. Note current `Final RM`.
2. Add row to `Decisions`: AF ID, Person, Original RM (from `Src - Personnel Budget`), New RM, Type ("Reporting line"), Rationale (user instruction), Source.
3. Update X's row in `Employee Master (Final)`: change `Final RM`, set `Changes Applied` to include "RM changed".
4. Update `org_chart.html`: remove X from A's `<ul>`, add to B's. If B is a leaf, convert to a `<details>` block.
5. Add new row to **Version** tab (newest at top): bumped version, today's date, updated by, summary, tabs touched.
6. Sense check: `Employee Master (Final)` Active count = active rows in `Src - Personnel Budget` (minus exits, minus role-changed-to-teaching, minus AF313/AF314).
7. Update APPROACH.md: bump "Current version" and "Last updated"; add Changelog entry.
8. Push all three to outputs via `present_files`.
9. Tell Akshay to upload the workbook + APPROACH.md to Drive.

---

## Changelog

Newest at top. One line per session summarizing what changed.

- **2026-04-26 v1.0** — Initial workbook build (13 tabs incl. Version). Major decisions: JNV PMU Sr Mgr wraps Urvashi/Lorina; Mushahid + 6 telecallers → Juned; Shivangi Trivedi → Savin (with Subhas); Abhinav → Deepansh team; APC required only where named (27 phantom TBHs removed); Centre Reassigned status added; Vacancy List flat with subject filters; Budget Corrections split from Teachers Audit; Centre Master sticky header + first column; mobile dropdown menu. APPROACH.md created as session protocol with user-handled Drive uploads.

---

## Known limitations

- **Source files may have moved on.** If Personnel Budget or Teacher Recruitment has been edited since the snapshot date in the Version tab, re-pull before making decisions.
- **Cell colour reading is fragile.** If the recruitment sheet owner re-themes, the colour map breaks. Validate one or two known teachers' levels after any re-pull.
- **AF ID coverage incomplete.** 28 teachers in recruitment have no AF ID (likely new joiners). Listed in Data Issues.
- **TBH subject info missing.** All TBH001-TBH116 rows have empty designation/level columns. Listed in Data Issues.
