---
description: "Pre-flight check for deal review slides. Reads PPT time, TSV output, Price list, Context, and Offer insights from a supplier Google Sheet and runs checks before copy-pasting to Excel/think-cell."
allowed-tools: Bash
argument-hint: "<google_sheet_url>"
---

# Pre-flight Slide Check

Run this before copying data to Excel and updating think-cell.

The skill **only makes changes in the Google Sheet** — never in Excel or PowerPoint.

---

## Quick Start (for new teammates)

**Step 1 — Clone the repo**
```bash
git clone <repo_url> ~/skills/picnic-claude-skills
cd ~/skills/picnic-claude-skills
```

**Step 2 — Run the setup script**
```bash
chmod +x setup.sh && ./setup.sh
```
This installs dependencies for every skill in the repo, creates all `~/.claude/skills/` symlinks automatically, and checks that Google Sheets credentials are present. If credentials are missing, follow the auth instructions it prints.

**Step 3 — Use the skill**

In Claude Code, type:
```
/NT-slides-check <google_sheet_url>
```
That's it. The skill pulls the latest checks from git automatically every time it runs.

---

## Contributor's Workflow

The `checks/` folder is the plugin registry. Every `.py` file in it that subclasses `CheckTemplate` is automatically discovered and run. Adding a new check is a three-step loop:

**1. Write your check**

Create `checks/check_NN_your_name.py` (next unused integer for `NN`). Copy this minimal template:

```python
from typing import Optional
from core.check_template import CheckTemplate, Finding, AuditContext

class MyCheck(CheckTemplate):
    id         = 21           # next unused integer — never reuse
    name       = "My check"
    sheet_name = "PPT time"   # tab this check concerns
    severity   = "WARNING"    # "ERROR" or "WARNING"

    def run(self, dm, ctx: AuditContext) -> Optional[Finding]:
        # dm.ppt_time, dm.tsv, dm.deal_sheet, etc. — all pre-fetched, no API calls
        if some_condition:
            return Finding("WARNING", "PPT time", "Describe the problem here.")
        ctx.ok_count += 1
        return None  # None = check passed
```

See `SKILL.md → Contributor Guide` (at the bottom) for the full reference including how to add an auto-fix.

**2. Test it locally**
```bash
cd ~/skills/nt-slides-check
poetry run python audit.py <any_sheet_id>
```
Run it against a real sheet and confirm the check behaves correctly — fix any issues before moving on.

**3. Ask before pushing**

Only once the check is confirmed working locally, ask: "Should I push this to origin/main?"

Push only after explicit confirmation:
```bash
git add checks/check_NN_my_name.py
git commit -m "Add check NN: describe what it catches"
git push
```

Because `audit.py` runs `git pull --ff-only` at startup, every colleague's next invocation will include your new check automatically — no setup required on their end.

---

## Step 1: Get the sheet ID

Arguments provided: `$ARGUMENTS`

Extract the Google Sheets ID from the URL (part between `/d/` and the next `/`).
If no argument was provided, ask the user for the Google Sheet URL.

## Step 2: Run the audit

Run with:
```bash
cd ~/.claude/skills/NT-slides-check && poetry run python audit.py REPLACE_WITH_SHEET_ID
```

The audit script (`audit.py`) auto-syncs via `git pull` before running, then dynamically loads all checks from `checks/`.

**Legacy reference** (only needed if audit.py is unavailable):

```python
import sys, json
sys.path.insert(0, "/home/picnic/.claude/skills/picnic-gsheet")
from gsheet_auth import get_credentials
import gspread
from datetime import date, timedelta

SHEET_ID = "REPLACE_WITH_SHEET_ID"

creds = get_credentials()
gc = gspread.authorize(creds)
sh = gc.open_by_key(SHEET_ID)
sheet_title = sh.title

# ── Single batch read — all tabs in one HTTP request ──────────────────────
# Shelf analysis: rows 1-3 only (keyweek lookup) — not the entire sheet.
# Context: rows 1-82 only — L4L check uses rows 80-81, indices 79-80.
_batch = sh.values_batch_get(
    ranges=[
        "'Shelf analysis'!1:3",
        "'PPT time'",
        "'TSV output'",
        "'Price list'",
        "Context!1:82",
        "'Offer insights'",
        "'PPT context'",
        "'Deal sheet'",
    ],
    params={"valueRenderOption": "FORMATTED_VALUE"},
)
_vr = _batch.get("valueRanges", [])

def _get_vr(i):
    return _vr[i].get("values", []) if i < len(_vr) else []

_shelf_rows    = _get_vr(0)
ppt_time       = _get_vr(1)
tsv            = _get_vr(2)
price_list     = _get_vr(3)
context        = _get_vr(4)
offer_insights = _get_vr(5)
ppt_ctx        = _get_vr(6)
deal_sheet     = _get_vr(7)

# Second read: PPT time formulas (different render option — cannot batch with values above)
ws_ppt       = sh.worksheet("PPT time")
ppt_formulas = ws_ppt.get_all_values(value_render_option="FORMULA")

# Current keyweek: find 'Offer' label in Shelf analysis row 3, read year+week from row 2 above it.
current_keyweek = None
if len(_shelf_rows) > 2:
    for _j, _v in enumerate(_shelf_rows[2]):
        if _v.strip().lower() == "offer":
            _kw_raw = _shelf_rows[1][_j].strip() if _j < len(_shelf_rows[1]) else ""
            try:
                current_keyweek = int(_kw_raw)
            except ValueError:
                pass
            break

# Fallback: date-based (last completed ISO week)
if current_keyweek is None:
    _ref = date.today() - timedelta(days=7)
    _y, _w, _ = _ref.isocalendar()
    current_keyweek = int(f"{_y}{_w:02d}")

tsv_header    = tsv[0]
supplier_col  = tsv_header.index("Contract_party") if "Contract_party" in tsv_header else None
offer_col     = tsv_header.index("Offer_ID")       if "Offer_ID"       in tsv_header else None
supplier_name = tsv[1][supplier_col] if supplier_col is not None and len(tsv) > 1 else "?"
offer_id      = tsv[1][offer_col]    if offer_col    is not None and len(tsv) > 1 else "?"

# Stable fix IDs — never renumber; map to Fix handlers below
FIX_CLEAR_DEALPOINTS  = 1
FIX_BLANK_N3_FUTURE   = 2
FIX_EXTEND_DATES      = 3
FIX_FILL_METRIC_GAPS  = 4
FIX_SET_B5_BENCHMARK  = 5
FIX_SET_PAYMENT_DAYS  = 6

findings    = []   # dicts: severity, sheet, message, fix_id, fix_data
auto_fixed  = []   # messages for actions applied automatically without user input
ok_count    = 0

def _finding(severity, sheet, message, fix_id=None, fix_data=None):
    findings.append({
        "severity": severity,
        "sheet":    sheet,
        "message":  message,
        "fix_id":   fix_id,
        "fix_data": fix_data,
    })

def _ok():
    global ok_count
    ok_count += 1


# ── CHECK 1 — Dealpoints ──────────────────────────────────────────────────
#
# For each week in PPT time where Dealpoint = TRUE, count how many articles
# in Price list have "Key week valid from" equal to that week.
# Flag the dealpoint if that count is < 40% of total articles in TSV output.
#
# PPT time layout (0-indexed rows):
#   Row 0: week numbers (e.g. 202201, 202202, ...)
#   Row 2: "Dealpoint" — TRUE or empty per week

week_row      = ppt_time[0]   # week number per column
dealpoint_row = ppt_time[2]   # "Dealpoint" row

# Total articles in current assessment (TSV output, excluding header)
total_articles = len([r for r in tsv[1:] if any(r)])

# Price list: find "Key week valid from" column
pl_header = price_list[0]
kw_col = None
for i, h in enumerate(pl_header):
    if "key week valid from" in h.lower():
        kw_col = i
        break

dealpoint_weeks_found = []
suspicious_dealpoints = []

if kw_col is None:
    results.append(("ERROR", "Dealpoints",
        "Column 'Key week valid from' not found in Price list tab. Cannot check dealpoints.", None))
else:
    # Build a count of articles per valid-from week from Price list
    week_counts = {}
    for row in price_list[1:]:
        if not any(row):
            continue
        wk = row[kw_col].strip()
        if wk:
            week_counts[wk] = week_counts.get(wk, 0) + 1

    for col_i, dp_val in enumerate(dealpoint_row):
        if dp_val != "TRUE":
            continue
        wk = week_row[col_i].strip() if col_i < len(week_row) else ""
        if not wk:
            continue
        dealpoint_weeks_found.append(wk)
        if wk == str(current_keyweek):
            continue  # current week always has a dealpoint — skip suspicious check
        count = week_counts.get(wk, 0)
        threshold = total_articles * 0.4
        if count < threshold:
            suspicious_dealpoints.append({
                "week": wk,
                "count": count,
                "total": total_articles,
                "col_index": col_i  # 0-based, for fix
            })

    if not dealpoint_weeks_found:
        results.append(("OK", "Dealpoints", "No dealpoints found.", None))
    elif not suspicious_dealpoints:
        results.append(("OK", "Dealpoints",
            f"{len(dealpoint_weeks_found)} dealpoint(s) found, all valid "
            f"(>= 40% of {total_articles} articles): {', '.join(dealpoint_weeks_found)}.", None))
    else:
        lines = []
        for d in suspicious_dealpoints:
            pct = round(100 * d['count'] / d['total']) if d['total'] else 0
            lines.append(f"  Week {d['week']}: {d['count']} of {d['total']} articles ({pct}%) — likely not a real dealpoint")
        valid = [w for w in dealpoint_weeks_found if w not in [d['week'] for d in suspicious_dealpoints]]
        msg = f"{len(suspicious_dealpoints)} suspicious dealpoint(s):\n" + "\n".join(lines)
        if valid:
            msg += f"\n  Valid dealpoints (>= 50%): {', '.join(valid)}"
        results.append(("WARNING", "Dealpoints", msg,
            {"fix": "clear_dealpoints", "suspicious": suspicious_dealpoints}))


# ── CHECK 3 — PPT time starts at week 202201 ─────────────────────────────
#
# Cell B1 in PPT time (ppt_time[0][1]) must equal "202201".
# If not, the time series is misaligned and all charts will be off.

b1_value = ppt_time[0][1].strip() if len(ppt_time) > 0 and len(ppt_time[0]) > 1 else ""
if b1_value != "202201":
    results.append(("ERROR", "PPT time start week",
        f"Cell B1 in PPT time is '{b1_value}' but must be '202201'. "
        "The time series is misaligned — charts will show incorrect weeks.", None))
else:
    results.append(("OK", "PPT time start week", "PPT time starts correctly at week 202201.", None))


# ── CHECK 2 — N3 line extends past current week ───────────────────────────
#
# PPT time layout (0-indexed rows):
#   Row 3: "Net 3 - all SKUs"
#   Row 6: "Net 3 - L4L"
#
# Flag if any week > current_keyweek has a non-empty N3 value.
# Use NLP L4L extent as the effective cutoff: if the TRUE range is stale and NLP
# extends further, N3 ending before the real current week should not be flagged.

n3_all_row = ppt_time[3] if len(ppt_time) > 3 else []
n3_l4l_row = ppt_time[6] if len(ppt_time) > 6 else []

# Determine effective N3 cutoff: max of current_keyweek and NLP L4L last data week
_nlp_l4l_early = ppt_time[7] if len(ppt_time) > 7 else []
_nlp_last_week = current_keyweek
for _i, _v in enumerate(_nlp_l4l_early):
    if _v.strip() and _i < len(week_row):
        try:
            _nlp_last_week = max(_nlp_last_week, int(week_row[_i]))
        except: pass
n3_cutoff_week = _nlp_last_week

future_weeks = []
last_n3_week = None

for col_i, wk in enumerate(week_row[1:], start=1):
    if not wk:
        continue
    try:
        wk_int = int(wk)
    except ValueError:
        continue
    val_all = n3_all_row[col_i] if col_i < len(n3_all_row) else ""
    val_l4l = n3_l4l_row[col_i] if col_i < len(n3_l4l_row) else ""
    if val_all or val_l4l:
        last_n3_week = wk_int
        if wk_int > n3_cutoff_week:
            future_weeks.append(col_i)

if future_weeks:
    results.append(("WARNING", "N3 line cutoff",
        f"Net 3 extends {len(future_weeks)} week(s) past current week "
        f"(last N3 week: {last_n3_week}, current week: {current_keyweek}). "
        "The chart line will run into the future.",
        {"fix": "blank_n3_future", "cols": future_weeks}))
else:
    results.append(("OK", "N3 line cutoff",
        f"Net 3 ends at or before current week ({current_keyweek}).", None))


# ── CHECK 4 — PPT time row 1 extends to week 202652 ──────────────────────
#
# Row 1 of PPT time is a TRANSPOSE of the Dates tab (column B).
# Check if 202652 is present anywhere in row 1.
# If not, the Dates tab needs more rows AND PPT time rows 3+ need formula extension.
#
# Column JA = column 261 (1-indexed). With col A as label and col B as the
# first week (202201), col JA holds week 202652 (5 years × 52 weeks = 260 cols).

TARGET_WEEK   = 202652
COL_JA_1IDX   = 261   # column JA, 1-indexed
COL_JA_0IDX   = 260   # column JA, 0-indexed (for Sheets API)

week_row_ints = []
for v in week_row[1:]:
    try: week_row_ints.append(int(v))
    except: week_row_ints.append(None)

last_week_in_row1 = max((w for w in week_row_ints if w), default=None)

if last_week_in_row1 is None:
    results.append(("ERROR", "PPT time full timeline",
        "No week numbers found in PPT time row 1.", None))
elif last_week_in_row1 >= TARGET_WEEK:
    results.append(("OK", "PPT time full timeline",
        f"PPT time row 1 reaches week {last_week_in_row1} (>= {TARGET_WEEK}).", None))
else:
    results.append(("ERROR", "PPT time full timeline",
        f"PPT time row 1 ends at week {last_week_in_row1}, missing weeks up to {TARGET_WEEK}. "
        f"Dates tab needs extending and PPT time formulas need copying to column JA.",
        {"fix": "extend_dates_and_ppt_time", "last_week": last_week_in_row1}))


# ── CHECK 5 — L4L % of net sales LTM ─────────────────────────────────────
#
# Context tab layout (0-indexed):
#   Row 79 (row 80): B80 = L4L date dropdown
#   Row 80 (row 81): E81 = % of net sales LTM for In L4L articles
#
# Thresholds: >= 75% OK, 70–74% WARNING, < 70% ERROR

l4l_date    = context[79][1].strip() if len(context) > 79 and len(context[79]) > 1 else ""
l4l_pct_raw = context[80][4].strip() if len(context) > 80 and len(context[80]) > 4 else ""

if not l4l_pct_raw:
    results.append(("ERROR", "L4L % of net sales",
        f"L4L % of net sales (Context E81) is empty. Check the L4L date in B80.", None))
else:
    try:
        l4l_pct = float(l4l_pct_raw.replace("%", "").replace(",", ".").strip())
        if l4l_pct >= 75:
            results.append(("OK", "L4L % of net sales",
                f"L4L covers {l4l_pct:.0f}% of net sales LTM (date: {l4l_date}). Above 75% threshold.", None))
        elif l4l_pct >= 70:
            results.append(("WARNING", "L4L % of net sales",
                f"L4L covers {l4l_pct:.0f}% of net sales LTM (date: {l4l_date}). "
                f"Below 75% threshold — consider moving the L4L date forward.", None))
        else:
            results.append(("ERROR", "L4L % of net sales",
                f"L4L covers only {l4l_pct:.0f}% of net sales LTM (date: {l4l_date}). "
                f"Well below 75% threshold — L4L date must be adjusted.", None))
    except ValueError:
        results.append(("ERROR", "L4L % of net sales",
            f"Could not parse L4L % value: '{l4l_pct_raw}'", None))


# ── CHECK: Benchmark article count consistency ────────────────────────────
# TSV output column "In_benchmark" TRUE count must equal Offer insights E7.
# E7 = the expected number of benchmark articles entered manually.
# Find the column by header name — layout varies between sheets.

_benchmark_col_idx = None
for _i, _h in enumerate(tsv_header):
    if _h.strip().lower() == "in_benchmark":
        _benchmark_col_idx = _i
        break

if _benchmark_col_idx is None:
    results.append(("WARNING", "Benchmark article count",
        "Column 'In_benchmark' not found in TSV output. Cannot verify benchmark count.", None))
    _benchmark_true = None
else:
    _benchmark_true = sum(
        1 for row in tsv[1:] if any(row) and len(row) > _benchmark_col_idx and row[_benchmark_col_idx].strip().upper() == "TRUE"
    )

if _benchmark_true is not None:
    _oi_e7_raw = offer_insights[6][4].strip() if len(offer_insights) > 6 and len(offer_insights[6]) > 4 else ""
    try:
        _oi_e7 = int(_oi_e7_raw)
    except (ValueError, TypeError):
        _oi_e7 = None

    if _oi_e7 is None:
        results.append(("WARNING", "Benchmark article count",
            f"Offer insights E7 is empty or non-numeric ('{_oi_e7_raw}'). Cannot verify benchmark count.", None))
    elif _benchmark_true != _oi_e7:
        results.append(("WARNING", "Benchmark article count",
            f"In_benchmark TRUE count in TSV ({_benchmark_true}) does not match Offer insights E7 ({_oi_e7}). "
            "Fix manually — check which articles are flagged In_benchmark in TSV output.", None))
    else:
        results.append(("OK", "Benchmark article count",
            f"In_benchmark count matches Offer insights E7: {_benchmark_true} articles.", None))


# ── CHECK 6 — Offer insights: >= 7 articles selected ─────────────────────
#
# Scan ALL rows for "include in deck" label — independent of "Offers type:" row.
# Count TRUE values in that column below the header row.
# Flag if fewer than 7 are selected.

_deck_col = None
_deck_header_row = None
for _ri, _row in enumerate(offer_insights):
    for _ci, _val in enumerate(_row):
        if "include in deck" in _val.strip().lower():
            _deck_col = _ci
            _deck_header_row = _ri
            break
    if _deck_col is not None:
        break

if _deck_col is None:
    results.append(("WARNING", "Offer insights selection",
        "Column 'Include in deck?' not found in Offer insights. Cannot check article selection.", None))
else:
    true_count = sum(
        1 for row in offer_insights[_deck_header_row + 1:]
        if len(row) > _deck_col and row[_deck_col].strip().upper() == "TRUE"
    )
    if true_count < 5:
        results.append(("WARNING", "Offer insights selection",
            f"Only {true_count} article(s) selected in Offer insights ('Include in deck?' column). "
            f"At least 5 are needed to populate the PPT time chart.", None))
    else:
        results.append(("OK", "Offer insights selection",
            f"{true_count} articles selected in Offer insights (>= 7).", None))


# ── CHECK 7 — PPT time data continuity (rows 4+) ─────────────────────────
#
# All rows from row 4 (0-indexed row 3) onwards that have any data should be
# continuously filled up to current_keyweek. Row 3 (Dealpoint) is intentionally
# excluded — cleared dealpoints should stay cleared.

cur_col = None
for i, v in enumerate(week_row):
    try:
        if int(v.strip()) == current_keyweek:
            cur_col = i
            break
    except: pass

gap_rows = []
if cur_col is not None:
    for row_idx in range(3, len(ppt_time)):  # 0-indexed row 3 = sheet row 4
        row_data = ppt_time[row_idx]
        fml_row  = ppt_formulas[row_idx] if row_idx < len(ppt_formulas) else []
        row_label = row_data[0].strip() if row_data else ""
        # A cell counts as present if it has a non-empty value OR a formula (starts with =)
        def cell_present(c, _rd=row_data, _fr=fml_row):
            has_val = c < len(_rd) and _rd[c].strip()
            has_fml = c < len(_fr) and str(_fr[c]).startswith("=")
            return bool(has_val or has_fml)
        last_filled_col = None
        for c in range(cur_col, 0, -1):
            if cell_present(c):
                last_filled_col = c
                break
        if last_filled_col is None or last_filled_col >= cur_col:
            continue  # already filled to current week, or no data at all
        if not any(cell_present(c) for c in range(1, cur_col + 1)):
            continue  # completely empty row, skip
        last_week = None
        if last_filled_col < len(week_row):
            try: last_week = int(week_row[last_filled_col])
            except: pass
        gap_rows.append({
            "row_name": row_label or f"Row {row_idx + 1}",
            "row_0idx": row_idx,
            "last_filled_col": last_filled_col,
            "last_week": last_week,
            "gap_weeks": cur_col - last_filled_col,
        })

if not gap_rows:
    results.append(("OK", "PPT time data continuity",
        f"All rows (row 4+) filled up to current week ({current_keyweek}).", None))
else:
    lines = [f"  {g['row_name']}: last data week {g['last_week']}, {g['gap_weeks']} week(s) missing"
             for g in gap_rows]
    results.append(("WARNING", "PPT time data continuity",
        f"{len(gap_rows)} row(s) have gaps before week {current_keyweek}:\n" + "\n".join(lines),
        {"fix": "fill_metric_gaps", "gaps": gap_rows, "cur_col": cur_col}))


# ── CHECK: PPT time current week ─────────────────────────────────────────
# If NLP - L4L (row 8, 0-indexed 7) has data beyond cur_col, the current
# week marker (row 3 TRUE range) needs extending.
_nlp_l4l = ppt_time[7] if len(ppt_time) > 7 else []
if cur_col is not None:
    _nlp_future_cols = [i for i in range(cur_col + 1, len(_nlp_l4l)) if _nlp_l4l[i].strip()]
    if _nlp_future_cols:
        _last_nlp_wk = week_row[_nlp_future_cols[-1]] if _nlp_future_cols[-1] < len(week_row) else "?"
        results.append(("ERROR", "PPT time current week",
            f"NLP - L4L has data up to week {_last_nlp_wk} but current week is set to {current_keyweek}. "
            "Extend the TRUE range in PPT time row 3 to the actual current week.",
            None))
    else:
        results.append(("OK", "PPT time current week",
            f"NLP - L4L data stops at current week ({current_keyweek}).", None))


# ── CHECK A — PPT context: header fields B1–B4 ───────────────────────────
# B1–B4 populate the slide header area. All four should be non-empty.
_header_cells = [(0, 1, "B1"), (1, 1, "B2"), (2, 1, "B3"), (3, 1, "B4")]
empty_headers = []
for _row_idx, _col_idx, _ref in _header_cells:
    _val = ppt_ctx[_row_idx][_col_idx].strip() if len(ppt_ctx) > _row_idx and len(ppt_ctx[_row_idx]) > _col_idx else ""
    if not _val:
        empty_headers.append(_ref)

if empty_headers:
    results.append(("WARNING", "PPT context header fields",
        f"Header cell(s) {', '.join(empty_headers)} in PPT context are empty. These populate the slide header.", None))
else:
    results.append(("OK", "PPT context header fields",
        "Header fields B1–B4 are all filled.", None))


# ── CHECK B — PPT context: L2 overview (A12:E31) ─────────────────────────
# This block feeds the supplier KPI table in the slides (L2 overview).
# ERROR only if the entire block is empty — partial fill is expected.
_l2_non_empty = sum(
    1
    for _ri in range(11, 31)
    if _ri < len(ppt_ctx)
    for _ci in range(5)
    if _ci < len(ppt_ctx[_ri]) and ppt_ctx[_ri][_ci].strip()
)

if _l2_non_empty == 0:
    results.append(("ERROR", "PPT context L2 overview",
        "The L2 overview section (A12:E31) is completely empty. The KPI table in the slides will be blank.", None))
else:
    results.append(("OK", "PPT context L2 overview",
        f"L2 overview section has data ({_l2_non_empty} non-empty cells).", None))


# ── CHECK 8 — PPT context: benchmark name in B5 ──────────────────────────
# The benchmark name should be in PPT context B5. If empty, the suggested
# value comes from Offer insights E6.
b5_value = ppt_ctx[4][1].strip()           if len(ppt_ctx) > 4         and len(ppt_ctx[4]) > 1         else ""
oi_e6    = offer_insights[5][4].strip()    if len(offer_insights) > 5  and len(offer_insights[5]) > 4  else ""

if not b5_value:
    if oi_e6:
        results.append(("ERROR", "PPT context benchmark name",
            f"Cell B5 in PPT context is empty. Offer insights E6 suggests: '{oi_e6}'.",
            {"fix": "set_b5_benchmark", "suggested": oi_e6}))
    else:
        results.append(("ERROR", "PPT context benchmark name",
            "Cell B5 in PPT context is empty and Offer insights E6 is also empty. Provide a benchmark name manually.",
            {"fix": "set_b5_benchmark", "suggested": None}))
else:
    results.append(("OK", "PPT context benchmark name",
        f"Benchmark name: '{b5_value}'.", None))


# ── CHECK 9 — PPT context: B35 vs N35 competitor names ───────────────────
# B35 = index [34][1], N35 = index [34][13]
b35_value = ppt_ctx[34][1].strip()  if len(ppt_ctx) > 34 and len(ppt_ctx[34]) > 1  else ""
n35_value = ppt_ctx[34][13].strip() if len(ppt_ctx) > 34 and len(ppt_ctx[34]) > 13 else ""

if not b35_value and not n35_value:
    results.append(("WARNING", "PPT context competitor names",
        "Both B35 and N35 in PPT context are empty. Set competitor names for the competitive analysis.", None))
elif b35_value == n35_value:
    results.append(("ERROR", "PPT context competitor names",
        f"B35 and N35 have the same value ('{b35_value}'). They must be different competitors (e.g. 'AH' and 'LY'). Fix manually in Offer insights — do not overwrite B35/N35 directly.",
        None))
else:
    results.append(("OK", "PPT context competitor names",
        f"Competitors: B35='{b35_value}', N35='{n35_value}'.", None))


# ── CHECK 10 — PPT context: competitive section (A35:W55) ────────────────
# Covers the full competitive data block (rows 36–55, 0-indexed 35–54).
# Left block:  col A (idx 0)  = promo group name, cols B–J (1–9) = data, col K (10) = net sales
# Right block: col M (idx 12) = promo group name, cols N–V (13–21) = data, col W (22) = net sales
# If a group name is present, data columns and net sales must be filled.
# If no group name but data is present, net sales must still be filled.

comp_issues = []
for _ri in range(35, 55):  # 0-indexed 35–54 = sheet rows 36–55
    if _ri >= len(ppt_ctx):
        break
    _row = ppt_ctx[_ri]
    if not any(_row):
        continue

    _left_group = _row[0].strip()  if len(_row) > 0  else ""
    _left_data  = any(_row[i].strip() for i in range(1, 10)  if i < len(_row))
    _left_ns    = _row[10].strip() if len(_row) > 10 else ""

    _right_group = _row[12].strip() if len(_row) > 12 else ""
    _right_data  = any(_row[i].strip() for i in range(13, 22) if i < len(_row))
    _right_ns    = _row[22].strip() if len(_row) > 22 else ""

    if _left_group:
        if not _left_data:
            comp_issues.append(f"Row {_ri + 1} ('{_left_group}'): no data in cols B–J")
        elif not _left_ns:
            comp_issues.append(f"Row {_ri + 1} ('{_left_group}'): net sales (col K) is empty")
    elif _left_data and not _left_ns:
        comp_issues.append(f"Row {_ri + 1}: data in cols B–J but net sales (col K) is empty")

    if _right_group:
        if not _right_data:
            comp_issues.append(f"Row {_ri + 1} ('{_right_group}'): no data in cols N–V")
        elif not _right_ns:
            comp_issues.append(f"Row {_ri + 1} ('{_right_group}'): net sales (col W) is empty")
    elif _right_data and not _right_ns:
        comp_issues.append(f"Row {_ri + 1}: data in cols N–V but net sales (col W) is empty")

if comp_issues:
    results.append(("WARNING", "PPT context competitive section",
        f"{len(comp_issues)} issue(s) in competitive section (rows 36–55):\n" +
        "\n".join(f"  {line}" for line in comp_issues), None))
else:
    results.append(("OK", "PPT context competitive section",
        "Competitive section: all rows with group names have data and net sales.", None))


# ── CHECK 15 — PPT time: N3 L4L stable between dealpoints ────────────────
# Net 3 - L4L (0-indexed row 6) should not change between consecutive dealpoints.
# Mid-period changes indicate a price update not tied to a dealpoint.

n3_l4l_row  = ppt_time[6] if len(ppt_time) > 6 else []
dp_cols_all = [c for c, v in enumerate(dealpoint_row) if v == "TRUE"]
last_n3_col = max((c for c, v in enumerate(n3_l4l_row) if v.strip()), default=None)

if last_n3_col is None or not dp_cols_all:
    results.append(("OK", "PPT time N3 L4L stability",
        "No N3 L4L data or no dealpoints — stability check skipped.", None))
else:
    boundaries   = dp_cols_all + [last_n3_col + 1]
    unstable_segs = []
    for i in range(len(boundaries) - 1):
        seg_start, seg_end = boundaries[i], boundaries[i + 1]
        vals = [(c, n3_l4l_row[c].strip()) for c in range(seg_start, seg_end)
                if c < len(n3_l4l_row) and n3_l4l_row[c].strip()]
        if len(set(v for _, v in vals)) <= 1:
            continue  # stable or empty segment
        first_val = vals[0][1]
        change_col, change_val = next((c, v) for c, v in vals if v != first_val)
        dp_wk      = week_row[seg_start].strip() if seg_start < len(week_row) else "?"
        change_wk  = week_row[change_col].strip() if change_col  < len(week_row) else "?"
        prev_wk    = week_row[change_col - 1].strip() if 0 < change_col - 1 < len(week_row) else "?"
        next_dp_wk = week_row[boundaries[i+1]].strip() if boundaries[i+1] < len(week_row) else "end"
        unstable_segs.append(
            f"  Between dealpoints {dp_wk}–{next_dp_wk}: "
            f"{dp_wk}–{prev_wk} at {first_val}, then changes at {change_wk} to {change_val}")

    if unstable_segs:
        results.append(("WARNING", "PPT time N3 L4L stability",
            f"Net 3 - L4L changes within {len(unstable_segs)} inter-dealpoint period(s) — "
            "check Price list for mid-period price changes:\n" + "\n".join(unstable_segs), None))
    else:
        results.append(("OK", "PPT time N3 L4L stability",
            "Net 3 - L4L is stable within all inter-dealpoint periods.", None))


# ── Find Deal sheet column indices from header labels ─────────────────────
# Row 4 (0-indexed 3) has column labels: 'Deal', 'LY', 'AH', 'Jumbo', 'Target'
# Row 2 (0-indexed 1) has: 'Selected offer'
# Columns can shift when benchmark columns are added or removed — look up by label.

_hdr4 = deal_sheet[3] if len(deal_sheet) > 3 else []
_hdr2 = deal_sheet[1] if len(deal_sheet) > 1 else []

def _find_col_by_label(row, label):
    for i, v in enumerate(row):
        if v.strip().lower() == label.lower():
            return i
    return None

col_target    = _find_col_by_label(_hdr4, "Target")
col_ly        = _find_col_by_label(_hdr4, "LY")
col_deal      = _find_col_by_label(_hdr4, "Deal")
col_sel_offer = _find_col_by_label(_hdr2, "Selected offer")

# Payment days row: find by col C label 'Payment days'
_payment_row_0idx = None
for _ri, _row in enumerate(deal_sheet):
    if len(_row) > 2 and _row[2].strip().lower() == "payment days":
        _payment_row_0idx = _ri
        break


# ── CHECK 11 — Deal sheet: target column filled for key rows ─────────────
# 'Target' column found by header label in row 4.
# Rows checked: Margin (%), Coverage, Budget (€k), GMD impact.
TARGET_ROWS = [(10, 11), (14, 15), (15, 16), (22, 23)]

def deal_label(row_0idx):
    row = deal_sheet[row_0idx] if row_0idx < len(deal_sheet) else []
    col_c = row[2].strip() if len(row) > 2 else ""
    if col_c:
        return col_c
    # No label in col C — try to derive from previous row's label + unit (col D)
    col_d = row[3].strip() if len(row) > 3 else ""
    if col_d and row_0idx > 0:
        prev_row = deal_sheet[row_0idx - 1] if row_0idx - 1 < len(deal_sheet) else []
        prev_c = prev_row[2].strip() if len(prev_row) > 2 else ""
        if prev_c:
            return f"{prev_c} ({col_d})"
    col_a = row[0].strip() if len(row) > 0 else ""
    return col_a or f"Row {row_0idx + 1}"

if col_target is None:
    results.append(("ERROR", "Deal sheet target column",
        "Column 'Target' not found in Deal sheet row 4. Cannot check target values.", None))
else:
    empty_target_rows = []
    for row_0idx, row_1idx in TARGET_ROWS:
        val = deal_sheet[row_0idx][col_target].strip() if len(deal_sheet) > row_0idx and len(deal_sheet[row_0idx]) > col_target else ""
        if not val:
            empty_target_rows.append(f"Row {row_1idx} ('{deal_label(row_0idx)}')")

    if empty_target_rows:
        results.append(("ERROR", "Deal sheet target column",
            f"{len(empty_target_rows)} key target cell(s) are empty:\n" +
            "\n".join(f"  {r}" for r in empty_target_rows), None))
    else:
        results.append(("OK", "Deal sheet target column",
            "All key target cells (rows 11, 15, 16, 23) are filled.", None))


# ── CHECK 12 — Deal sheet: LY benchmark filled ────────────────────────────
# "Offers type" label is always in row 1 or 2; value is the cell immediately to its right.
# LY column found by header label 'LY' in row 4.
offers_type = ""
for _row in deal_sheet[:2]:
    for _ci, _val in enumerate(_row):
        if "offers type" in _val.strip().lower() and _ci + 1 < len(_row):
            offers_type = _row[_ci + 1].strip()
            break
    if offers_type:
        break

if "promo" in offers_type.lower() and "shelf" in offers_type.lower():
    ly_check_rows = list(range(5, 22))   # rows 6–22
elif "shelf" in offers_type.lower():
    ly_check_rows = list(range(5, 11))   # rows 6–11
else:
    ly_check_rows = []

if not offers_type:
    results.append(("ERROR", "Deal sheet LY benchmark",
        "Cell L1 in Deal sheet is empty — cannot determine offer type.", None))
elif col_ly is None:
    results.append(("ERROR", "Deal sheet LY benchmark",
        "Column 'LY' not found in Deal sheet row 4. Cannot check LY benchmark.", None))
else:
    empty_ly = []
    for row_0idx in ly_check_rows:
        val = deal_sheet[row_0idx][col_ly].strip() if len(deal_sheet) > row_0idx and len(deal_sheet[row_0idx]) > col_ly else ""
        if not val:
            empty_ly.append(f"Row {row_0idx + 1} ('{deal_label(row_0idx)}')")

    if empty_ly:
        results.append(("WARNING", "Deal sheet LY benchmark",
            f"LY benchmark missing for {len(empty_ly)} row(s) (offer type: {offers_type}):\n" +
            "\n".join(f"  {r}" for r in empty_ly), None))
    else:
        results.append(("OK", "Deal sheet LY benchmark",
            f"LY benchmark fully filled (offer type: {offers_type}).", None))


# ── CHECK E — Deal sheet: deal column filled ──────────────────────────────
# 'Deal' column found by header label in row 4. Same row range as LY benchmark.

if offers_type and col_deal is None:
    results.append(("ERROR", "Deal sheet deal column",
        "Column 'Deal' not found in Deal sheet row 4. Cannot check deal values.", None))
elif offers_type:
    empty_deal_e = []
    for row_0idx in ly_check_rows:
        val = deal_sheet[row_0idx][col_deal].strip() if len(deal_sheet) > row_0idx and len(deal_sheet[row_0idx]) > col_deal else ""
        if not val:
            empty_deal_e.append(f"Row {row_0idx + 1} ('{deal_label(row_0idx)}')")

    if empty_deal_e:
        results.append(("WARNING", "Deal sheet deal column",
            f"Deal column missing for {len(empty_deal_e)} row(s) (offer type: {offers_type}):\n" +
            "\n".join(f"  {r}" for r in empty_deal_e), None))
    else:
        results.append(("OK", "Deal sheet deal column",
            f"Deal column fully filled (offer type: {offers_type}).", None))


# ── CHECK 13 — Deal sheet: selected offer filled ──────────────────────────
# 'Selected offer' column found by header label in row 2.
# If Promo & Shelf: check row 4 (0-indexed 3). If Shelf only: check row 3 (0-indexed 2).

if "promo" in offers_type.lower() and "shelf" in offers_type.lower():
    sel_row_0idx, sel_row_label = 3, "Promo offer row"
elif "shelf" in offers_type.lower():
    sel_row_0idx, sel_row_label = 2, "Shelf offer row"
else:
    sel_row_0idx, sel_row_label = None, None

if sel_row_0idx is None:
    results.append(("ERROR", "Deal sheet selected offer",
        "Cell L1 is empty — cannot determine which row to check for selected offer.", None))
elif col_sel_offer is None:
    results.append(("ERROR", "Deal sheet selected offer",
        "Column 'Selected offer' not found in Deal sheet row 2. Cannot check selected offer.", None))
else:
    sel_val = deal_sheet[sel_row_0idx][col_sel_offer].strip() if len(deal_sheet) > sel_row_0idx and len(deal_sheet[sel_row_0idx]) > col_sel_offer else ""
    if not sel_val:
        results.append(("ERROR", "Deal sheet selected offer",
            f"No selected offer in {sel_row_label} (offer type: {offers_type}). Fill it in manually.", None))
    else:
        results.append(("OK", "Deal sheet selected offer",
            f"Selected offer: '{sel_val}' ({sel_row_label}).", None))


# ── CHECK 14 — Deal sheet: payment days filled ────────────────────────────
# Row found by col C label 'Payment days'. Value read from the 'Deal' column.

if _payment_row_0idx is None:
    results.append(("ERROR", "Deal sheet payment days",
        "Row 'Payment days' not found in col C of Deal sheet.", None))
elif col_deal is None:
    results.append(("ERROR", "Deal sheet payment days",
        "Column 'Deal' not found in Deal sheet row 4. Cannot check payment days.", None))
else:
    payment_days = deal_sheet[_payment_row_0idx][col_deal].strip() if len(deal_sheet[_payment_row_0idx]) > col_deal else ""
    if not payment_days:
        results.append(("ERROR", "Deal sheet payment days",
            f"Payment days (row {_payment_row_0idx + 1}) is empty.",
            {"fix": "set_payment_days"}))
    else:
        results.append(("OK", "Deal sheet payment days",
            f"Payment days: {payment_days}.", None))


# ── Print report ──────────────────────────────────────────────────────────
# All findings collected above — print in two passes: errors then warnings.
# OK results are never printed; ok_count is incremented by _ok() during checks.

BORDER = "━" * 54

errors   = [f for f in findings if f["severity"] == "ERROR"]
warnings = [f for f in findings if f["severity"] == "WARNING"]

def _fmt_finding(f, icon):
    fix_part = f" (Fix {f['fix_id']})" if f["fix_id"] is not None else ""
    lines = f["message"].split("\n")
    first = f"  {icon} [Tab: {f['sheet']}] {lines[0]}{fix_part}"
    rest  = ["        " + l for l in lines[1:]]
    return "\n".join([first] + rest)

print(BORDER)
print(f"📋  NT-SLIDE-TABS-CHECK — {supplier_name}")
print(f"    Sheet: {sheet_title}  |  Offer: {offer_id}  |  Week: {current_keyweek}")
print(BORDER)
print()

if not findings:
    print("✅  All checks passed.")
else:
    if errors:
        print("❌  ERRORS")
        for f in errors:
            print(_fmt_finding(f, "❌"))
        print()
    if warnings:
        print("⚠️   WARNINGS")
        for f in warnings:
            print(_fmt_finding(f, "⚠️ "))
        print()
    print(BORDER)
    print(f"Audit Complete: {len(errors)} Errors, {len(warnings)} Warnings ({ok_count} checks passed).")

# Emit fix metadata — keyed by stable integer fix_id (as string)
fixable = {str(f["fix_id"]): f["fix_data"]
           for f in findings if f["fix_id"] is not None and f["fix_data"] is not None}
if fixable:
    print()
    print("__FIXABLE__" + json.dumps(fixable))
```

## Step 3: Summarize issues and wait for instructions

The report groups findings by severity. Each finding includes `[Tab: X]` to show which tab is affected, and `(Fix N)` where N is the stable fix_id for auto-fixable issues.

After showing the report, if there are any findings, immediately list them as a concise "To fix" summary — one line per issue, using the same `[Tab: X]` prefix.

Example format:
```
To fix:
[Tab: PPT context] B5 benchmark name — set to '2025' (from Offer insights E6)? (Fix 5)
[Tab: Deal sheet] Payment days empty — how many days? (Fix 6)
[Tab: PPT time] Dealpoints: 202344 (20%), 202418 (25%) — clear them? (Fix 1)
[Tab: TSV output] Benchmark count mismatch (6 vs 8) — fix manually
```

For dealpoints, always include the % of SKUs with a price change at each week.

The user replies with answers mapped to Fix IDs (e.g. "Fix 1 yes / Fix 5 2025 / Fix 6 30").
- Items without a fix_id must be fixed manually; user will say skip or not address them
- After receiving the reply, collect ALL fix operations and dispatch in the minimum number of API calls (see Batch Dispatch below)

If there are no findings, just stop — no summary needed.

### Fix A — Clear suspicious dealpoints

**Important:** Always re-read PPT time row 1 to get current column positions — stored `col_index` values may be stale if B1 was fixed in the same session (B1 change shifts the TRANSPOSE row).

Adds entries to `value_updates` (one per dealpoint week). Previously used a loop of `update_cell()` calls — now a single range per week, all batched together.

```python
ws_ppt = sh.worksheet("PPT time")
week_row_live = ws_ppt.row_values(1)  # re-read current positions
for d in suspicious_dealpoints_to_clear:
    try:
        col_1idx = week_row_live.index(str(d["week"])) + 1
    except ValueError:
        print(f"  Week {d['week']} not found in row 1 — skipping.")
        continue
    value_updates.append({
        "range":  f"'PPT time'!{gspread.utils.rowcol_to_a1(3, col_1idx)}",
        "values": [[""]]
    })
    print(f"  Queued: clear dealpoint week {d['week']} (col {col_1idx}).")
```

### Fix B — Blank N3 after current week

Blank rows 4 and 7 (1-indexed) in `PPT time` for all columns in `future_weeks`. Adds entries to `value_updates` — folded into the cross-fix batch dispatch.

```python
for col_i in future_col_indices:
    value_updates.append({"range": f"'PPT time'!{gspread.utils.rowcol_to_a1(4, col_i + 1)}", "values": [[""]]})
    value_updates.append({"range": f"'PPT time'!{gspread.utils.rowcol_to_a1(7, col_i + 1)}", "values": [[""]]})
print(f"  Queued: blank N3 for {len(future_col_indices)} week(s) after week {current_keyweek}.")
```

### Fix C — Fix B1 in PPT time

Adds one entry to `value_updates`.

```python
value_updates.append({"range": "'PPT time'!B1", "values": [[202201]]})
print("  Queued: set PPT time B1 = 202201.")
```

### Fix D — Extend Dates tab + PPT time to week 202652

Two steps, run in sequence:

**Step 1: Extend the Dates tab**

```python
from datetime import date, timedelta

def next_iso_week(yyyyww):
    year, week = yyyyww // 100, yyyyww % 100
    jan4 = date(year, 1, 4)
    monday = jan4 - timedelta(days=jan4.weekday()) + timedelta(weeks=week - 1)
    nxt = monday + timedelta(weeks=1)
    y, w, _ = nxt.isocalendar()
    return int(f"{y}{w:02d}")

ws_dates = sh.worksheet("Dates")
dates_vals = ws_dates.get_all_values()
last_row = len(dates_vals)          # current last row, 1-indexed
last_week = int(dates_vals[-1][1])  # last week number in col B

rows_to_add = []
current = last_week
while True:
    nxt = next_iso_week(current)
    if nxt > 202652:
        break
    year = nxt // 100
    wk   = nxt % 100
    period_num = (wk - 1) // 4 + 1
    period = int(f"{year}{period_num:02d}")
    n = last_row + len(rows_to_add) + 1  # this row's 1-indexed sheet row number
    rows_to_add.append([
        year,
        nxt,
        period,
        f"=IF(C{n}=C{n-1},,RIGHT(C{n},2))",
        f'=IF(D{n}<>"","P"&VALUE(D{n}),)',
        f"=B{n}>'PPT time'!$B$1"
    ])
    current = nxt

ws_dates.append_rows(rows_to_add, value_input_option="USER_ENTERED")
```

Confirm: "Extended Dates tab by [N] rows (up to week 202652). Rows 1 and 2 of PPT time will update automatically via TRANSPOSE."

**Step 2: Extend PPT time rows 3+ to column JA**

Column JA = column 261 (1-indexed) = index 260 (0-indexed, exclusive end = 261).

Find the rightmost column that has any non-empty value in rows 3+ of PPT time, then copyPaste its formulas across to column JA:

```python
ws_ppt   = sh.worksheet("PPT time")
ppt_vals = ws_ppt.get_all_values()

# Find last populated column in rows 3+ (0-indexed row 2+)
last_col_0 = 0
for row in ppt_vals[2:]:
    for c in range(len(row) - 1, -1, -1):
        if row[c].strip():
            last_col_0 = max(last_col_0, c)
            break

COL_JA_0 = 260  # column JA, 0-indexed
max_row   = len(ppt_vals)  # actual row count — never hardcode this

# Ensure the sheet has enough columns before copying (copyPaste to a non-existent
# column silently does nothing — always expand first if short)
if ws_ppt.col_count <= COL_JA_0:
    cols_needed = COL_JA_0 + 1 - ws_ppt.col_count
    ws_ppt.add_cols(cols_needed)
    print(f"  Expanded PPT time by {cols_needed} column(s) to reach column JA.")

api_requests.append({
    "copyPaste": {
        "source": {
            "sheetId": ws_ppt.id,
            "startRowIndex": 2,           # row 3 (0-indexed)
            "endRowIndex":   max_row,
            "startColumnIndex": last_col_0,
            "endColumnIndex":   last_col_0 + 1
        },
        "destination": {
            "sheetId": ws_ppt.id,
            "startRowIndex": 2,
            "endRowIndex":   max_row,
            "startColumnIndex": last_col_0 + 1,
            "endColumnIndex":   COL_JA_0 + 1   # exclusive → includes JA
        },
        "pasteType":        "PASTE_FORMULA",
        "pasteOrientation": "NORMAL"
    }
})
print(f"  Queued: extend PPT time formulas from col {last_col_0} to column JA.")
```

### Fix E — Re-extend metric row formulas to current week

For each gap row, copyPaste its last filled cell forward to `cur_col`. Appends to `api_requests` — dispatched in the shared batch.

```python
ws_ppt = sh.worksheet("PPT time")
for g in gap_rows_to_fix:
    if g["last_filled_col"] is None:
        continue
    api_requests.append({
        "copyPaste": {
            "source": {
                "sheetId": ws_ppt.id,
                "startRowIndex": g["row_0idx"],
                "endRowIndex":   g["row_0idx"] + 1,
                "startColumnIndex": g["last_filled_col"],
                "endColumnIndex":   g["last_filled_col"] + 1,
            },
            "destination": {
                "sheetId": ws_ppt.id,
                "startRowIndex": g["row_0idx"],
                "endRowIndex":   g["row_0idx"] + 1,
                "startColumnIndex": g["last_filled_col"] + 1,
                "endColumnIndex":   cur_col + 1,
            },
            "pasteType":        "PASTE_FORMULA",
            "pasteOrientation": "NORMAL",
        }
    })
print(f"  Queued: re-extend formulas for {len(gap_rows_to_fix)} row(s) to week {current_keyweek}.")
```

### Fix F — Set benchmark name in PPT context B5

Check the `suggested` field in the fix data:
- If `suggested` is not None: ask "Set benchmark name to '[suggested]' (from Offer insights E6)? (yes/no)"
  - If yes: use the suggested value
  - If no: ask "What benchmark name should I use instead?"
- If `suggested` is None: ask "What benchmark name should I use? (e.g. '2025' or 'LTM')"

Adds one entry to `value_updates`.

```python
value_updates.append({"range": "'PPT context'!B5", "values": [[benchmark_name]]})
print(f"  Queued: set PPT context B5 = '{benchmark_name}'.")
```

### Fix G — Set payment days in Deal sheet

Ask the user: "How many payment days should I fill in?"

Adds one entry to `value_updates` using the dynamic row/col from Check 20.

```python
cell = gspread.utils.rowcol_to_a1(_payment_row_0idx + 1, col_deal + 1)
value_updates.append({"range": f"'Deal sheet'!{cell}", "values": [[payment_days_value]]})
print(f"  Queued: set Deal sheet payment days = {payment_days_value}.")
```

### Batch Dispatch — send all queued operations

Run this after all selected fixes have added to `value_updates` and `api_requests`. Fix D Step 1 (`append_rows`) must already have run before this block if selected.

```python
print("🚀 Applying all fixes in one batch...")

# Step 1 (if Fix D selected): append_rows already executed above — cannot be batched
# Step 2: all cell value writes across all sheets
if value_updates:
    sh.values_batch_update(value_updates, params={"valueInputOption": "USER_ENTERED"})

# Step 3: all structural operations (copyPaste formula extensions)
if api_requests:
    sh.batch_update({"requests": api_requests})

print("✅ Done")
```

**Ordering constraint:** Fix D Step 1 (`append_rows`) must execute before the `api_requests` dispatch, because the copyPaste in Step 2 reads the row count of the just-extended sheet.

## Notes

- Tab names are identical across all supplier sheets — this skill works for any supplier.
- If a tab is not found, report as ERROR: "Tab '[name]' missing — is this the right sheet?"
- The skill never modifies the Excel template or PowerPoint.
- More checks will be added to this skill over time.

---

## Contributor Guide

### How to add a new check

1. **Create a new file** in `checks/` named `check_NN_your_name.py` where `NN` is the next unused integer.

2. **Subclass `CheckTemplate`** and set the four required class attributes:

```python
from typing import Optional
from core.check_template import CheckTemplate, Finding, AuditContext

class MyNewCheck(CheckTemplate):
    id         = 21           # next unused integer — never reuse or renumber
    name       = "My check"   # shown in the report
    sheet_name = "PPT time"   # primary tab this check concerns
    severity   = "WARNING"    # "ERROR" or "WARNING"

    def run(self, dm, ctx: AuditContext) -> Optional[Finding]:
        # Read from dm.* and ctx.* — never call the Sheets API here.
        value = dm.ppt_time[0][1]  # example: read a cell from cached data
        if value != "expected":
            return Finding("WARNING", "PPT time", f"Value is '{value}', expected 'expected'.")
        ctx.ok_count += 1
        return None  # None = check passed
```

3. **Add a fix** (optional) by overriding `fix()`. Queue operations via `wq` — never call the API directly:

```python
    def fix(self, fix_data: dict, wq, dm) -> None:
        # wq.add_value(range_a1, [[value]])        — write a cell value
        # wq.add_structural(copyPaste_request)      — copy formulas
        # wq.add_append(ws, rows)                   — append rows to a tab
        wq.add_value("'PPT time'!B99", [["fixed"]])
        print("  Queued: set B99 = 'fixed'.")
```

   Then include `fix_id` and `fix_data` in the `Finding` you return from `run()`:

```python
        from core.constants import FIX_MY_NEW_CHECK  # add to core/constants.py
        return Finding("WARNING", "PPT time", "...",
                       fix_id=FIX_MY_NEW_CHECK,
                       fix_data={"my_key": "my_value"})
```

4. **Register the fix ID** in `core/constants.py` (add `FIX_MY_NEW_CHECK = 21`). Never renumber existing IDs.

5. **Test locally** — run against a real sheet and confirm it works:
```bash
cd ~/skills/nt-slides-check
poetry run python audit.py <any_sheet_id>
```

6. **Ask before pushing** — only after the check is confirmed working, ask the user: "Should I push this to origin/main?" Then commit and push on confirmation:
```bash
git add checks/check_NN_my_name.py core/constants.py
git commit -m "Add check NN: my check description"
git push
```

All teammates will get the new check automatically on their next run (the `git pull` sync at startup).

### Key constraints

| Rule | Why |
|---|---|
| `run()` reads only `dm.*` and `ctx.*` | DataManager fetches all data in 2 API calls at startup; adding reads inside checks would break the performance model |
| `fix()` writes only to `wq` | WriteQueue dispatches everything in ≤3 API calls; direct writes would bypass batching |
| Never renumber existing `id` values | SKILL.md fix instructions (Fix 1, Fix 6, etc.) reference these by number |
| One class per file, file named `check_NN_*.py` | Auto-discovery relies on this naming convention |
