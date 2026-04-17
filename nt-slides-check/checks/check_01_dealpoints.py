from typing import Optional
from datetime import date, timedelta
import gspread.utils

from core.check_template import CheckTemplate, Finding, AuditContext
from core.constants import FIX_CLEAR_DEALPOINTS


class DealPointsCheck(CheckTemplate):
    id         = 1
    name       = "Dealpoints"
    sheet_name = "PPT time"
    severity   = "WARNING"

    def run(self, dm, ctx: AuditContext) -> Optional[Finding]:
        week_row      = ctx.week_row
        dealpoint_row = ctx.dealpoint_row
        total_articles = len([r for r in dm.tsv[1:] if any(r)])

        pl_header = dm.price_list[0] if dm.price_list else []
        kw_col = next((i for i, h in enumerate(pl_header) if "key week valid from" in h.lower()), None)

        if kw_col is None:
            return Finding("ERROR", "Price list",
                "Column 'Key week valid from' not found. Cannot check dealpoints.")

        week_counts: dict[str, int] = {}
        for row in dm.price_list[1:]:
            if not any(row):
                continue
            wk = row[kw_col].strip()
            if wk:
                week_counts[wk] = week_counts.get(wk, 0) + 1

        dealpoint_weeks_found = []
        suspicious = []
        for col_i, dp_val in enumerate(dealpoint_row):
            if dp_val != "TRUE":
                continue
            wk = week_row[col_i].strip() if col_i < len(week_row) else ""
            if not wk:
                continue
            dealpoint_weeks_found.append(wk)
            if wk == str(ctx.current_keyweek):
                continue
            count = week_counts.get(wk, 0)
            if count < total_articles * 0.4:
                suspicious.append({"week": wk, "count": count, "total": total_articles, "col_index": col_i})

        if not suspicious:
            ctx.ok_count += 1
            return None

        lines = []
        for d in suspicious:
            pct = round(100 * d["count"] / d["total"]) if d["total"] else 0
            lines.append(f"  Week {d['week']}: {d['count']} of {d['total']} articles ({pct}%) — likely not a real dealpoint")
        valid = [w for w in dealpoint_weeks_found if w not in {d["week"] for d in suspicious}]
        msg = f"{len(suspicious)} suspicious dealpoint(s):\n" + "\n".join(lines)
        if valid:
            msg += f"\n  Valid dealpoints (>= 40%): {', '.join(valid)}"
        return Finding("WARNING", "PPT time", msg,
                       fix_id=FIX_CLEAR_DEALPOINTS,
                       fix_data={"fix": "clear_dealpoints", "suspicious": suspicious})

    def fix(self, fix_data: dict, wq, dm) -> None:
        for d in fix_data.get("suspicious", []):
            wk = str(d["week"])
            try:
                col_1idx = dm.ppt_time[0].index(wk) + 1
            except ValueError:
                print(f"  Week {wk} not found in cached row 1 — skipping.")
                continue
            wq.add_value(f"'PPT time'!{gspread.utils.rowcol_to_a1(3, col_1idx)}", [[""]])
            print(f"  Queued: clear dealpoint week {wk} (col {col_1idx}).")
