from typing import Optional

from core.check_template import CheckTemplate, Finding, AuditContext
from core.constants import FIX_FILL_METRIC_GAPS


class PptContinuityCheck(CheckTemplate):
    id         = 8
    name       = "PPT time data continuity"
    sheet_name = "PPT time"
    severity   = "WARNING"

    def run(self, dm, ctx: AuditContext) -> Optional[Finding]:
        cur_col = ctx.cur_col
        if cur_col is None:
            ctx.ok_count += 1
            return None

        week_row = ctx.week_row
        gap_rows = []

        for row_idx in range(3, len(dm.ppt_time)):
            row_data = dm.ppt_time[row_idx]
            fml_row  = dm.ppt_formulas[row_idx] if row_idx < len(dm.ppt_formulas) else []
            row_label = row_data[0].strip() if row_data else ""

            def cell_present(c, _rd=row_data, _fr=fml_row):
                has_val = c < len(_rd) and _rd[c].strip()
                has_fml = c < len(_fr) and str(_fr[c]).startswith("=")
                return bool(has_val or has_fml)

            if not any(cell_present(c) for c in range(1, cur_col + 1)):
                continue

            last_filled = None
            for c in range(cur_col, 0, -1):
                if cell_present(c):
                    last_filled = c
                    break

            if last_filled is None or last_filled >= cur_col:
                continue

            last_week = None
            if last_filled < len(week_row):
                try:
                    last_week = int(week_row[last_filled])
                except (ValueError, TypeError):
                    pass

            gap_rows.append({
                "row_name": row_label or f"Row {row_idx + 1}",
                "row_0idx": row_idx,
                "last_filled_col": last_filled,
                "last_week": last_week,
                "gap_weeks": cur_col - last_filled,
            })

        if not gap_rows:
            ctx.ok_count += 1
            return None

        lines = [
            f"  {g['row_name']}: last data week {g['last_week']}, {g['gap_weeks']} week(s) missing"
            for g in gap_rows
        ]
        return Finding("WARNING", "PPT time",
            f"{len(gap_rows)} row(s) have gaps before week {ctx.current_keyweek}:\n" + "\n".join(lines),
            fix_id=FIX_FILL_METRIC_GAPS,
            fix_data={"fix": "fill_metric_gaps", "gaps": gap_rows, "cur_col": cur_col})

    def fix(self, fix_data: dict, wq, dm) -> None:
        cur_col = fix_data.get("cur_col")
        for g in fix_data.get("gaps", []):
            lfc = g.get("last_filled_col")
            if lfc is None:
                continue
            wq.add_structural({
                "copyPaste": {
                    "source": {
                        "sheetId": dm.ws_ppt.id,
                        "startRowIndex": g["row_0idx"], "endRowIndex": g["row_0idx"] + 1,
                        "startColumnIndex": lfc, "endColumnIndex": lfc + 1,
                    },
                    "destination": {
                        "sheetId": dm.ws_ppt.id,
                        "startRowIndex": g["row_0idx"], "endRowIndex": g["row_0idx"] + 1,
                        "startColumnIndex": lfc + 1, "endColumnIndex": cur_col + 1,
                    },
                    "pasteType": "PASTE_FORMULA", "pasteOrientation": "NORMAL",
                }
            })
        print(f"  Queued: re-extend {len(fix_data.get('gaps', []))} row(s) to current week.")
