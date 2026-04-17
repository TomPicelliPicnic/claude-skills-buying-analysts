from typing import Optional
from datetime import date, timedelta

from core.check_template import CheckTemplate, Finding, AuditContext
from core.constants import FIX_EXTEND_DATES

TARGET_WEEK = 202652
COL_JA_0IDX = 260   # column JA, 0-indexed


class DatesExtensionCheck(CheckTemplate):
    id         = 4
    name       = "PPT time full timeline to 202652"
    sheet_name = "PPT time"
    severity   = "ERROR"

    def run(self, dm, ctx: AuditContext) -> Optional[Finding]:
        week_row_ints = []
        for v in ctx.week_row[1:]:
            try:
                week_row_ints.append(int(v))
            except (ValueError, TypeError):
                week_row_ints.append(None)

        last_week = max((w for w in week_row_ints if w), default=None)
        if last_week is None:
            return Finding("ERROR", "PPT time", "No week numbers found in PPT time row 1.")
        if last_week >= TARGET_WEEK:
            ctx.ok_count += 1
            return None
        return Finding("ERROR", "PPT time",
            f"Row 1 ends at week {last_week}, missing weeks up to {TARGET_WEEK}. "
            "Dates tab needs extending and PPT time formulas need copying to column JA.",
            fix_id=FIX_EXTEND_DATES,
            fix_data={"fix": "extend_dates_and_ppt_time", "last_week": last_week})

    def fix(self, fix_data: dict, wq, dm) -> None:
        ws_dates  = dm.sh.worksheet("Dates")
        dates_all = ws_dates.get_all_values()
        last_row  = len(dates_all)
        last_week = int(dates_all[-1][1])

        rows_to_add = []
        current = last_week
        while True:
            nxt = _next_iso_week(current)
            if nxt > TARGET_WEEK:
                break
            yr  = nxt // 100
            wk  = nxt % 100
            per = int(f"{yr}{(wk - 1) // 4 + 1:02d}")
            n   = last_row + len(rows_to_add) + 1
            rows_to_add.append([
                yr, nxt, per,
                f"=IF(C{n}=C{n-1},,RIGHT(C{n},2))",
                f'=IF(D{n}<>"","P"&VALUE(D{n}),)',
                f"=B{n}>'PPT time'!$B$1",
            ])
            current = nxt

        if not rows_to_add:
            print("  Dates already extends to 202652 — skipping formula extension.")
            return

        wq.add_append(ws_dates, rows_to_add)
        print(f"  Queued: extend Dates by {len(rows_to_add)} rows (up to 202652).")

        last_col_0 = 0
        for row in dm.ppt_time[2:]:
            for c in range(len(row) - 1, -1, -1):
                if row[c].strip():
                    last_col_0 = max(last_col_0, c)
                    break

        max_row = len(dm.ppt_time)

        if dm.ws_ppt.col_count <= COL_JA_0IDX:
            cols_needed = COL_JA_0IDX + 1 - dm.ws_ppt.col_count
            dm.ws_ppt.add_cols(cols_needed)
            print(f"  Expanded PPT time by {cols_needed} column(s) to reach JA.")

        wq.add_structural({
            "copyPaste": {
                "source": {
                    "sheetId": dm.ws_ppt.id,
                    "startRowIndex": 2, "endRowIndex": max_row,
                    "startColumnIndex": last_col_0, "endColumnIndex": last_col_0 + 1,
                },
                "destination": {
                    "sheetId": dm.ws_ppt.id,
                    "startRowIndex": 2, "endRowIndex": max_row,
                    "startColumnIndex": last_col_0 + 1, "endColumnIndex": COL_JA_0IDX + 1,
                },
                "pasteType": "PASTE_FORMULA", "pasteOrientation": "NORMAL",
            }
        })
        print(f"  Queued: extend PPT time formulas from col {last_col_0} to JA ({max_row} data rows).")


def _next_iso_week(yyyyww: int) -> int:
    year, week = yyyyww // 100, yyyyww % 100
    jan4    = date(year, 1, 4)
    monday  = jan4 - timedelta(days=jan4.weekday()) + timedelta(weeks=week - 1)
    nxt     = monday + timedelta(weeks=1)
    y, w, _ = nxt.isocalendar()
    return int(f"{y}{w:02d}")
