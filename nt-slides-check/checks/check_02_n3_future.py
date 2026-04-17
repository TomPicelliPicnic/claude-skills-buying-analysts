from typing import Optional
import gspread.utils

from core.check_template import CheckTemplate, Finding, AuditContext
from core.constants import FIX_BLANK_N3_FUTURE


class N3FutureCheck(CheckTemplate):
    id         = 2
    name       = "N3 line extends past current week"
    sheet_name = "PPT time"
    severity   = "WARNING"

    def run(self, dm, ctx: AuditContext) -> Optional[Finding]:
        week_row   = ctx.week_row
        n3_all_row = dm.ppt_time[3] if len(dm.ppt_time) > 3 else []
        n3_l4l_row = dm.ppt_time[6] if len(dm.ppt_time) > 6 else []

        # Effective cutoff: max of current_keyweek and NLP L4L last data week
        nlp_l4l = dm.ppt_time[7] if len(dm.ppt_time) > 7 else []
        nlp_last = ctx.current_keyweek
        for i, v in enumerate(nlp_l4l):
            if v.strip() and i < len(week_row):
                try:
                    nlp_last = max(nlp_last, int(week_row[i]))
                except ValueError:
                    pass
        cutoff = nlp_last

        future_cols = []
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
                if wk_int > cutoff:
                    future_cols.append(col_i)

        if not future_cols:
            ctx.ok_count += 1
            return None
        return Finding("WARNING", "PPT time",
            f"Net 3 extends {len(future_cols)} week(s) past current week "
            f"(last N3 week: {last_n3_week}, current week: {ctx.current_keyweek}). "
            "The chart line will run into the future.",
            fix_id=FIX_BLANK_N3_FUTURE,
            fix_data={"fix": "blank_n3_future", "cols": future_cols})

    def fix(self, fix_data: dict, wq, dm) -> None:
        cols = fix_data.get("cols", [])
        for col_i in cols:
            wq.add_value(f"'PPT time'!{gspread.utils.rowcol_to_a1(4, col_i + 1)}", [[""]])
            wq.add_value(f"'PPT time'!{gspread.utils.rowcol_to_a1(7, col_i + 1)}", [[""]])
        print(f"  Queued: blank N3 for {len(cols)} week(s) after current week.")
