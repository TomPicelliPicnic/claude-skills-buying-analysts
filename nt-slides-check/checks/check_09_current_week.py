from typing import Optional

from core.check_template import CheckTemplate, Finding, AuditContext


class CurrentWeekCheck(CheckTemplate):
    id         = 9
    name       = "PPT time current week marker"
    sheet_name = "PPT time"
    severity   = "ERROR"

    def run(self, dm, ctx: AuditContext) -> Optional[Finding]:
        if ctx.cur_col is None:
            ctx.ok_count += 1
            return None

        nlp_l4l = dm.ppt_time[7] if len(dm.ppt_time) > 7 else []
        future_cols = [
            i for i in range(ctx.cur_col + 1, len(nlp_l4l))
            if nlp_l4l[i].strip()
        ]
        if not future_cols:
            ctx.ok_count += 1
            return None

        last_wk = ctx.week_row[future_cols[-1]] if future_cols[-1] < len(ctx.week_row) else "?"
        return Finding("ERROR", "PPT time",
            f"NLP - L4L has data up to week {last_wk} but current week is set to {ctx.current_keyweek}. "
            "Extend the TRUE range in PPT time row 3 to the actual current week.")
