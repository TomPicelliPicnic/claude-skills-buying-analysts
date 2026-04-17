from typing import Optional

from core.check_template import CheckTemplate, Finding, AuditContext


class CurrentWeekDealPointCheck(CheckTemplate):
    """Auto-fix: current week must never carry a dealpoint. Clears it immediately."""
    id         = 10
    name       = "Current week dealpoint (auto-fix)"
    sheet_name = "PPT time"
    severity   = "ERROR"

    def run(self, dm, ctx: AuditContext) -> Optional[Finding]:
        cur_col = ctx.cur_col
        if cur_col is None:
            ctx.ok_count += 1
            return None

        dp_val = ctx.dealpoint_row[cur_col].strip() if cur_col < len(ctx.dealpoint_row) else ""
        if dp_val != "TRUE":
            ctx.ok_count += 1
            return None

        try:
            col_1idx = dm.ppt_time[0].index(str(ctx.current_keyweek)) + 1
            dm.ws_ppt.update_cell(3, col_1idx, "")
            ctx.auto_fixed.append(f"Cleared dealpoint for current week {ctx.current_keyweek}")
        except (ValueError, Exception) as e:
            return Finding("ERROR", "PPT time",
                f"Dealpoint is TRUE for current week ({ctx.current_keyweek}) and auto-clear failed: {e}")
        ctx.ok_count += 1
        return None
