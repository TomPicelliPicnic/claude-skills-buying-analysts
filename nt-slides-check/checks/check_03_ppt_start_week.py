from typing import Optional

from core.check_template import CheckTemplate, Finding, AuditContext

FIX_PPT_START_WEEK = 11


class PptStartWeekCheck(CheckTemplate):
    id               = 3
    name             = "PPT time starts at week 202201"
    sheet_name       = "PPT time"
    severity         = "ERROR"
    auto_fix         = True
    auto_fix_message = "Set PPT time B1 = 202201"
    handles_fix_id   = FIX_PPT_START_WEEK

    def run(self, dm, ctx: AuditContext) -> Optional[Finding]:
        b1 = dm.ppt_time[0][1].strip() if len(dm.ppt_time) > 0 and len(dm.ppt_time[0]) > 1 else ""
        if b1 != "202201":
            return Finding("ERROR", "PPT time",
                f"Cell B1 is '{b1}' but must be '202201'. "
                "The time series is misaligned — charts will show incorrect weeks.",
                fix_id=FIX_PPT_START_WEEK, fix_data={})
        ctx.ok_count += 1
        return None

    def fix(self, fix_data: dict, wq, dm) -> None:
        wq.add_value("'PPT time'!B1", [["202201"]])
        print("  Queued: set PPT time B1 = 202201.")
