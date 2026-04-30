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
        b1_val     = dm.ppt_time[0][1].strip()         if len(dm.ppt_time)    > 0 and len(dm.ppt_time[0])    > 1 else ""
        b1_formula = str(dm.ppt_formulas[0][1]).strip() if len(dm.ppt_formulas) > 0 and len(dm.ppt_formulas[0]) > 1 else ""
        # Flag if value is wrong OR if value looks right but is stored as text (formula starts with ')
        if b1_val != "202201" or b1_formula.startswith("'"):
            display = f"'{b1_val}" if b1_formula.startswith("'") else b1_val
            return Finding("ERROR", "PPT time",
                f"Cell B1 is '{display}' but must be numeric 202201. "
                "Storing it as text breaks the TRANSPOSE formula in row 1.",
                fix_id=FIX_PPT_START_WEEK, fix_data={})
        ctx.ok_count += 1
        return None

    def fix(self, fix_data: dict, wq, dm) -> None:
        wq.add_formula("'PPT time'!B1", [[202201]])
        print("  Queued: set PPT time B1 = 202201.")
