from typing import Optional

from core.check_template import CheckTemplate, Finding, AuditContext
from core.constants import FIX_SET_B5_BENCHMARK


class BenchmarkNameCheck(CheckTemplate):
    id         = 13
    name       = "PPT context benchmark name in B5"
    sheet_name = "PPT context"
    severity   = "ERROR"

    def run(self, dm, ctx: AuditContext) -> Optional[Finding]:
        b5   = dm.ppt_ctx[4][1].strip()        if len(dm.ppt_ctx) > 4        and len(dm.ppt_ctx[4]) > 1        else ""
        oi_e6 = dm.offer_insights[5][4].strip() if len(dm.offer_insights) > 5 and len(dm.offer_insights[5]) > 4 else ""

        if b5:
            ctx.ok_count += 1
            return None
        if oi_e6:
            return Finding("ERROR", "PPT context",
                f"Cell B5 is empty. Offer insights E6 suggests: '{oi_e6}'.",
                fix_id=FIX_SET_B5_BENCHMARK,
                fix_data={"fix": "set_b5_benchmark", "suggested": oi_e6})
        return Finding("ERROR", "PPT context",
            "Cell B5 is empty and Offer insights E6 is also empty. Provide a benchmark name manually.",
            fix_id=FIX_SET_B5_BENCHMARK,
            fix_data={"fix": "set_b5_benchmark", "suggested": None})

    def fix(self, fix_data: dict, wq, dm) -> None:
        value = fix_data.get("value") or fix_data.get("suggested") or ""
        wq.add_value("'PPT context'!B5", [[value]])
        print(f"  Queued: set PPT context B5 = '{value}'.")
