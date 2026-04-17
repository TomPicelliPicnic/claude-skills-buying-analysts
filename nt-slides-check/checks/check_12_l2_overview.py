from typing import Optional

from core.check_template import CheckTemplate, Finding, AuditContext


class L2OverviewCheck(CheckTemplate):
    id         = 12
    name       = "PPT context L2 overview section"
    sheet_name = "PPT context"
    severity   = "ERROR"

    def run(self, dm, ctx: AuditContext) -> Optional[Finding]:
        non_empty = sum(
            1
            for ri in range(11, 31)
            if ri < len(dm.ppt_ctx)
            for ci in range(5)
            if ci < len(dm.ppt_ctx[ri]) and dm.ppt_ctx[ri][ci].strip()
        )
        if non_empty == 0:
            return Finding("ERROR", "PPT context",
                "The L2 overview section (A12:E31) is completely empty. The KPI table in the slides will be blank.")
        ctx.ok_count += 1
        return None
