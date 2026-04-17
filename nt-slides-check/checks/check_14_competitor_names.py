from typing import Optional

from core.check_template import CheckTemplate, Finding, AuditContext


class CompetitorNamesCheck(CheckTemplate):
    id         = 14
    name       = "PPT context competitor names B35 vs N35"
    sheet_name = "PPT context"
    severity   = "ERROR"

    def run(self, dm, ctx: AuditContext) -> Optional[Finding]:
        b35 = dm.ppt_ctx[34][1].strip()  if len(dm.ppt_ctx) > 34 and len(dm.ppt_ctx[34]) > 1  else ""
        n35 = dm.ppt_ctx[34][13].strip() if len(dm.ppt_ctx) > 34 and len(dm.ppt_ctx[34]) > 13 else ""

        if not b35 and not n35:
            return Finding("WARNING", "PPT context",
                "Both B35 and N35 are empty. Set competitor names for the competitive analysis.")
        if b35 == n35:
            return Finding("ERROR", "PPT context",
                f"B35 and N35 have the same value ('{b35}'). They must be different competitors. "
                "Fix manually in Offer insights — do not overwrite B35/N35 directly.")
        ctx.ok_count += 1
        return None
