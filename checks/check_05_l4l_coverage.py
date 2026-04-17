from typing import Optional

from core.check_template import CheckTemplate, Finding, AuditContext


class L4LCoverageCheck(CheckTemplate):
    id         = 5
    name       = "L4L % of net sales LTM"
    sheet_name = "Context"
    severity   = "ERROR"

    def run(self, dm, ctx: AuditContext) -> Optional[Finding]:
        l4l_date    = dm.context[79][1].strip() if len(dm.context) > 79 and len(dm.context[79]) > 1 else ""
        l4l_pct_raw = dm.context[80][4].strip() if len(dm.context) > 80 and len(dm.context[80]) > 4 else ""

        if not l4l_pct_raw:
            return Finding("ERROR", "Context",
                "L4L % of net sales (E81) is empty. Check the L4L date in B80.")
        try:
            pct = float(l4l_pct_raw.replace("%", "").replace(",", ".").strip())
        except ValueError:
            return Finding("ERROR", "Context", f"Could not parse L4L % value: '{l4l_pct_raw}'")

        if pct >= 75:
            ctx.ok_count += 1
            return None
        if pct >= 70:
            return Finding("WARNING", "Context",
                f"L4L covers {pct:.0f}% of net sales LTM (date: {l4l_date}). "
                "Below 75% threshold — consider moving the L4L date forward.")
        return Finding("ERROR", "Context",
            f"L4L covers only {pct:.0f}% of net sales LTM (date: {l4l_date}). "
            "Well below 75% threshold — L4L date must be adjusted.")
