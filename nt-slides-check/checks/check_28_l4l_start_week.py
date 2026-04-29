from typing import Optional

from core.check_template import CheckTemplate, Finding, AuditContext
from checks.check_26_offer_week_staleness import _weeks_apart


def _get_tsv_col(tsv: list, col_name: str) -> Optional[str]:
    """Return the first data row value for a TSV column by header name."""
    if not tsv or len(tsv) < 2:
        return None
    header = tsv[0]
    for ci, h in enumerate(header):
        if h.strip() == col_name:
            return tsv[1][ci].strip() if ci < len(tsv[1]) else ""
    return None


class L4LStartWeekCheck(CheckTemplate):
    id         = 28
    name       = "L4L start week valid (≥52 weeks before offer, before deal week)"
    sheet_name = "Context"
    severity   = "WARNING"

    def run(self, dm, ctx: AuditContext) -> Optional[Finding]:
        # L4L start week from Context B80
        l4l_raw = dm.context[79][1].strip() if len(dm.context) > 79 and len(dm.context[79]) > 1 else ""
        if not l4l_raw or not l4l_raw.isdigit() or len(l4l_raw) != 6:
            ctx.ok_count += 1
            return None

        l4l_week = int(l4l_raw)

        offer_raw = _get_tsv_col(dm.tsv, "Key_offer_week")
        deal_raw  = _get_tsv_col(dm.tsv, "Key_deal_week")

        try:
            offer_week = int(offer_raw) if offer_raw else None
        except ValueError:
            offer_week = None

        try:
            deal_week = int(deal_raw) if deal_raw else None
        except ValueError:
            deal_week = None

        issues = []

        if offer_week:
            lag = _weeks_apart(l4l_week, offer_week)
            if lag < 52:
                issues.append(
                    f"L4L start week {l4l_week} is only {lag} week(s) before offer week "
                    f"{offer_week} — should be at least 52 weeks (1 year) to ensure "
                    f"sufficient LTM coverage."
                )

        if deal_week:
            if l4l_week >= deal_week:
                issues.append(
                    f"L4L start week {l4l_week} is not before deal week {deal_week} — "
                    f"the L4L period must start before the deal."
                )

        if not issues:
            ctx.ok_count += 1
            return None

        return Finding("WARNING", "Context", " ".join(issues))
