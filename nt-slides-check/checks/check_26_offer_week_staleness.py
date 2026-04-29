from datetime import datetime
from typing import Optional

from core.check_template import CheckTemplate, Finding, AuditContext


def _yyyyww_to_date(yyyyww: int):
    y, w = divmod(yyyyww, 100)
    return datetime.strptime(f"{y}-W{w:02d}-1", "%G-W%V-%u").date()


def _weeks_apart(older: int, newer: int) -> int:
    """Return (newer - older) in full ISO weeks. Positive means newer is later."""
    return (_yyyyww_to_date(newer) - _yyyyww_to_date(older)).days // 7


def _max_article_shelf_week(article_shelf_weeks: list) -> Optional[int]:
    """Return the highest valid ISO week number (YYYYWW) from Article shelf col C."""
    max_week = None
    for row in article_shelf_weeks:
        if not row:
            continue
        v = row[0].strip()
        if len(v) == 6 and v.isdigit():
            ww = int(v)
            if 1 <= ww % 100 <= 53:
                if max_week is None or ww > max_week:
                    max_week = ww
    return max_week


class OfferWeekStalenessCheck(CheckTemplate):
    id         = 26
    name       = "Offer week not stale vs Article shelf"
    sheet_name = "TSV output"
    severity   = "WARNING"

    def run(self, dm, ctx: AuditContext) -> Optional[Finding]:
        tsv = dm.tsv
        if not tsv or len(tsv) < 2:
            ctx.ok_count += 1
            return None

        header = tsv[0]
        kow_col = next((i for i, h in enumerate(header) if h.strip() == "Key_offer_week"), None)
        if kow_col is None:
            ctx.ok_count += 1
            return None

        offer_week_raw = tsv[1][kow_col].strip() if kow_col < len(tsv[1]) else ""
        try:
            offer_week = int(offer_week_raw)
        except ValueError:
            ctx.ok_count += 1
            return None

        max_shelf_week = _max_article_shelf_week(dm.article_shelf_weeks)
        if max_shelf_week is None:
            ctx.ok_count += 1
            return None

        lag = _weeks_apart(offer_week, max_shelf_week)
        if lag <= 1:
            ctx.ok_count += 1
            return None

        return Finding(
            "WARNING", "TSV output",
            f"Offer week {offer_week} is {lag} week(s) behind the latest week in "
            f"Article shelf ({max_shelf_week}). Re-run the tool to update the TSV export.",
        )
