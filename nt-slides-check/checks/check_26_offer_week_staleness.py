import re
from datetime import datetime
from typing import Optional

from core.check_template import CheckTemplate, Finding, AuditContext


def _yyyyww_to_date(yyyyww: int):
    y, w = divmod(yyyyww, 100)
    return datetime.strptime(f"{y}-W{w:02d}-1", "%G-W%V-%u").date()


def _weeks_apart(older: int, newer: int) -> int:
    """Return (newer - older) in full ISO weeks. Positive means newer is later."""
    return (_yyyyww_to_date(newer) - _yyyyww_to_date(older)).days // 7


def _find_max_shelf_week(shelf_rows: list) -> Optional[int]:
    """Return the highest valid ISO week number (YYYYWW) found in Shelf analysis rows."""
    max_week = None
    for row in shelf_rows:
        for val in row:
            v = val.strip()
            if re.fullmatch(r'20[2-9]\d{3}', v):
                try:
                    ww = int(v)
                    if 1 <= ww % 100 <= 53:
                        if max_week is None or ww > max_week:
                            max_week = ww
                except ValueError:
                    pass
    return max_week


class OfferWeekStalenessCheck(CheckTemplate):
    id         = 26
    name       = "Offer week not stale vs Shelf analysis"
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

        max_shelf_week = _find_max_shelf_week(dm.shelf_rows)
        if max_shelf_week is None:
            ctx.ok_count += 1
            return None

        lag = _weeks_apart(offer_week, max_shelf_week)
        if lag <= 1:
            ctx.ok_count += 1
            return None

        return Finding(
            "WARNING", "TSV output",
            f"Offer week {offer_week} is {lag} weeks behind the latest week in "
            f"Shelf analysis ({max_shelf_week}). The TSV may be based on an outdated "
            f"export — re-run the tool to update.",
        )
