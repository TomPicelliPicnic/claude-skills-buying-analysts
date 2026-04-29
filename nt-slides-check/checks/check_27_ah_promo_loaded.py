from typing import Optional

from core.check_template import CheckTemplate, Finding, AuditContext


def _find_col(header_row: list, label: str) -> Optional[int]:
    """Return 0-based column index of the first cell matching label (case-insensitive)."""
    for ci, val in enumerate(header_row):
        if val.strip().lower() == label.strip().lower():
            return ci
    return None


def _find_header_row(rows: list, *labels) -> Optional[int]:
    """Return 0-based row index of the first row containing ALL of the given labels."""
    for ri, row in enumerate(rows):
        row_vals = [v.strip().lower() for v in row]
        if all(label.strip().lower() in row_vals for label in labels):
            return ri
    return None


class AhPromoLoadedCheck(CheckTemplate):
    id         = 27
    name       = "AH-Promo plan loaded (Price match filled for all GTINs)"
    sheet_name = "AH-Promo"
    severity   = "WARNING"

    def run(self, dm, ctx: AuditContext) -> Optional[Finding]:
        rows = dm.ah_promo
        if not rows:
            ctx.ok_count += 1
            return None  # Tab absent — AH not a comparator for this offer

        header_ri = _find_header_row(rows, "CU GTIN", "Price match")
        if header_ri is None:
            ctx.ok_count += 1
            return None  # Unexpected layout — skip silently

        header = rows[header_ri]
        gtin_col  = _find_col(header, "CU GTIN")
        price_col = _find_col(header, "Price match")

        if gtin_col is None or price_col is None:
            ctx.ok_count += 1
            return None

        missing = 0
        total   = 0
        for row in rows[header_ri + 1:]:
            gtin = row[gtin_col].strip() if gtin_col < len(row) else ""
            if not gtin:
                continue
            total += 1
            price_match = row[price_col].strip() if price_col < len(row) else ""
            if not price_match:
                missing += 1

        if missing == 0:
            ctx.ok_count += 1
            return None

        return Finding(
            "WARNING", "AH-Promo",
            f"{missing} of {total} article(s) have a CU GTIN but no Price match entry "
            f"in AH-Promo. Load the AH promo plan to populate the Price match column.",
        )
