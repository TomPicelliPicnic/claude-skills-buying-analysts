from typing import Optional

from core.check_template import CheckTemplate, Finding, AuditContext
from checks.check_24_new_listing_names import _EXPECTED as _NEW_PATTERN


class NewListingSellingPriceCheck(CheckTemplate):
    id         = 25
    name       = "New listing selling prices are filled in"
    sheet_name = "TSV output"
    severity   = "ERROR"

    def run(self, dm, ctx: AuditContext) -> Optional[Finding]:
        tsv_header = dm.tsv[0] if dm.tsv else []

        col_id = next(
            (i for i, h in enumerate(tsv_header) if h.strip().lower() == "article id"),
            None,
        )
        col_price = next(
            (i for i, h in enumerate(tsv_header) if h.strip().lower() == "net_selling_price"),
            None,
        )

        if col_id is None or col_price is None:
            missing = []
            if col_id is None:
                missing.append("'Article ID'")
            if col_price is None:
                missing.append("'Net_selling_price'")
            return Finding("ERROR", "TSV output",
                f"Column(s) {', '.join(missing)} not found. Cannot verify new listing selling prices.")

        missing_price = []
        for row in dm.tsv[1:]:
            if not any(row):
                continue
            art_id = row[col_id].strip() if len(row) > col_id else ""
            if not _NEW_PATTERN.match(art_id):
                continue
            price_raw = row[col_price].strip() if len(row) > col_price else ""
            try:
                price = float(price_raw.replace(",", "."))
            except (ValueError, TypeError):
                price = 0.0
            if price == 0.0:
                missing_price.append(art_id)

        if missing_price:
            return Finding("ERROR", "TSV output",
                f"New listing(s) {', '.join(missing_price)} have no Net_selling_price (0 or empty). "
                "Update the selling price in TSV output — new listings are not auto-filled like regular articles.")

        ctx.ok_count += 1
        return None
