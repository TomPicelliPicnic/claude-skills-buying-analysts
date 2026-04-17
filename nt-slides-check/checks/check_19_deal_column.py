from typing import Optional

from core.check_template import CheckTemplate, Finding, AuditContext
from checks.check_17_deal_target import _deal_label


class DealColumnCheck(CheckTemplate):
    id         = 19
    name       = "Deal sheet deal column"
    sheet_name = "Deal sheet"
    severity   = "WARNING"

    def run(self, dm, ctx: AuditContext) -> Optional[Finding]:
        offers_type = ctx.offers_type
        hdr4 = dm.deal_sheet[3] if len(dm.deal_sheet) > 3 else []
        col_deal = next((i for i, v in enumerate(hdr4) if v.strip().lower() == "deal"), None)

        if not offers_type:
            ctx.ok_count += 1
            return None
        if col_deal is None:
            return Finding("ERROR", "Deal sheet",
                "Column 'Deal' not found in row 4. Cannot check deal values.")

        if "promo" in offers_type.lower() and "shelf" in offers_type.lower():
            rows = list(range(5, 22))
        elif "shelf" in offers_type.lower():
            rows = list(range(5, 11))
        else:
            ctx.ok_count += 1
            return None

        empty = []
        for row_0idx in rows:
            val = dm.deal_sheet[row_0idx][col_deal].strip() if len(dm.deal_sheet) > row_0idx and len(dm.deal_sheet[row_0idx]) > col_deal else ""
            if not val:
                empty.append(f"Row {row_0idx + 1} ('{_deal_label(dm.deal_sheet, row_0idx)}')")

        if empty:
            return Finding("WARNING", "Deal sheet",
                f"Deal column missing for {len(empty)} row(s) (offer type: {offers_type}):\n" +
                "\n".join(f"  {r}" for r in empty))
        ctx.ok_count += 1
        return None
