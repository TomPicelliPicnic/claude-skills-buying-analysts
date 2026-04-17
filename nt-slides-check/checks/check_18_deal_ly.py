from typing import Optional

from core.check_template import CheckTemplate, Finding, AuditContext
from checks.check_17_deal_target import _deal_label


class DealLYCheck(CheckTemplate):
    id         = 18
    name       = "Deal sheet LY benchmark"
    sheet_name = "Deal sheet"
    severity   = "WARNING"

    def run(self, dm, ctx: AuditContext) -> Optional[Finding]:
        offers_type = ctx.offers_type
        hdr4 = dm.deal_sheet[3] if len(dm.deal_sheet) > 3 else []
        col_ly = next((i for i, v in enumerate(hdr4) if v.strip().lower() == "ly"), None)

        if not offers_type:
            return Finding("ERROR", "Deal sheet",
                "Offer type not found in rows 1–2 — cannot determine which rows to check.")
        if col_ly is None:
            return Finding("ERROR", "Deal sheet",
                "Column 'LY' not found in row 4. Cannot check LY benchmark.")

        if "promo" in offers_type.lower() and "shelf" in offers_type.lower():
            rows = list(range(5, 22))
        elif "shelf" in offers_type.lower():
            rows = list(range(5, 11))
        else:
            ctx.ok_count += 1
            return None

        empty = []
        for row_0idx in rows:
            val = dm.deal_sheet[row_0idx][col_ly].strip() if len(dm.deal_sheet) > row_0idx and len(dm.deal_sheet[row_0idx]) > col_ly else ""
            if not val:
                empty.append(f"Row {row_0idx + 1} ('{_deal_label(dm.deal_sheet, row_0idx)}')")

        if empty:
            return Finding("WARNING", "Deal sheet",
                f"LY benchmark missing for {len(empty)} row(s) (offer type: {offers_type}):\n" +
                "\n".join(f"  {r}" for r in empty))
        ctx.ok_count += 1
        return None
