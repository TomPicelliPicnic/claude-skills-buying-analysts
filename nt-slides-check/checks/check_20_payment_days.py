from typing import Optional

from core.check_template import CheckTemplate, Finding, AuditContext
from core.constants import FIX_SET_PAYMENT_DAYS
import gspread.utils


class PaymentDaysCheck(CheckTemplate):
    id         = 20
    name       = "Deal sheet payment days"
    sheet_name = "Deal sheet"
    severity   = "ERROR"

    def run(self, dm, ctx: AuditContext) -> Optional[Finding]:
        hdr4 = dm.deal_sheet[3] if len(dm.deal_sheet) > 3 else []
        col_deal = next((i for i, v in enumerate(hdr4) if v.strip().lower() == "deal"), None)

        payment_row = None
        for ri, row in enumerate(dm.deal_sheet):
            if len(row) > 2 and row[2].strip().lower() == "payment days":
                payment_row = ri
                break

        if payment_row is None:
            return Finding("ERROR", "Deal sheet", "Row 'Payment days' not found in col C.")
        if col_deal is None:
            return Finding("ERROR", "Deal sheet",
                "Column 'Deal' not found in row 4. Cannot check payment days.")

        val = dm.deal_sheet[payment_row][col_deal].strip() if len(dm.deal_sheet[payment_row]) > col_deal else ""
        if not val:
            return Finding("ERROR", "Deal sheet",
                f"Payment days (row {payment_row + 1}) is empty.",
                fix_id=FIX_SET_PAYMENT_DAYS,
                fix_data={"fix": "set_payment_days", "row_0idx": payment_row, "col_idx": col_deal})
        ctx.ok_count += 1
        return None

    def fix(self, fix_data: dict, wq, dm) -> None:
        value   = fix_data.get("value", 30)
        row_idx = fix_data.get("row_0idx")
        col_idx = fix_data.get("col_idx")
        if row_idx is None or col_idx is None:
            print("  ERROR: Fix 20 missing row/col indices.")
            return
        cell = gspread.utils.rowcol_to_a1(row_idx + 1, col_idx + 1)
        wq.add_value(f"'Deal sheet'!{cell}", [[value]])
        print(f"  Queued: set Deal sheet {cell} = {value}.")
