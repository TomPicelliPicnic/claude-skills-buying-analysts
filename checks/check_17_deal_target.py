from typing import Optional

from core.check_template import CheckTemplate, Finding, AuditContext

TARGET_ROWS = [(10, 11), (14, 15), (15, 16), (22, 23)]


def _deal_label(deal_sheet, row_0idx):
    row = deal_sheet[row_0idx] if row_0idx < len(deal_sheet) else []
    col_c = row[2].strip() if len(row) > 2 else ""
    if col_c:
        return col_c
    col_d = row[3].strip() if len(row) > 3 else ""
    if col_d and row_0idx > 0:
        prev = deal_sheet[row_0idx - 1] if row_0idx - 1 < len(deal_sheet) else []
        prev_c = prev[2].strip() if len(prev) > 2 else ""
        if prev_c:
            return f"{prev_c} ({col_d})"
    return row[0].strip() if len(row) > 0 else f"Row {row_0idx + 1}"


class DealTargetCheck(CheckTemplate):
    id         = 17
    name       = "Deal sheet target column"
    sheet_name = "Deal sheet"
    severity   = "ERROR"

    def run(self, dm, ctx: AuditContext) -> Optional[Finding]:
        hdr4 = dm.deal_sheet[3] if len(dm.deal_sheet) > 3 else []
        col_target = next((i for i, v in enumerate(hdr4) if v.strip().lower() == "target"), None)

        if col_target is None:
            return Finding("ERROR", "Deal sheet",
                "Column 'Target' not found in row 4. Cannot check target values.")

        empty = []
        for row_0idx, row_1idx in TARGET_ROWS:
            val = dm.deal_sheet[row_0idx][col_target].strip() if len(dm.deal_sheet) > row_0idx and len(dm.deal_sheet[row_0idx]) > col_target else ""
            if not val:
                empty.append(f"Row {row_1idx} ('{_deal_label(dm.deal_sheet, row_0idx)}')")

        if empty:
            return Finding("ERROR", "Deal sheet",
                f"{len(empty)} key target cell(s) are empty:\n" +
                "\n".join(f"  {r}" for r in empty))
        ctx.ok_count += 1
        return None
