from typing import Optional

from core.check_template import CheckTemplate, Finding, AuditContext
from core.constants import FIX_LTM_FIXED_DATES


def _col_letter(col_0idx: int) -> str:
    result = ""
    n = col_0idx + 1
    while n:
        n, r = divmod(n - 1, 26)
        result = chr(65 + r) + result
    return result


def _find_ltm_fixed_dates(shelf_rows: list) -> tuple[str, str, str, str]:
    """
    Scan Shelf analysis rows 1-8, cols > Z (index 26+) for 'LTM Fixed'.
    Returns (startdate_ref, enddate_ref, startdate_val, enddate_val) where
    refs are A1-notation cell references for use in formulas.
    The cell directly above LTM Fixed = enddate, one more above = startdate.
    Returns ("", "", "", "") if not found or dates missing.
    """
    for ri in range(min(8, len(shelf_rows))):
        row = shelf_rows[ri]
        for ci in range(26, len(row)):
            if row[ci].strip().lower() == "ltm fixed":
                end_ri   = ri - 1
                start_ri = ri - 2
                if end_ri < 0 or start_ri < 0:
                    continue
                end_row   = shelf_rows[end_ri]   if end_ri   < len(shelf_rows) else []
                start_row = shelf_rows[start_ri] if start_ri < len(shelf_rows) else []
                end_val   = end_row[ci].strip()   if ci < len(end_row)   else ""
                start_val = start_row[ci].strip() if ci < len(start_row) else ""
                # Skip if values aren't numeric (e.g. label columns like "End week:")
                if not end_val or not start_val:
                    continue
                try:
                    int(start_val)
                    int(end_val)
                except ValueError:
                    continue
                col_ref       = _col_letter(ci)
                end_ref   = f"'Shelf analysis'!{col_ref}{end_ri + 1}"
                start_ref = f"'Shelf analysis'!{col_ref}{start_ri + 1}"
                return start_ref, end_ref, start_val, end_val
    return "", "", "", ""


class LtmFixedDatesCheck(CheckTemplate):
    id         = 22
    name       = "LTM Fixed dates in PPT context C5/D5"
    sheet_name = "PPT context"
    severity   = "ERROR"

    def run(self, dm, ctx: AuditContext) -> Optional[Finding]:
        b5 = dm.ppt_ctx[4][1].strip() if len(dm.ppt_ctx) > 4 and len(dm.ppt_ctx[4]) > 1 else ""

        if b5.strip().lower() != "ltm fixed":
            ctx.ok_count += 1
            return None  # Benchmark is not LTM Fixed — nothing to check

        c5 = dm.ppt_ctx[4][2].strip() if len(dm.ppt_ctx[4]) > 2 else ""
        d5 = dm.ppt_ctx[4][3].strip() if len(dm.ppt_ctx[4]) > 3 else ""

        start_ref, end_ref, start_val, end_val = _find_ltm_fixed_dates(dm.shelf_rows)

        if not start_ref:
            return Finding("WARNING", "PPT context",
                "Benchmark is LTM Fixed but no LTM Fixed dates found in Shelf analysis (cols > Z, rows 1-8).")

        missing = []
        if not c5:
            missing.append(f"C5 (startdate, should be {start_val})")
        if not d5:
            missing.append(f"D5 (enddate, should be {end_val})")

        if not missing:
            ctx.ok_count += 1
            return None

        return Finding("ERROR", "PPT context",
            f"Benchmark is LTM Fixed but {' and '.join(missing)} is empty. "
            f"Add start and enddate of LTM Fixed to ensure the benchmark line appears in the graph.",
            fix_id=FIX_LTM_FIXED_DATES,
            fix_data={"start_ref": start_ref, "end_ref": end_ref,
                      "start_val": start_val, "end_val": end_val,
                      "c5_empty": not c5, "d5_empty": not d5})

    def fix(self, fix_data: dict, wq, dm) -> None:
        print("  Add start and enddate of LTM fixed to ensure benchmark line in graph.")
        if fix_data.get("c5_empty"):
            wq.add_formula("'PPT context'!C5", [[f"={fix_data['start_ref']}"]])
            print(f"  Queued: set PPT context C5 = {fix_data['start_ref']} ('{fix_data['start_val']}').")
        if fix_data.get("d5_empty"):
            wq.add_formula("'PPT context'!D5", [[f"={fix_data['end_ref']}"]])
            print(f"  Queued: set PPT context D5 = {fix_data['end_ref']} ('{fix_data['end_val']}').")
