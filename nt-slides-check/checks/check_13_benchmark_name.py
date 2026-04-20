from typing import Optional

from core.check_template import CheckTemplate, Finding, AuditContext
from core.constants import FIX_SET_B5_BENCHMARK


def _col_letter(col_0idx: int) -> str:
    result = ""
    n = col_0idx + 1
    while n:
        n, r = divmod(n - 1, 26)
        result = chr(65 + r) + result
    return result


def _find_benchmark_cell(offer_insights: list) -> tuple[str, str]:
    """
    Scan first 15 rows of Offer insights for a 'Benchmark' column header.
    Returns (cell_ref, value) of the cell immediately below it, e.g. ("E6", "AH").
    Returns ("", "") if not found.
    """
    for ri in range(min(15, len(offer_insights))):
        for ci, val in enumerate(offer_insights[ri]):
            if val.strip().lower() == "benchmark":
                data_ri = ri + 1
                if data_ri < len(offer_insights) and ci < len(offer_insights[data_ri]):
                    cell_ref = f"{_col_letter(ci)}{data_ri + 1}"
                    value = offer_insights[data_ri][ci].strip()
                    return cell_ref, value
    return "", ""


class BenchmarkNameCheck(CheckTemplate):
    id         = 13
    name       = "PPT context benchmark name in B5"
    sheet_name = "PPT context"
    severity   = "ERROR"

    def run(self, dm, ctx: AuditContext) -> Optional[Finding]:
        b5 = dm.ppt_ctx[4][1].strip() if len(dm.ppt_ctx) > 4 and len(dm.ppt_ctx[4]) > 1 else ""

        if b5:
            ctx.ok_count += 1
            return None

        cell_ref, suggested = _find_benchmark_cell(dm.offer_insights)

        if cell_ref:
            return Finding("ERROR", "PPT context",
                f"Cell B5 is empty. Offer insights {cell_ref} suggests: '{suggested}'.",
                fix_id=FIX_SET_B5_BENCHMARK,
                fix_data={"cell_ref": cell_ref, "suggested": suggested})
        return Finding("ERROR", "PPT context",
            "Cell B5 is empty and no 'Benchmark' header found in Offer insights (first 15 rows). Set B5 manually.",
            fix_id=FIX_SET_B5_BENCHMARK,
            fix_data={"cell_ref": None, "suggested": None})

    def fix(self, fix_data: dict, wq, dm) -> None:
        cell_ref = fix_data.get("cell_ref")
        suggested = fix_data.get("suggested") or fix_data.get("value") or ""
        if cell_ref:
            formula = f"='Offer insights'!{cell_ref}"
            wq.add_formula("'PPT context'!B5", [[formula]])
            print(f"  Queued: set PPT context B5 = formula referencing Offer insights {cell_ref} ('{suggested}').")
        elif suggested:
            wq.add_value("'PPT context'!B5", [[suggested]])
            print(f"  Queued: set PPT context B5 = '{suggested}' (static, no Benchmark header found).")
        else:
            print("  Skipped: no cell reference or value available — set B5 manually.")
