from typing import Optional

from core.check_template import CheckTemplate, Finding, AuditContext
from core.constants import FIX_L2_L3_CATEGORY


def _col_letter(col_0idx: int) -> str:
    result = ""
    n = col_0idx + 1
    while n:
        n, r = divmod(n - 1, 26)
        result = chr(65 + r) + result
    return result


def _find_l2_l3_from_context(context: list) -> dict:
    """
    Locate 'Margin comparison: Brand in L2' anchor in Context tab and derive
    cell references for B8/D8/F8 in PPT context.

    Returns dict with keys:
        b8_ref, d8_ref, f8_ref  — A1-notation Context refs (empty string if absent)
        b8_val, d8_val, f8_val  — display values for messages
        tool                    — 'old', 'new', or 'unknown'

    Old tool layout:
      Row after anchor contains 'L2' label at some col, chosen value to its right.
      b8_ref → the 'L2' label cell, d8_ref → the value cell to its right.

    New tool layout:
      'Select: L2 or L3' label → value in next row same col → b8_ref
      'Select: specific L2' label → value in next row same col → d8_ref
      'Select: specific L3' label → value in next row same col → f8_ref
    """
    anchor_ri = None
    for ri, row in enumerate(context):
        for val in row:
            if "margin comparison" in val.strip().lower() and "brand in l2" in val.strip().lower():
                anchor_ri = ri
                break
        if anchor_ri is not None:
            break

    if anchor_ri is None:
        return None

    search_end = min(anchor_ri + 30, len(context))
    result = {"b8_ref": "", "d8_ref": "", "f8_ref": "",
              "b8_val": "", "d8_val": "", "f8_val": "", "tool": "unknown"}

    # New tool: look for 'Select:' markers first
    new_tool_found = False
    for ri in range(anchor_ri + 1, search_end):
        row = context[ri]
        for ci, val in enumerate(row):
            v = val.strip().lower()
            if "select: l2 or l3" in v:
                new_tool_found = True
                data_ri = ri + 1
                if data_ri < len(context) and ci < len(context[data_ri]):
                    result["b8_ref"] = f"Context!{_col_letter(ci)}{data_ri + 1}"
                    result["b8_val"] = context[data_ri][ci].strip()
            elif "select: specific l2" in v:
                new_tool_found = True
                data_ri = ri + 1
                if data_ri < len(context) and ci < len(context[data_ri]):
                    result["d8_ref"] = f"Context!{_col_letter(ci)}{data_ri + 1}"
                    result["d8_val"] = context[data_ri][ci].strip()
            elif "select: specific l3" in v:
                new_tool_found = True
                data_ri = ri + 1
                if data_ri < len(context) and ci < len(context[data_ri]):
                    result["f8_ref"] = f"Context!{_col_letter(ci)}{data_ri + 1}"
                    result["f8_val"] = context[data_ri][ci].strip()

    if new_tool_found:
        result["tool"] = "new"
        return result

    # Old tool: find cell with value exactly 'L2'; chosen category is one col to the right
    for ri in range(anchor_ri + 1, search_end):
        row = context[ri]
        for ci, val in enumerate(row):
            if val.strip() == "L2" and ci + 1 < len(row):
                col = _col_letter(ci)
                col_next = _col_letter(ci + 1)
                sheet_row = ri + 1
                result["b8_ref"] = f"Context!{col}{sheet_row}"
                result["b8_val"] = "L2"
                result["d8_ref"] = f"Context!{col_next}{sheet_row}"
                result["d8_val"] = row[ci + 1].strip()
                result["tool"] = "old"
                return result

    return result


class L2L3CategoryCheck(CheckTemplate):
    id         = 23
    name       = "L2/L3 category labels in PPT context A8:F8"
    sheet_name = "PPT context"
    severity   = "ERROR"

    _STATIC = {0: "L2 or L3", 2: "L2:", 4: "L3:"}

    def run(self, dm, ctx: AuditContext) -> Optional[Finding]:
        ppt_ctx = dm.ppt_ctx
        row8 = ppt_ctx[7] if len(ppt_ctx) > 7 else []

        def _cell(col):
            return row8[col].strip() if col < len(row8) else ""

        static_issues = []
        for col, expected in self._STATIC.items():
            actual = _cell(col)
            if actual != expected:
                static_issues.append(f"{chr(65 + col)}8 should be '{expected}' (is '{actual}')")

        dynamic_issues = []
        if not _cell(1):
            dynamic_issues.append("B8 (L2 or L3 selector) is empty")
        if not _cell(3):
            dynamic_issues.append("D8 (specific L2 value) is empty")

        all_issues = static_issues + dynamic_issues
        if not all_issues:
            ctx.ok_count += 1
            return None

        ctx_info = _find_l2_l3_from_context(dm.context)
        fix_available = ctx_info is not None and ctx_info.get("tool") != "unknown"

        msg = (
            f"PPT context row 8 (L2/L3 category) has {len(all_issues)} issue(s): "
            + "; ".join(all_issues) + "."
        )
        if fix_available:
            msg += (f" Suggested: B8='{ctx_info['b8_val']}', D8='{ctx_info['d8_val']}'"
                    f" (tool: {ctx_info['tool']})")
        else:
            msg += " Set A8:F8 manually."

        return Finding(
            "ERROR", "PPT context", msg,
            fix_id=FIX_L2_L3_CATEGORY if fix_available else None,
            fix_data=ctx_info if fix_available else None,
        )

    def fix(self, fix_data: dict, wq, dm) -> None:
        tool = fix_data.get("tool", "unknown")
        print(f"  Setting PPT context row 8 labels (tool: {tool}).")

        wq.add_value("'PPT context'!A8", [["L2 or L3"]])
        print("  Queued: A8 = 'L2 or L3'.")
        wq.add_value("'PPT context'!C8", [["L2:"]])
        print("  Queued: C8 = 'L2:'.")
        wq.add_value("'PPT context'!E8", [["L3:"]])
        print("  Queued: E8 = 'L3:'.")

        b8_ref = fix_data.get("b8_ref") or ""
        d8_ref = fix_data.get("d8_ref") or ""
        f8_ref = fix_data.get("f8_ref") or ""

        if b8_ref:
            wq.add_formula("'PPT context'!B8", [[f"={b8_ref}"]])
            print(f"  Queued: B8 = formula ={b8_ref} ('{fix_data.get('b8_val', '')}').")
        else:
            print("  Skipped B8: no Context cell reference found.")

        if d8_ref:
            wq.add_formula("'PPT context'!D8", [[f"={d8_ref}"]])
            print(f"  Queued: D8 = formula ={d8_ref} ('{fix_data.get('d8_val', '')}').")
        else:
            print("  Skipped D8: no Context cell reference found.")

        if f8_ref:
            wq.add_formula("'PPT context'!F8", [[f"={f8_ref}"]])
            print(f"  Queued: F8 = formula ={f8_ref} ('{fix_data.get('f8_val', '')}').")
        else:
            print("  Skipped F8: no L3 reference (expected for old tool or L2 choice).")
