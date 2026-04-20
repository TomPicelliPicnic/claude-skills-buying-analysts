from typing import Optional

from core.check_template import CheckTemplate, Finding, AuditContext
from core.constants import FIX_L2_L3_CATEGORY


def _find_l2_l3_from_context(context: list) -> dict:
    """
    Locate 'Margin comparison: Brand in L2' anchor in Context tab and derive
    the B8/D8/F8 values for PPT context.

    Old tool layout (after anchor):
      - A cell with value 'L2' → value to its right is the chosen L2 category.
      - B8 = 'L2', D8 = chosen L2 value, F8 = ''

    New tool layout (after anchor):
      - Row with 'Select: L2 or L3' → next row col = B8 value ('L2' or 'L3')
      - Row with 'Select: specific L2' → next row col = D8 value
      - Row with 'Select: specific L3' → next row col = F8 value

    Returns dict with keys: b8, d8, f8, tool ('old'/'new'/'unknown')
    Returns None if anchor not found.
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

    # Search within the next 30 rows
    search_end = min(anchor_ri + 30, len(context))
    result = {"b8": "", "d8": "", "f8": "", "tool": "unknown"}

    # Check for new tool markers first
    new_tool_found = False
    for ri in range(anchor_ri + 1, search_end):
        row = context[ri]
        for ci, val in enumerate(row):
            v = val.strip().lower()
            if "select: l2 or l3" in v:
                new_tool_found = True
                # Value is in the next row, same column
                if ri + 1 < len(context) and ci < len(context[ri + 1]):
                    result["b8"] = context[ri + 1][ci].strip()
            elif "select: specific l2" in v:
                new_tool_found = True
                if ri + 1 < len(context) and ci < len(context[ri + 1]):
                    result["d8"] = context[ri + 1][ci].strip()
            elif "select: specific l3" in v:
                new_tool_found = True
                if ri + 1 < len(context) and ci < len(context[ri + 1]):
                    result["f8"] = context[ri + 1][ci].strip()

    if new_tool_found:
        result["tool"] = "new"
        return result

    # Old tool: find a cell whose value is exactly 'L2'; value to its right is the chosen category
    for ri in range(anchor_ri + 1, search_end):
        row = context[ri]
        for ci, val in enumerate(row):
            if val.strip() == "L2":
                if ci + 1 < len(row):
                    chosen = row[ci + 1].strip()
                    result["b8"] = "L2"
                    result["d8"] = chosen
                    result["f8"] = ""
                    result["tool"] = "old"
                    return result

    return result


class L2L3CategoryCheck(CheckTemplate):
    id         = 23
    name       = "L2/L3 category labels in PPT context A8:F8"
    sheet_name = "PPT context"
    severity   = "ERROR"

    # Expected static labels in PPT context row 8 (0-indexed row 7)
    _STATIC = {0: "L2 or L3", 2: "L2:", 4: "L3:"}

    def run(self, dm, ctx: AuditContext) -> Optional[Finding]:
        ppt_ctx = dm.ppt_ctx

        row8 = ppt_ctx[7] if len(ppt_ctx) > 7 else []

        def _cell(col):
            return row8[col].strip() if col < len(row8) else ""

        # Check static labels
        static_issues = []
        for col, expected in self._STATIC.items():
            actual = _cell(col)
            if actual != expected:
                col_letter = chr(65 + col)
                static_issues.append(f"{col_letter}8 should be '{expected}' (is '{actual}')")

        # Check dynamic values (B8, D8)
        b8 = _cell(1)
        d8 = _cell(3)

        dynamic_issues = []
        if not b8:
            dynamic_issues.append("B8 (L2 or L3 selector) is empty")
        if not d8:
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
            msg += f" (fix available from Context tab — tool: {ctx_info['tool']})"
        else:
            msg += " Set A8:F8 manually."

        return Finding(
            "ERROR", "PPT context", msg,
            fix_id=FIX_L2_L3_CATEGORY if fix_available else None,
            fix_data=ctx_info if fix_available else None,
        )

    def fix(self, fix_data: dict, wq, dm) -> None:
        b8 = fix_data.get("b8") or ""
        d8 = fix_data.get("d8") or ""
        f8 = fix_data.get("f8") or ""
        tool = fix_data.get("tool", "unknown")

        print(f"  Setting PPT context row 8 labels (tool: {tool}).")
        wq.add_value("'PPT context'!A8", [["L2 or L3"]])
        print("  Queued: A8 = 'L2 or L3'.")
        wq.add_value("'PPT context'!C8", [["L2:"]])
        print("  Queued: C8 = 'L2:'.")
        wq.add_value("'PPT context'!E8", [["L3:"]])
        print("  Queued: E8 = 'L3:'.")

        if b8:
            wq.add_value("'PPT context'!B8", [[b8]])
            print(f"  Queued: B8 = '{b8}'.")
        else:
            print("  Skipped B8: no L2/L3 selector value found in Context tab.")

        if d8:
            wq.add_value("'PPT context'!D8", [[d8]])
            print(f"  Queued: D8 = '{d8}'.")
        else:
            print("  Skipped D8: no specific L2 value found in Context tab.")

        if f8:
            wq.add_value("'PPT context'!F8", [[f8]])
            print(f"  Queued: F8 = '{f8}'.")
        else:
            print("  Skipped F8: no specific L3 value found (expected for old tool or L2 choice).")
