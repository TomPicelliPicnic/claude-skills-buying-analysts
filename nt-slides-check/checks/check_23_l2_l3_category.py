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


def _find_brand_row(ppt_ctx: list) -> Optional[int]:
    """Return 0-based row index of the row containing 'Brand' in col A, or None."""
    for ri, row in enumerate(ppt_ctx):
        if row and row[0].strip().lower() == "brand":
            return ri
    return None


def _find_l2_l3_from_context(context: list) -> dict:
    """
    Locate 'Margin comparison: Brand in L2' anchor in Context tab and derive
    cell references for the row above 'Brand' in PPT context.

    Returns dict with keys:
        b8_ref, d8_ref, f8_ref  — A1-notation Context refs (empty string if absent)
        b8_val, d8_val, f8_val  — display values for messages
        tool                    — 'old', 'new', or 'unknown'

    Old tool: 'L2' label + chosen value appear on the same row as 'Brand' in Context.
    New tool: 'Select: L2 or L3', 'Select: specific L2', 'Select: specific L3' rows
              each have their value in the row below the label.
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

    # Old tool: find the 'Brand' row in Context (col B = index 1), then read the L2
    # label + chosen value that appear on that same row (cols 12 and 13).
    for ri in range(anchor_ri + 1, search_end):
        row = context[ri]
        if len(row) > 1 and row[1].strip() == "Brand":
            # 'L2' label is at col 12, chosen category at col 13 (on the Brand row)
            if len(row) > 13:
                result["b8_ref"] = f"Context!{_col_letter(12)}{ri + 1}"
                result["b8_val"] = row[12].strip()
                result["d8_ref"] = f"Context!{_col_letter(13)}{ri + 1}"
                result["d8_val"] = row[13].strip()
                result["tool"] = "old"
                return result

    return result


class L2L3CategoryCheck(CheckTemplate):
    id         = 23
    name       = "L2/L3 category labels in PPT context (row above Brand)"
    sheet_name = "PPT context"
    severity   = "ERROR"
    auto_fix   = True

    # Static labels expected in the row above 'Brand' — col indices 0, 2, 4
    _STATIC = {0: "L2 or L3", 2: "L2:", 4: "L3:"}

    def run(self, dm, ctx: AuditContext) -> Optional[Finding]:
        ppt_ctx = dm.ppt_ctx

        brand_row_idx = _find_brand_row(ppt_ctx)
        if brand_row_idx is None or brand_row_idx == 0:
            ctx.ok_count += 1
            return None  # No Brand header found — sheet type doesn't have this section

        target_row_idx = brand_row_idx - 1
        row_above = ppt_ctx[target_row_idx] if target_row_idx < len(ppt_ctx) else []

        def _cell(col):
            return row_above[col].strip() if col < len(row_above) else ""

        static_issues = []
        for col, expected in self._STATIC.items():
            actual = _cell(col)
            if actual != expected:
                static_issues.append(f"{chr(65 + col)}{brand_row_idx} should be '{expected}' (is '{actual}')")

        dynamic_issues = []
        if not _cell(1):
            dynamic_issues.append(f"B{brand_row_idx} (L2 or L3 selector) is empty")
        if not _cell(3):
            dynamic_issues.append(f"D{brand_row_idx} (specific L2 value) is empty")

        all_issues = static_issues + dynamic_issues
        if not all_issues:
            ctx.ok_count += 1
            return None

        ctx_info = _find_l2_l3_from_context(dm.context)
        fix_available = ctx_info is not None and ctx_info.get("tool") != "unknown"

        # Include brand_row_idx in fix_data so fix() can target the right row dynamically
        if fix_available:
            ctx_info["brand_row_idx"] = brand_row_idx

        msg = (
            f"PPT context row above Brand ({len(all_issues)} issue(s)): "
            + "; ".join(all_issues) + "."
        )
        if fix_available:
            msg += f" Suggested: B={ctx_info['b8_val']}, D={ctx_info['d8_val']} (tool: {ctx_info['tool']})"

        return Finding(
            "ERROR", "PPT context", msg,
            fix_id=FIX_L2_L3_CATEGORY if fix_available else None,
            fix_data=ctx_info if fix_available else None,
        )

    def fix(self, fix_data: dict, wq, dm) -> None:
        tool = fix_data.get("tool", "unknown")
        # Row above Brand: brand_row_idx (0-based) == 1-based sheet row of the target row
        brand_row_idx = fix_data.get("brand_row_idx", 8)
        row = brand_row_idx  # 1-based sheet row for cell refs

        print(f"  Setting PPT context row {row} labels (row above Brand, tool: {tool}).")
        wq.add_value(f"'PPT context'!A{row}", [["L2 or L3"]])
        print(f"  Queued: A{row} = 'L2 or L3'.")
        wq.add_value(f"'PPT context'!C{row}", [["L2:"]])
        print(f"  Queued: C{row} = 'L2:'.")
        wq.add_value(f"'PPT context'!E{row}", [["L3:"]])
        print(f"  Queued: E{row} = 'L3:'.")

        b8_ref = fix_data.get("b8_ref") or ""
        d8_ref = fix_data.get("d8_ref") or ""
        f8_ref = fix_data.get("f8_ref") or ""

        if b8_ref:
            wq.add_formula(f"'PPT context'!B{row}", [[f"={b8_ref}"]])
            print(f"  Queued: B{row} = formula ={b8_ref} ('{fix_data.get('b8_val', '')}').")
        else:
            print(f"  Skipped B{row}: no Context cell reference found.")

        if d8_ref:
            wq.add_formula(f"'PPT context'!D{row}", [[f"={d8_ref}"]])
            print(f"  Queued: D{row} = formula ={d8_ref} ('{fix_data.get('d8_val', '')}').")
        else:
            print(f"  Skipped D{row}: no Context cell reference found.")

        if f8_ref:
            wq.add_formula(f"'PPT context'!F{row}", [[f"={f8_ref}"]])
            print(f"  Queued: F{row} = formula ={f8_ref} ('{fix_data.get('f8_val', '')}').")
        else:
            print(f"  Skipped F{row}: no L3 reference (expected for old tool or L2 choice).")
