from typing import Optional

from core.check_template import CheckTemplate, Finding, AuditContext

_GUARD = "<=MAX('Article shelf'!C:C)"


def _col_letter(col_0idx: int) -> str:
    result = ""
    n = col_0idx + 1
    while n > 0:
        n, rem = divmod(n - 1, 26)
        result = chr(65 + rem) + result
    return result


def _extract_if_true_branch(f: str) -> str:
    """Extract the true branch of =IF(cond, TRUE_BRANCH, false) by counting parens."""
    try:
        start = f.index("(")
    except ValueError:
        return f[1:]
    depth = 0
    splits = []
    for i, c in enumerate(f[start + 1:], start + 1):
        if c == "(":
            depth += 1
        elif c == ")":
            if depth == 0:
                break
            depth -= 1
        elif c == "," and depth == 0:
            splits.append(i)
    if len(splits) >= 2:
        return f[splits[0] + 1 : splits[1]]
    return f[1:]


class N3FutureCheck(CheckTemplate):
    id               = 2
    name             = "PPT time rows 4+ capped at Article shelf max week"
    sheet_name       = "PPT time"
    severity         = "ERROR"
    auto_fix         = True
    auto_fix_message = "Added Article shelf week cap to PPT time rows 4+"

    def _unwrapped(self, dm):
        result = []
        for row_idx in range(3, len(dm.ppt_formulas)):
            for col_idx, cell_formula in enumerate(dm.ppt_formulas[row_idx]):
                if col_idx == 0:
                    continue
                f = str(cell_formula).strip()
                if not f.startswith("="):
                    continue
                if _GUARD in f:
                    continue
                # Cells guarded by the L4L check (check 29) — skip to avoid conflict
                if ">=Context!" in f or f.upper().startswith("=IFERROR("):
                    continue
                col_letter = _col_letter(col_idx)
                a1 = f"'PPT time'!{col_letter}{row_idx + 1}"
                # Extract inner formula when re-wrapping an existing IF guard
                inner = _extract_if_true_branch(f) if f.upper().startswith("=IF(") else f[1:]
                result.append((inner, a1, col_letter))
        return result

    def run(self, dm, ctx: AuditContext) -> Optional[Finding]:
        cells = self._unwrapped(dm)
        if not cells:
            ctx.ok_count += 1
            return None
        return Finding("ERROR", "PPT time",
            f"{len(cells)} formula cell(s) in rows 4+ are not capped at the Article shelf max week.",
            fix_data={})

    def fix(self, fix_data: dict, wq, dm) -> None:
        cells = self._unwrapped(dm)
        for inner, a1, col_letter in cells:
            wrapped = f"=IF({col_letter}$1{_GUARD},{inner},\"\")"
            wq.add_formula(a1, [[wrapped]])
        if cells:
            print(f"  Queued: Article shelf week cap for {len(cells)} formula cell(s) in PPT time rows 4+.")
