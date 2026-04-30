from typing import Optional

from core.check_template import CheckTemplate, Finding, AuditContext

FIX_L4L_ROW_GUARD = 10

# PPT time rows (0-indexed) that must be blank before the L4L start week
_L4L_ROWS = {6: "Net 3 - L4L", 7: "NLP - L4L"}

# PPT time rows (0-indexed) that depend on L4L rows and need IFERROR wrapping
_IFERROR_ROWS = {8: "Margin L4L"}

_ERROR_MARKERS = {"#div/0!", "#value!", "#ref!", "#n/a", "#name?", "#num!", "#null!"}


def _find_l4l_week_and_ref(context: list):
    """Return (week_int, cell_ref) from the row below 'Like for Like' in Context.

    cell_ref is e.g. 'B82' — the Sheets A1 address of the L4L week cell.
    Returns (None, None) if not found.
    """
    for ri, row in enumerate(context):
        for val in row:
            if "like for like" in val.strip().lower():
                data_ri = ri + 1
                if data_ri < len(context) and len(context[data_ri]) > 1:
                    raw = context[data_ri][1].strip()
                    if raw.isdigit() and len(raw) == 6:
                        return int(raw), f"B{data_ri + 1}"  # +1 for 1-indexed Sheets row
    return None, None


def _col_letter(col_0idx: int) -> str:
    """Convert a 0-based column index to an A1 column letter (A, B, ..., AA, ...)."""
    result = ""
    n = col_0idx + 1
    while n > 0:
        n, rem = divmod(n - 1, 26)
        result = chr(65 + rem) + result
    return result


class L4LRowGuardCheck(CheckTemplate):
    id             = 29
    name           = "PPT time L4L rows blank before L4L start week"
    sheet_name     = "PPT time"
    severity       = "WARNING"
    handles_fix_id = FIX_L4L_ROW_GUARD

    def run(self, dm, ctx: AuditContext) -> Optional[Finding]:
        l4l_week, _ = _find_l4l_week_and_ref(dm.context)
        if l4l_week is None:
            ctx.ok_count += 1
            return None

        # Find the column index of the L4L start week in the week row
        l4l_col = None
        for i, v in enumerate(ctx.week_row):
            try:
                if int(v.strip()) == l4l_week:
                    l4l_col = i
                    break
            except (ValueError, AttributeError):
                pass

        if l4l_col is None:
            ctx.ok_count += 1
            return None

        # Count non-empty values in L4L rows for columns strictly before l4l_col
        problem_rows = []
        for row_idx, row_label in _L4L_ROWS.items():
            row_data = dm.ppt_time[row_idx] if len(dm.ppt_time) > row_idx else []
            pre_vals = [
                c for c in range(1, l4l_col)
                if c < len(row_data) and row_data[c].strip()
            ]
            if pre_vals:
                problem_rows.append(f"{row_label} ({len(pre_vals)} week(s))")

        # Also check IFERROR rows for error values
        error_rows = []
        for row_idx, row_label in _IFERROR_ROWS.items():
            row_data = dm.ppt_time[row_idx] if len(dm.ppt_time) > row_idx else []
            err_count = sum(
                1 for c in range(1, len(row_data))
                if row_data[c].strip().lower() in _ERROR_MARKERS
            )
            if err_count:
                error_rows.append(f"{row_label} ({err_count} error(s))")

        if not problem_rows and not error_rows:
            ctx.ok_count += 1
            return None

        parts = []
        if problem_rows:
            parts.append(
                f"Data found before L4L start week {l4l_week} in: {', '.join(problem_rows)}."
            )
        if error_rows:
            parts.append(
                f"Formula errors in: {', '.join(error_rows)} — likely caused by L4L guard returning empty string."
            )
        parts.append("Fix will wrap formula cells with IF / IFERROR guards.")

        return Finding(
            "WARNING", "PPT time",
            " ".join(parts),
            fix_id=FIX_L4L_ROW_GUARD,
            fix_data={},
        )

    def fix(self, fix_data: dict, wq, dm) -> None:
        l4l_week, l4l_cell_ref = _find_l4l_week_and_ref(dm.context)
        if l4l_week is None or l4l_cell_ref is None:
            print("  L4L week not found in Context — skipping fix.")
            return

        for row_idx, row_label in _L4L_ROWS.items():
            formulas = dm.ppt_formulas[row_idx] if len(dm.ppt_formulas) > row_idx else []
            updates = []
            for col_idx, cell_formula in enumerate(formulas):
                if col_idx == 0:
                    continue  # skip label column A
                f = str(cell_formula).strip()
                if not f.startswith("="):
                    continue  # empty or plain value — leave alone
                # Skip if already wrapped with this guard
                if ">=Context!" in f:
                    continue
                col_letter = _col_letter(col_idx)
                wrapped = (
                    f"=IF({col_letter}$1>=Context!{l4l_cell_ref},"
                    f"{f[1:]},"  # strip leading "="
                    f'"")'
                )
                sheet_row = row_idx + 1  # 1-indexed
                a1 = f"'PPT time'!{col_letter}{sheet_row}"
                updates.append({"range": a1, "values": [[wrapped]]})

            if updates:
                for u in updates:
                    wq.add_formula(u["range"], u["values"])
                print(f"  Queued {len(updates)} formula update(s) for {row_label}.")

        for row_idx, row_label in _IFERROR_ROWS.items():
            formulas = dm.ppt_formulas[row_idx] if len(dm.ppt_formulas) > row_idx else []
            updates = []
            for col_idx, cell_formula in enumerate(formulas):
                if col_idx == 0:
                    continue
                f = str(cell_formula).strip()
                if not f.startswith("="):
                    continue
                # Skip if already wrapped with IFERROR
                if f.upper().startswith("=IFERROR("):
                    continue
                col_letter = _col_letter(col_idx)
                wrapped = f"=IFERROR({f[1:]},\"\")"
                sheet_row = row_idx + 1
                a1 = f"'PPT time'!{col_letter}{sheet_row}"
                updates.append({"range": a1, "values": [[wrapped]]})

            if updates:
                for u in updates:
                    wq.add_formula(u["range"], u["values"])
                print(f"  Queued {len(updates)} IFERROR wrap(s) for {row_label}.")
