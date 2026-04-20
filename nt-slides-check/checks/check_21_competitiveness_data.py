from typing import Optional

from core.check_template import CheckTemplate, Finding, AuditContext


# ── Helpers ───────────────────────────────────────────────────────────────────

def _parse_num(val: str) -> Optional[float]:
    if not val or not val.strip():
        return None
    v = val.strip().replace(",", ".").replace("%", "").replace("+", "")
    try:
        return float(v)
    except ValueError:
        return None


def _bad_metric(val: str) -> tuple[bool, str]:
    """Return (is_bad, reason) for a depth/pressure cell."""
    if not val or not val.strip():
        return True, "empty"
    v = val.strip()
    if v.startswith("#"):
        return True, f"Excel error ({v})"
    num = _parse_num(v)
    if num is not None and num == 0:
        return True, "zero"
    return False, ""


def _parse_structure(ppt_ctx: list, anchor_row_idx: int) -> list[dict]:
    """
    Dynamically map both Competitiveness blocks from their header rows.

    Returns a list of block dicts:
        group_col   — col index of the promo group label
        comp1_name  — e.g. 'LY' or 'AH'
        comp2_name  — e.g. 'Offer'
        metrics     — {'freq': {'comp1': ci, 'comp2': ci},
                        'depth': ..., 'pressure': ...}
    """
    anchor_row = ppt_ctx[anchor_row_idx]

    # Both block start columns carry 'Competitiveness'
    block_cols = [
        ci for ci, v in enumerate(anchor_row)
        if "competitiveness" in v.strip().lower()
    ]

    # Find the section-label row (contains "freq") within next 4 rows
    section_row_idx = None
    for ri in range(anchor_row_idx + 1, min(anchor_row_idx + 5, len(ppt_ctx))):
        if any("freq" in v.strip().lower() for v in ppt_ctx[ri]):
            section_row_idx = ri
            break

    if section_row_idx is None:
        return []

    section_row = ppt_ctx[section_row_idx]

    # Col-label row is the one immediately after the section row that does NOT
    # contain section keywords — try +1, +2, -1 until we find it
    col_label_row: list = []
    for offset in (1, -1, 2):
        candidate_idx = section_row_idx + offset
        if not (0 <= candidate_idx < len(ppt_ctx)):
            continue
        candidate = ppt_ctx[candidate_idx]
        non_empty = [v.strip().lower() for v in candidate if v.strip()]
        if non_empty and not any(
            kw in v for v in non_empty for kw in ("freq", "depth", "pressure", "competitiveness")
        ):
            col_label_row = candidate
            break

    blocks = []
    for block_col in block_cols:
        window_end = block_col + 14

        freq_col = depth_col = pressure_col = None
        for ci in range(block_col + 1, min(window_end, len(section_row))):
            v = section_row[ci].strip().lower() if ci < len(section_row) else ""
            if "freq" in v and freq_col is None:
                freq_col = ci
            elif "depth" in v and depth_col is None:
                depth_col = ci
            elif "pressure" in v and pressure_col is None:
                pressure_col = ci

        if None in (freq_col, depth_col, pressure_col):
            continue

        comp1_name = col_label_row[freq_col].strip() if freq_col < len(col_label_row) else "Comp1"
        comp2_name = col_label_row[freq_col + 1].strip() if freq_col + 1 < len(col_label_row) else "Comp2"

        blocks.append({
            "group_col":  block_col,
            "comp1_name": comp1_name or "Comp1",
            "comp2_name": comp2_name or "Comp2",
            "metrics": {
                "freq":     {"comp1": freq_col,     "comp2": freq_col + 1},
                "depth":    {"comp1": depth_col,    "comp2": depth_col + 1},
                "pressure": {"comp1": pressure_col, "comp2": pressure_col + 1},
            },
        })

    return blocks


# ── Check ─────────────────────────────────────────────────────────────────────

class CompetitivenessDataCheck(CheckTemplate):
    id         = 21
    name       = "Competitiveness data integrity"
    sheet_name = "PPT context"
    severity   = "ERROR"

    def run(self, dm, ctx: AuditContext) -> Optional[Finding]:
        ppt_ctx = dm.ppt_ctx

        # ── Locate anchor ─────────────────────────────────────────────────────
        anchor_row_idx = None
        for ri, row in enumerate(ppt_ctx):
            if row and "competitiveness" in row[0].strip().lower():
                anchor_row_idx = ri
                break

        if anchor_row_idx is None:
            ctx.ok_count += 1
            return None  # Section absent — not an error for this sheet type

        # ── Parse column structure ────────────────────────────────────────────
        blocks = _parse_structure(ppt_ctx, anchor_row_idx)
        if not blocks:
            return Finding("WARNING", "PPT context",
                "Competitiveness section found but column structure could not be parsed.")

        # ── Find first data row ───────────────────────────────────────────────
        data_start = anchor_row_idx + 4
        for ri in range(anchor_row_idx + 1, min(anchor_row_idx + 6, len(ppt_ctx))):
            if ppt_ctx[ri] and ppt_ctx[ri][0].strip().lower() == "promo group":
                data_start = ri + 2
                break

        # ── Validate each data row ────────────────────────────────────────────
        issues: list[str] = []
        consecutive_empty = 0

        for ri in range(data_start, len(ppt_ctx)):
            row = ppt_ctx[ri]

            if not any(v.strip() for v in row):
                consecutive_empty += 1
                if consecutive_empty >= 2:
                    break
                continue
            consecutive_empty = 0

            for block in blocks:
                group_col  = block["group_col"]
                group_name = row[group_col].strip() if group_col < len(row) else ""

                if not group_name or group_name.lower() == "total":
                    continue

                for comp_key in ("comp1", "comp2"):
                    comp_name = block[f"{comp_key}_name"]
                    freq_col  = block["metrics"]["freq"][comp_key]
                    freq_val  = row[freq_col].strip() if freq_col < len(row) else ""
                    freq_num  = _parse_num(freq_val)

                    if freq_num is None or freq_num <= 0:
                        continue  # No frequency — nothing to validate

                    for metric in ("depth", "pressure"):
                        col = block["metrics"][metric][comp_key]
                        val = row[col].strip() if col < len(row) else ""
                        bad, reason = _bad_metric(val)
                        if bad:
                            fix_tab = f"{comp_name}-promo"
                            issues.append(
                                f"Group '{group_name}' has {comp_name} Frequency "
                                f"({freq_val}) but {comp_name} {metric.capitalize()} "
                                f"is {reason}. Fix the formula in the {fix_tab} tab."
                            )

        if not issues:
            ctx.ok_count += 1
            return None

        return Finding(
            "ERROR", "PPT context",
            f"{len(issues)} competitiveness data integrity issue(s):\n" +
            "\n".join(f"  {i}" for i in issues),
        )
