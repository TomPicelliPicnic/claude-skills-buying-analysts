from typing import Optional

from core.check_template import CheckTemplate, Finding, AuditContext


class N3StabilityCheck(CheckTemplate):
    id         = 16
    name       = "PPT time N3 L4L stable between dealpoints"
    sheet_name = "PPT time"
    severity   = "WARNING"

    def run(self, dm, ctx: AuditContext) -> Optional[Finding]:
        n3_l4l = dm.ppt_time[6] if len(dm.ppt_time) > 6 else []
        dp_cols = [c for c, v in enumerate(ctx.dealpoint_row) if v == "TRUE"]
        last_n3_col = max((c for c, v in enumerate(n3_l4l) if v.strip()), default=None)

        if last_n3_col is None or not dp_cols:
            ctx.ok_count += 1
            return None

        boundaries  = dp_cols + [last_n3_col + 1]
        unstable    = []
        week_row    = ctx.week_row

        for i in range(len(boundaries) - 1):
            seg_start, seg_end = boundaries[i], boundaries[i + 1]
            vals = [
                (c, n3_l4l[c].strip()) for c in range(seg_start, seg_end)
                if c < len(n3_l4l) and n3_l4l[c].strip()
            ]
            if len({v for _, v in vals}) <= 1:
                continue
            first_val = vals[0][1]
            change_col, change_val = next((c, v) for c, v in vals if v != first_val)
            dp_wk      = week_row[seg_start].strip()      if seg_start      < len(week_row) else "?"
            change_wk  = week_row[change_col].strip()     if change_col     < len(week_row) else "?"
            prev_wk    = week_row[change_col - 1].strip() if 0 < change_col - 1 < len(week_row) else "?"
            next_dp_wk = week_row[boundaries[i + 1]].strip() if boundaries[i + 1] < len(week_row) else "end"
            unstable.append(
                f"  Between dealpoints {dp_wk}–{next_dp_wk}: "
                f"{dp_wk}–{prev_wk} at {first_val}, then changes at {change_wk} to {change_val}"
            )

        if unstable:
            return Finding("WARNING", "PPT time",
                f"Net 3 - L4L changes within {len(unstable)} inter-dealpoint period(s) — "
                "check Price list for mid-period price changes:\n" + "\n".join(unstable))
        ctx.ok_count += 1
        return None
