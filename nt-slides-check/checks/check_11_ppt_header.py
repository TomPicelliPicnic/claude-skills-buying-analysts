from typing import Optional

from core.check_template import CheckTemplate, Finding, AuditContext


class PptHeaderCheck(CheckTemplate):
    id         = 11
    name       = "PPT context header fields B1–B4"
    sheet_name = "PPT context"
    severity   = "WARNING"

    def run(self, dm, ctx: AuditContext) -> Optional[Finding]:
        empty = []
        for row_idx, col_idx, ref in [(0, 1, "B1"), (1, 1, "B2"), (2, 1, "B3"), (3, 1, "B4")]:
            val = dm.ppt_ctx[row_idx][col_idx].strip() if len(dm.ppt_ctx) > row_idx and len(dm.ppt_ctx[row_idx]) > col_idx else ""
            if not val:
                empty.append(ref)
        if empty:
            return Finding("WARNING", "PPT context",
                f"Header cell(s) {', '.join(empty)} are empty. These populate the slide header.")
        ctx.ok_count += 1
        return None
