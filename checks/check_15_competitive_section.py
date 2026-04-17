from typing import Optional

from core.check_template import CheckTemplate, Finding, AuditContext


class CompetitiveSectionCheck(CheckTemplate):
    id         = 15
    name       = "PPT context competitive section"
    sheet_name = "PPT context"
    severity   = "WARNING"

    def run(self, dm, ctx: AuditContext) -> Optional[Finding]:
        issues = []
        for ri in range(35, 55):
            if ri >= len(dm.ppt_ctx):
                break
            row = dm.ppt_ctx[ri]
            if not any(row):
                continue

            left_group = row[0].strip()  if len(row) > 0  else ""
            left_data  = any(row[i].strip() for i in range(1, 10)  if i < len(row))
            left_ns    = row[10].strip() if len(row) > 10 else ""

            right_group = row[12].strip() if len(row) > 12 else ""
            right_data  = any(row[i].strip() for i in range(13, 22) if i < len(row))
            right_ns    = row[22].strip() if len(row) > 22 else ""

            if left_group:
                if not left_data:
                    issues.append(f"Row {ri + 1} ('{left_group}'): no data in cols B–J")
                elif not left_ns:
                    issues.append(f"Row {ri + 1} ('{left_group}'): net sales (col K) is empty")
            elif left_data and not left_ns:
                issues.append(f"Row {ri + 1}: data in cols B–J but net sales (col K) is empty")

            if right_group:
                if not right_data:
                    issues.append(f"Row {ri + 1} ('{right_group}'): no data in cols N–V")
                elif not right_ns:
                    issues.append(f"Row {ri + 1} ('{right_group}'): net sales (col W) is empty")
            elif right_data and not right_ns:
                issues.append(f"Row {ri + 1}: data in cols N–V but net sales (col W) is empty")

        if issues:
            return Finding("WARNING", "PPT context",
                f"{len(issues)} issue(s) in competitive section (rows 36–55):\n" +
                "\n".join(f"  {l}" for l in issues))
        ctx.ok_count += 1
        return None
