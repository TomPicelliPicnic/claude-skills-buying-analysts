import re
from typing import Optional

from core.check_template import CheckTemplate, Finding, AuditContext

_EXPECTED = re.compile(r"^new\d+$", re.IGNORECASE)
_ANY_NEW   = re.compile(r"^new", re.IGNORECASE)


class NewListingNamesCheck(CheckTemplate):
    id         = 24
    name       = "New listing Article IDs are uniquely numbered"
    sheet_name = "TSV output"
    severity   = "ERROR"

    def run(self, dm, ctx: AuditContext) -> Optional[Finding]:
        tsv_header = dm.tsv[0] if dm.tsv else []
        col = next(
            (i for i, h in enumerate(tsv_header) if h.strip().lower() == "article id"),
            None,
        )
        if col is None:
            return Finding("ERROR", "TSV output",
                "Column 'Article ID' not found. Cannot verify new listing names.")

        new_ids = [
            row[col].strip()
            for row in dm.tsv[1:]
            if any(row) and len(row) > col and _ANY_NEW.match(row[col].strip())
        ]

        if not new_ids:
            ctx.ok_count += 1
            return None

        bad = [v for v in new_ids if not _EXPECTED.match(v)]
        duplicates = [v for v in set(new_ids) if new_ids.count(v) > 1]

        issues = []
        if bad:
            unique_bad = sorted(set(bad))
            issues.append(
                f"not following 'new1/new2/…' format: {', '.join(unique_bad)}"
            )
        if duplicates:
            issues.append(
                f"duplicate entries: {', '.join(sorted(set(duplicates)))} "
                f"(appears {max(new_ids.count(v) for v in duplicates)}× each)"
            )

        if issues:
            return Finding("ERROR", "TSV output",
                "New listing Article IDs are incorrectly named — "
                + "; ".join(issues)
                + ". Rename them to new1, new2, new3, … in the TSV output sheet.")

        ctx.ok_count += 1
        return None
