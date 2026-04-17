from typing import Optional

from core.check_template import CheckTemplate, Finding, AuditContext


class BenchmarkCountCheck(CheckTemplate):
    id         = 6
    name       = "Benchmark article count consistency"
    sheet_name = "TSV output"
    severity   = "WARNING"

    def run(self, dm, ctx: AuditContext) -> Optional[Finding]:
        tsv_header = dm.tsv[0] if dm.tsv else []
        col = next((i for i, h in enumerate(tsv_header) if h.strip().lower() == "in_benchmark"), None)

        if col is None:
            return Finding("WARNING", "TSV output",
                "Column 'In_benchmark' not found. Cannot verify benchmark count.")

        true_count = sum(
            1 for row in dm.tsv[1:]
            if any(row) and len(row) > col and row[col].strip().upper() == "TRUE"
        )

        e7_raw = dm.offer_insights[6][4].strip() if len(dm.offer_insights) > 6 and len(dm.offer_insights[6]) > 4 else ""
        try:
            e7 = int(e7_raw)
        except (ValueError, TypeError):
            return Finding("WARNING", "Offer insights",
                f"E7 is empty or non-numeric ('{e7_raw}'). Cannot verify benchmark count.")

        if true_count != e7:
            return Finding("WARNING", "TSV output",
                f"In_benchmark TRUE count in TSV ({true_count}) does not match Offer insights E7 ({e7}). "
                "Fix manually — check which articles are flagged In_benchmark in TSV output.")
        ctx.ok_count += 1
        return None
