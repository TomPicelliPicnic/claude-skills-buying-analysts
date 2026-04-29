from typing import Optional

from core.check_template import CheckTemplate, Finding, AuditContext
from core.constants import FIX_BENCHMARK_TSV


class BenchmarkCountCheck(CheckTemplate):
    id         = 6
    name       = "Benchmark article count consistency"
    sheet_name = "TSV output"
    severity   = "WARNING"
    handles_fix_id = FIX_BENCHMARK_TSV

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
            n_rows = len([r for r in dm.tsv[1:] if any(r)])
            return Finding("WARNING", "TSV output",
                f"In_benchmark TRUE count in TSV ({true_count}) does not match Offer insights E7 ({e7}). "
                "Reset columns I/J (Shelf analysis spill) and P/R/V formulas?",
                fix_id=FIX_BENCHMARK_TSV,
                fix_data={"n_rows": n_rows})
        ctx.ok_count += 1
        return None

    def fix(self, fix_data: dict, wq, dm) -> None:
        import gspread.utils

        n_rows = fix_data.get("n_rows") or len([r for r in dm.tsv[1:] if any(r)])
        if n_rows == 0:
            print("  No data rows in TSV output — skipping.")
            return
        last_row = n_rows + 1  # 1-indexed last data row

        headers = dm.tsv[0] if dm.tsv else []

        def _col(name):
            idx = next((i for i, h in enumerate(headers) if h.strip() == name), None)
            if idx is None:
                raise ValueError(f"Header '{name}' not found in TSV output row 1")
            # rowcol_to_a1(1, n) always ends with "1"; strip it to get the column letter(s)
            return gspread.utils.rowcol_to_a1(1, idx + 1)[:-1]

        try:
            c_article = _col("Article ID")   # spill target col 1 (I)
            c_gtin    = _col("CU_GTIN")       # spill target col 2 (J), used in lookup
            c_offer   = _col("Offer_ID")      # match key (G), used in lookup
            c_net1    = _col("Net_1_price")                  # P
            c_net3    = _col("Net_3_price_ex_lumpsum")       # R
            c_net3le  = _col("Net_3_price_LE")               # V
        except ValueError as e:
            print(f"  ERROR: {e} — skipping fix.")
            return

        # Clear spill target columns so the array formula finds empty cells
        wq.add_clear(f"'TSV output'!{c_article}2:{c_gtin}{last_row + 5}")

        # Spill formula: fills c_article (Article ID) and c_gtin (CU_GTIN) from Shelf analysis
        wq.add_formula(f"'TSV output'!{c_article}2", [["={'Shelf analysis'!B8:C}"]])

        # Lookup formulas for net price columns — match on CU_GTIN × Offer_ID
        def _lookup(stacked_col, row):
            return (
                f"=IF({c_gtin}{row}>0,"
                f"IFNA(INDEX('Offers stacked'!{stacked_col}:{stacked_col},"
                f"MATCH(1,({c_gtin}{row}='Offers stacked'!B:B)*({c_offer}{row}='Offers stacked'!E:E),0))"
                f"),)"
            )

        rows = range(2, last_row + 1)
        wq.add_formula(f"'TSV output'!{c_net1}2:{c_net1}{last_row}",  [[_lookup("F", r)] for r in rows])
        wq.add_formula(f"'TSV output'!{c_net3}2:{c_net3}{last_row}",  [[_lookup("G", r)] for r in rows])
        wq.add_formula(f"'TSV output'!{c_net3le}2:{c_net3le}{last_row}", [[_lookup("H", r)] for r in rows])

        print(
            f"  Queued: clear {c_article}:{c_gtin} + spill formula + "
            f"update {c_net1}/{c_net3}/{c_net3le} for {n_rows} rows."
        )
