from typing import Optional

from core.check_template import CheckTemplate, Finding, AuditContext

MIN_DECK_ARTICLES = 5


class DeckSelectionCheck(CheckTemplate):
    id         = 7
    name       = "Offer insights: articles selected for deck"
    sheet_name = "Offer insights"
    severity   = "WARNING"

    def run(self, dm, ctx: AuditContext) -> Optional[Finding]:
        deck_col = None
        deck_header_row = None
        for ri, row in enumerate(dm.offer_insights):
            for ci, val in enumerate(row):
                if "include in deck" in val.strip().lower():
                    deck_col = ci
                    deck_header_row = ri
                    break
            if deck_col is not None:
                break

        if deck_col is None:
            return Finding("WARNING", "Offer insights",
                "Column 'Include in deck?' not found. Cannot check article selection.")

        true_count = sum(
            1 for row in dm.offer_insights[deck_header_row + 1:]
            if len(row) > deck_col and row[deck_col].strip().upper() == "TRUE"
        )
        if true_count < MIN_DECK_ARTICLES:
            return Finding("WARNING", "Offer insights",
                f"Only {true_count} article(s) selected ('Include in deck?' column). "
                f"At least {MIN_DECK_ARTICLES} are needed to populate the PPT time chart.")
        ctx.ok_count += 1
        return None
