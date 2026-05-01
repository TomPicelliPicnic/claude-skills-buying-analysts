import re
from typing import Optional, List, Tuple

from core.check_template import CheckTemplate, Finding, AuditContext
from core.constants import FIX_ARTICLE_NAME_LENGTH

MAX_LEN_EXISTING = 50  # first 'Article name' block — existing articles
MAX_LEN_NEW      = 43  # second 'Article name' block — new listings


def _abbreviate(name: str) -> str:
    s = re.sub(r'\bgram\b',  'g',  name, flags=re.IGNORECASE)
    s = re.sub(r'\bliter\b', 'l',  s,    flags=re.IGNORECASE)
    s = re.sub(r'\bstuks\b', 'st', s,    flags=re.IGNORECASE)
    return s


def _remove_unit_space(name: str) -> str:
    """'350 ml' → '350ml', '30 g' → '30g', etc."""
    return re.sub(r'(\d+)\s+(ml|cl|g|l|st|kg)\b', r'\1\2', name, flags=re.IGNORECASE)


def _drop_brand(name: str) -> str:
    parts = name.split(' ', 1)
    return parts[1].strip() if len(parts) > 1 else name


def _propose(name: str, max_len: int) -> str:
    s = _abbreviate(name)
    if len(s) <= max_len:
        return s
    s = _remove_unit_space(s)
    if len(s) <= max_len:
        return s
    return _drop_brand(s)


def _find_blocks(offer_insights) -> List[List[Tuple[int, str, str]]]:
    """
    Scan Offer insights column C for 'Article name' headers.
    For each, skip blank row(s) then collect up to 15 data rows.
    Returns list of blocks: each block = [(row_0idx, article_id, name), ...].
    """
    blocks = []
    i = 0
    while i < len(offer_insights):
        row = offer_insights[i]
        col_c = row[2].strip() if len(row) > 2 else ""
        if col_c.lower() == "article name":
            j = i + 1
            while j < len(offer_insights) and not any(offer_insights[j]):
                j += 1
            block = []
            k = j
            while k < len(offer_insights) and k < j + 15:
                r = offer_insights[k]
                if not any(r):
                    break
                article_id = r[1].strip() if len(r) > 1 else ""
                name       = r[2].strip() if len(r) > 2 else ""
                block.append((k, article_id, name))
                k += 1
            if block:
                blocks.append(block)
        i += 1
    return blocks


class ArticleNameLengthCheck(CheckTemplate):
    id         = 30
    name       = "Article name length (≤47 chars)"
    sheet_name = "Offer insights"
    severity   = "WARNING"
    handles_fix_id = FIX_ARTICLE_NAME_LENGTH

    def run(self, dm, ctx: AuditContext) -> Optional[Finding]:
        blocks = _find_blocks(dm.offer_insights)
        if not blocks:
            ctx.ok_count += 1
            return None

        # --- Assortment lookup (block 0 only — existing articles) ---
        assortment  = getattr(dm, 'assortment', [])
        asst_header = assortment[0] if assortment else []

        # Article name column: look up by header, fall back to col E (index 4)
        asst_name_col = 4
        for i, h in enumerate(asst_header):
            if h.strip().lower() == "article name":
                asst_name_col = i
                break

        # Article ID is in col A (index 0) per sheet convention
        asst_by_id: dict = {}
        for row_0idx, row in enumerate(assortment[1:], start=1):
            aid = str(row[0]).strip() if len(row) > 0 else ""
            if aid:
                asst_by_id[aid] = row_0idx

        # --- Block 0: existing articles → fixable via Assortment ---
        fixable = []
        for _, article_id, name in blocks[0]:
            if not name or len(name) <= MAX_LEN_EXISTING:
                continue
            proposed  = _propose(name, MAX_LEN_EXISTING)
            asst_row  = asst_by_id.get(article_id)
            fixable.append({
                "article_id":          article_id,
                "current_name":        name,
                "proposed_name":       proposed,
                "assortment_row_1idx": (asst_row + 1) if asst_row is not None else None,
                "assortment_col_1idx": asst_name_col + 1,
            })

        # --- Block 1+: new listings → manual fix only ---
        manual = []
        for block in blocks[1:]:
            for _, _, name in block:
                if name and len(name) > MAX_LEN_NEW:
                    manual.append(name)

        if not fixable and not manual:
            ctx.ok_count += 1
            return None

        parts = []
        if fixable:
            lines = []
            for a in fixable:
                plen  = len(a['proposed_name'])
                clen  = len(a['current_name'])
                note  = "" if plen <= MAX_LEN_EXISTING else f"  ← still {plen} chars, shorten further manually"
                lines.append(
                    f"  '{a['current_name']}' ({clen})\n"
                    f"    → '{a['proposed_name']}' ({plen}){note}"
                )
            parts.append(
                f"{len(fixable)} article name(s) exceed {MAX_LEN_EXISTING} chars "
                f"— fix will update Assortment col E:\n" + "\n".join(lines)
            )
        if manual:
            lines = [f"  '{n}' ({len(n)} chars)" for n in manual]
            parts.append(
                f"{len(manual)} new listing name(s) exceed {MAX_LEN_NEW} chars "
                f"— fix manually in current offer tab:\n" + "\n".join(lines)
            )

        return Finding(
            "WARNING", "Offer insights",
            "\n".join(parts),
            fix_id   = FIX_ARTICLE_NAME_LENGTH if fixable else None,
            fix_data = {"articles": fixable}   if fixable else None,
        )

    def fix(self, fix_data: dict, wq, dm) -> None:
        import gspread.utils
        for a in fix_data.get("articles", []):
            row = a.get("assortment_row_1idx")
            col = a.get("assortment_col_1idx")
            if not row or not col:
                print(f"  Skipped '{a['article_id']}': not found in Assortment.")
                continue
            cell = gspread.utils.rowcol_to_a1(row, col)
            wq.add_value(f"'Assortment'!{cell}", [[a["proposed_name"]]])
            print(f"  Queued: Assortment {cell} = '{a['proposed_name']}' (was '{a['current_name']}')")
