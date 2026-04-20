#!/usr/bin/env python3
"""
NT Slides pre-flight audit.

Usage:
    poetry run python audit.py <SHEET_ID> [--fix '{"1": {...}, "6": {...}}']
    poetry run python audit.py <SHEET_ID>          # audit only
"""
import sys, json, subprocess, pathlib

REPO_DIR   = pathlib.Path(__file__).parent
GSHEET_DIR = pathlib.Path.home() / ".claude/skills/picnic-gsheet"
sys.path.insert(0, str(GSHEET_DIR))
sys.path.insert(0, str(REPO_DIR))

from datetime import date, timedelta
from gsheet_auth import get_credentials
import gspread

from core.data_manager  import DataManager
from core.write_queue   import WriteQueue
from core.check_template import AuditContext
from core.report        import print_report
from checks             import load_checks


def _sync():
    """Pull latest checks from git before running the audit."""
    result = subprocess.run(
        ["git", "-C", str(REPO_DIR), "pull", "--ff-only"],
        capture_output=True, text=True,
    )
    if result.returncode == 0:
        msg = result.stdout.strip() or "Already up to date."
        print(f"🔄 Sync: {msg}")
    else:
        print(f"⚠️  git pull failed (running with local version): {result.stderr.strip()}")


def _build_context(sh, dm) -> AuditContext:
    shelf_rows = dm.shelf_rows
    ppt_time   = dm.ppt_time
    tsv        = dm.tsv
    deal_sheet = dm.deal_sheet

    # Current keyweek: find 'Offer' label in Shelf analysis row 3
    current_keyweek = None
    if len(shelf_rows) > 2:
        for j, v in enumerate(shelf_rows[2]):
            if v.strip().lower() == "offer":
                raw = shelf_rows[1][j].strip() if j < len(shelf_rows[1]) else ""
                try:
                    current_keyweek = int(raw)
                except ValueError:
                    pass
                break
    if current_keyweek is None:
        ref = date.today() - timedelta(days=7)
        y, w, _ = ref.isocalendar()
        current_keyweek = int(f"{y}{w:02d}")

    tsv_header   = tsv[0] if tsv else []
    supplier_col = tsv_header.index("Contract_party") if "Contract_party" in tsv_header else None
    offer_col    = tsv_header.index("Offer_ID")       if "Offer_ID"       in tsv_header else None
    supplier_name = tsv[1][supplier_col] if supplier_col is not None and len(tsv) > 1 else "?"
    offer_id      = tsv[1][offer_col]    if offer_col    is not None and len(tsv) > 1 else "?"

    week_row      = ppt_time[0] if ppt_time else []
    dealpoint_row = ppt_time[2] if len(ppt_time) > 2 else []

    cur_col = None
    for i, v in enumerate(week_row):
        try:
            if int(v.strip()) == current_keyweek:
                cur_col = i
                break
        except (ValueError, AttributeError):
            pass

    offers_type = ""
    for row in deal_sheet[:2]:
        for ci, val in enumerate(row):
            if "offers type" in val.strip().lower() and ci + 1 < len(row):
                offers_type = row[ci + 1].strip()
                break
        if offers_type:
            break

    return AuditContext(
        sh              = sh,
        current_keyweek = current_keyweek,
        supplier_name   = supplier_name,
        offer_id        = offer_id,
        sheet_title     = sh.title,
        week_row        = week_row,
        dealpoint_row   = dealpoint_row,
        cur_col         = cur_col,
        offers_type     = offers_type,
    )


def main():
    args = sys.argv[1:]
    if not args:
        print("Usage: audit.py <SHEET_ID> [--fix '{\"1\": {}, ...}']")
        sys.exit(1)

    sheet_id = args[0]
    fix_instructions = {}
    if "--fix" in args:
        fix_idx = args.index("--fix")
        if fix_idx + 1 < len(args):
            try:
                fix_instructions = json.loads(args[fix_idx + 1])
            except json.JSONDecodeError as e:
                print(f"Invalid --fix JSON: {e}")
                sys.exit(1)

    creds = get_credentials()
    gc    = gspread.authorize(creds)
    sh    = gc.open_by_key(sheet_id)

    if fix_instructions:
        # Fast path: skip sync and checks — just load data and apply fixes
        dm     = DataManager(sh)
        checks = load_checks()
        check_map = {c.id: c for c in checks}
        wq = WriteQueue()
        for fix_id_str, fix_args in fix_instructions.items():
            check = check_map.get(int(fix_id_str))
            if check:
                check.fix(fix_args, wq, dm)
            else:
                print(f"  Unknown fix ID: {fix_id_str}")
        wq.dispatch(sh)
        return

    _sync()

    dm  = DataManager(sh)
    ctx = _build_context(sh, dm)

    checks   = load_checks()
    findings = []
    auto_wq  = WriteQueue()
    for check in checks:
        result = check.run(dm, ctx)
        if result is None:
            continue
        if check.auto_fix and result.fix_data is not None:
            check.fix(result.fix_data, auto_wq, dm)
            ctx.auto_fixed.append(f"[{result.sheet}] {result.message}")
        else:
            findings.append(result)

    if not auto_wq.is_empty:
        auto_wq.dispatch(sh)

    print_report(findings, ctx)


if __name__ == "__main__":
    main()
