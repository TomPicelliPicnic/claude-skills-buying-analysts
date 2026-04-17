import json
from core.check_template import Finding, AuditContext

BORDER = "━" * 54


def _fmt_finding(f: Finding, icon: str) -> str:
    fix_part = f" (Fix {f.fix_id})" if f.fix_id is not None else ""
    lines = f.message.split("\n")
    first = f"  {icon} [Tab: {f.sheet}] {lines[0]}{fix_part}"
    rest  = ["        " + l for l in lines[1:]]
    return "\n".join([first] + rest)


def print_report(findings: list[Finding], ctx: AuditContext) -> None:
    errors   = [f for f in findings if f.severity == "ERROR"]
    warnings = [f for f in findings if f.severity == "WARNING"]

    print(BORDER)
    print(f"📋  NT-SLIDE-TABS-CHECK — {ctx.supplier_name}")
    print(f"    Sheet: {ctx.sheet_title}  |  Offer: {ctx.offer_id}  |  Week: {ctx.current_keyweek}")
    print(BORDER)
    print()

    if ctx.auto_fixed:
        print("⚡  AUTO-FIXED")
        for msg in ctx.auto_fixed:
            print(f"  ⚡ {msg}")
        print()

    if not findings:
        print("✅  All checks passed.")
    else:
        if errors:
            print("❌  ERRORS")
            for f in errors:
                print(_fmt_finding(f, "❌"))
            print()
        if warnings:
            print("⚠️   WARNINGS")
            for f in warnings:
                print(_fmt_finding(f, "⚠️ "))
            print()
        print(BORDER)
        print(f"Audit Complete: {len(errors)} Errors, {len(warnings)} Warnings ({ctx.ok_count} checks passed).")

    fixable = {
        str(f.fix_id): f.fix_data
        for f in findings
        if f.fix_id is not None and f.fix_data is not None
    }
    if fixable:
        print()
        print("__FIXABLE__" + json.dumps(fixable))
