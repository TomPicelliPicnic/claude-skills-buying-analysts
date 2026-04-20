from abc import ABC, abstractmethod
from dataclasses import dataclass, field
from typing import Optional


@dataclass
class Finding:
    severity: str               # "ERROR" or "WARNING"
    sheet: str                  # tab name — shown as [Tab: X] in report
    message: str
    fix_id: Optional[int] = None
    fix_data: Optional[dict] = None


@dataclass
class AuditContext:
    """Derived values computed once in audit.py and shared across all checks."""
    sh: object                  # gspread Spreadsheet (needed by some fix handlers)
    current_keyweek: int
    supplier_name: str
    offer_id: str
    sheet_title: str
    week_row: list              # PPT time row 0 — used by ~8 checks
    dealpoint_row: list         # PPT time row 2 — used by checks 1, 10, 16
    cur_col: Optional[int]      # column index of current_keyweek in week_row
    offers_type: str            # from Deal sheet rows 1–2
    ok_count: int = 0
    auto_fixed: list = field(default_factory=list)


class CheckTemplate(ABC):
    """
    Subclass this in checks/check_NN_name.py to add a new check.

    Required class attributes:
        id         int   — stable integer, never reuse (gaps are fine)
        name       str   — shown in report
        sheet_name str   — primary tab this check concerns
        severity   str   — max severity: "ERROR" or "WARNING"

    Optional class attributes:
        auto_fix   bool  — if True, fix() is called automatically during the audit
                           when the check fails and fix_data is available

    Rules:
        - run() must never call the Sheets API — read only from dm.* and ctx.*
        - fix() must never call the API directly — queue via wq.add_value /
          wq.add_structural / wq.add_append
        - Return None from run() (and do ctx.ok_count += 1) to signal a pass
    """
    id: int
    name: str
    sheet_name: str
    severity: str
    auto_fix: bool = False

    @abstractmethod
    def run(self, dm, ctx: AuditContext) -> Optional[Finding]:
        ...

    def fix(self, fix_data: dict, wq, dm) -> None:
        pass
