class WriteQueue:
    """Accumulate all fix operations during selection; dispatch in ≤3 API calls."""

    def __init__(self):
        self._raw     = []   # static cell values → RAW (no server-side parsing)
        self._struct  = []   # copyPaste / structural batchUpdate requests
        self._appends = []   # (ws, rows) — must execute before structural ops

    def add_value(self, range_a1: str, values: list) -> None:
        self._raw.append({"range": range_a1, "values": values})

    def add_structural(self, request: dict) -> None:
        self._struct.append(request)

    def add_append(self, ws, rows: list) -> None:
        self._appends.append((ws, rows))

    @property
    def is_empty(self) -> bool:
        return not (self._raw or self._struct or self._appends)

    def dispatch(self, sh) -> None:
        if self.is_empty:
            return
        print("🚀 Applying all fixes in one batch...")
        # Step 1: append_rows — new rows must exist before any copyPaste reads them
        for ws, rows in self._appends:
            ws.append_rows(rows, value_input_option="USER_ENTERED")
        # Step 2: all cell value writes — RAW skips server-side formula parsing
        if self._raw:
            sh.values_batch_update(self._raw, params={"valueInputOption": "RAW"})
        # Step 3: structural ops (copyPaste formula extensions)
        if self._struct:
            sh.batch_update({"requests": self._struct})
        print("✅ Done")
