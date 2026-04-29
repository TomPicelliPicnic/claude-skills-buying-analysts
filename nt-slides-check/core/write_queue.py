import threading


class WriteQueue:
    """Accumulate all fix operations during selection; dispatch in ≤3 API calls."""

    def __init__(self):
        self._raw          = []   # static cell values → RAW (no server-side parsing)
        self._clears       = []   # ranges to truly clear before formula writes
        self._user_entered = []   # formulas → USER_ENTERED (server parses =... syntax)
        self._struct       = []   # copyPaste / structural batchUpdate requests
        self._appends      = []   # (ws, rows) — must execute before structural ops

    def add_value(self, range_a1: str, values: list) -> None:
        self._raw.append({"range": range_a1, "values": values})

    def add_clear(self, range_a1: str) -> None:
        """Truly empty a range (not just set to '') before formula writes."""
        self._clears.append(range_a1)

    def add_formula(self, range_a1: str, values: list) -> None:
        """Use for formulas (=...) — dispatched with USER_ENTERED so Sheets parses them."""
        self._user_entered.append({"range": range_a1, "values": values})

    def add_structural(self, request: dict) -> None:
        self._struct.append(request)

    def add_append(self, ws, rows: list) -> None:
        self._appends.append((ws, rows))

    @property
    def is_empty(self) -> bool:
        return not (self._raw or self._clears or self._user_entered or self._struct or self._appends)

    def dispatch(self, sh) -> None:
        if self.is_empty:
            return
        print("🚀 Applying all fixes in one batch...")

        # Step 1: append_rows must finish before structural copyPaste (reads the new rows)
        for ws, rows in self._appends:
            ws.append_rows(rows, value_input_option="USER_ENTERED")

        # Steps 2-5 split into three independent groups — run in parallel:
        #   Group A — raw values (independent of everything)
        #   Group B — clear then formula (clear must precede user_entered; both independent of A/C)
        #   Group C — structural copyPaste (independent of A/B; depends on append, done above)

        def _raw():
            if self._raw:
                sh.values_batch_update({"valueInputOption": "RAW", "data": self._raw})

        def _clear_then_formula():
            if self._clears:
                sh.values_batch_clear(body={"ranges": self._clears})
            if self._user_entered:
                sh.values_batch_update({"valueInputOption": "USER_ENTERED", "data": self._user_entered})

        def _struct():
            if self._struct:
                sh.batch_update({"requests": self._struct})

        threads = [
            threading.Thread(target=_raw,               daemon=True),
            threading.Thread(target=_clear_then_formula, daemon=True),
            threading.Thread(target=_struct,             daemon=True),
        ]
        for t in threads:
            t.start()
        for t in threads:
            t.join()

        print("✅ Done")
