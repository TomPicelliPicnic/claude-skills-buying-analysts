class DataManager:
    """Fetches every required tab in exactly 2 API calls. Zero network calls during checks."""

    def __init__(self, sh):
        self.sh     = sh
        self.ws_ppt = sh.worksheet("PPT time")
        self._load()

    def _load(self):
        # Call 1: FORMATTED_VALUE for all tabs.
        # PPT time scoped to A1:JA20 (261 cols × 20 rows) — full-sheet reads of
        # a 261-col sheet are the main latency driver; this is ~50× less data.
        _batch = self.sh.values_batch_get(
            ranges=[
                "'Shelf analysis'!1:3",
                "'PPT time'!A1:JA20",
                "'TSV output'",
                "'Price list'",
                "Context!1:82",
                "'Offer insights'",
                "'PPT context'",
                "'Deal sheet'",
            ],
            params={"valueRenderOption": "FORMATTED_VALUE"},
        )
        _vr = _batch.get("valueRanges", [])

        def _g(i):
            return _vr[i].get("values", []) if i < len(_vr) else []

        self.shelf_rows     = _g(0)
        self.ppt_time       = _g(1)
        self.tsv            = _g(2)
        self.price_list     = _g(3)
        self.context        = _g(4)
        self.offer_insights = _g(5)
        self.ppt_ctx        = _g(6)
        self.deal_sheet     = _g(7)

        # Call 2: FORMULA render for PPT time — scoped identically.
        # Cannot be combined with call 1 (different valueRenderOption).
        self.ppt_formulas = self.ws_ppt.get("A1:JA20", value_render_option="FORMULA")
