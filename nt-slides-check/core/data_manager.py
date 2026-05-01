import threading


class DataManager:
    """Fetches every required tab in 2 parallel API calls. Zero network calls during checks."""

    def __init__(self, sh):
        self.sh     = sh
        self.ws_ppt = sh.worksheet("PPT time")
        self._load()

    def _load(self):
        # Call 1 (FORMATTED_VALUE) and Call 2 (FORMULA) run in parallel — they use different
        # valueRenderOption so cannot be combined, but they're fully independent.
        _formula_box = [None]

        def _fetch_formulas():
            _formula_box[0] = self.ws_ppt.get("A1:JA60", value_render_option="FORMULA")

        formula_thread = threading.Thread(target=_fetch_formulas, daemon=True)
        formula_thread.start()

        _batch = self.sh.values_batch_get(
            ranges=[
                "'Shelf analysis'!1:8",
                "'PPT time'!A1:JA60",
                "'TSV output'",
                "'Price list'",
                "Context!1:200",
                "'Offer insights'",
                "'PPT context'",
                "'Deal sheet'",
                "'Article shelf'!C:C",  # Week column only — for offer-week staleness check
                "'AH-Promo'!A:Z",       # CU GTIN + Price match columns — for promo plan check
                "'Assortment'!A:F",     # Article ID (col B) + Article name (col E) — for name length check
            ],
            params={"valueRenderOption": "FORMATTED_VALUE"},
        )

        formula_thread.join()
        self.ppt_formulas = _formula_box[0]

        _vr = _batch.get("valueRanges", [])

        def _g(i):
            return _vr[i].get("values", []) if i < len(_vr) else []

        self.shelf_rows          = _g(0)
        self.ppt_time            = _g(1)
        self.tsv                 = _g(2)
        self.price_list          = _g(3)
        self.context             = _g(4)
        self.offer_insights      = _g(5)
        self.ppt_ctx             = _g(6)
        self.deal_sheet          = _g(7)
        self.article_shelf_weeks = _g(8)  # [[week], [week], ...] from Article shelf col C
        self.ah_promo            = _g(9)  # AH-Promo rows A:Z — for promo plan check
        self.assortment          = _g(10) # Assortment A:F — article ID (col B) + name (col E)
