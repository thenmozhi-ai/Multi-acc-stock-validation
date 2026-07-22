# Stock Validation Dashboard

A Streamlit app that reconciles **Expected Stock** (from Product Master / SOH / the
marketplace's own StockValidation export) against **live marketplace stock** for Lazada, Shopee,
TikTok, and Zalora — flags mismatches, and exports a formatted, colour-coded Excel workbook.

## What it does

1. You drag-and-drop all your report files into one uploader (or use the per-file uploaders in
   the sidebar — either works).
2. The app **auto-detects** what each file is from its filename (and, if that's ambiguous, from
   its columns) and shows you a confirmation table before running anything.
3. It builds an **Expected Stock** figure per SKU (preferring the StockValidation CSV's own
   `Expected Stock` column, cross-checked against Product Master / SOH when those are uploaded),
   compares it to each marketplace's live stock, and assigns a remark:
   - `NOT IN <MARKETPLACE>` — SKU missing from that marketplace's own export
   - `UPDATE 0` — Expected Stock is 0 but the marketplace still shows stock > 0
   - `MISMATCH STOCK` — Expected Stock and Marketplace Stock disagree
   - `GOOD` — they match
   - `REMOVE MAX` (optional toggle) — `Max Stock` is 0 while real stock exists (a listing cap is
     silently blocking sales)
4. It shows a live summary dashboard (KPIs + charts) in the browser.
5. It exports everything to a single `.xlsx` with a `Summary` sheet plus one sheet per
   marketplace actually uploaded, colour-coded, frozen header, autofiltered.

## Supported input files

| File | How it's detected |
|---|---|
| Product Master | filename contains `product master` / `productmaster` |
| ALL Report | filename starts with `ALL` (functionally the same role as Product Master — merged) |
| SOH Report | filename contains `soh` (e.g. `SOHbySKU...xls`) |
| Lazada Price & Stock Report | filename contains `pricestock` |
| Shopee Mass Update Report | filename contains `mass_update_sales_info` |
| TikTok Batch Edit Report | filename contains `batchedit` (may be split into ACTIVE + INACTIVE files — both are auto-merged) |
| Zalora Stock Report | filename contains `sellerstocktemplate` |
| Lazada Stock Validation Report | filename contains `stockvalidation-lazada` (or `stockvalidation` + `lazada`) |
| Shopee Stock Validation Report | filename contains `stockvalidation-shopee` |
| TikTok Stock Validation Report | filename contains `stockvalidation-tiktok` |
| Zalora Stock Validation Report | filename contains `stockvalidation-zalora` |

None of these are individually required — upload whatever subset you have. A marketplace only
gets a sheet in the output once **both** its StockValidation CSV and its own marketplace stock
file are present.

## Project structure

```
stock-validation-app/
├── app.py                    # Streamlit entrypoint / UI
├── requirements.txt
├── .streamlit/
│   └── config.toml           # theme
└── src/
    ├── __init__.py
    ├── constants.py           # column aliases, colours, styling constants
    ├── file_detection.py      # filename/column-based classification of uploads
    ├── readers.py             # per-report-type pandas readers (handles header offsets, the
    │                          #   Shopee activePane bug, variable Lazada column counts, etc.)
    ├── validation.py          # Expected Stock resolution + remark logic
    └── excel_export.py        # openpyxl workbook builder (Summary + per-marketplace sheets)
```

## Running locally

```bash
pip install -r requirements.txt
streamlit run app.py
```

## Notes / assumptions

- Column names inside each report can vary slightly between exports, so every reader matches
  columns **by case-insensitive substring** (e.g. anything containing `seller sku` or `sku`)
  rather than by fixed position — this is the same defensive approach used in the underlying
  stock-validation skill this app is based on.
- SKUs are always treated as strings (to preserve leading zeros) and are stripped of whitespace
  before matching.
- Shopee bundle SKUs (containing `+`, e.g. `4975479496295+THE246`) are excluded from mismatch
  reporting — they're combo listings that don't map 1:1 to a single SKU.
- Nothing is ever written back to your uploaded files; the app only reads them in-memory.
