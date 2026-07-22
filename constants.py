"""Shared constants: marketplace list, remark labels, colours, styling."""

MARKETPLACES = ["Lazada", "Shopee", "TikTok", "Zalora"]

# File categories recognised by the detector. Each maps to a human label used in the UI.
CATEGORY_LABELS = {
    "product_master": "Product Master",
    "all_report": "ALL Report",
    "soh_report": "SOH Report",
    "mkt_lazada": "Lazada Price & Stock Report",
    "mkt_shopee": "Shopee Mass Update Report",
    "mkt_tiktok": "TikTok Batch Edit Report",
    "mkt_zalora": "Zalora Stock Report",
    "sv_lazada": "Lazada Stock Validation Report",
    "sv_shopee": "Shopee Stock Validation Report",
    "sv_tiktok": "TikTok Stock Validation Report",
    "sv_zalora": "Zalora Stock Validation Report",
}

# Remarks
REMARK_NOT_IN = "NOT IN {mkt}"
REMARK_UPDATE_0 = "UPDATE 0"
REMARK_MISMATCH = "MISMATCH STOCK"
REMARK_GOOD = "GOOD"
REMARK_REMOVE_MAX = "REMOVE MAX"

# Fill colours (hex, no leading '#') for openpyxl PatternFill, keyed by remark family
REMARK_FILL_COLORS = {
    "GOOD": "C6EFCE",
    "UPDATE 0": "FFEB9C",
    "MISMATCH STOCK": "FFC7CE",
    "NOT IN": "D9D9D9",       # prefix match for "NOT IN <MARKETPLACE>"
    "REMOVE MAX": "BDD7EE",
}

# Streamlit-side (CSS) colours, mirroring the Excel fills, for the on-screen preview
REMARK_STREAMLIT_COLORS = {
    "GOOD": "#C6EFCE",
    "UPDATE 0": "#FFEB9C",
    "MISMATCH STOCK": "#FFC7CE",
    "NOT IN": "#D9D9D9",
    "REMOVE MAX": "#BDD7EE",
}

HEADER_FILL_COLOR = "1F3864"   # navy
HEADER_FONT_COLOR = "FFFFFF"   # white
HEADER_FONT_NAME = "Arial"
HEADER_FONT_SIZE = 11
BODY_FONT_NAME = "Arial"
BODY_FONT_SIZE = 10
ALT_ROW_FILL_COLOR = "F2F2F2"

# Candidate column-name substrings (case-insensitive) used to locate columns by content
# rather than fixed position, since exports vary between runs.
COLUMN_ALIASES = {
    "sku": ["seller sku", "sellersku", "seller_sku", "sku"],
    "item_title": ["item title", "product name", "title"],
    "expected_stock": ["expected stock"],
    "master_stock": ["master stock"],
    "max_stock": ["max stock"],
    "status": ["status"],
    "quantity": ["quantity", "stock", "total"],
}

# Output workbook sheet name for the dashboard
SUMMARY_SHEET_NAME = "Summary"
