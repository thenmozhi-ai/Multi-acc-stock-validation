"""
Classifies uploaded files into known report categories using filename keywords,
falling back to a light column-sniff when the filename alone is ambiguous.

Detection is intentionally permissive: it never raises on an unrecognised file, it just
leaves it out of the mapping (surfaced to the user as "unrecognised").
"""
from __future__ import annotations

from dataclasses import dataclass, field
from typing import Optional

import pandas as pd

from src.constants import CATEGORY_LABELS


@dataclass
class DetectedFiles:
    """Holds the classified uploads. Single-file categories store one UploadedFile;
    tiktok marketplace stock can legitimately be split ACTIVE/INACTIVE, so it's a list."""
    product_master: Optional[object] = None
    all_report: Optional[object] = None
    soh_report: Optional[object] = None
    mkt_lazada: Optional[object] = None
    mkt_shopee: Optional[object] = None
    mkt_tiktok: list = field(default_factory=list)
    mkt_zalora: Optional[object] = None
    mkt_zalora_status: Optional[object] = None  # optional SellerStatusTemplate
    sv_lazada: Optional[object] = None
    sv_shopee: Optional[object] = None
    sv_shopee_delist: Optional[object] = None
    sv_tiktok: Optional[object] = None
    sv_zalora: Optional[object] = None
    unrecognised: list = field(default_factory=list)

    def as_summary_rows(self) -> list[dict]:
        """Flat list of {category, label, filename} for the confirmation table."""
        rows = []

        def add(cat_key, file_obj):
            if file_obj is None:
                return
            rows.append(
                {
                    "Detected as": CATEGORY_LABELS.get(cat_key, cat_key),
                    "Filename": getattr(file_obj, "name", str(file_obj)),
                }
            )

        add("product_master", self.product_master)
        add("all_report", self.all_report)
        add("soh_report", self.soh_report)
        add("mkt_lazada", self.mkt_lazada)
        add("mkt_shopee", self.mkt_shopee)
        for f in self.mkt_tiktok:
            add("mkt_tiktok", f)
        add("mkt_zalora", self.mkt_zalora)
        if self.mkt_zalora_status is not None:
            rows.append(
                {
                    "Detected as": "Zalora Status Template (optional)",
                    "Filename": getattr(self.mkt_zalora_status, "name", "?"),
                }
            )
        add("sv_lazada", self.sv_lazada)
        add("sv_shopee", self.sv_shopee)
        if self.sv_shopee_delist is not None:
            rows.append(
                {
                    "Detected as": "Shopee Stock Validation (DELIST subset)",
                    "Filename": getattr(self.sv_shopee_delist, "name", "?"),
                }
            )
        add("sv_tiktok", self.sv_tiktok)
        add("sv_zalora", self.sv_zalora)
        return rows


def _name(f) -> str:
    return getattr(f, "name", str(f)).lower()


def detect_files(uploaded_files: list) -> DetectedFiles:
    """Classify a flat list of uploaded file objects (Streamlit UploadedFile-like:
    must expose .name) into a DetectedFiles bundle."""
    det = DetectedFiles()

    for f in uploaded_files:
        n = _name(f)

        # --- StockValidation CSVs (check these before the generic marketplace files,
        # since e.g. "stockvalidation-lazada" also contains "lazada") ---
        if "stockvalidation" in n.replace("_", "").replace(" ", ""):
            if "lazada" in n:
                det.sv_lazada = f
            elif "shopee" in n:
                if "delist" in n:
                    det.sv_shopee_delist = f
                else:
                    det.sv_shopee = f
            elif "tiktok" in n:
                det.sv_tiktok = f
            elif "zalora" in n:
                det.sv_zalora = f
            else:
                det.unrecognised.append(f)
            continue

        # --- Marketplace stock/status files ---
        if "pricestock" in n:
            det.mkt_lazada = f
            continue
        if "mass_update_sales_info" in n or "massupdatesalesinfo" in n:
            det.mkt_shopee = f
            continue
        if "batchedit" in n:
            det.mkt_tiktok.append(f)
            continue
        if "sellerstocktemplate" in n:
            det.mkt_zalora = f
            continue
        if "sellerstatustemplate" in n:
            det.mkt_zalora_status = f
            continue

        # --- Reference / master files ---
        if "productmaster" in n.replace(" ", "").replace("_", ""):
            det.product_master = f
            continue
        if "soh" in n:
            det.soh_report = f
            continue
        if n.startswith("all") or n.startswith("all_") or n.startswith("all-") or n.startswith("all "):
            det.all_report = f
            continue

        det.unrecognised.append(f)

    return det


def sniff_columns(f) -> list[str]:
    """Best-effort peek at a file's header row, used only for diagnostics /
    disambiguation messages shown to the user -- never required for detection to work."""
    try:
        name = _name(f)
        f.seek(0)
        if name.endswith(".csv"):
            df = pd.read_csv(f, nrows=0)
        else:
            df = pd.read_excel(f, nrows=0)
        f.seek(0)
        return list(df.columns)
    except Exception:
        try:
            f.seek(0)
        except Exception:
            pass
        return []
