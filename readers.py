"""
Per-report-type readers. Each returns a pandas DataFrame with a normalised
`SKU` (string) column plus whatever else that reader extracts.

Reads are defensive by design: columns are located by case-insensitive substring
match rather than fixed position/name, since exports vary between runs (extra
columns, renamed headers, etc.) -- mirroring the approach used in the underlying
stock-validation skill.
"""
from __future__ import annotations

import io
import re
import zipfile
from typing import Optional

import pandas as pd

from src.constants import COLUMN_ALIASES


# --------------------------------------------------------------------------- #
# Column-matching helpers
# --------------------------------------------------------------------------- #

def _find_column(df: pd.DataFrame, aliases: list[str]) -> Optional[str]:
    """Return the first column in df whose name contains any of the given
    case-insensitive substrings, longest alias first (so 'seller sku' wins over 'sku')."""
    cols = list(df.columns)
    for alias in sorted(aliases, key=len, reverse=True):
        for c in cols:
            if alias in str(c).strip().lower():
                return c
    return None


def find_col(df: pd.DataFrame, key: str) -> Optional[str]:
    """Look up a column by semantic key (see COLUMN_ALIASES)."""
    return _find_column(df, COLUMN_ALIASES.get(key, [key]))


def normalize_sku_series(s: pd.Series) -> pd.Series:
    return s.astype(str).str.strip()


def clean_skus(df: pd.DataFrame, sku_col: str) -> pd.DataFrame:
    df = df.copy()
    df[sku_col] = normalize_sku_series(df[sku_col])
    df = df[df[sku_col].notna() & (df[sku_col] != "") & (df[sku_col].str.lower() != "nan")]
    return df


def is_bundle_sku(sku: str) -> bool:
    """Shopee combo/bundle SKUs contain a '+' and don't map 1:1 to a single SKU."""
    return "+" in str(sku)


# --------------------------------------------------------------------------- #
# Shopee activePane bug patch
# --------------------------------------------------------------------------- #

def _patch_shopee_workbook(file_bytes: bytes) -> io.BytesIO:
    """
    Some Shopee mass_update_sales_info exports ship with an invalid
    `activePane` attribute inside xl/worksheets/sheetN.xml's <pane> element,
    which makes openpyxl/pandas choke on load. Patch it out in-memory by
    rewriting the offending attribute before parsing.
    """
    try:
        src = zipfile.ZipFile(io.BytesIO(file_bytes), "r")
    except zipfile.BadZipFile:
        # Not actually a zip-based xlsx (e.g. already CSV-like) -- return as-is.
        return io.BytesIO(file_bytes)

    out_buffer = io.BytesIO()
    with zipfile.ZipFile(out_buffer, "w", zipfile.ZIP_DEFLATED) as dst:
        for item in src.infolist():
            data = src.read(item.filename)
            if item.filename.startswith("xl/worksheets/") and item.filename.endswith(".xml"):
                text = data.decode("utf-8", errors="ignore")
                # Strip an invalid activePane reference (e.g. activePane="topLeft" combined
                # with a malformed pane split) that some Shopee exports emit.
                text = re.sub(r'\s*activePane="[^"]*"', "", text)
                data = text.encode("utf-8")
            dst.writestr(item, data)
    src.close()
    out_buffer.seek(0)
    return out_buffer


def _read_excel_any(f, **kwargs) -> pd.DataFrame:
    """Read an uploaded file as Excel regardless of whether it arrives as a
    Streamlit UploadedFile, file path, or raw bytes buffer."""
    if hasattr(f, "read"):
        f.seek(0)
        data = f.read()
        f.seek(0)
    else:
        data = f
    return pd.read_excel(io.BytesIO(data), **kwargs)


# --------------------------------------------------------------------------- #
# StockValidation CSV reader (shared shape across all 4 marketplaces)
# --------------------------------------------------------------------------- #

def read_stock_validation_csv(f) -> pd.DataFrame:
    """
    Reads a stockValidation-<marketplace>.csv file. Expected (flexibly-matched)
    columns: SKU / Seller SKU, Item Title, Expected Stock, Max Stock (optional),
    Status (optional), Master Stock (optional).
    """
    if hasattr(f, "seek"):
        f.seek(0)
    df = pd.read_csv(f)

    sku_col = find_col(df, "sku")
    if sku_col is None:
        raise ValueError("Could not find a SKU / Seller SKU column in this StockValidation file.")
    exp_col = find_col(df, "expected_stock")
    if exp_col is None:
        raise ValueError("Could not find an 'Expected Stock' column in this StockValidation file.")

    out = pd.DataFrame()
    out["SKU"] = df[sku_col]
    out["Expected Stock"] = pd.to_numeric(df[exp_col], errors="coerce").fillna(0)

    title_col = find_col(df, "item_title")
    out["Item Title"] = df[title_col] if title_col else ""

    max_col = find_col(df, "max_stock")
    out["Max Stock"] = pd.to_numeric(df[max_col], errors="coerce") if max_col else pd.NA

    status_col = find_col(df, "status")
    out["Status"] = df[status_col] if status_col else ""

    master_col = find_col(df, "master_stock")
    if master_col:
        out["Master Stock"] = pd.to_numeric(df[master_col], errors="coerce")

    out = clean_skus(out, "SKU")
    return out


# --------------------------------------------------------------------------- #
# Marketplace stock readers
# --------------------------------------------------------------------------- #

def read_lazada_pricestock(f) -> pd.DataFrame:
    """Header row 0, skip rows 1-3. Column count can vary (15 vs 16 -- some exports
    add a Barcode column) so columns are matched by name, not position."""
    df = _read_excel_any(f, header=0, skiprows=[1, 2, 3])
    sku_col = find_col(df, "sku")
    qty_col = find_col(df, "quantity")
    status_col = find_col(df, "status")
    if sku_col is None or qty_col is None:
        raise ValueError("Could not locate SellerSKU / Quantity columns in the Lazada pricestock file.")
    out = pd.DataFrame()
    out["SKU"] = df[sku_col]
    out["Marketplace Stock"] = pd.to_numeric(df[qty_col], errors="coerce").fillna(0)
    out["Marketplace Status"] = df[status_col] if status_col else ""
    out = clean_skus(out, "SKU")
    out = out.groupby("SKU", as_index=False).agg(
        {"Marketplace Stock": "sum", "Marketplace Status": "first"}
    )
    return out


def read_shopee_mass_update(f) -> pd.DataFrame:
    """Patches the activePane bug, then header row 2, skip rows 3-5.
    Matches SKU, falling back to Parent SKU when SKU is blank."""
    if hasattr(f, "seek"):
        f.seek(0)
    data = f.read() if hasattr(f, "read") else f
    patched = _patch_shopee_workbook(data)
    df = pd.read_excel(patched, header=2, skiprows=[3, 4, 5])

    sku_col = _find_column(df, ["sku"])
    parent_col = _find_column(df, ["parent sku"])
    qty_col = find_col(df, "quantity")
    if sku_col is None or qty_col is None:
        raise ValueError("Could not locate SKU / Stock columns in the Shopee mass_update file.")

    out = pd.DataFrame()
    sku_series = df[sku_col]
    if parent_col is not None:
        sku_series = sku_series.where(sku_series.notna() & (sku_series.astype(str).str.strip() != ""), df[parent_col])
    out["SKU"] = sku_series
    out["Marketplace Stock"] = pd.to_numeric(df[qty_col], errors="coerce").fillna(0)
    out = clean_skus(out, "SKU")

    # Exclude bundle SKUs (contain '+') -- combo listings, not 1:1 with a single SKU.
    out = out[~out["SKU"].apply(is_bundle_sku)]
    out = out.groupby("SKU", as_index=False).agg({"Marketplace Stock": "sum"})
    return out


def read_shopee_delist(f) -> set:
    """Reads a Shopee _DELIST export and returns the set of SKUs it contains
    (used to mark those SKUs INACTIVE relative to the full export)."""
    if hasattr(f, "seek"):
        f.seek(0)
    data = f.read() if hasattr(f, "read") else f
    patched = _patch_shopee_workbook(data)
    df = pd.read_excel(patched, header=2, skiprows=[3, 4, 5])
    sku_col = _find_column(df, ["sku"])
    if sku_col is None:
        return set()
    skus = normalize_sku_series(df[sku_col])
    return set(skus[skus.notna() & (skus != "")])


def read_tiktok_batchedit_single(f) -> pd.DataFrame:
    """Header row 2, skip rows 3-4."""
    df = _read_excel_any(f, header=2, skiprows=[3, 4])
    sku_col = find_col(df, "sku")
    qty_col = find_col(df, "quantity")
    if sku_col is None or qty_col is None:
        raise ValueError("Could not locate Seller SKU / Quantity columns in the TikTok batchedit file.")
    out = pd.DataFrame()
    out["SKU"] = df[sku_col]
    out["Marketplace Stock"] = pd.to_numeric(df[qty_col], errors="coerce").fillna(0)
    out = clean_skus(out, "SKU")
    return out


def read_tiktok_batchedit(files: list) -> pd.DataFrame:
    """Accepts one or two TikTok batchedit files (a single combined export, or a
    split ACTIVE + INACTIVE pair). If two, concatenates and sums duplicate SKUs."""
    frames = [read_tiktok_batchedit_single(f) for f in files]
    combined = pd.concat(frames, ignore_index=True) if frames else pd.DataFrame(columns=["SKU", "Marketplace Stock"])
    if combined.empty:
        return combined
    combined = combined.groupby("SKU", as_index=False).agg({"Marketplace Stock": "sum"})
    return combined


def read_zalora_stock(f, status_f=None) -> pd.DataFrame:
    """Header row 0. Optional SellerStatusTemplate file appends Zalora_Status."""
    df = _read_excel_any(f, header=0)
    sku_col = find_col(df, "sku")
    qty_col = find_col(df, "quantity")
    if sku_col is None or qty_col is None:
        raise ValueError("Could not locate SellerSku / Quantity columns in the Zalora stock file.")
    out = pd.DataFrame()
    out["SKU"] = df[sku_col]
    out["Marketplace Stock"] = pd.to_numeric(df[qty_col], errors="coerce").fillna(0)
    out = clean_skus(out, "SKU")

    if status_f is not None:
        try:
            sdf = _read_excel_any(status_f, header=0)
            s_sku_col = find_col(sdf, "sku")
            s_status_col = find_col(sdf, "status")
            if s_sku_col and s_status_col:
                sdf2 = pd.DataFrame(
                    {"SKU": normalize_sku_series(sdf[s_sku_col]), "Marketplace Status": sdf[s_status_col]}
                )
                out = out.merge(sdf2, on="SKU", how="left")
        except Exception:
            pass

    if "Marketplace Status" not in out.columns:
        out["Marketplace Status"] = ""

    out = out.groupby("SKU", as_index=False).agg(
        {"Marketplace Stock": "sum", "Marketplace Status": "first"}
    )
    return out


# --------------------------------------------------------------------------- #
# Reference file readers (Product Master / ALL report / SOH) -- optional,
# used only as a presence-check cross-reference against Expected Stock.
# --------------------------------------------------------------------------- #

def read_reference_file(f) -> pd.DataFrame:
    """Reads a Product Master / ALL report / SOH file generically: whatever SKU +
    quantity-like columns it can find. Used only for an optional presence-check,
    so failures here are non-fatal."""
    try:
        name = getattr(f, "name", "").lower()
        if hasattr(f, "seek"):
            f.seek(0)
        if name.endswith(".csv"):
            df = pd.read_csv(f)
        else:
            df = _read_excel_any(f)
        sku_col = find_col(df, "sku")
        if sku_col is None:
            return pd.DataFrame(columns=["SKU"])
        out = pd.DataFrame()
        out["SKU"] = normalize_sku_series(df[sku_col])
        out = out[out["SKU"].notna() & (out["SKU"] != "")]
        return out.drop_duplicates(subset="SKU")
    except Exception:
        return pd.DataFrame(columns=["SKU"])
