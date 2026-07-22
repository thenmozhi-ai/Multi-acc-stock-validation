"""
Core validation logic: resolves Expected Stock and assigns a remark per SKU per
marketplace, following the same priority rules as the underlying stock-validation
skill (stock-only -- no listing-status recommendation).
"""
from __future__ import annotations

import pandas as pd

from src.constants import (
    REMARK_GOOD,
    REMARK_MISMATCH,
    REMARK_NOT_IN,
    REMARK_REMOVE_MAX,
    REMARK_UPDATE_0,
)


def cross_check_reference(sv_df: pd.DataFrame, reference_df: pd.DataFrame) -> dict:
    """Presence-check only: what fraction of StockValidation SKUs are also found in
    the optional Product Master / SOH reference file. Does NOT change Expected Stock --
    the StockValidation CSV's own 'Expected Stock' column is already correctly derived
    downstream, per the underlying skill."""
    if reference_df is None or reference_df.empty or sv_df.empty:
        return {"checked": False}
    ref_skus = set(reference_df["SKU"])
    sv_skus = set(sv_df["SKU"])
    found = sv_skus & ref_skus
    return {
        "checked": True,
        "sv_sku_count": len(sv_skus),
        "found_in_reference": len(found),
        "match_rate": (len(found) / len(sv_skus)) if sv_skus else 0.0,
    }


def validate_marketplace(
    marketplace: str,
    sv_df: pd.DataFrame,
    mkt_df: pd.DataFrame,
    flag_max_zero: bool = False,
) -> pd.DataFrame:
    """
    Joins a marketplace's StockValidation export against its own live-stock export
    and assigns a remark per SKU, in priority order:

      1. SKU not found in the marketplace's own stock export -> NOT IN <MARKETPLACE>
      2. (optional) Max Stock == 0 and (Expected or Marketplace stock) > 0 -> REMOVE MAX
      3. Expected Stock == 0 and Marketplace Stock > 0 -> UPDATE 0
      4. Expected Stock != Marketplace Stock -> MISMATCH STOCK
      5. otherwise -> GOOD
    """
    if sv_df is None or sv_df.empty:
        return pd.DataFrame()

    merged = sv_df.merge(mkt_df, on="SKU", how="left", suffixes=("", "_mkt"))
    if "Marketplace Stock" not in merged.columns:
        merged["Marketplace Stock"] = pd.NA
    if "Marketplace Status" not in merged.columns:
        merged["Marketplace Status"] = ""

    def remark_row(row):
        if pd.isna(row["Marketplace Stock"]):
            return REMARK_NOT_IN.format(mkt=marketplace.upper())

        expected = row.get("Expected Stock", 0) or 0
        mkt_stock = row["Marketplace Stock"] or 0

        if flag_max_zero:
            max_stock = row.get("Max Stock")
            if pd.notna(max_stock) and max_stock == 0 and (expected > 0 or mkt_stock > 0):
                return REMARK_REMOVE_MAX

        if expected == 0 and mkt_stock > 0:
            return REMARK_UPDATE_0
        if expected != mkt_stock:
            return REMARK_MISMATCH
        return REMARK_GOOD

    merged["Difference"] = merged["Marketplace Stock"].fillna(0) - merged.get("Expected Stock", 0)
    merged["Remark"] = merged.apply(remark_row, axis=1)
    merged["Marketplace"] = marketplace

    cols = [
        "Marketplace",
        "SKU",
        "Item Title",
        "Expected Stock",
        "Marketplace Stock",
        "Difference",
    ]
    if "Max Stock" in merged.columns:
        cols.append("Max Stock")
    if "Marketplace Status" in merged.columns:
        cols.append("Marketplace Status")
    cols.append("Remark")
    cols = [c for c in cols if c in merged.columns]
    return merged[cols]


def apply_shopee_delist(mkt_df: pd.DataFrame, delist_skus: set) -> pd.DataFrame:
    """Marks SKUs found in the Shopee _DELIST export as INACTIVE, everything else ACTIVE."""
    if not delist_skus or mkt_df.empty:
        return mkt_df
    mkt_df = mkt_df.copy()
    mkt_df["Marketplace Status"] = mkt_df["SKU"].apply(
        lambda s: "INACTIVE" if s in delist_skus else "ACTIVE"
    )
    return mkt_df


def summarize(results: dict) -> pd.DataFrame:
    """Builds the KPI summary table (one row per marketplace) from the
    {marketplace: result_df} dict produced by validate_marketplace calls."""
    rows = []
    for mkt, df in results.items():
        if df is None or df.empty:
            continue
        total = len(df)
        good = (df["Remark"] == REMARK_GOOD).sum()
        mismatch = (df["Remark"] == REMARK_MISMATCH).sum()
        update0 = (df["Remark"] == REMARK_UPDATE_0).sum()
        remove_max = (df["Remark"] == REMARK_REMOVE_MAX).sum()
        not_in = df["Remark"].str.startswith("NOT IN").sum()
        issues = total - good
        accuracy = (good / total * 100) if total else 0.0
        rows.append(
            {
                "Marketplace": mkt,
                "Total SKUs": total,
                "GOOD": int(good),
                "MISMATCH STOCK": int(mismatch),
                "UPDATE 0": int(update0),
                "REMOVE MAX": int(remove_max),
                "NOT IN MARKETPLACE": int(not_in),
                "Total Issues": int(issues),
                "Accuracy %": round(accuracy, 1),
            }
        )
    return pd.DataFrame(rows)
