"""
Stock Validation Dashboard
Streamlit entrypoint. See README.md for the full file list this app understands.
"""
from __future__ import annotations

import datetime as dt
import os
import sys

# --------------------------------------------------------------------------- #
# Make sure this app's own directory is on sys.path, regardless of what
# working directory Streamlit Cloud (or any other host) launches from.
# This is what fixes "ModuleNotFoundError: No module named 'src'".
# --------------------------------------------------------------------------- #
_APP_DIR = os.path.dirname(os.path.abspath(__file__))
if _APP_DIR not in sys.path:
    sys.path.insert(0, _APP_DIR)

import pandas as pd
import streamlit as st

try:
    from src import excel_export, readers, validation
    from src.constants import CATEGORY_LABELS, MARKETPLACES, REMARK_STREAMLIT_COLORS
    from src.file_detection import DetectedFiles, detect_files
except ModuleNotFoundError as e:
    st.error(
        "Could not import the `src` package. This almost always means the `src/` "
        "folder (with its 5 .py files, including `__init__.py`) wasn't fully pushed "
        "to your GitHub repo, or your Streamlit Cloud app's 'Main file path' isn't "
        "pointing at this app.py's own repo root.\n\n"
        f"Underlying error: `{e}`\n\n"
        "**Check on GitHub.com** that your repo has this exact layout:\n\n"
        "```\n"
        "your-repo/\n"
        "├── app.py\n"
        "├── requirements.txt\n"
        "└── src/\n"
        "    ├── __init__.py\n"
        "    ├── constants.py\n"
        "    ├── file_detection.py\n"
        "    ├── readers.py\n"
        "    ├── validation.py\n"
        "    └── excel_export.py\n"
        "```\n\n"
        "If `src/` is missing or any file inside it shows 0 files when you open the "
        "folder on GitHub, re-upload that folder (drag-and-drop uploads sometimes "
        "silently skip empty or nested files) and reboot the app from "
        "**Manage app → Reboot**."
    )
    st.stop()

st.set_page_config(page_title="Stock Validation Dashboard", page_icon="📦", layout="wide")


# --------------------------------------------------------------------------- #
# Sidebar: uploads
# --------------------------------------------------------------------------- #

st.sidebar.title("📦 Stock Validation")
st.sidebar.caption(
    "Upload any subset of your reports below. Files are matched automatically by "
    "filename — you don't need to sort them yourself."
)

uploaded = st.sidebar.file_uploader(
    "Drop all report files here",
    accept_multiple_files=True,
    type=["csv", "xlsx", "xls"],
)

flag_max_zero = st.sidebar.checkbox(
    "Flag REMOVE MAX (Max Stock = 0 while real stock exists)",
    value=False,
    help="Off by default. When on, rows where the StockValidation file's 'Max Stock' "
    "column is 0 but Expected or Marketplace stock is > 0 get a REMOVE MAX remark "
    "instead of the usual mismatch remark — this catches listing caps silently "
    "blocking sales.",
)

run_clicked = st.sidebar.button("▶ Run Validation", type="primary", use_container_width=True)


# --------------------------------------------------------------------------- #
# Header
# --------------------------------------------------------------------------- #

st.title("📦 Stock Validation Dashboard")
st.caption(
    "Reconciles Expected Stock against live marketplace stock for Lazada, Shopee, "
    "TikTok, and Zalora, flags mismatches, and exports a formatted Excel workbook."
)

if not uploaded:
    st.info("👈 Upload your report files in the sidebar to get started.")
    with st.expander("What files does this app recognise?"):
        st.markdown(
            """
| File | Detected from filename containing |
|---|---|
| Product Master | `product master` |
| ALL Report | filename starting with `ALL` |
| SOH Report | `soh` |
| Lazada Price & Stock Report | `pricestock` |
| Shopee Mass Update Report | `mass_update_sales_info` |
| TikTok Batch Edit Report | `batchedit` (ACTIVE + INACTIVE files auto-merged) |
| Zalora Stock Report | `sellerstocktemplate` |
| Lazada Stock Validation Report | `stockvalidation` + `lazada` |
| Shopee Stock Validation Report | `stockvalidation` + `shopee` |
| TikTok Stock Validation Report | `stockvalidation` + `tiktok` |
| Zalora Stock Validation Report | `stockvalidation` + `zalora` |
            """
        )
    st.stop()


# --------------------------------------------------------------------------- #
# Detection + confirmation table
# --------------------------------------------------------------------------- #

detected: DetectedFiles = detect_files(uploaded)

st.subheader("1. Detected files")
summary_rows = detected.as_summary_rows()
if summary_rows:
    st.dataframe(pd.DataFrame(summary_rows), use_container_width=True, hide_index=True)
else:
    st.warning("No recognised files yet — check filenames against the table above.")

if detected.unrecognised:
    with st.expander(f"⚠️ {len(detected.unrecognised)} file(s) not recognised"):
        for f in detected.unrecognised:
            st.write(f"- `{f.name}`")
        st.caption(
            "These were not matched to any known report type by filename and were "
            "ignored. Rename them to include a recognisable keyword (see the table "
            "in the empty-state above) and re-upload if they should be included."
        )

marketplaces_ready = []
for mkt, sv_attr, mkt_attr in [
    ("Lazada", "sv_lazada", "mkt_lazada"),
    ("Shopee", "sv_shopee", "mkt_shopee"),
    ("TikTok", "sv_tiktok", "mkt_tiktok"),
    ("Zalora", "sv_zalora", "mkt_zalora"),
]:
    sv_file = getattr(detected, sv_attr)
    mkt_file = getattr(detected, mkt_attr)
    has_mkt = bool(mkt_file) if isinstance(mkt_file, list) else mkt_file is not None
    if sv_file is not None and has_mkt:
        marketplaces_ready.append(mkt)

if marketplaces_ready:
    st.success(f"Ready to validate: **{', '.join(marketplaces_ready)}**")
else:
    st.warning(
        "No marketplace has both a Stock Validation Report and its matching stock "
        "file yet — upload both files for at least one marketplace to run validation."
    )

if not run_clicked:
    st.caption("Set your options in the sidebar, then click **▶ Run Validation**.")
    st.stop()

if not marketplaces_ready:
    st.error("Can't run validation yet — see the warning above.")
    st.stop()


# --------------------------------------------------------------------------- #
# Run validation
# --------------------------------------------------------------------------- #

with st.spinner("Reading files and validating stock..."):
    errors = []
    reference_frames = []

    if detected.product_master is not None:
        try:
            reference_frames.append(readers.read_reference_file(detected.product_master))
        except Exception as e:
            errors.append(f"Product Master: {e}")
    if detected.all_report is not None:
        try:
            reference_frames.append(readers.read_reference_file(detected.all_report))
        except Exception as e:
            errors.append(f"ALL Report: {e}")
    if detected.soh_report is not None:
        try:
            reference_frames.append(readers.read_reference_file(detected.soh_report))
        except Exception as e:
            errors.append(f"SOH Report: {e}")

    reference_df = (
        pd.concat(reference_frames, ignore_index=True).drop_duplicates(subset="SKU")
        if reference_frames
        else pd.DataFrame(columns=["SKU"])
    )

    results = {}
    cross_checks = {}

    # Lazada
    if "Lazada" in marketplaces_ready:
        try:
            sv = readers.read_stock_validation_csv(detected.sv_lazada)
            mkt = readers.read_lazada_pricestock(detected.mkt_lazada)
            cross_checks["Lazada"] = validation.cross_check_reference(sv, reference_df)
            results["Lazada"] = validation.validate_marketplace("Lazada", sv, mkt, flag_max_zero)
        except Exception as e:
            errors.append(f"Lazada: {e}")

    # Shopee
    if "Shopee" in marketplaces_ready:
        try:
            sv = readers.read_stock_validation_csv(detected.sv_shopee)
            mkt = readers.read_shopee_mass_update(detected.mkt_shopee)
            if detected.sv_shopee_delist is not None:
                delist_skus = readers.read_shopee_delist(detected.sv_shopee_delist)
                mkt = validation.apply_shopee_delist(mkt, delist_skus)
            cross_checks["Shopee"] = validation.cross_check_reference(sv, reference_df)
            results["Shopee"] = validation.validate_marketplace("Shopee", sv, mkt, flag_max_zero)
        except Exception as e:
            errors.append(f"Shopee: {e}")

    # TikTok
    if "TikTok" in marketplaces_ready:
        try:
            sv = readers.read_stock_validation_csv(detected.sv_tiktok)
            mkt = readers.read_tiktok_batchedit(detected.mkt_tiktok)
            cross_checks["TikTok"] = validation.cross_check_reference(sv, reference_df)
            results["TikTok"] = validation.validate_marketplace("TikTok", sv, mkt, flag_max_zero)
        except Exception as e:
            errors.append(f"TikTok: {e}")

    # Zalora
    if "Zalora" in marketplaces_ready:
        try:
            sv = readers.read_stock_validation_csv(detected.sv_zalora)
            mkt = readers.read_zalora_stock(detected.mkt_zalora, detected.mkt_zalora_status)
            cross_checks["Zalora"] = validation.cross_check_reference(sv, reference_df)
            results["Zalora"] = validation.validate_marketplace("Zalora", sv, mkt, flag_max_zero)
        except Exception as e:
            errors.append(f"Zalora: {e}")

if errors:
    st.error("Some marketplaces couldn't be processed:")
    for e in errors:
        st.write(f"- {e}")

if not results:
    st.stop()

summary_df = validation.summarize(results)

# --------------------------------------------------------------------------- #
# Dashboard
# --------------------------------------------------------------------------- #

st.subheader("2. Summary dashboard")

total_skus = int(summary_df["Total SKUs"].sum()) if not summary_df.empty else 0
total_issues = int(summary_df["Total Issues"].sum()) if not summary_df.empty else 0
overall_accuracy = (
    round((summary_df["GOOD"].sum() / total_skus) * 100, 1) if total_skus else 0.0
)

kpi_cols = st.columns(4)
kpi_cols[0].metric("Total SKUs validated", f"{total_skus:,}")
kpi_cols[1].metric("Marketplaces covered", len(results))
kpi_cols[2].metric("Total issues flagged", f"{total_issues:,}")
kpi_cols[3].metric("Overall accuracy", f"{overall_accuracy}%")

if reference_frames:
    with st.expander("Reference file cross-check (Product Master / ALL / SOH presence-check)"):
        for mkt, cc in cross_checks.items():
            if cc.get("checked"):
                st.write(
                    f"**{mkt}**: {cc['found_in_reference']}/{cc['sv_sku_count']} SKUs found "
                    f"in the uploaded reference file(s) ({cc['match_rate']*100:.1f}%)."
                )

st.dataframe(summary_df, use_container_width=True, hide_index=True)

if not summary_df.empty:
    chart_df = summary_df.set_index("Marketplace")[
        ["GOOD", "MISMATCH STOCK", "UPDATE 0", "REMOVE MAX", "NOT IN MARKETPLACE"]
    ]
    st.bar_chart(chart_df)


# --------------------------------------------------------------------------- #
# Per-marketplace detail tabs
# --------------------------------------------------------------------------- #

st.subheader("3. Marketplace detail")


def _style_remark(val: str) -> str:
    color = REMARK_STREAMLIT_COLORS.get(val)
    if not color and isinstance(val, str) and val.startswith("NOT IN"):
        color = REMARK_STREAMLIT_COLORS["NOT IN"]
    return f"background-color: {color}" if color else ""


tabs = st.tabs(list(results.keys()))
for tab, mkt in zip(tabs, results.keys()):
    with tab:
        df = results[mkt]
        remark_filter = st.multiselect(
            f"Filter {mkt} remarks",
            options=sorted(df["Remark"].unique().tolist()),
            default=[],
            key=f"filter_{mkt}",
        )
        view_df = df[df["Remark"].isin(remark_filter)] if remark_filter else df
        display_cols = [c for c in view_df.columns if c != "Marketplace"]
        styled = view_df[display_cols].style.map(_style_remark, subset=["Remark"])
        st.dataframe(styled, use_container_width=True, hide_index=True)


# --------------------------------------------------------------------------- #
# Export
# --------------------------------------------------------------------------- #

st.subheader("4. Export")

run_meta = {
    "Generated": dt.datetime.now().strftime("%Y-%m-%d %H:%M"),
    "Marketplaces": ", ".join(results.keys()),
    "REMOVE MAX flagging": "On" if flag_max_zero else "Off",
}
workbook_bytes = excel_export.build_workbook(summary_df, results, run_meta)

st.download_button(
    label="⬇ Download Excel workbook",
    data=workbook_bytes,
    file_name=f"Stock_Validation_{dt.datetime.now().strftime('%Y%m%d_%H%M')}.xlsx",
    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    type="primary",
    use_container_width=True,
)
