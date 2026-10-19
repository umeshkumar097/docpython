"""Employee folder filter.

Filter the rows of the doc-status Excel (Phase, State, Batch, Employer, dates,
pasted IDs ...) and copy / ZIP only the export folders whose name is the
employee ID of a matching row.

Run:  streamlit run app.py
"""

from __future__ import annotations

import hashlib
import io
import subprocess
import sys
from pathlib import Path

import pandas as pd
import streamlit as st

import folder_filter as ff

ZIP_DOWNLOAD_LIMIT = 500 * 1024 * 1024  # bigger ZIPs stay on disk only
FILTERS_PER_ROW = 3

st.set_page_config(page_title="Employee folder filter", layout="wide")


@st.cache_data(show_spinner="Reading Excel ...")
def read_excel(data: bytes) -> dict[str, pd.DataFrame]:
    return ff.load_workbook(io.BytesIO(data))


def widget_key(*parts: object) -> str:
    """Stable key that changes when a widget's options change, so a stale
    selection is dropped instead of breaking the widget."""
    return "w" + hashlib.sha1("|".join(map(str, parts)).encode()).hexdigest()[:12]


def open_in_file_manager(path: Path) -> None:
    if sys.platform == "darwin":
        subprocess.run(["open", str(path)], check=False)
    elif sys.platform.startswith("win"):
        subprocess.run(["explorer", str(path)], check=False)
    else:
        subprocess.run(["xdg-open", str(path)], check=False)


def to_xlsx_bytes(df: pd.DataFrame) -> bytes:
    buf = io.BytesIO()
    df.to_excel(buf, index=False, engine="openpyxl")
    return buf.getvalue()


def column_filter_widget(df: pd.DataFrame, col: str, ignore_case: bool) -> ff.ColumnFilter:
    series = df[col]
    if ff.is_date_column(col, series):
        days = ff.parse_dates(series).dropna().dt.date
        lo, hi = days.min(), days.max()
        picked = st.date_input(
            col, value=(lo, hi), min_value=lo, max_value=hi, format="DD/MM/YYYY",
            key=widget_key("date", col, lo, hi),
        )
        picked = tuple(picked) if isinstance(picked, (tuple, list)) else (picked,)
        start = picked[0] if picked and picked[0] > lo else None
        end = picked[1] if len(picked) > 1 and picked[1] < hi else None
        return ff.ColumnFilter(col, date_from=start, date_to=end)

    options = ff.value_options(series, ignore_case)
    if len(options) > ff.MAX_LIST_VALUES:
        text = st.text_input(
            f"{col} contains", key=widget_key("contains", col),
            help="Comma se alag karo; koi bhi ek match ho to row aayegi.",
        )
        return ff.ColumnFilter(col, contains=[t for t in text.split(",") if t.strip()])

    labels = {key: f"{label}  ({count})" for key, label, count in options}
    picked = st.multiselect(
        col, list(labels), format_func=labels.__getitem__, placeholder="Any value",
        key=widget_key("values", col, ignore_case, tuple(labels)),
    )
    return ff.ColumnFilter(col, values=picked)


# --------------------------------------------------------------------------
# 1-2. Inputs
# --------------------------------------------------------------------------

st.title("Employee folder filter")
st.caption(
    "Excel mein filter lagao: jitne employees filter mein aayenge, sirf unke "
    "folders copy / ZIP honge. Original export folder kabhi change nahi hota."
)

with st.sidebar:
    st.header("1. Excel")
    upload = st.file_uploader("Doc status Excel", type=["xlsx", "xlsm"])

    st.header("2. Export folder")
    folder_raw = st.text_input(
        "Folder path",
        placeholder="/Users/you/Downloads/Export-2026-09-23 (1)",
        help=(
            "Mac: Finder mein folder par right-click karo, Option key dabaye rakho, "
            "'Copy ... as Pathname' chuno aur yahan paste karo. Folder ko Terminal "
            "mein drag karke bhi path mil jata hai."
        ),
    )

if upload is None:
    st.info("Left side se Excel file upload karo.")
    st.stop()

sheets = read_excel(upload.getvalue())

with st.sidebar:
    sheet_names = st.multiselect("Sheets", list(sheets), default=list(sheets))
    if not sheet_names:
        st.warning("Kam se kam ek sheet chuno.")
        st.stop()
    df = ff.combine_sheets(sheets, sheet_names)

    columns = [c for c in df.columns if c != ff.SHEET_COLUMN]
    guessed = ff.guess_id_column(columns)
    id_col = st.selectbox(
        "Employee ID column", columns, index=columns.index(guessed) if guessed else 0,
        help="Is column ki value folder ke naam se match hoti hai.",
    )
    ignore_case = st.checkbox(
        "Ignore case and extra spaces", value=True,
        help="'Home credit ' aur 'HOME CREDIT' ek hi value maane jayenge.",
    )

folder: Path | None = None
index: dict[str, list[Path]] | None = None
entries: list[Path] = []
if folder_raw.strip():
    folder = ff.clean_path_input(folder_raw)
    if folder.is_dir():
        index = ff.scan_export(folder)
        entries = sorted({p for paths in index.values() for p in paths})
        st.sidebar.success(f"{len(entries)} folders / files found")
    else:
        st.sidebar.error(f"Folder nahi mila: {folder}")

# --------------------------------------------------------------------------
# 3. Filters
# --------------------------------------------------------------------------

st.subheader("3. Filters")
filter_cols = st.multiselect(
    "Filter by columns", list(df.columns), placeholder="Columns chuno (Phase, State, Batch ...)",
    help="Alag columns ke filters AND hote hain; ek column mein chuni values OR hoti hain.",
)

filters: list[ff.ColumnFilter] = []
for start in range(0, len(filter_cols), FILTERS_PER_ROW):
    row_cols = st.columns(FILTERS_PER_ROW)
    for slot, col in zip(row_cols, filter_cols[start:start + FILTERS_PER_ROW]):
        with slot:
            filters.append(column_filter_widget(df, col, ignore_case))

left, right = st.columns([2, 1])
with left:
    pasted = ff.parse_id_list(st.text_area(
        "Employee IDs (optional)", height=90,
        placeholder="80013618, 80011485 ...  (ek line mein ek, ya comma / space se alag)",
    ))
with right:
    folder_state = st.radio(
        "Folder status", ["All", "Folder found", "Folder missing"], horizontal=True,
        disabled=index is None,
        help="Export folder ka path dene ke baad kaam karta hai.",
    )

keys = ff.id_keys(df, id_col)
mask = ff.filter_mask(df, filters, ignore_case)
if pasted:
    mask &= keys.isin(pasted)
has_folder = keys.map(lambda k: index is not None and k in index)
if index is not None and folder_state == "Folder found":
    mask &= has_folder
elif index is not None and folder_state == "Folder missing":
    mask &= ~has_folder

rows = df[mask]
row_keys = keys[mask]
unique_ids = [k for k in dict.fromkeys(row_keys) if k]

# --------------------------------------------------------------------------
# Result
# --------------------------------------------------------------------------

stats = st.columns(4)
stats[0].metric("Rows matching", len(rows))
stats[1].metric("Employee IDs", len(unique_ids))
if index is not None:
    found_ids = [k for k in unique_ids if k in index]
    stats[2].metric("Folders found", len(found_ids))
    stats[3].metric("Folders missing", len(unique_ids) - len(found_ids))

view = rows.copy()
if index is not None:
    view.insert(
        view.columns.get_loc(id_col) + 1, ff.FOLDER_FOUND_COLUMN,
        has_folder[mask].map({True: "Yes", False: "No"}),
    )
st.dataframe(view, hide_index=True, width="stretch", height=380)

if pasted:
    unknown = sorted(set(pasted) - set(keys))
    if unknown:
        st.warning(f"{len(unknown)} pasted IDs Excel mein nahi hain: {', '.join(unknown)}")

if index is not None:
    missing_view = view[view[ff.FOLDER_FOUND_COLUMN] == "No"]
    if len(missing_view):
        with st.expander(f"Excel mein hain par folder nahi mila ({len(missing_view)})"):
            st.dataframe(missing_view, hide_index=True, width="stretch")
    all_keys = set(keys)
    orphans = [p.name for p in entries if not ff.entry_keys(p) & all_keys]
    if orphans:
        with st.expander(f"Folder hai par Excel ({id_col}) mein ID nahi mili ({len(orphans)})"):
            st.write(", ".join(orphans))

st.download_button(
    "Download filtered list (.xlsx)", to_xlsx_bytes(view), file_name="filtered_list.xlsx",
    disabled=rows.empty, on_click="ignore",
)

# --------------------------------------------------------------------------
# 4. Copy / ZIP
# --------------------------------------------------------------------------

st.subheader("4. Copy / download folders")
if index is None:
    st.info("Folders copy karne ke liye left side mein export folder ka path do.")
    st.stop()

sources = {p for k in unique_ids for p in index.get(k, [])}
out_raw = st.text_input(
    "Save to (new folder)", value=str(ff.default_output_dir(folder)),
    key=widget_key("out", folder),
    help="Ye folder naya banega; pehle se ho to naam ke aage (2), (3) lag jayega.",
)
make_zip = st.checkbox("Also make a ZIP", value=True)

selection = hashlib.sha1("|".join(sorted(unique_ids)).encode()).hexdigest()
if st.button(f"Copy {len(sources)} folders", type="primary", disabled=not sources):
    bar = st.progress(0.0, text="Copying ...")
    result = ff.export_selection(
        rows, id_col, index, ff.clean_path_input(out_raw), make_zip,
        progress=lambda i, total, name: bar.progress(i / total, text=f"{i}/{total}  {name}"),
    )
    bar.empty()
    st.session_state["export"] = (selection, result)

done = st.session_state.get("export")
if done and done[0] == selection:
    result = done[1]
    st.success(
        f"{len(result.copied)} folders copied to {result.out_dir}"
        + (f"  |  ZIP: {result.zip_path}" if result.zip_path else "")
    )
    if result.missing_ids:
        st.warning(f"{len(result.missing_ids)} IDs ka folder nahi mila (filtered_list.xlsx mein list hai).")
    buttons = st.columns([1, 1, 3])
    if buttons[0].button("Open folder"):
        open_in_file_manager(result.out_dir)
    if result.zip_path and result.zip_path.stat().st_size <= ZIP_DOWNLOAD_LIMIT:
        with result.zip_path.open("rb") as fh:
            buttons[1].download_button(
                "Download ZIP", fh, file_name=result.zip_path.name,
                mime="application/zip", on_click="ignore",
            )
    elif result.zip_path:
        buttons[1].caption("ZIP bada hai; upar diye path par saved hai.")
