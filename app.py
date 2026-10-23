"""Employee documents filter.

1. Upload the doc-status Excel / CSV and the documents ZIP (one folder per
   Employee ID).
   Only employees whose folder is in the ZIP are kept; other rows are skipped.
2. Filter by Batch, Language, Education and Month (plus any other column).
3. Tick the rows you want and download one ZIP with only those employees'
   folders, in the same structure as the uploaded ZIP.

Run:  streamlit run app.py
"""

from __future__ import annotations

import hashlib
import io
import zipfile
from pathlib import Path
from typing import Callable
from urllib.parse import urlparse

import pandas as pd
import streamlit as st

import folder_filter as ff

SELECT_COLUMN = "Select"
FILES_COLUMN = "Files in ZIP"
# A browser download is built in memory; bigger selections are saved straight to disk.
BROWSER_ZIP_LIMIT = 1_000_000_000
# Where "From this computer" looks for ZIPs (a browser upload holds the whole
# file in memory twice over, so large ZIPs are read from disk instead).
LARGE_ZIP_FOLDERS = ("Downloads", "Desktop", "Documents")

# The name of the documents ZIP and a function that opens it for zipfile.
ZipChoice = tuple[str | None, Callable[[], object] | None]

st.set_page_config(page_title="Employee documents filter", layout="wide")


@st.cache_data(show_spinner="Reading Excel ...")
def read_table(data: bytes, filename: str) -> dict[str, pd.DataFrame]:
    return ff.load_table(data, filename)


def widget_key(*parts: object) -> str:
    """Stable key that changes when a widget's options change, so a stale
    selection is dropped instead of breaking the widget."""
    return "w" + hashlib.sha1("|".join(map(str, parts)).encode()).hexdigest()[:12]


def size_text(n: int) -> str:
    """Decimal units, as macOS Finder shows them (5,000,000,000 bytes = 5.00 GB)."""
    if n >= 1_000_000_000:
        return f"{n / 1_000_000_000:,.2f} GB"
    if n >= 1_000_000:
        return f"{n / 1_000_000:,.1f} MB"
    return f"{n / 1000:,.0f} KB"


def values_widget(label: str, series: pd.Series, ignore_case: bool, months: bool = False) -> list[str]:
    options = ff.value_options(series, ignore_case, chronological=months)
    labels = {key: f"{text}  ({count})" for key, text, count in options}
    return st.multiselect(
        label, list(labels), format_func=labels.__getitem__, placeholder="All",
        key=widget_key("values", label, series.name, ignore_case, tuple(labels)),
    )


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
    if series.map(lambda v: ff.clean_value(v, ignore_case)).nunique() > ff.MAX_LIST_VALUES:
        text = st.text_input(
            f"{col} contains", key=widget_key("contains", col),
            help="Separate terms with commas; a row matches if it contains any of them.",
        )
        return ff.ColumnFilter(col, contains=[t for t in text.split(",") if t.strip()])
    return ff.ColumnFilter(col, values=values_widget(col, series, ignore_case))


def is_local_session() -> bool:
    """True when the browser runs on the same computer as the app, so the
    ZIP can be written straight into this computer's Downloads folder."""
    return (urlparse(st.context.url or "").hostname or "") in ("localhost", "127.0.0.1", "::1")


def upload_zip(label: str) -> ZipChoice:
    upload = st.file_uploader(label, type=["zip"])
    if upload is None:
        return None, None
    # A private stream per use: reruns and the download never share a read position.
    return upload.name, lambda: io.BytesIO(upload.getvalue())


def pick_zip_on_disk() -> ZipChoice:
    found = ff.find_zip_files(Path.home() / folder for folder in LARGE_ZIP_FOLDERS)
    picked = st.selectbox(
        "ZIP in Downloads, Desktop or Documents", found, index=None, placeholder="Choose a ZIP",
        format_func=lambda p: f"{p.name}  ({size_text(p.stat().st_size)}, {p.parent.name})",
    )
    typed = st.text_input(
        "Or paste the ZIP's full path",
        help="Mac: right-click the ZIP, hold the Option key, choose 'Copy ... as Pathname' "
        "and paste it here.",
    )
    path = ff.clean_path_input(typed) if typed.strip() else picked
    if path is None:
        return None, None
    if not path.is_file():
        st.error(f"File not found: {path}")
        return None, None
    return path.name, lambda: path


# --------------------------------------------------------------------------
# 1. Upload
# --------------------------------------------------------------------------

st.title("Employee documents filter")
st.caption(
    "Upload the Excel / CSV and the documents ZIP, apply filters and select rows: only the "
    "selected employees' folders are downloaded in one ZIP, with the same structure."
)

st.subheader("1. Upload")
local = is_local_session()
up_left, up_right = st.columns(2)
table_upload = up_left.file_uploader("Doc status Excel / CSV", type=["xlsx", "xlsm", "csv"])
with up_right:
    if local and st.radio(
        "Documents ZIP (one folder per Employee ID)",
        ["Upload (up to 2 GB)", "From this computer (large ZIP, 5 GB and more)"],
        horizontal=True,
        help="A ZIP on this computer is read straight from disk, so its size does not matter.",
    ).startswith("From"):
        zip_name, zip_file = pick_zip_on_disk()
    else:
        zip_name, zip_file = upload_zip(
            "ZIP file (up to 2 GB)" if local else "Documents ZIP (one folder per Employee ID)"
        )

if table_upload is None:
    st.info("Upload an Excel or CSV file to start.")
    st.stop()

sheets = read_table(table_upload.getvalue(), table_upload.name)

with st.expander("Settings: sheets, Employee ID column, filter columns"):
    set_left, set_right = st.columns(2)
    with set_left:
        sheet_names = st.multiselect("Sheets", list(sheets), default=list(sheets))
    if not sheet_names:
        st.warning("Select at least one sheet.")
        st.stop()
    df = ff.combine_sheets(sheets, sheet_names)
    columns = [c for c in df.columns if c != ff.SHEET_COLUMN]
    guessed_id = ff.guess_id_column(columns)
    with set_right:
        id_col = st.selectbox(
            "Employee ID column", columns, index=columns.index(guessed_id) if guessed_id else 0,
            help="Values in this column are matched with the folder names in the ZIP.",
        )
        ignore_case = st.checkbox(
            "Ignore case, spaces, _ and -", value=True,
            help="'UNDERGRADUATE', 'Under Graduate' and 'Under_GRADUATE' count as the same value.",
        )
    st.caption("Which sheet column each filter uses:")
    quick_columns: dict[str, str | None] = {}
    for slot, name in zip(st.columns(len(ff.QUICK_FILTERS)), ff.QUICK_FILTERS):
        guessed = ff.guess_filter_column(df, name)
        choices = [None, *columns]
        quick_columns[name] = slot.selectbox(
            f"{name} column", choices, index=choices.index(guessed),
            format_func=lambda c: "(none)" if c is None else c, key=widget_key("qcol", name, *columns),
        )

keys = ff.id_keys(df, id_col)

index: ff.ZipIndex | None = None
if zip_file is not None:
    try:
        with zipfile.ZipFile(zip_file()) as zf:
            index = ff.index_zip(zf, known_ids=set(keys))
    except (zipfile.BadZipFile, OSError):
        st.error("This ZIP file cannot be opened. Create the ZIP again and choose it.")

if index is not None:
    # Only employees whose folder is in the ZIP go further; the rest is skipped.
    excel_ids = set(keys)
    matched = keys.map(lambda k: k in index.by_id)
    skipped = df[~matched]
    df, keys = df[matched], keys[matched]
    st.success(
        f"Matched with the ZIP: {len(df)} employees. "
        f"{len(skipped)} Excel rows skipped (no folder in the ZIP)."
    )
    orphans = [f.name for f in index.folders if not f.keys & excel_ids]
    info_left, info_right = st.columns(2)
    if len(skipped):
        with info_left.expander(f"Skipped Excel rows ({len(skipped)})"):
            st.dataframe(skipped, hide_index=True, width="stretch")
    if orphans:
        with info_right.expander(f"ZIP folders not found in the Excel ({len(orphans)})"):
            st.write(", ".join(orphans))
    if df.empty:
        st.error(
            f"No folder in the ZIP matches the Excel column '{id_col}'. "
            "Choose the correct Employee ID column in Settings."
        )
        st.stop()

# --------------------------------------------------------------------------
# 2. Filters
# --------------------------------------------------------------------------

st.subheader("2. Filters")
mask = pd.Series(True, index=df.index)
for slot, (name, col) in zip(st.columns(len(quick_columns)), quick_columns.items()):
    with slot:
        if col is None:
            st.multiselect(name, [], placeholder="Choose a column in Settings", disabled=True)
            continue
        values = ff.month_values(col, df[col]) if name == "Month" else df[col]
        picked = values_widget(name, values, ignore_case, months=name == "Month")
        mask &= ff.values_mask(values, picked, ignore_case)

with st.expander("More filters (any column, Employee IDs)"):
    extra_cols = st.multiselect(
        "Filter by columns", list(df.columns), placeholder="Choose columns",
        help="Different filters combine with AND; values picked in one filter combine with OR.",
    )
    extra_filters: list[ff.ColumnFilter] = []
    for start in range(0, len(extra_cols), 3):
        for slot, col in zip(st.columns(3), extra_cols[start:start + 3]):
            with slot:
                extra_filters.append(column_filter_widget(df, col, ignore_case))
    mask &= ff.filter_mask(df, extra_filters, ignore_case)

    pasted = ff.parse_id_list(st.text_area(
        "Employee IDs", height=90,
        placeholder="80013618, 80011485 ...  (one per line, or separated by commas / spaces)",
    ))
if pasted:
    mask &= keys.isin(pasted)
    unknown = sorted(set(pasted) - set(keys))
    if unknown:
        st.warning(f"{len(unknown)} pasted IDs did not match: {', '.join(unknown)}")

rows = df[mask]

# --------------------------------------------------------------------------
# 3. Select
# --------------------------------------------------------------------------

st.subheader("3. Select")
signature = hashlib.sha1(",".join(map(str, rows.index)).encode()).hexdigest()
if st.session_state.get("table_sig") != signature:
    st.session_state.update(table_sig=signature, select_all=True, table_reset=0)
buttons = st.container(horizontal=True)
if buttons.button("Select all"):
    st.session_state.update(select_all=True, table_reset=st.session_state.table_reset + 1)
if buttons.button("Clear all"):
    st.session_state.update(select_all=False, table_reset=st.session_state.table_reset + 1)

# The sheet may have its own "Select" / "Files in ZIP" columns: the app's get a free name.
select_col = ff.unique_name(SELECT_COLUMN, rows.columns)
files_col = ff.unique_name(FILES_COLUMN, [*rows.columns, select_col])
table = rows.copy()
table.insert(0, select_col, st.session_state.select_all)
if index is not None:
    table.insert(
        table.columns.get_loc(id_col) + 1, files_col,
        keys[rows.index].map(lambda k: sum(f.files for f in index.by_id[k])),
    )
edited = st.data_editor(
    table, hide_index=True, width="stretch", height=420,
    disabled=[c for c in table.columns if c != select_col],
    column_config={select_col: st.column_config.CheckboxColumn(select_col, width="small")},
    key=widget_key("table", signature, st.session_state.table_reset),
)
selected = rows.loc[edited.index[edited[select_col].astype(bool)]]
selected_ids = list(dict.fromkeys(k for k in keys[selected.index] if k))
folders = (
    list(dict.fromkeys(f for k in selected_ids for f in index.by_id.get(k, [])))
    if index is not None else []
)

stats = st.columns(4)
stats[0].metric("Employees after filters", len(rows))
stats[1].metric("Selected", len(selected))
if index is not None:
    stats[2].metric("Documents", sum(f.files for f in folders))
    stats[3].metric("Data size", size_text(sum(f.size for f in folders)))

# --------------------------------------------------------------------------
# 4. Download
# --------------------------------------------------------------------------

st.subheader("4. Download")
if index is None:
    st.info("Choose the documents ZIP to download the selected folders here.")
else:
    def build_download() -> io.BytesIO:
        out = io.BytesIO()
        with zipfile.ZipFile(zip_file()) as zf:
            ff.build_zip(zf, index, folders, out)
        return out

    total_size = sum(f.size for f in folders)
    file_name = f"{zip_name.rsplit('.', 1)[0]}_selected_{len(folders)}.zip"
    summary = f"{len(folders)} folders, {size_text(total_size)}"
    fits = total_size <= BROWSER_ZIP_LIMIT
    actions = st.container(horizontal=True)
    save = local and actions.button(
        f"Save ZIP to Downloads ({summary})", type="primary", disabled=not folders
    )
    actions.download_button(
        f"Download ZIP ({summary})", data=build_download, file_name=file_name,
        mime="application/zip", type="secondary" if local else "primary",
        disabled=not folders or not fits, on_click="ignore",
    )
    if not fits:
        st.caption(
            "Selections over 1 GB are too large for a browser download: "
            + ("use Save ZIP to Downloads." if local
               else "select fewer employees, or run the app on your computer.")
        )
    else:
        st.caption("A large ZIP can take a little while to prepare after you click.")

    selection = hashlib.sha1("|".join(f.name for f in folders).encode()).hexdigest()
    if save:
        bar = st.progress(0.0, text="Saving ZIP ...")
        step = max(1, sum(f.files for f in folders) // 200)

        def show(done: int, total: int) -> None:
            if done == total or done % step == 0:
                bar.progress(done / total, text=f"Saving ZIP ... {done:,} / {total:,} files")

        with zipfile.ZipFile(zip_file()) as zf:
            saved = ff.save_zip(zf, index, folders, Path.home() / "Downloads" / file_name, show)
        bar.empty()
        st.session_state["saved_zip"] = (selection, saved)
    last_save = st.session_state.get("saved_zip")
    if last_save and last_save[0] == selection:
        st.success(f"Saved: {last_save[1]}")
