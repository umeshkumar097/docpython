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

import pandas as pd
import streamlit as st

import folder_filter as ff

SELECT_COLUMN = "Select"
FILES_COLUMN = "Files in ZIP"

st.set_page_config(page_title="Employee documents filter", layout="wide")


@st.cache_data(show_spinner="Reading Excel ...")
def read_table(data: bytes, filename: str) -> dict[str, pd.DataFrame]:
    return ff.load_table(data, filename)


def widget_key(*parts: object) -> str:
    """Stable key that changes when a widget's options change, so a stale
    selection is dropped instead of breaking the widget."""
    return "w" + hashlib.sha1("|".join(map(str, parts)).encode()).hexdigest()[:12]


def size_text(n: int) -> str:
    return f"{n / 1_048_576:,.1f} MB" if n >= 1_048_576 else f"{n / 1024:,.0f} KB"


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
            help="Comma se alag karo; koi bhi ek match ho to row aayegi.",
        )
        return ff.ColumnFilter(col, contains=[t for t in text.split(",") if t.strip()])
    return ff.ColumnFilter(col, values=values_widget(col, series, ignore_case))


def zip_bytes(upload) -> io.BytesIO:
    # A private stream per use: reruns and the download never share a read position.
    return io.BytesIO(upload.getvalue())


# --------------------------------------------------------------------------
# 1. Upload
# --------------------------------------------------------------------------

st.title("Employee documents filter")
st.caption(
    "Excel / CSV aur documents ki ZIP upload karo, filter lagao, rows select karo: "
    "sirf selected employees ke folders ek ZIP mein download honge, same structure mein."
)

st.subheader("1. Upload")
up_left, up_right = st.columns(2)
table_upload = up_left.file_uploader("Doc status Excel / CSV", type=["xlsx", "xlsm", "csv"])
zip_upload = up_right.file_uploader(
    "Documents ZIP (har employee ka folder = Employee ID)", type=["zip"]
)

if table_upload is None:
    st.info("Excel ya CSV file upload karo.")
    st.stop()

sheets = read_table(table_upload.getvalue(), table_upload.name)

with st.expander("Settings: sheets, Employee ID column, filter columns"):
    set_left, set_right = st.columns(2)
    with set_left:
        sheet_names = st.multiselect("Sheets", list(sheets), default=list(sheets))
    if not sheet_names:
        st.warning("Kam se kam ek sheet chuno.")
        st.stop()
    df = ff.combine_sheets(sheets, sheet_names)
    columns = [c for c in df.columns if c != ff.SHEET_COLUMN]
    guessed_id = ff.guess_id_column(columns)
    with set_right:
        id_col = st.selectbox(
            "Employee ID column", columns, index=columns.index(guessed_id) if guessed_id else 0,
            help="Is column ki value ZIP ke folder ke naam se match hoti hai.",
        )
        ignore_case = st.checkbox(
            "Ignore case, spaces, _ and -", value=True,
            help="'UNDERGRADUATE', 'Under Graduate' aur 'Under_GRADUATE' ek hi value maane jayenge.",
        )
    st.caption("Har filter kis column se bane (sheet ke column names):")
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
if zip_upload is not None:
    try:
        with zipfile.ZipFile(zip_bytes(zip_upload)) as zf:
            index = ff.index_zip(zf, known_ids=set(keys))
    except zipfile.BadZipFile:
        st.error("Ye ZIP file khul nahi rahi. Dobara ZIP banakar upload karo.")

if index is not None:
    # Only employees whose folder is in the ZIP go further; the rest is skipped.
    excel_ids = set(keys)
    matched = keys.map(lambda k: k in index.by_id)
    skipped = df[~matched]
    df, keys = df[matched], keys[matched]
    st.success(
        f"ZIP se match: {len(df)} employees. "
        f"{len(skipped)} Excel rows skip hui (unka folder ZIP mein nahi hai)."
    )
    orphans = [f.name for f in index.folders if not f.keys & excel_ids]
    info_left, info_right = st.columns(2)
    if len(skipped):
        with info_left.expander(f"Skipped Excel rows ({len(skipped)})"):
            st.dataframe(skipped, hide_index=True, width="stretch")
    if orphans:
        with info_right.expander(f"ZIP folders jinki ID Excel mein nahi ({len(orphans)})"):
            st.write(", ".join(orphans))
    if df.empty:
        st.error(
            f"ZIP ka koi folder Excel ke '{id_col}' column se match nahi hua. "
            "Settings mein sahi Employee ID column chuno."
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
            st.multiselect(name, [], placeholder="Settings mein column chuno", disabled=True)
            continue
        values = ff.month_values(col, df[col]) if name == "Month" else df[col]
        picked = values_widget(name, values, ignore_case, months=name == "Month")
        mask &= ff.values_mask(values, picked, ignore_case)

with st.expander("More filters (koi bhi column, Employee IDs)"):
    extra_cols = st.multiselect(
        "Filter by columns", list(df.columns), placeholder="Columns chuno",
        help="Alag filters AND hote hain; ek filter mein chuni values OR hoti hain.",
    )
    extra_filters: list[ff.ColumnFilter] = []
    for start in range(0, len(extra_cols), 3):
        for slot, col in zip(st.columns(3), extra_cols[start:start + 3]):
            with slot:
                extra_filters.append(column_filter_widget(df, col, ignore_case))
    mask &= ff.filter_mask(df, extra_filters, ignore_case)

    pasted = ff.parse_id_list(st.text_area(
        "Employee IDs", height=90,
        placeholder="80013618, 80011485 ...  (ek line mein ek, ya comma / space se alag)",
    ))
if pasted:
    mask &= keys.isin(pasted)
    unknown = sorted(set(pasted) - set(keys))
    if unknown:
        st.warning(f"{len(unknown)} pasted IDs match nahi hue: {', '.join(unknown)}")

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

table = rows.copy()
table.insert(0, SELECT_COLUMN, st.session_state.select_all)
if index is not None:
    table.insert(
        table.columns.get_loc(id_col) + 1, FILES_COLUMN,
        keys[rows.index].map(lambda k: sum(f.files for f in index.by_id[k])),
    )
edited = st.data_editor(
    table, hide_index=True, width="stretch", height=420,
    disabled=[c for c in table.columns if c != SELECT_COLUMN],
    column_config={SELECT_COLUMN: st.column_config.CheckboxColumn(SELECT_COLUMN, width="small")},
    key=widget_key("table", signature, st.session_state.table_reset),
)
selected = rows.loc[edited.index[edited[SELECT_COLUMN].astype(bool)]]
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
    st.info("Documents ki ZIP upload karo, phir yahan se selected folders ki ZIP milegi.")
else:
    def build_download() -> io.BytesIO:
        out = io.BytesIO()
        with zipfile.ZipFile(zip_bytes(zip_upload)) as zf:
            ff.build_zip(zf, index, folders, out)
        return out

    stem = zip_upload.name.rsplit(".", 1)[0]
    st.download_button(
        f"Download ZIP ({len(folders)} folders, {size_text(sum(f.size for f in folders))})",
        data=build_download, file_name=f"{stem}_selected_{len(folders)}.zip",
        mime="application/zip", type="primary", disabled=not folders, on_click="ignore",
    )
    st.caption("Badi ZIP banne mein click ke baad thoda time lag sakta hai.")

st.download_button(
    "Download selected list (.xlsx)", ff.selected_list_xlsx(selected),
    file_name="selected_list.xlsx", disabled=selected.empty, on_click="ignore",
)
