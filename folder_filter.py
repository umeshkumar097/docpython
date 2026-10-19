"""Core logic: read the status Excel, filter its rows, find the matching
employee-ID folders in an export and copy / zip only those.

Kept free of any UI code so it can be tested on its own (tests/) and reused
from a script.
"""

from __future__ import annotations

import re
import shutil
from dataclasses import dataclass, field
from datetime import date
from pathlib import Path
from typing import BinaryIO, Callable, Iterable

import pandas as pd

BLANK_LABEL = "(blank)"
SHEET_COLUMN = "Sheet"
FOLDER_FOUND_COLUMN = "Folder found"
ID_COLUMN_HINTS = ("employee id", "emp id", "employee code", "emp code", "employeeid", "empid")

# A column with more distinct values than this gets a "contains" search box
# instead of a value list (names, emails, ID numbers ...).
MAX_LIST_VALUES = 500


# --------------------------------------------------------------------------
# Excel
# --------------------------------------------------------------------------


def load_workbook(source: str | Path | BinaryIO) -> dict[str, pd.DataFrame]:
    """Every sheet of the workbook, all cells as text.

    Reading as text keeps IDs exactly as typed (no 8.0013618e7, no lost
    leading zeros). Fully empty rows are dropped, header names trimmed.
    """
    sheets = pd.read_excel(source, sheet_name=None, dtype=str, engine="openpyxl")
    out: dict[str, pd.DataFrame] = {}
    for name, df in sheets.items():
        df = df.dropna(how="all").reset_index(drop=True)
        df.columns = [str(c).strip() for c in df.columns]
        out[str(name)] = df
    return out


def combine_sheets(sheets: dict[str, pd.DataFrame], names: Iterable[str]) -> pd.DataFrame:
    """Stack the chosen sheets (columns are unioned) with a leading Sheet column."""
    frames = []
    for name in names:
        df = sheets[name].copy()
        if SHEET_COLUMN in df.columns:
            df = df.rename(columns={SHEET_COLUMN: f"{SHEET_COLUMN} (original)"})
        df.insert(0, SHEET_COLUMN, name)
        frames.append(df)
    if not frames:
        return pd.DataFrame()
    return pd.concat(frames, ignore_index=True, sort=False)


def guess_id_column(columns: Iterable[str]) -> str | None:
    """The column that holds employee IDs, preferring the main one over MD_ copies."""
    cols = list(columns)
    for wanted in ID_COLUMN_HINTS:
        exact = [c for c in cols if c.strip().lower() == wanted]
        if exact:
            return exact[0]
    loose = [c for c in cols if any(h in c.lower() for h in ID_COLUMN_HINTS)]
    loose.sort(key=lambda c: c.lower().startswith("md_"))
    return loose[0] if loose else None


# --------------------------------------------------------------------------
# IDs and the export folder
# --------------------------------------------------------------------------


def normalize_id(value: object) -> str:
    """Comparable form of an employee ID: '080013618', '80013618.0' and
    ' 80013618 ' all become '80013618'. Empty for blanks."""
    if value is None or (not isinstance(value, str) and pd.isna(value)):
        return ""
    text = str(value).strip()
    if text.lower() in ("", "nan", "none", "n/a", "na", "-"):
        return ""
    text = re.sub(r"\.0+$", "", text)
    text = re.sub(r"\s+", "", text)
    if text.isdigit():
        text = text.lstrip("0") or "0"
    return text.upper()


def parse_id_list(text: str) -> list[str]:
    """IDs pasted by the user: one per line, or separated by commas / spaces / tabs."""
    seen: dict[str, None] = {}
    for part in re.split(r"[\s,;]+", text or ""):
        key = normalize_id(part)
        if key:
            seen[key] = None
    return list(seen)


def entry_keys(path: Path) -> set[str]:
    """IDs a folder (or file) in the export answers to.

    '80013618' -> {'80013618'}; '80013618 - Somanath' and '80013618_photo.jpg'
    also answer to their leading number.
    """
    name = path.name if path.is_dir() else path.stem
    keys = {normalize_id(name)}
    leading = re.match(r"\s*(\d+)", name)
    if leading:
        keys.add(normalize_id(leading.group(1)))
    keys.discard("")
    return keys


def clean_path_input(raw: str) -> Path:
    """A path typed or pasted by the user: tolerates surrounding quotes, '~'
    and the backslash escapes macOS Terminal adds when a folder is dropped in."""
    text = (raw or "").strip().strip("'\"").strip()
    path = Path(text).expanduser()
    if not path.exists() and "\\" in text:
        unescaped = Path(re.sub(r"\\(.)", r"\1", text)).expanduser()
        if unescaped.exists():
            return unescaped
    return path


def scan_export(root: Path) -> dict[str, list[Path]]:
    """ID -> the folders / files directly inside `root` named after it."""
    index: dict[str, list[Path]] = {}
    for entry in sorted(root.iterdir(), key=lambda p: p.name):
        if entry.name.startswith((".", "~$")):
            continue  # .DS_Store, Excel lock files
        for key in entry_keys(entry):
            index.setdefault(key, []).append(entry)
    return index


# --------------------------------------------------------------------------
# Filters
# --------------------------------------------------------------------------


def clean_value(value: object, ignore_case: bool) -> str:
    """Value used for matching: extra spaces collapsed, optionally case-folded,
    so 'Home credit ' and 'HOME CREDIT' can count as the same choice."""
    if value is None or (not isinstance(value, str) and pd.isna(value)):
        return ""
    text = " ".join(str(value).split())
    return text.casefold() if ignore_case else text


def value_options(series: pd.Series, ignore_case: bool) -> list[tuple[str, str, int]]:
    """(key, label, count) for every distinct value, sorted by label.

    The label is the most common original spelling; the blank key is ''.
    """
    keys = series.map(lambda v: clean_value(v, ignore_case))
    shown = series.map(lambda v: clean_value(v, False))
    options = []
    for key, group in shown.groupby(keys, sort=False):
        label = group.value_counts().index[0] if key else BLANK_LABEL
        options.append((key, label, len(group)))
    options.sort(key=lambda o: (o[0] == "", o[1].casefold(), o[1]))
    return options


def parse_dates(series: pd.Series) -> pd.Series:
    """ISO dates first (what Excel date cells turn into), then Indian day-first
    text like 08/07/2026. Unparseable cells become NaT."""
    iso = pd.to_datetime(series, format="ISO8601", errors="coerce")
    rest = iso.isna() & series.notna()
    if rest.any():
        iso.loc[rest] = pd.to_datetime(
            series[rest], format="mixed", dayfirst=True, errors="coerce"
        )
    return iso


def is_date_column(name: str, series: pd.Series) -> bool:
    if "date" not in name.lower() and "dob" not in name.lower():
        return False
    filled = series.dropna()
    filled = filled[filled.map(lambda v: clean_value(v, True) not in ("", "n/a", "na"))]
    if filled.empty:
        return False
    return parse_dates(filled).notna().mean() >= 0.8


@dataclass
class ColumnFilter:
    """One column's condition. Several filters combine with AND; the values
    picked inside one filter combine with OR."""

    column: str
    values: list[str] = field(default_factory=list)  # keys from value_options
    contains: list[str] = field(default_factory=list)  # any of these substrings
    date_from: date | None = None
    date_to: date | None = None

    def is_active(self) -> bool:
        return bool(self.values or self.contains or self.date_from or self.date_to)


def filter_mask(df: pd.DataFrame, filters: Iterable[ColumnFilter], ignore_case: bool) -> pd.Series:
    mask = pd.Series(True, index=df.index)
    for f in filters:
        if not f.is_active() or f.column not in df.columns:
            continue
        col = df[f.column]
        if f.values:
            wanted = set(f.values)
            mask &= col.map(lambda v: clean_value(v, ignore_case) in wanted)
        if f.contains:
            terms = [t.casefold() for t in (clean_value(t, False) for t in f.contains) if t]
            if terms:
                mask &= col.map(lambda v: any(t in clean_value(v, True) for t in terms))
        if f.date_from or f.date_to:
            days = parse_dates(col).dt.date
            ok = days.notna()
            if f.date_from:
                ok &= days.map(lambda d: d is not None and not pd.isna(d) and d >= f.date_from)
            if f.date_to:
                ok &= days.map(lambda d: d is not None and not pd.isna(d) and d <= f.date_to)
            mask &= ok
    return mask


def id_keys(df: pd.DataFrame, id_column: str) -> pd.Series:
    return df[id_column].map(normalize_id)


# --------------------------------------------------------------------------
# Export
# --------------------------------------------------------------------------


@dataclass
class ExportResult:
    out_dir: Path
    copied: list[Path]
    missing_ids: list[str]
    report: Path
    zip_path: Path | None = None


def default_output_dir(source: Path) -> Path:
    """Next to the export; `unique_dir` adds (2), (3) ... on later runs."""
    return source.parent / f"{source.name}_filtered"


def unique_dir(path: Path) -> Path:
    """`path`, or `path (2)`, `path (3)` ... whichever does not exist yet."""
    candidate, n = path, 2
    while candidate.exists() or candidate.with_name(candidate.name + ".zip").exists():
        candidate = path.with_name(f"{path.name} ({n})")
        n += 1
    return candidate


def export_selection(
    rows: pd.DataFrame,
    id_column: str,
    index: dict[str, list[Path]],
    out_dir: Path,
    make_zip: bool = False,
    progress: Callable[[int, int, str], None] | None = None,
) -> ExportResult:
    """Copy the folders of every employee in `rows` into `out_dir`, write
    filtered_list.xlsx next to them and optionally zip the lot.

    Nothing in the source export is changed.
    """
    out_dir = unique_dir(out_dir)
    out_dir.mkdir(parents=True)

    keys = list(dict.fromkeys(k for k in id_keys(rows, id_column) if k))
    missing = [k for k in keys if k not in index]
    sources = list(dict.fromkeys(p for k in keys for p in index.get(k, [])))

    copied: list[Path] = []
    for i, src in enumerate(sources, start=1):
        dest = out_dir / src.name
        if src.is_dir():
            shutil.copytree(src, dest)
        else:
            shutil.copy2(src, dest)
        copied.append(dest)
        if progress:
            progress(i, len(sources), src.name)

    report = out_dir / "filtered_list.xlsx"
    listed = rows.copy()
    listed[FOLDER_FOUND_COLUMN] = id_keys(rows, id_column).map(
        lambda k: "Yes" if k in index else "No"
    )
    with pd.ExcelWriter(report, engine="openpyxl") as writer:
        listed.to_excel(writer, sheet_name="Filtered rows", index=False)
        if missing:
            pd.DataFrame({id_column: missing}).to_excel(
                writer, sheet_name="Missing folders", index=False
            )

    zip_path = None
    if make_zip:
        zip_path = Path(
            shutil.make_archive(str(out_dir), "zip", root_dir=out_dir.parent, base_dir=out_dir.name)
        )
    return ExportResult(out_dir, copied, missing, report, zip_path)
