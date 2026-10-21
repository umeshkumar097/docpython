"""Core logic: read the doc-status Excel / CSV, filter its rows, find each
employee's folder inside the uploaded documents ZIP and write a new ZIP with
only the selected employees' folders (same paths as the original).

Kept free of any UI code so it can be tested on its own (tests/).
"""

from __future__ import annotations

import io
import re
import shutil
import zipfile
from pathlib import Path
from dataclasses import dataclass, field
from datetime import date
from typing import BinaryIO, Callable, Iterable

import pandas as pd

BLANK_LABEL = "(blank)"
SHEET_COLUMN = "Sheet"
ID_COLUMN_HINTS = ("employee id", "emp id", "employee code", "emp code", "employeeid", "empid")

# A column with more distinct values than this gets a "contains" search box
# instead of a value list (names, emails, ID numbers ...).
MAX_LIST_VALUES = 500

# The four main filters and the column-name words that find their column,
# best tier first ("Batch" falls back to a Phase column, "Month" to a date).
QUICK_FILTERS: dict[str, tuple[tuple[str, ...], ...]] = {
    "Batch": (("batch",), ("phase",)),
    "Language": (("language",),),
    "Education": (("education", "qualification"), ("category",)),
    "Month": (("month",), ("date",)),
}

# Already-compressed formats are stored as-is in the output ZIP (fast, same size).
STORED_SUFFIXES = {
    ".jpg", ".jpeg", ".png", ".gif", ".webp", ".heic", ".pdf", ".zip", ".rar", ".7z",
    ".mp4", ".mov", ".docx", ".xlsx", ".pptx",
}
JUNK_NAMES = {".DS_Store", "Thumbs.db", "desktop.ini"}


# --------------------------------------------------------------------------
# Excel / CSV
# --------------------------------------------------------------------------


def _tidy(df: pd.DataFrame) -> pd.DataFrame:
    df = df.dropna(how="all").reset_index(drop=True)
    df.columns = [str(c).strip() for c in df.columns]
    return df


def load_table(data: bytes, filename: str) -> dict[str, pd.DataFrame]:
    """Sheets of an .xlsx, or one 'CSV' sheet, with every cell as text.

    Reading as text keeps IDs exactly as typed (no 8.0013618e7, no lost
    leading zeros). Fully empty rows are dropped, header names trimmed.
    """
    if filename.lower().endswith(".csv"):
        try:
            df = pd.read_csv(io.BytesIO(data), dtype=str, encoding="utf-8-sig")
        except UnicodeDecodeError:  # saved by Excel on Windows
            df = pd.read_csv(io.BytesIO(data), dtype=str, encoding="cp1252")
        return {"CSV": _tidy(df)}
    sheets = pd.read_excel(io.BytesIO(data), sheet_name=None, dtype=str, engine="openpyxl")
    return {str(name): _tidy(df) for name, df in sheets.items()}


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


def guess_filter_column(df: pd.DataFrame, filter_name: str) -> str | None:
    """Best column for one of QUICK_FILTERS: within the first tier of words
    that matches any column, the column with the most filled cells."""
    for words in QUICK_FILTERS[filter_name]:
        matches = [
            c for c in df.columns
            if c != SHEET_COLUMN and any(w in c.lower() for w in words)
        ]
        if matches:
            return max(matches, key=lambda c: df[c].map(lambda v: clean_value(v, True) != "").sum())
    return None


# --------------------------------------------------------------------------
# IDs
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


def name_keys(name: str) -> set[str]:
    """IDs a folder / file name answers to: '80013618' -> {'80013618'};
    '80013618 - Somanath' and '80013618_photo' also answer to their leading number."""
    keys = {normalize_id(name)}
    leading = re.match(r"\s*(\d+)", name)
    if leading:
        keys.add(normalize_id(leading.group(1)))
    keys.discard("")
    return keys


def id_keys(df: pd.DataFrame, id_column: str) -> pd.Series:
    return df[id_column].map(normalize_id)


# --------------------------------------------------------------------------
# Filters
# --------------------------------------------------------------------------


def clean_value(value: object, ignore_case: bool) -> str:
    """Value used for matching. With `ignore_case`, case, spaces, '_', '-' and
    '.' are ignored, so 'UNDERGRADUATE', 'Under Graduate' and 'Under_GRADUATE'
    are one choice."""
    if value is None or (not isinstance(value, str) and pd.isna(value)):
        return ""
    text = " ".join(str(value).split())
    return re.sub(r"[\s_.\-]+", "", text).casefold() if ignore_case else text


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


def month_values(name: str, series: pd.Series) -> pd.Series:
    """Values for the Month filter: a date column becomes 'July 2026'; a
    text column ('July 2026', 'Sept batch') is used as it is."""
    if not is_date_column(name, series):
        return series
    parsed = parse_dates(series)
    return parsed.dt.strftime("%B %Y").where(parsed.notna(), series)


def _month_sort_key(label: str) -> tuple[int, pd.Timestamp | None, str]:
    for fmt in ("%B %Y", "%b %Y"):
        when = pd.to_datetime(label, format=fmt, errors="coerce")
        if not pd.isna(when):
            return (0, when, label)
    return (1, None, label.casefold())


def value_options(
    series: pd.Series, ignore_case: bool, chronological: bool = False
) -> list[tuple[str, str, int]]:
    """(key, label, count) for every distinct value: blanks last, months in
    calendar order when `chronological`, otherwise A-Z.

    The label is the most common original spelling; the blank key is ''.
    """
    keys = series.map(lambda v: clean_value(v, ignore_case))
    shown = series.map(lambda v: clean_value(v, False))
    options = []
    for key, group in shown.groupby(keys, sort=False):
        label = group.value_counts().index[0] if key else BLANK_LABEL
        options.append((key, label, len(group)))
    if chronological:
        options.sort(key=lambda o: (o[0] == "", _month_sort_key(o[1])))
    else:
        options.sort(key=lambda o: (o[0] == "", o[1].casefold(), o[1]))
    return options


def values_mask(series: pd.Series, keys: Iterable[str], ignore_case: bool) -> pd.Series:
    wanted = set(keys)
    if not wanted:
        return pd.Series(True, index=series.index)
    return series.map(lambda v: clean_value(v, ignore_case) in wanted)


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
            mask &= values_mask(col, f.values, ignore_case)
        if f.contains:
            terms = [clean_value(t, True) for t in f.contains]
            terms = [t for t in terms if t]
            if terms:
                mask &= col.map(lambda v: any(t in clean_value(v, True) for t in terms))
        if f.date_from or f.date_to:
            days = parse_dates(col).dt.date
            ok = days.notna()
            if f.date_from:
                ok &= days.map(lambda d: not pd.isna(d) and d >= f.date_from)
            if f.date_to:
                ok &= days.map(lambda d: not pd.isna(d) and d <= f.date_to)
            mask &= ok
    return mask


# --------------------------------------------------------------------------
# Documents ZIP
# --------------------------------------------------------------------------


def _parts(info: zipfile.ZipInfo) -> list[str]:
    return [p for p in info.filename.replace("\\", "/").split("/") if p]


def is_junk(info: zipfile.ZipInfo) -> bool:
    """macOS / Windows clutter that is never an employee's document."""
    parts = _parts(info)
    return (
        not parts
        or parts[0] == "__MACOSX"
        or parts[-1] in JUNK_NAMES
        or parts[-1].startswith(("._", "~$"))
    )


@dataclass(eq=False)  # identity: the same folder object is chosen once
class ZipFolder:
    """One employee entry: a folder (or a single file) at the top of the ZIP,
    below any wrapper folder such as 'Export-2026-09-23 (1)/'."""

    name: str
    members: list[zipfile.ZipInfo]
    keys: set[str]

    @property
    def files(self) -> int:
        return sum(1 for m in self.members if not m.is_dir())

    @property
    def size(self) -> int:
        return sum(m.file_size for m in self.members)


@dataclass
class ZipIndex:
    root: str  # shared wrapper path ('Export-2026-09-23 (1)/') or ''
    root_dirs: list[zipfile.ZipInfo]  # the wrapper's own directory entries
    folders: list[ZipFolder]
    by_id: dict[str, list[ZipFolder]]


def index_zip(zf: zipfile.ZipFile, known_ids: Iterable[str] = ()) -> ZipIndex:
    """Find the employee folders inside the ZIP.

    Wrapper folders that hold everything ('Export-2026-09-23 (1)/') are
    looked through, unless the single folder is itself a known employee ID.
    """
    known = set(known_ids)
    infos = [i for i in zf.infolist() if not is_junk(i)]
    paths = [_parts(i) for i in infos]

    root: list[str] = []
    while True:
        depth = len(root)
        tops = {p[depth] for p in paths if len(p) > depth}
        if len(tops) != 1:
            break
        only = next(iter(tops))
        is_file = any(len(p) == depth + 1 and not i.is_dir() for p, i in zip(paths, infos))
        if is_file or name_keys(only) & known or not any(len(p) > depth + 1 for p in paths):
            break
        root.append(only)

    depth = len(root)
    root_dirs: list[zipfile.ZipInfo] = []
    groups: dict[str, list[tuple[list[str], zipfile.ZipInfo]]] = {}
    for p, info in zip(paths, infos):
        if len(p) <= depth:
            root_dirs.append(info)
        else:
            groups.setdefault(p[depth], []).append((p, info))

    folders: list[ZipFolder] = []
    by_id: dict[str, list[ZipFolder]] = {}
    for name in sorted(groups):
        members = groups[name]
        single_file = len(members) == 1 and len(members[0][0]) == depth + 1 and not members[0][1].is_dir()
        stem = name.rsplit(".", 1)[0] if single_file and "." in name else name
        folder = ZipFolder(name, [i for _, i in members], name_keys(stem))
        folders.append(folder)
        for key in folder.keys:
            by_id.setdefault(key, []).append(folder)
    return ZipIndex("".join(f"{r}/" for r in root), root_dirs, folders, by_id)


def _copy_info(info: zipfile.ZipInfo) -> zipfile.ZipInfo:
    """A fresh ZipInfo with the same name / date / permissions (writing
    mutates a ZipInfo, so the source's own objects are never reused)."""
    new = zipfile.ZipInfo(info.filename, info.date_time)
    new.external_attr = info.external_attr
    new.file_size = info.file_size
    suffix = "." + info.filename.rsplit(".", 1)[-1].lower() if "." in info.filename else ""
    new.compress_type = (
        zipfile.ZIP_STORED if info.is_dir() or suffix in STORED_SUFFIXES else zipfile.ZIP_DEFLATED
    )
    return new


def build_zip(
    zf: zipfile.ZipFile,
    index: ZipIndex,
    folders: Iterable[ZipFolder],
    out: BinaryIO,
    progress: Callable[[int, int], None] | None = None,
) -> int:
    """Write the chosen folders, with exactly their original paths, into
    `out` as a new ZIP. Returns the number of files written."""
    chosen = list(dict.fromkeys(folders))
    total = sum(f.files for f in chosen)
    written = 0
    with zipfile.ZipFile(out, "w", allowZip64=True) as dest:
        for info in index.root_dirs:
            if info.is_dir():
                dest.writestr(_copy_info(info), b"")
        for folder in chosen:
            for info in folder.members:
                new = _copy_info(info)
                if info.is_dir():
                    dest.writestr(new, b"")
                    continue
                with zf.open(info) as src, dest.open(new, "w") as dst:
                    shutil.copyfileobj(src, dst, 1 << 20)
                written += 1
                if progress:
                    progress(written, total)
    return written


def clean_path_input(raw: str) -> Path:
    """A path typed or pasted by the user: tolerates surrounding quotes, '~'
    and the backslash escapes macOS Terminal adds when a file is dropped in."""
    text = (raw or "").strip().strip("'\"").strip()
    path = Path(text).expanduser()
    if not path.exists() and "\\" in text:
        unescaped = Path(re.sub(r"\\(.)", r"\1", text)).expanduser()
        if unescaped.exists():
            return unescaped
    return path


def find_zip_files(folders: Iterable[Path], limit: int = 50) -> list[Path]:
    """ZIP files directly inside `folders` (e.g. Downloads, Desktop), newest
    first, leaving out the ZIPs this app itself saved."""
    found = []
    for folder in folders:
        if folder.is_dir():
            found += [
                p for p in folder.glob("*.zip")
                if p.is_file() and not p.name.startswith(".") and "_selected_" not in p.name
            ]
    return sorted(found, key=lambda p: p.stat().st_mtime, reverse=True)[:limit]


def unique_path(path: Path) -> Path:
    """`path`, or `name (2).zip`, `name (3).zip` ... whichever does not exist yet."""
    candidate, n = path, 2
    while candidate.exists():
        candidate = path.with_name(f"{path.stem} ({n}){path.suffix}")
        n += 1
    return candidate


def save_zip(
    zf: zipfile.ZipFile,
    index: ZipIndex,
    folders: Iterable[ZipFolder],
    target: Path,
    progress: Callable[[int, int], None] | None = None,
) -> Path:
    """Stream the chosen folders into a new ZIP file on disk (memory stays
    small however big the selection is). Returns the path written, which is
    `target` or a numbered variant if that name is taken."""
    target = unique_path(target)
    target.parent.mkdir(parents=True, exist_ok=True)
    part = target.with_name(target.name + ".part")
    try:
        with part.open("wb") as out:
            build_zip(zf, index, folders, out, progress)
        part.replace(target)
    finally:
        part.unlink(missing_ok=True)
    return target


def selected_list_xlsx(rows: pd.DataFrame) -> bytes:
    buf = io.BytesIO()
    rows.to_excel(buf, index=False, engine="openpyxl")
    return buf.getvalue()
