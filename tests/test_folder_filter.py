from datetime import date
from pathlib import Path
import sys
import zipfile

import pandas as pd
import pytest

sys.path.insert(0, str(Path(__file__).resolve().parents[1]))
import folder_filter as ff  # noqa: E402


@pytest.fixture
def workbook(tmp_path: Path) -> Path:
    path = tmp_path / "status.xlsx"
    tc = pd.DataFrame({
        "Registration Date": ["2026-07-08 16:49:59", "2026-07-20 10:00:00", "2026-08-02 09:00:00"],
        "Candidate Name": ["SOMANATH", "AMANDEEP", "KHUSHBU"],
        "Phase": ["Phase 1 August", "Batch 01", "Batch 01"],
        "Employer": ["Home credit ", "HOME CREDIT", "Other Co"],
        "Employee ID": [80013618, "80011485", "080608"],
        "MD_Employee ID": ["x", "y", "z"],
    })
    field = pd.DataFrame({
        "Registration Date": ["2026-05-30 08:51:44"],
        "Candidate Name": ["ANSHUMAN"],
        "Phase": ["June Phase 1"],
        "Employer": [None],
        "Employee ID": ["80004215"],
        "MD_Active Status": ["Active"],
    })
    with pd.ExcelWriter(path) as writer:
        tc.to_excel(writer, sheet_name="tc", index=False)
        field.to_excel(writer, sheet_name="field", index=False)
    return path


@pytest.fixture
def export(tmp_path: Path) -> Path:
    root = tmp_path / "Export-2026-09-23 (1)"
    for name in ["80013618", "80011485", "80004215 - Anshuman", "99999999"]:
        (root / name).mkdir(parents=True)
        (root / name / "photo.jpg").write_bytes(b"jpg")
    (root / "80608_sign.pdf").write_bytes(b"pdf")
    (root / ".DS_Store").write_bytes(b"")
    return root


def test_load_and_combine_keeps_ids_as_text(workbook: Path) -> None:
    sheets = ff.load_workbook(workbook)
    assert list(sheets) == ["tc", "field"]
    df = ff.combine_sheets(sheets, ["tc", "field"])
    assert list(df[ff.SHEET_COLUMN]) == ["tc", "tc", "tc", "field"]
    assert list(df["Employee ID"]) == ["80013618", "80011485", "080608", "80004215"]
    assert "MD_Active Status" in df.columns  # columns of both sheets are kept


def test_guess_id_column_prefers_main_column() -> None:
    assert ff.guess_id_column(["Name", "MD_Employee ID", "Employee ID"]) == "Employee ID"
    assert ff.guess_id_column(["MD_Employee ID ", "Name"]) == "MD_Employee ID "
    assert ff.guess_id_column(["Name", "Phone"]) is None


@pytest.mark.parametrize(
    ("raw", "expected"),
    [(" 80013618 ", "80013618"), ("80013618.0", "80013618"), ("080608", "80608"),
     (80013618, "80013618"), ("ab12", "AB12"), ("N/A", ""), (None, ""), (float("nan"), "")],
)
def test_normalize_id(raw: object, expected: str) -> None:
    assert ff.normalize_id(raw) == expected


def test_parse_id_list() -> None:
    assert ff.parse_id_list("80013618, 80011485\n80013618 080608\t") == ["80013618", "80011485", "80608"]


def test_scan_export_matches_names_and_leading_numbers(export: Path) -> None:
    index = ff.scan_export(export)
    assert [p.name for p in index["80013618"]] == ["80013618"]
    assert [p.name for p in index["80004215"]] == ["80004215 - Anshuman"]
    assert [p.name for p in index["80608"]] == ["80608_sign.pdf"]
    assert not any(p.name == ".DS_Store" for paths in index.values() for p in paths)


def test_clean_path_input_handles_terminal_escapes(export: Path) -> None:
    escaped = str(export).replace(" ", "\\ ").replace("(", "\\(").replace(")", "\\)")
    assert ff.clean_path_input(escaped) == export
    assert ff.clean_path_input(f"'{export}' ") == export


def test_value_options_group_case_and_spaces(workbook: Path) -> None:
    df = ff.combine_sheets(ff.load_workbook(workbook), ["tc", "field"])
    options = ff.value_options(df["Employer"], ignore_case=True)
    assert options == [("home credit", "Home credit", 2), ("other co", "Other Co", 1),
                       ("", ff.BLANK_LABEL, 1)]
    strict = ff.value_options(df["Employer"], ignore_case=False)
    assert [o[0] for o in strict] == ["HOME CREDIT", "Home credit", "Other Co", ""]


def test_filters_and_or_contains_dates(workbook: Path) -> None:
    df = ff.combine_sheets(ff.load_workbook(workbook), ["tc", "field"])

    def ids(*filters: ff.ColumnFilter) -> list[str]:
        return list(df[ff.filter_mask(df, filters, ignore_case=True)]["Employee ID"])

    assert ids() == list(df["Employee ID"])
    # OR inside a column
    assert ids(ff.ColumnFilter("Phase", values=["batch 01", "june phase 1"])) == ["80011485", "080608", "80004215"]
    # AND across columns
    assert ids(ff.ColumnFilter("Phase", values=["batch 01"]),
               ff.ColumnFilter("Employer", values=["home credit"])) == ["80011485"]
    assert ids(ff.ColumnFilter("Employer", values=[""])) == ["80004215"]
    assert ids(ff.ColumnFilter("Candidate Name", contains=["aman", " khush "])) == ["80011485", "080608"]
    assert ids(ff.ColumnFilter("Registration Date", date_from=date(2026, 7, 8),
                               date_to=date(2026, 7, 20))) == ["80013618", "80011485"]
    assert ff.is_date_column("Registration Date", df["Registration Date"])
    assert not ff.is_date_column("Phase", df["Phase"])


def test_parse_dates_keeps_iso_month_and_reads_day_first() -> None:
    parsed = ff.parse_dates(pd.Series(["2026-07-08 16:49:59", "08/07/2026", "N/A"]))
    assert parsed[0].date() == date(2026, 7, 8)
    assert parsed[1].date() == date(2026, 7, 8)
    assert pd.isna(parsed[2])


def test_export_copies_only_selected(workbook: Path, export: Path, tmp_path: Path) -> None:
    df = ff.combine_sheets(ff.load_workbook(workbook), ["tc", "field"])
    rows = df[df["Phase"].isin(["Phase 1 August", "Batch 01"])]
    index = ff.scan_export(export)
    seen: list[tuple[int, int]] = []

    result = ff.export_selection(
        rows, "Employee ID", index, ff.default_output_dir(export), make_zip=True,
        progress=lambda i, total, _name: seen.append((i, total)),
    )

    assert result.out_dir == tmp_path / "Export-2026-09-23 (1)_filtered"
    assert sorted(p.name for p in result.out_dir.iterdir()) == [
        "80011485", "80013618", "80608_sign.pdf", "filtered_list.xlsx"]
    assert (result.out_dir / "80013618" / "photo.jpg").read_bytes() == b"jpg"
    assert result.missing_ids == []
    assert seen == [(1, 3), (2, 3), (3, 3)]
    assert (export / "99999999").is_dir()  # source untouched

    report = pd.read_excel(result.report, dtype=str)
    assert list(report[ff.FOLDER_FOUND_COLUMN]) == ["Yes", "Yes", "Yes"]

    with zipfile.ZipFile(result.zip_path) as zf:
        names = zf.namelist()
    assert "Export-2026-09-23 (1)_filtered/80013618/photo.jpg" in names
    assert not any("99999999" in n for n in names)

    # a second run never overwrites the first
    again = ff.export_selection(rows.iloc[:1], "Employee ID", index, ff.default_output_dir(export))
    assert again.out_dir.name == "Export-2026-09-23 (1)_filtered (2)"
    assert again.zip_path is None


def test_export_reports_missing_folders(workbook: Path, tmp_path: Path) -> None:
    df = ff.combine_sheets(ff.load_workbook(workbook), ["tc"])
    empty_export = tmp_path / "empty"
    empty_export.mkdir()
    result = ff.export_selection(df, "Employee ID", ff.scan_export(empty_export), tmp_path / "out")
    assert result.copied == []
    assert result.missing_ids == ["80013618", "80011485", "80608"]
    missing = pd.read_excel(result.report, sheet_name="Missing folders", dtype=str)
    assert list(missing["Employee ID"]) == ["80013618", "80011485", "80608"]
