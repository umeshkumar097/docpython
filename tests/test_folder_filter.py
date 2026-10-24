from datetime import date
from pathlib import Path
import io
import sys
import zipfile

import pandas as pd
import pytest

sys.path.insert(0, str(Path(__file__).resolve().parents[1]))
import folder_filter as ff  # noqa: E402


def xlsx_bytes(sheets: dict[str, pd.DataFrame]) -> bytes:
    buf = io.BytesIO()
    with pd.ExcelWriter(buf) as writer:
        for name, df in sheets.items():
            df.to_excel(writer, sheet_name=name, index=False)
    return buf.getvalue()


@pytest.fixture
def df() -> pd.DataFrame:
    tc = pd.DataFrame({
        "Registration Date": ["2026-07-08 16:49:59", "2026-07-20 10:00:00", "2026-08-02 09:00:00"],
        "Candidate Name": ["SOMANATH", "AMANDEEP", "KHUSHBU"],
        "Phase": ["Phase 1 August", "Batch 01", "Batch 01"],
        "Employer": ["Home credit ", "HOME CREDIT", "Other Co"],
        "Employee ID": [80013618, "80011485", "080608"],
        "Selected Qualification": ["UNDERGRADUATE", "Under Graduate", "GRADUATE"],
        "Qual Document": ["Uploaded"] * 3,
        "MD_Dra Batch": [None, "Batch 01", "Batch 01"],
        "MD_Upload Month": ["July 2026", "July 2026", "June 2026"],
        "MD_Training Month": [None, None, "September 2026"],
        "MD_Training Language": [None, "Hindi", None],
    })
    field = pd.DataFrame({
        "Registration Date": ["2026-05-30 08:51:44"],
        "Candidate Name": ["ANSHUMAN"],
        "Phase": ["June Phase 1"],
        "Employer": [None],
        "Employee ID": ["80004215"],
        "Selected Qualification": ["Under_GRADUATE"],
        "MD_Dra Batch": ["Batch 04"],
        "MD_Upload Month": ["June 2026"],
    })
    sheets = ff.load_table(xlsx_bytes({"tc": tc, "field": field}), "status.xlsx")
    return ff.combine_sheets(sheets, ["tc", "field"])


def make_zip(files: dict[str, bytes], dirs: tuple[str, ...] = ()) -> zipfile.ZipFile:
    buf = io.BytesIO()
    with zipfile.ZipFile(buf, "w", zipfile.ZIP_DEFLATED) as zf:
        for d in dirs:
            zf.writestr(zipfile.ZipInfo(d), b"")
        for name, data in files.items():
            zf.writestr(name, data)
    buf.seek(0)
    return zipfile.ZipFile(buf)


EXPORT = "Export-2026-09-23 (1)/"


@pytest.fixture
def export_zip() -> zipfile.ZipFile:
    return make_zip(
        {
            f"{EXPORT}80013618/photo.jpg": b"jpg-1",
            f"{EXPORT}80013618/id_proof.pdf": b"pdf-1",
            f"{EXPORT}80011485/photo.jpg": b"jpg-2",
            f"{EXPORT}80011485/notes.txt": b"text " * 100,
            f"{EXPORT}80004215 - Anshuman/photo.jpg": b"jpg-3",
            f"{EXPORT}80608_sign.pdf": b"pdf-4",
            f"{EXPORT}99999999/photo.jpg": b"jpg-x",
            f"{EXPORT}.DS_Store": b"",
            "__MACOSX/Export-2026-09-23 (1)/._80013618": b"",
        },
        dirs=(EXPORT, f"{EXPORT}80013618/", f"{EXPORT}80011485/"),
    )


# -- Excel / CSV ------------------------------------------------------------


def test_load_and_combine_keeps_ids_as_text(df: pd.DataFrame) -> None:
    assert list(df[ff.SHEET_COLUMN]) == ["tc", "tc", "tc", "field"]
    assert list(df["Employee ID"]) == ["80013618", "80011485", "080608", "80004215"]
    assert "MD_Training Language" in df.columns  # columns of both sheets are kept


@pytest.mark.parametrize("encoding", ["utf-8-sig", "cp1252"])
def test_load_csv(encoding: str) -> None:
    data = "Employee ID,Phase,City\n080608,Batch 01,Pune\n80011485,,Café\n".encode(encoding)
    sheets = ff.load_table(data, "status.CSV")
    assert list(sheets) == ["CSV"]
    assert list(sheets["CSV"]["Employee ID"]) == ["080608", "80011485"]
    assert sheets["CSV"]["City"][1] == "Café"


def test_unique_name() -> None:
    assert ff.unique_name("Select", ["Phase"]) == "Select"
    assert ff.unique_name("Select", ["Select", "Select (2)"]) == "Select (3)"


def test_headers_that_only_differ_by_spaces_stay_separate() -> None:
    sheet = pd.DataFrame({"Employee ID": ["1"], "Phase": ["Batch 01"], "Phase ": ["Batch 02"],
                          "Sheet": ["mine"], "Sheet (original)": ["also mine"]})
    sheets = ff.load_table(xlsx_bytes({"tc": sheet}), "status.xlsx")
    assert list(sheets["tc"].columns) == ["Employee ID", "Phase", "Phase (2)", "Sheet", "Sheet (original)"]
    df = ff.combine_sheets(sheets, ["tc"])
    assert list(df.columns) == [ff.SHEET_COLUMN, "Employee ID", "Phase", "Phase (2)",
                                "Sheet (original) (2)", "Sheet (original)"]
    assert df["Sheet (original) (2)"][0] == "mine" and df["Phase (2)"][0] == "Batch 02"


def test_guess_id_column_prefers_main_column() -> None:
    assert ff.guess_id_column(["Name", "MD_Employee ID", "Employee ID"]) == "Employee ID"
    assert ff.guess_id_column(["MD_Employee ID ", "Name"]) == "MD_Employee ID "
    assert ff.guess_id_column(["Name", "Phone"]) is None


def test_guess_filter_columns(df: pd.DataFrame) -> None:
    assert ff.guess_filter_column(df, "Batch") == "MD_Dra Batch"
    assert ff.guess_filter_column(df, "Language") == "MD_Training Language"
    assert ff.guess_filter_column(df, "Education") == "Selected Qualification"
    assert ff.guess_filter_column(df, "Month") == "MD_Upload Month"  # most filled month column
    only_phase = pd.DataFrame({"Phase": ["Batch 01"], "Registration Date": ["2026-07-08"]})
    assert ff.guess_filter_column(only_phase, "Batch") == "Phase"
    assert ff.guess_filter_column(only_phase, "Month") == "Registration Date"
    assert ff.guess_filter_column(only_phase, "Language") is None


# -- IDs --------------------------------------------------------------------


@pytest.mark.parametrize(
    ("raw", "expected"),
    [(" 80013618 ", "80013618"), ("80013618.0", "80013618"), ("080608", "80608"),
     (80013618, "80013618"), ("ab12", "AB12"), ("N/A", ""), (None, ""), (float("nan"), "")],
)
def test_normalize_id(raw: object, expected: str) -> None:
    assert ff.normalize_id(raw) == expected


def test_parse_id_list() -> None:
    assert ff.parse_id_list("80013618, 80011485\n80013618 080608\t") == ["80013618", "80011485", "80608"]


def test_name_keys() -> None:
    assert ff.name_keys("80013618") == {"80013618"}
    assert ff.name_keys("80004215 - Anshuman") == {"80004215-ANSHUMAN", "80004215"}
    assert ff.name_keys("Photos") == {"PHOTOS"}


# -- Filters ----------------------------------------------------------------


def test_clean_value_ignores_case_spaces_and_separators() -> None:
    variants = ["UNDERGRADUATE", "Under Graduate", "Under_GRADUATE", " under-graduate "]
    assert {ff.clean_value(v, True) for v in variants} == {"undergraduate"}
    assert ff.clean_value("  Under   Graduate ", False) == "Under Graduate"
    assert ff.clean_value(None, True) == ""


def test_value_options(df: pd.DataFrame) -> None:
    assert ff.value_options(df["Selected Qualification"], ignore_case=True) == [
        ("graduate", "GRADUATE", 1), ("undergraduate", "UNDERGRADUATE", 3)]
    assert ff.value_options(df["Employer"], ignore_case=True) == [
        ("homecredit", "Home credit", 2), ("otherco", "Other Co", 1), ("", ff.BLANK_LABEL, 1)]
    strict = ff.value_options(df["Employer"], ignore_case=False)
    assert [o[0] for o in strict] == ["HOME CREDIT", "Home credit", "Other Co", ""]


def test_month_values_and_calendar_order(df: pd.DataFrame) -> None:
    months = ff.month_values("Registration Date", df["Registration Date"])
    assert list(months) == ["July 2026", "July 2026", "August 2026", "May 2026"]
    assert [o[1] for o in ff.value_options(months, True, chronological=True)] == [
        "May 2026", "July 2026", "August 2026"]
    text = pd.Series(["Sept batch", "July 2026", None, "Jun 2026"])
    assert ff.month_values("MD_Upload Month", text) is text
    assert [o[1] for o in ff.value_options(text, True, chronological=True)] == [
        "Jun 2026", "July 2026", "Sept batch", ff.BLANK_LABEL]


def test_values_mask(df: pd.DataFrame) -> None:
    ids = lambda m: list(df[m]["Employee ID"])  # noqa: E731
    assert ids(ff.values_mask(df["Selected Qualification"], ["undergraduate"], True)) == [
        "80013618", "80011485", "80004215"]
    assert ids(ff.values_mask(df["MD_Dra Batch"], [], True)) == list(df["Employee ID"])
    assert ids(ff.values_mask(df["MD_Dra Batch"], [""], True)) == ["80013618"]


def test_filters_and_or_contains_dates(df: pd.DataFrame) -> None:
    def ids(*filters: ff.ColumnFilter) -> list[str]:
        return list(df[ff.filter_mask(df, filters, ignore_case=True)]["Employee ID"])

    assert ids() == list(df["Employee ID"])
    assert ids(ff.ColumnFilter("Phase", values=["batch01", "junephase1"])) == ["80011485", "080608", "80004215"]
    assert ids(ff.ColumnFilter("Phase", values=["batch01"]),
               ff.ColumnFilter("Employer", values=["homecredit"])) == ["80011485"]
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


# -- ZIP --------------------------------------------------------------------


def test_index_zip_looks_through_wrapper_and_skips_junk(export_zip: zipfile.ZipFile) -> None:
    index = ff.index_zip(export_zip)
    assert index.root == EXPORT
    assert [f.name for f in index.folders] == [
        "80004215 - Anshuman", "80011485", "80013618", "80608_sign.pdf", "99999999"]
    folder = index.by_id["80013618"][0]
    assert (folder.files, folder.size) == (2, 10)
    assert index.by_id["80004215"][0].name == "80004215 - Anshuman"
    assert index.by_id["80608"][0].name == "80608_sign.pdf"
    assert not any("__MACOSX" in m.filename or ".DS_Store" in m.filename
                   for f in index.folders for m in f.members)


def test_index_zip_without_wrapper_and_single_employee() -> None:
    flat = ff.index_zip(make_zip({"80013618/a.jpg": b"1", "80011485/b.jpg": b"2"}))
    assert flat.root == ""
    assert sorted(flat.by_id) == ["80011485", "80013618"]

    single = make_zip({"80013618/a.jpg": b"1", "80013618/b.pdf": b"2"})
    assert ff.index_zip(single, known_ids={"80013618"}).by_id["80013618"][0].files == 2
    # without knowing the ID the lone folder is treated as a wrapper
    assert ff.index_zip(single).root == "80013618/"


def test_build_zip_keeps_paths_of_selected_only(export_zip: zipfile.ZipFile) -> None:
    index = ff.index_zip(export_zip)
    chosen = index.by_id["80013618"] + index.by_id["80011485"] + index.by_id["80013618"]
    out = io.BytesIO()
    seen: list[tuple[int, int]] = []

    assert ff.build_zip(export_zip, index, chosen, out, progress=lambda i, t: seen.append((i, t))) == 4

    with zipfile.ZipFile(out) as result:
        assert sorted(result.namelist()) == sorted([
            EXPORT, f"{EXPORT}80011485/", f"{EXPORT}80011485/notes.txt", f"{EXPORT}80011485/photo.jpg",
            f"{EXPORT}80013618/", f"{EXPORT}80013618/id_proof.pdf", f"{EXPORT}80013618/photo.jpg",
        ])
        assert result.read(f"{EXPORT}80013618/id_proof.pdf") == b"pdf-1"
        assert result.read(f"{EXPORT}80011485/notes.txt") == b"text " * 100
        assert result.getinfo(f"{EXPORT}80013618/photo.jpg").compress_type == zipfile.ZIP_STORED
        assert result.getinfo(f"{EXPORT}80011485/notes.txt").compress_type == zipfile.ZIP_DEFLATED
        assert result.testzip() is None
    assert seen[-1] == (4, 4)
    # the source archive is still readable afterwards
    assert export_zip.read(f"{EXPORT}80013618/photo.jpg") == b"jpg-1"


def test_save_zip_streams_to_disk_without_overwriting(export_zip: zipfile.ZipFile, tmp_path: Path) -> None:
    index = ff.index_zip(export_zip)
    target = tmp_path / "Downloads" / "export_selected_1.zip"
    seen: list[tuple[int, int]] = []

    first = ff.save_zip(export_zip, index, index.by_id["80013618"], target,
                        progress=lambda i, t: seen.append((i, t)))
    second = ff.save_zip(export_zip, index, index.by_id["80011485"], target)

    assert first == target
    assert second == tmp_path / "Downloads" / "export_selected_1 (2).zip"
    assert seen == [(1, 2), (2, 2)]
    assert sorted(p.name for p in target.parent.iterdir()) == [
        "export_selected_1 (2).zip", "export_selected_1.zip"]  # no .part left behind
    with zipfile.ZipFile(first) as result:
        assert result.read(f"{EXPORT}80013618/photo.jpg") == b"jpg-1"


def test_save_zip_removes_partial_file_on_failure(export_zip: zipfile.ZipFile, tmp_path: Path) -> None:
    index = ff.index_zip(export_zip)

    def boom(_done: int, _total: int) -> None:
        raise RuntimeError("disk full")

    with pytest.raises(RuntimeError):
        ff.save_zip(export_zip, index, index.by_id["80013618"], tmp_path / "out.zip", boom)
    assert list(tmp_path.iterdir()) == []


def test_clean_path_input_handles_quotes_tilde_and_terminal_escapes(tmp_path: Path) -> None:
    zip_path = tmp_path / "Export-2026-09-23 (1).zip"
    zip_path.write_bytes(b"")
    escaped = str(zip_path).replace(" ", "\\ ").replace("(", "\\(").replace(")", "\\)")
    assert ff.clean_path_input(escaped) == zip_path
    assert ff.clean_path_input(f"'{zip_path}' ") == zip_path
    assert ff.clean_path_input(f'"{zip_path}"') == zip_path
    assert ff.clean_path_input("~/x.zip") == Path.home() / "x.zip"


def test_find_zip_files_newest_first_without_own_outputs(tmp_path: Path) -> None:
    import os

    downloads, desktop = tmp_path / "Downloads", tmp_path / "Desktop"
    downloads.mkdir()
    desktop.mkdir()
    for i, path in enumerate([downloads / "old.zip", desktop / "new.zip", downloads / "notes.txt",
                              downloads / "export_selected_49.zip", downloads / ".hidden.zip"]):
        path.write_bytes(b"")
        os.utime(path, (1_000_000 + i, 1_000_000 + i))
    found = ff.find_zip_files([downloads, desktop, tmp_path / "Missing"])
    assert [p.name for p in found] == ["new.zip", "old.zip"]


# -- Matching real-world ZIP / Excel layouts ------------------------------------


IDS = ["80013618", "80011485", "80004215"]


def docs(prefix: str, folder: str) -> dict[str, bytes]:
    return {f"{prefix}{folder}/photo.jpg": b"jpg", f"{prefix}{folder}/id_proof.pdf": b"pdf"}


def matched_paths(zf: zipfile.ZipFile, excel_ids: list[str]) -> list[str]:
    ids = pd.Series(excel_ids)
    index = ff.index_zip(zf, known_ids=ff.known_ids(ids))
    keys = ff.match_keys(ids, index.by_id)
    return sorted("/".join(f.path) for k in keys for f in index.by_id.get(k, []))


@pytest.mark.parametrize(
    ("layout", "expected"),
    [
        ({**docs(EXPORT, IDS[0]), **docs(EXPORT, IDS[1]), **docs(EXPORT, IDS[2]), "doc_status.xlsx": b"x"},
         [f"{EXPORT}{i}" for i in sorted(IDS)]),
        ({**docs(f"{EXPORT}tc/", IDS[0]), **docs(f"{EXPORT}tc/", IDS[1]), **docs(f"{EXPORT}field/", IDS[2])},
         [f"{EXPORT}field/{IDS[2]}", f"{EXPORT}tc/{IDS[1]}", f"{EXPORT}tc/{IDS[0]}"]),
        ({**docs(EXPORT, f"SOMANATH_{IDS[0]}"), **docs(EXPORT, f"AMANDEEP {IDS[1]}"), **docs(EXPORT, IDS[2])},
         [f"{EXPORT}{IDS[2]}", f"{EXPORT}AMANDEEP {IDS[1]}", f"{EXPORT}SOMANATH_{IDS[0]}"]),
        ({**docs("Export (1)/", IDS[0]), **docs("Export (2)/", IDS[1]), **docs("Export (2)/", IDS[2])},
         [f"Export (1)/{IDS[0]}", f"Export (2)/{IDS[2]}", f"Export (2)/{IDS[1]}"]),
        ({f"Photos/{i}.jpg": b"jpg" for i in IDS} | {f"Signatures/{i}_sign.png": b"png" for i in IDS},
         sorted([f"Photos/{i}.jpg" for i in IDS] + [f"Signatures/{i}_sign.png" for i in IDS])),
    ],
    ids=["excel-next-to-export", "sheet-subfolders", "name-around-id", "two-exports", "grouped-by-type"],
)
def test_employee_folders_are_found_at_any_depth(layout: dict[str, bytes], expected: list[str]) -> None:
    assert matched_paths(make_zip(layout), IDS) == sorted(expected)


def test_short_ids_inside_folder_names_match_but_never_the_export_folder() -> None:
    zf = make_zip({**docs(EXPORT, "VIBHA_8670"), **docs(EXPORT, IDS[0]), **docs(EXPORT, "OTHER_99")})
    # 2026 and 23 appear in the export folder's own name; 99 is in a folder name below it
    assert matched_paths(zf, ["8670", IDS[0], "2026", "23"]) == [f"{EXPORT}{IDS[0]}", f"{EXPORT}VIBHA_8670"]
    index = ff.index_zip(zf, known_ids={"8670", "2026", "23", IDS[0]})
    assert index.root == EXPORT and "2026" not in index.by_id and "23" not in index.by_id


def test_excel_ids_with_prefix_or_scientific_notation_match() -> None:
    zf = make_zip({**docs(EXPORT, IDS[0]), **docs(EXPORT, IDS[1])})
    assert matched_paths(zf, [f"HC-{IDS[0]}", "8.0011485E+07"]) == [f"{EXPORT}{IDS[1]}", f"{EXPORT}{IDS[0]}"]


def test_photo_names_inside_a_folder_never_look_like_employees() -> None:
    zf = make_zip({f"{EXPORT}6076774/IMG_20260923_123456.jpg": b"x", f"{EXPORT}{IDS[0]}/a.jpg": b"x"})
    index = ff.index_zip(zf, known_ids={"123456", IDS[0]})
    assert sorted("/".join(f.path) for f in index.folders) == [f"{EXPORT}6076774", f"{EXPORT}{IDS[0]}"]
    assert "123456" not in index.by_id


def test_guess_id_column_reads_header_variants() -> None:
    assert ff.guess_id_column(["Registration Date", "Employee_ID"]) == "Employee_ID"
    assert ff.guess_id_column(["Name", "Emp. Code"]) == "Emp. Code"
    assert ff.guess_id_column(["Name", "EMPLOYEE NO"]) == "EMPLOYEE NO"


def test_best_id_column_follows_the_zip_folder_names() -> None:
    df = pd.DataFrame({"Employee ID": IDS, "Application No": ["6076774", "6076656", "6076683"],
                       "Pincode": ["400708", "122015", "560100"]})
    zip_keys = ff.zip_name_keys(make_zip({**docs(EXPORT, "6076774"), **docs(EXPORT, "6076656")}))
    assert ff.best_id_column(df, list(df.columns), zip_keys) == ("Application No", 2)
    assert ff.count_matches(df["Employee ID"], zip_keys) == 0


def test_build_zip_keeps_nested_paths_and_only_needed_parent_folders() -> None:
    zf = make_zip({**docs(f"{EXPORT}tc/", IDS[0]), **docs(f"{EXPORT}field/", IDS[2])},
                  dirs=(EXPORT, f"{EXPORT}tc/", f"{EXPORT}field/"))
    index = ff.index_zip(zf, known_ids=set(IDS))
    out = io.BytesIO()
    ff.build_zip(zf, index, index.by_id[IDS[0]], out)
    with zipfile.ZipFile(out) as result:
        assert sorted(result.namelist()) == [
            EXPORT, f"{EXPORT}tc/", f"{EXPORT}tc/{IDS[0]}/id_proof.pdf", f"{EXPORT}tc/{IDS[0]}/photo.jpg"]


def test_selected_list_xlsx(df: pd.DataFrame) -> None:
    back = pd.read_excel(io.BytesIO(ff.selected_list_xlsx(df.iloc[:2])), dtype=str)
    assert list(back["Employee ID"]) == ["80013618", "80011485"]
