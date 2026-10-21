# Employee documents filter

Upload the doc-status Excel / CSV and the documents ZIP (one folder per Employee ID), filter by
**Batch / Language / Education / Month**, select rows, and download only the selected employees'
folders in a new ZIP, with exactly the same structure as the ZIP you uploaded.

```
export.zip (upload)                       export_selected_49.zip (download)
  Export-2026-09-23 (1)/                    Export-2026-09-23 (1)/
    80013618/  photo.jpg  id_proof.pdf        80013618/  photo.jpg  id_proof.pdf
    80011485/  ...                            ...  (selected employees only)
    ...
```

## Run it (Mac)

1. Python 3 is required (check with `python3 --version` in Terminal; install it from python.org if
   it is missing).
2. Double-click `run_mac.command`. If macOS blocks it the first time, right-click it and choose
   **Open**. The first run installs the packages, then the app opens in your browser.

Or from Terminal (Mac / Windows / Linux):

```bash
python3 -m venv .venv
source .venv/bin/activate          # Windows: .venv\Scripts\activate
pip install -r requirements.txt
streamlit run app.py --server.address localhost
```

## How to use it

1. **Upload** the doc-status Excel (`.xlsx`) or CSV, and choose the documents ZIP:
   - **Upload (up to 2 GB)**: pick the ZIP in the browser.
   - **From this computer (large ZIP, 5 GB and more)**: choose a ZIP from your Downloads, Desktop
     or Documents folder, or paste its full path (Finder: right-click the ZIP, hold Option,
     **Copy ... as Pathname**). The ZIP is read straight from disk, so its size does not matter.
     This option appears when the app runs on your own computer.
   - All sheets (for example `tc` and `field`) are loaded together.
   - **Only employees whose folder is in the ZIP are kept.** Other Excel rows are skipped (listed
     under "Skipped Excel rows"), and ZIP folders whose ID is not in the Excel are listed separately.
2. **Filters**: Batch, Language, Education and Month. Each filter lists the sheet's values with
   their counts, like an Excel column filter. Several values in one filter = OR; different
   filters = AND.
   - Settings lets you choose which column each filter uses. Defaults: Batch = `MD_Dra Batch`,
     Language = the first `..._Language` column, Education = `Selected Qualification`,
     Month = `MD_Upload Month` (choosing a date column turns dates into months such as
     "July 2026").
   - With "Ignore case, spaces, _ and -" on, `UNDERGRADUATE`, `Under Graduate` and
     `Under_GRADUATE` are one value.
   - "More filters" adds any other column, or a pasted list of Employee IDs.
3. **Select**: every row starts ticked; untick the ones you do not want (Select all / Clear all are
   available too). The counters show employees, documents and size of the selection.
4. **Download** the selected employees' folders with their original paths:
   - **Save ZIP to Downloads** (on your own computer) writes the ZIP straight into your Downloads
     folder, with a progress bar. Any size works; an existing file is never overwritten.
   - **Download ZIP** downloads through the browser, for selections up to 1 GB.
   - **Download selected list (.xlsx)** gives the selected rows as an Excel file.

## How folders are matched to Employee IDs

- The folder name is the value of the `Employee ID` column. Leading zeros, spaces and a trailing
  `.0` are ignored (`080608` = `80608`).
- A folder such as `80013618 - Somanath` or a file such as `80013618_photo.jpg` matches by its
  leading number.
- The ZIP may or may not have an outer folder (`Export-2026-09-23 (1)/`). Clutter such as
  `__MACOSX` and `.DS_Store` is ignored.
- If the folders are named after a different number, choose that column as the
  **Employee ID column** in Settings.

## Data

When run on your computer (`run_mac.command` or the command above) the app only listens on
`localhost`; files never leave the machine. Excel, CSV and ZIP files are in `.gitignore` and never
committed.

## Large ZIPs

- A browser upload is held in the app's memory, about 2x the ZIP's size (a 1 GB upload uses about
  2 GB of RAM), so browser uploads are limited to 2 GB (`.streamlit/config.toml`).
- **From this computer** + **Save ZIP to Downloads** do not load the ZIP into memory: a 5 GB ZIP
  with 950 employee folders was opened in under a second and fully saved in about 70 seconds,
  with the app using under 200 MB of RAM. You need free disk space for the new ZIP (up to the
  size of the selection).

## Deploying on Streamlit Community Cloud (share.streamlit.io)

- The repository is private, so Streamlit needs access to your private GitHub repositories,
  otherwise it does not appear in the repository search. You can also use **Paste GitHub URL**
  with the link to `app.py`.
- Branch: the repository's default branch. Main file path: `app.py`.
- On the cloud, the uploaded Excel and documents (Aadhaar, PAN copies) are sent to Streamlit's
  servers: keep the app **private** (only invited emails can open it). Cloud memory is limited
  (roughly 1-2.7 GB), so on the cloud only browser uploads are available and a ZIP over about
  1 GB can crash there. For 5 GB ZIPs, run the app on your computer.

## Tests

```bash
pip install pytest
pytest -q
```
