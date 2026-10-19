# Employee folder filter

Doc-status Excel mein filter lagao (Phase, Residential State, Batch, Employer,
Registration Date, pasted Employee IDs ...) aur export folder se **sirf unhi
employees ke folders** ek naye folder / ZIP mein copy karo. Jitne filters,
utna hi data. Original export folder kabhi change nahi hota.

```
Export-2026-09-23 (1)/            <- portal se aaya export (folder = Employee ID)
  80013618/  80011485/  ...
Export-2026-09-23 (1)_filtered/   <- app banata hai: sirf filtered employees
  80013618/  ...  filtered_list.xlsx
Export-2026-09-23 (1)_filtered.zip
```

## Chalana (Mac)

1. Python 3 chahiye (Terminal mein `python3 --version`; na ho to python.org se install karo).
2. `run_mac.command` par double-click karo. Pehli baar macOS roke to right-click, **Open**.
   Pehli baar packages install hote hain, phir browser mein app khulta hai.

Ya Terminal se (Mac / Windows / Linux):

```bash
python3 -m venv .venv
source .venv/bin/activate          # Windows: .venv\Scripts\activate
pip install -r requirements.txt
streamlit run app.py
```

## Use kaise karein

1. **Excel**: doc status `.xlsx` upload karo. Saari sheets (jaise `tc`, `field`) ek saath aati hain;
   `Sheet` column se sheet ka filter bhi lag sakta hai.
2. **Export folder**: folder ka path paste karo. Finder mein folder par right-click, Option key
   dabaye rakho, **Copy "…" as Pathname**. (Folder ko Terminal mein drag karke bhi path milta hai;
   wo `\ ` wala path bhi chalega.)
3. **Filters**: "Filter by columns" mein columns chuno; har column ka apna filter aata hai.
   - Alag columns = AND (Phase = Batch 01 **aur** State = HARYANA).
   - Ek column ke andar chuni values = OR (Batch 01 **ya** Batch 02).
   - Date columns (Registration Date) par date range, bahut zyada values wale columns
     (Name, Email) par "contains" search.
   - "Ignore case and extra spaces" on ho to `Home credit `, `HOME CREDIT` ek hi value hain.
   - Employee IDs box mein IDs paste karke sirf unhe le sakte ho.
   - Folder status: sirf jinka folder mila / sirf jinka folder missing hai.
4. **Copy / download**: "Copy N folders" dabao. Naya folder (+ ZIP) export ke bagal mein banta hai,
   saath mein `filtered_list.xlsx` (filtered rows + "Folder found" Yes/No, aur missing IDs ki sheet).
   500 MB tak ka ZIP browser se bhi download ho jata hai.

## Folder aur Employee ID kaise match hote hain

- Folder ka naam = `Employee ID` column ki value. Leading zeros, spaces aur `.0` ignore hote hain
  (`080608` = `80608`).
- `80013618 - Somanath` jaise folder ya `80013618_photo.jpg` jaisi files bhi apne shuru wale number
  se match hoti hain.
- Agar folders kisi aur number se bane hain (jaise registration / application number), sidebar mein
  **Employee ID column** badal kar wo column chuno.
- App dikhata hai: Excel mein hai par folder nahi mila, aur folder hai par Excel mein ID nahi mili.

## Data

App sirf isi computer par khulta hai (`localhost`). Excel, ZIP aur `_filtered` folders
`.gitignore` mein hain, kabhi git mein commit nahi hote.

## Tests

```bash
pip install pytest
pytest -q
```
