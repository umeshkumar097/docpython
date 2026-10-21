# Employee documents filter

Doc-status Excel / CSV aur documents ki ZIP (har employee ka folder = Employee ID) upload
karo, **Batch / Language / Education / Month** filter lagao, rows select karo aur sirf
selected employees ke folders ek nayi ZIP mein download karo, bilkul usi structure mein
jaisi upload ki thi.

```
export.zip (upload)                       export_selected_49.zip (download)
  Export-2026-09-23 (1)/                    Export-2026-09-23 (1)/
    80013618/  photo.jpg  id_proof.pdf        80013618/  photo.jpg  id_proof.pdf
    80011485/  ...                            ...  (sirf selected employees)
    ...
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

1. **Upload**: doc status Excel (`.xlsx`) ya CSV, aur documents ki ZIP.
   - Saari sheets (jaise `tc`, `field`) ek saath aati hain.
   - **Sirf wahi employees aage aate hain jinka folder ZIP mein Employee ID se match hota hai.**
     Baaki Excel rows skip ho jaati hain ("Skipped Excel rows" mein dikhti hain), aur ZIP ke jin
     folders ki ID Excel mein nahi hai wo bhi alag dikhte hain.
2. **Filters**: Batch, Language, Education, Month. Har filter mein sheet ki values unke count ke
   saath aati hain (Excel ke column filter jaisa). Ek filter mein kai values = OR, alag filters = AND.
   - Settings mein badal sakte ho ki kaunsa filter kis column se bane. Default:
     Batch = `MD_Dra Batch`, Language = `MD_..._Language`, Education = `Selected Qualification`,
     Month = `MD_Upload Month` (date column chuno to "July 2026" jaise months ban jaate hain).
   - "Ignore case, spaces, _ and -" on ho to `UNDERGRADUATE`, `Under Graduate`, `Under_GRADUATE`
     ek hi value hain.
   - "More filters" mein koi bhi aur column, ya Employee IDs paste karke filter.
3. **Select**: table mein sab rows pehle se ticked hoti hain; jinko nahi chahiye unka tick hatao
   (Select all / Clear all bhi hai). Upar dikhta hai kitne employees, kitne documents, kitna MB.
4. **Download**: "Download ZIP" se selected employees ke folders, same paths ke saath.
   "Download selected list (.xlsx)" se unki Excel list.

## Folder aur Employee ID kaise match hote hain

- Folder ka naam = `Employee ID` column ki value. Leading zeros, spaces aur `.0` ignore hote hain
  (`080608` = `80608`).
- `80013618 - Somanath` jaise folder ya `80013618_photo.jpg` jaisi files bhi apne shuru wale number
  se match hoti hain.
- ZIP ke andar ek bahari folder ho (`Export-2026-09-23 (1)/`) ya na ho, dono chalte hain.
  `__MACOSX` aur `.DS_Store` jaisi faltu files ignore hoti hain.
- Agar folders kisi aur number se bane hain, Settings mein **Employee ID column** badlo.

## Data

App sirf isi computer par khulta hai (`localhost`); files kahin upload nahi hoti. ZIP 4 GB tak
upload ho sakti hai (`.streamlit/config.toml`). Excel, CSV aur ZIP `.gitignore` mein hain, kabhi
git mein commit nahi hote.

## Tests

```bash
pip install pytest
pytest -q
```
