# Manual Request Register – Report Generator

Streamlit app that turns the original Manual Request Register Excel report into the
formatted daily issue sheet: **Upload Excel → Generate Report → Download Excel**.

## What it does
1. Reads the uploaded `.xlsx` and finds the sheet/header row with the mandatory columns
   (Day, Work centre, Item, Day Plan Qty, Contractor, ENV, TAG). CTN, BLEND, PROMO, etc. are dropped.
2. Cleans quantities written as text (e.g. `  6,00,000`) into numbers.
3. Picks the ENV and TAG conversion factor for each row from its Work centre
   (`conversion_factors.json`) and writes `ENV in Kgs = (ENV / factor) × 105%`,
   `TAG in Kgs = (TAG / factor) × 105%` as Excel formulas.
4. Adds the four blank columns: Issued Env in Pcs, Issued Env in kgs, Issued Tag in Pcs, Issued Tag in kgs.
5. Adds the TOTAL row (sum of ENV kg, sum of TAG kg, and the combined total across J:M), as in the reference sheet.
6. Produces exactly one `.xlsx` file to download.

Rows whose Work centre isn't recognised, or whose quantity isn't a number, are kept, but the kg
cell is left blank, highlighted red with a cell note, and listed in the app with the input row number.

## Work centre → conversion factor mapping

| Work centre (input) | ENV factor | TAG factor |
|---|---|---|
| C250, C250-A, C250 AF | CONSTANTA-250 – 1740 | Constanta Tag – 10780 |
| CONSTANTA ENV, ENV BT, ENV BT 25, ENV F/G/H/HF/M | Constanta Envelope – 1525 | Constanta Tag – 10780 |
| MD20 ENV, MD20 ENV GARANT, MD20 4GM TAG | MD-20 Crimp Envelope – 1700 | MD-20 Tag – 12600 |
| PERFECTA ENV | Perfecta Envelope – 1520 | Perfecta Tag – 10780 |
| MAISA ENV | Maisa Envelope – 1410 | Maisa Tag – 9080 |
| MD HS … PAPER / MD HS … FOIL | 1425 / 1190 | – |
| CONSTANTA … FAMILY | Constanta Family Size – 6320 | 10780 |
| CONSTANTA TAG BIG | – | Constanta Tag Big – 6040 |
| PK-45 | – | 10780 |

Factors come from the machine-line conversion round-off chart. The mapping was checked against the
formulas in the 20-10, 19-11 and 21-11 sheets of the reference workbook.

To change a factor or add a work centre, edit `conversion_factors.json` (`factors` for values,
`rules` for regex patterns matched against the Work centre, first match wins). Factors can also be
changed for a single session in the app's "Conversion factors" panel.

## Run locally
```
pip install -r requirements.txt
streamlit run app.py
```

## Tests
```
pip install pytest
python -m pytest
```

## Deploy on Streamlit Community Cloud
New app → this repository → branch `main` → main file `app.py`.
