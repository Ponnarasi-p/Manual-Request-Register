"""Streamlit front end: Upload Excel -> Generate Report -> Download Excel."""
from __future__ import annotations

import copy
import hashlib
import re
from datetime import date, datetime

import pandas as pd
import streamlit as st

import processor as P

st.set_page_config(page_title="Manual Request Register", page_icon="📦", layout="wide")


@st.cache_data
def base_config() -> dict:
    return P.load_config()


def sheet_date(name: str) -> date | None:
    m = re.search(r"(\d{1,2})[-./](\d{1,2})[-./](\d{4})", name or "")
    if not m:
        return None
    try:
        return datetime(int(m[3]), int(m[2]), int(m[1])).date()
    except ValueError:
        return None


st.title("Manual Request Register – Report Generator")
st.caption(
    "Upload the original Manual Request Register (.xlsx). The app keeps the required columns, "
    "calculates ENV and TAG in Kgs using each machine line's conversion factor (+5%), adds the "
    "four blank Issued columns, and gives you one formatted Excel file."
)

# ---------------------------------------------------------------- conversion factors
config = copy.deepcopy(base_config())
with st.expander("Conversion factors (pieces per kg)", expanded=False):
    st.write(
        "Values come from `conversion_factors.json` in the repository. Changes made here apply to "
        "this session only; update the JSON file to change them permanently."
    )
    factor_df = pd.DataFrame(
        [{"Machine line / category": k, "Conversion factor": v} for k, v in config["factors"].items()]
    )
    edited = st.data_editor(
        factor_df,
        disabled=["Machine line / category"],
        hide_index=True,
        use_container_width=True,
        column_config={"Conversion factor": st.column_config.NumberColumn(min_value=1, step=1, format="%d")},
        key="factor_editor",
    )
    config["factors"] = {
        row["Machine line / category"]: float(row["Conversion factor"])
        for _, row in edited.iterrows()
        if pd.notna(row["Conversion factor"]) and row["Conversion factor"] > 0
    }
    config["uplift_percent"] = st.number_input(
        "Allowance (%) applied to the kg value", min_value=100.0, max_value=200.0,
        value=float(config.get("uplift_percent", 105)), step=1.0,
    )
    st.markdown("**Work centre → factor rules** (first match wins)")
    st.dataframe(
        pd.DataFrame(
            [
                {
                    "Work centre starts with / pattern": r["pattern"],
                    "ENV factor": r.get("env") or "—",
                    "TAG factor": r.get("tag") or "—",
                    "Examples": r.get("note", ""),
                }
                for r in config["rules"]
            ]
        ),
        hide_index=True,
        use_container_width=True,
    )

try:
    P.validate_config(config)
except ValueError as exc:
    st.error(f"Conversion factor settings are invalid: {exc}")
    st.stop()

# ---------------------------------------------------------------- upload
uploaded = st.file_uploader("1. Upload the Manual Request Register", type=["xlsx"])
if not uploaded:
    st.info("Waiting for an .xlsx file.")
    st.stop()

data = uploaded.getvalue()
digest = hashlib.sha1(data).hexdigest()

try:
    wb = P.open_workbook(data)
except P.InputError as exc:
    st.error(str(exc))
    st.stop()

valid_sheets = P.candidate_sheets(wb)
if not valid_sheets:
    header = P._find_header(wb.worksheets[0])
    found = header[1] if header else {}
    missing = [n for n, _ in P.REQUIRED_COLUMNS if n not in found]
    st.error(
        "No sheet in this workbook has all the mandatory columns "
        f"({', '.join(n for n, _ in P.REQUIRED_COLUMNS)}). "
        f"First sheet is missing: {', '.join(missing) or 'n/a'}."
    )
    st.stop()

col_a, col_b = st.columns([2, 1])
with col_a:
    sheet = st.selectbox(
        "2. Sheet to process",
        valid_sheets,
        help="Only sheets that contain all mandatory columns are listed.",
    )
with col_b:
    report_date = st.date_input("Report date (used for sheet and file name)", value=sheet_date(sheet) or date.today(),
                                format="DD/MM/YYYY")

if st.button("3. Generate Report", type="primary"):
    try:
        result = P.process(data, config, sheet_name=sheet)
        title = report_date.strftime("%d-%m-%Y")
        st.session_state["report"] = {
            "key": (digest, sheet),
            "result": result,
            "xlsx": P.build_workbook(result, sheet_title=title, report_date=title),
            "filename": f"Manual Request Register {title}.xlsx",
        }
    except P.InputError as exc:
        st.session_state.pop("report", None)
        st.error(str(exc))
        st.stop()

report = st.session_state.get("report")
if not report or report["key"] != (digest, sheet):
    st.stop()

result: P.Result = report["result"]

# ---------------------------------------------------------------- results
m1, m2, m3, m4 = st.columns(4)
m1.metric("Rows", len(result.rows))
m2.metric("ENV total (kg)", f"{result.total_env_kg:,.0f}")
m3.metric("TAG total (kg)", f"{result.total_tag_kg:,.0f}")
m4.metric("Rows flagged", len({i.row for i in result.issues if i.level == "error"}))

if result.issues:
    errors = [i for i in result.issues if i.level == "error"]
    if errors:
        st.warning(
            f"{len(errors)} value(s) could not be calculated. They are left blank and highlighted red "
            "(with a note) in the Excel file."
        )
    st.dataframe(
        pd.DataFrame(
            [{"Input row": i.row, "Type": i.level.title(), "Item": i.item, "Problem": i.message} for i in result.issues]
        ),
        hide_index=True,
        use_container_width=True,
    )
else:
    st.success("All rows processed. Every ENV / TAG quantity was matched to a conversion factor.")

preview = pd.DataFrame(
    [
        {
            "Work centre": r.work_centre,
            "Item": r.item,
            "Day Plan Qty CFC": r.day_plan_qty,
            "Contractor": r.contractor,
            "ENV": r.env,
            "TAG": r.tag,
            "ENV factor": r.env_factor,
            "TAG factor": r.tag_factor,
            "ENV in Kgs": None if r.env_kg is None else round(r.env_kg, 2),
            "TAG in Kgs": None if r.tag_kg is None else round(r.tag_kg, 2),
        }
        for r in result.rows
    ]
)
st.subheader("Preview")
st.caption("Factors are shown here for checking only; the Excel file has exactly the 12 report columns.")
st.dataframe(preview, hide_index=True, use_container_width=True)

st.download_button(
    "⬇️ Download Excel",
    data=report["xlsx"],
    file_name=report["filename"],
    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    type="primary",
)
