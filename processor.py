"""Core logic: read a Manual Request Register workbook and build the formatted report.

Kept free of Streamlit so it can be tested and reused on its own.
"""
from __future__ import annotations

import io
import json
import re
from dataclasses import dataclass, field
from pathlib import Path
from typing import Any

from openpyxl import Workbook, load_workbook
from openpyxl.comments import Comment
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter

CONFIG_PATH = Path(__file__).with_name("conversion_factors.json")

# Input columns we keep, in output order, with the header spellings we accept.
REQUIRED_COLUMNS: list[tuple[str, tuple[str, ...]]] = [
    ("Day", ("day",)),
    ("Work centre", ("work centre", "work center", "workcentre", "workcenter", "work centre name")),
    ("Item", ("item", "item description", "item name")),
    ("Day Plan Qty", ("day plan qty", "day plan qty cfc", "day plan quantity", "day plan qty (cfc)")),
    ("Contractor", ("contractor", "contractor name")),
    ("ENV", ("env", "envelope", "env pcs", "env in pcs")),
    ("TAG", ("tag", "tag pcs", "tag in pcs")),
]

OUTPUT_HEADERS = [
    "Day",
    "Work centre",
    "Item",
    "Day Plan Qty CFC",
    "Contractor",
    "ENV",
    "TAG",
    "ENV in Kgs",
    "TAG in Kgs",
    "Issued Env in Pcs",
    "Issued Env in kgs",
    "Issued Tag in Pcs",
    "Issued Tag in kgs",
]

# Widths based on the 23-11-2026 reference sheet, widened slightly for the longer headings.
COLUMN_WIDTHS = [10, 20, 46, 10, 14, 11, 11, 9, 9, 10, 10, 10, 10]

HEADER_SCAN_ROWS = 15


class InputError(Exception):
    """The uploaded file can't be processed at all (bad file, no usable sheet, missing columns)."""


@dataclass
class Issue:
    row: int | None  # Excel row number in the uploaded sheet
    level: str  # "error" (value not calculated) or "warning"
    message: str
    item: str = ""


@dataclass
class ReportRow:
    source_row: int
    day: Any
    work_centre: Any
    item: Any
    day_plan_qty: Any
    contractor: Any
    env: Any  # number, None, or original text when not numeric
    tag: Any
    env_factor: float | None = None
    tag_factor: float | None = None
    env_kg: float | None = None
    tag_kg: float | None = None
    env_flag: str | None = None  # reason the ENV kg could not be calculated
    tag_flag: str | None = None


@dataclass
class Result:
    sheet_name: str
    rows: list[ReportRow]
    issues: list[Issue] = field(default_factory=list)
    uplift: float = 1.05

    @property
    def total_env_kg(self) -> float:
        return sum(r.env_kg or 0 for r in self.rows)

    @property
    def total_tag_kg(self) -> float:
        return sum(r.tag_kg or 0 for r in self.rows)

    @property
    def error_count(self) -> int:
        return sum(1 for i in self.issues if i.level == "error")


# ---------------------------------------------------------------- config


def load_config(path: Path | str = CONFIG_PATH) -> dict:
    with open(path, encoding="utf-8") as fh:
        config = json.load(fh)
    validate_config(config)
    return config


def validate_config(config: dict) -> None:
    factors = config.get("factors") or {}
    for name, value in factors.items():
        if not isinstance(value, (int, float)) or value <= 0:
            raise ValueError(f"Conversion factor '{name}' must be a positive number (got {value!r}).")
    for rule in config.get("rules") or []:
        re.compile(rule["pattern"], re.IGNORECASE)
        for key in ("env", "tag"):
            ref = rule.get(key)
            if ref is not None and ref not in factors:
                raise ValueError(f"Rule '{rule['pattern']}' refers to unknown factor '{ref}'.")


def match_rule(work_centre: Any, config: dict) -> dict | None:
    name = _clean_text(work_centre)
    if not name:
        return None
    for rule in config.get("rules") or []:
        if re.search(rule["pattern"], name, re.IGNORECASE):
            return rule
    return None


# ---------------------------------------------------------------- parsing helpers


def _clean_text(value: Any) -> str:
    if value is None:
        return ""
    text = str(value).replace("\xa0", " ")
    return re.sub(r"\s+", " ", text).strip()


def _norm_header(value: Any) -> str:
    return re.sub(r"[^a-z0-9()]+", " ", _clean_text(value).lower()).strip()


def parse_quantity(value: Any) -> tuple[float | None, bool]:
    """Return (number, ok). Blank -> (None, True). Handles '  6,00,000' style text."""
    if value is None:
        return None, True
    if isinstance(value, bool):
        return None, False
    if isinstance(value, (int, float)):
        return float(value), True
    text = _clean_text(value).replace(",", "").replace(" ", "")
    if text in ("", "-", "–"):
        return None, True
    try:
        return float(text), True
    except ValueError:
        return None, False


def _as_number(value: float | None) -> int | float | None:
    if value is None:
        return None
    return int(value) if float(value).is_integer() else value


def _find_header(ws) -> tuple[int, dict[str, int]] | None:
    """Locate the header row and map each required column to its 1-based index."""
    max_row = min(ws.max_row or 0, HEADER_SCAN_ROWS)
    best: tuple[int, dict[str, int]] | None = None
    for r in range(1, max_row + 1):
        normalised = {}
        for c in range(1, (ws.max_column or 0) + 1):
            key = _norm_header(ws.cell(r, c).value)
            if key and key not in normalised:
                normalised[key] = c
        found = {}
        for canonical, aliases in REQUIRED_COLUMNS:
            for alias in aliases:
                if alias in normalised:
                    found[canonical] = normalised[alias]
                    break
        if best is None or len(found) > len(best[1]):
            best = (r, found)
        if len(found) == len(REQUIRED_COLUMNS):
            return r, found
    return best


def candidate_sheets(wb) -> list[str]:
    """Sheets whose header row contains every required column."""
    out = []
    for ws in wb.worksheets:
        header = _find_header(ws)
        if header and len(header[1]) == len(REQUIRED_COLUMNS):
            out.append(ws.title)
    return out


def open_workbook(data: bytes):
    try:
        return load_workbook(io.BytesIO(data), data_only=True, read_only=False)
    except Exception as exc:  # zipfile.BadZipFile, InvalidFileException, KeyError ...
        raise InputError(
            "This file could not be read as an Excel workbook. Please upload the original "
            "Manual Request Register in .xlsx format."
        ) from exc


# ---------------------------------------------------------------- processing


def process(data: bytes, config: dict | None = None, sheet_name: str | None = None) -> Result:
    config = config or load_config()
    validate_config(config)
    wb = open_workbook(data)

    if sheet_name is None:
        sheets = candidate_sheets(wb)
        sheet_name = sheets[0] if sheets else wb.worksheets[0].title
    if sheet_name not in wb.sheetnames:
        raise InputError(f"Sheet '{sheet_name}' was not found in the workbook.")
    ws = wb[sheet_name]

    header = _find_header(ws)
    found = header[1] if header else {}
    missing = [name for name, _ in REQUIRED_COLUMNS if name not in found]
    if missing:
        raise InputError(
            f"Sheet '{sheet_name.strip()}' is missing the mandatory column(s): {', '.join(missing)}. "
            f"Expected headings: {', '.join(n for n, _ in REQUIRED_COLUMNS)}."
        )
    header_row = header[0]

    uplift = float(config.get("uplift_percent", 105)) / 100
    factors = config["factors"]
    result = Result(sheet_name=sheet_name, rows=[], uplift=uplift)

    for r in range(header_row + 1, (ws.max_row or header_row) + 1):
        raw = {name: ws.cell(r, col).value for name, col in found.items()}
        if all(_clean_text(v) == "" for v in raw.values()):
            continue  # completely empty row
        item_text = _clean_text(raw["Item"])
        if _is_total_row(raw):
            continue  # an existing TOTAL line in the input; we build our own

        row = ReportRow(
            source_row=r,
            day=_clean_text(raw["Day"]) or None,
            work_centre=_clean_text(raw["Work centre"]) or None,
            item=_clean_text(raw["Item"]) or None,
            day_plan_qty=None,
            contractor=_clean_text(raw["Contractor"]) or None,
            env=None,
            tag=None,
        )

        if not item_text:
            result.issues.append(Issue(r, "warning", "Item name is blank.", item_text))
        if not row.work_centre:
            result.issues.append(Issue(r, "warning", "Work centre is blank.", item_text))

        qty, ok = parse_quantity(raw["Day Plan Qty"])
        if ok:
            row.day_plan_qty = _as_number(qty)
        else:
            row.day_plan_qty = _clean_text(raw["Day Plan Qty"])
            result.issues.append(
                Issue(r, "warning", f"Day Plan Qty '{row.day_plan_qty}' is not a number; copied as text.", item_text)
            )

        rule = match_rule(row.work_centre, config)
        for kind in ("env", "tag"):
            label = kind.upper()
            source = raw[label]
            value, ok = parse_quantity(source)
            if not ok:
                setattr(row, kind, _clean_text(source))
                setattr(row, f"{kind}_flag", f"{label} quantity '{_clean_text(source)}' is not a number")
                result.issues.append(
                    Issue(r, "error", f"{label} quantity '{_clean_text(source)}' is not a number; {label} in Kgs left blank.", item_text)
                )
                continue
            setattr(row, kind, _as_number(value))
            if value is None:
                continue  # nothing to convert; row is still kept
            if value < 0:
                result.issues.append(Issue(r, "warning", f"{label} quantity is negative ({_as_number(value)}).", item_text))

            factor_name = rule.get(kind) if rule else None
            if factor_name is None:
                if rule is None:
                    reason = f"Work centre '{row.work_centre or '(blank)'}' is not recognised"
                else:
                    reason = f"No {label} conversion factor is configured for work centre '{row.work_centre}'"
                setattr(row, f"{kind}_flag", reason)
                result.issues.append(Issue(r, "error", f"{reason}; {label} in Kgs left blank.", item_text))
                continue
            factor = float(factors[factor_name])
            setattr(row, f"{kind}_factor", factor)
            setattr(row, f"{kind}_kg", value / factor * uplift)

        result.rows.append(row)

    if not result.rows:
        raise InputError(f"Sheet '{sheet_name.strip()}' has the required headings but no data rows.")
    return result


def _is_total_row(raw: dict) -> bool:
    texts = [_clean_text(v).upper() for v in raw.values() if isinstance(v, str)]
    return any(t in ("TOTAL", "GRAND TOTAL") for t in texts) and not _clean_text(raw["Item"])


# ---------------------------------------------------------------- output workbook

THIN = Side(style="thin", color="000000")
BORDER = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)
HEADER_FILL = PatternFill("solid", fgColor="D9E1F2")
FLAG_FILL = PatternFill("solid", fgColor="FFC7CE")
PCS_FORMAT = '_ * #,##0_ ;_ * \\-#,##0_ ;_ * "-"??_ ;_ @_ '
KG_FORMAT = "#,##0"
QTY_FORMAT = "#,##0"


def _fmt_factor(factor: float) -> str:
    return str(int(factor)) if float(factor).is_integer() else repr(factor)


def _fmt_uplift(uplift: float) -> str:
    pct = uplift * 100
    return f"{int(round(pct))}%" if abs(pct - round(pct)) < 1e-9 else f"{pct}%"


def build_workbook(result: Result, sheet_title: str | None = None) -> bytes:
    wb = Workbook()
    ws = wb.active
    ws.title = _safe_sheet_title(sheet_title or result.sheet_name)

    base = Font(name="Calibri", size=9)
    kg_font = Font(name="Calibri", size=10)
    header_font = Font(name="Calibri", size=10, bold=True)
    center = Alignment(horizontal="center", vertical="center")
    left = Alignment(horizontal="left", vertical="center")

    for c, title in enumerate(OUTPUT_HEADERS, start=1):
        cell = ws.cell(1, c, title)
        cell.font = header_font
        cell.fill = HEADER_FILL
        cell.border = BORDER
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
    ws.row_dimensions[1].height = 30

    uplift_txt = _fmt_uplift(result.uplift)
    first = 2
    for i, row in enumerate(result.rows):
        r = first + i
        values = [row.day, row.work_centre, row.item, row.day_plan_qty, row.contractor, row.env, row.tag]
        for c, v in enumerate(values, start=1):
            ws.cell(r, c, v)

        for kind, qty_col, kg_col in (("env", "F", 8), ("tag", "G", 9)):
            qty = getattr(row, kind)
            factor = getattr(row, f"{kind}_factor")
            flag = getattr(row, f"{kind}_flag")
            cell = ws.cell(r, kg_col)
            if factor is not None and isinstance(qty, (int, float)):
                cell.value = f"=({qty_col}{r}/{_fmt_factor(factor)})*{uplift_txt}"
            elif flag:
                cell.fill = FLAG_FILL
                cell.comment = Comment(flag + ". Please check the conversion factor.", "Report Generator")
                if not isinstance(qty, (int, float)) and qty not in (None, ""):
                    ws.cell(r, 6 if kind == "env" else 7).fill = FLAG_FILL

        # Columns J:M (issued quantities) are deliberately left empty.
        for c in range(1, len(OUTPUT_HEADERS) + 1):
            cell = ws.cell(r, c)
            cell.border = BORDER
            cell.font = kg_font if c in (8, 9) else base
            if c in (2, 3):
                cell.alignment = left
            elif c in (8, 9) or c >= 10:
                cell.alignment = Alignment(vertical="center")
            else:
                cell.alignment = center
        ws.cell(r, 4).number_format = QTY_FORMAT
        ws.cell(r, 6).number_format = PCS_FORMAT
        ws.cell(r, 7).number_format = PCS_FORMAT
        ws.cell(r, 8).number_format = KG_FORMAT
        ws.cell(r, 9).number_format = KG_FORMAT
        ws.row_dimensions[r].height = 22

    last = first + len(result.rows) - 1
    total = last + 1
    bold = Font(name="Calibri", size=11, bold=True)
    ws.merge_cells(start_row=total, start_column=4, end_row=total, end_column=7)
    ws.cell(total, 4, "TOTAL").font = bold
    ws.cell(total, 4).alignment = center
    ws.cell(total, 8, f"=SUM(H{first}:H{last})")
    ws.cell(total, 9, f"=SUM(I{first}:I{last})")
    ws.merge_cells(start_row=total, start_column=10, end_row=total, end_column=13)
    ws.cell(total, 10, f"=H{total}+I{total}")
    for c in range(4, 14):
        cell = ws.cell(total, c)
        cell.border = BORDER
        if c >= 8:
            cell.font = bold
            cell.number_format = KG_FORMAT
    ws.cell(total, 10).alignment = center
    ws.row_dimensions[total].height = 22

    for c, width in enumerate(COLUMN_WIDTHS, start=1):
        ws.column_dimensions[get_column_letter(c)].width = width
    ws.freeze_panes = "A2"
    ws.page_setup.orientation = "landscape"
    ws.page_setup.fitToWidth = 1
    ws.page_setup.fitToHeight = 0
    ws.sheet_properties.pageSetUpPr.fitToPage = True
    ws.print_title_rows = "1:1"

    buf = io.BytesIO()
    wb.save(buf)
    return buf.getvalue()


def _safe_sheet_title(title: str) -> str:
    cleaned = re.sub(r"[\[\]:*?/\\]", "-", (title or "Report").strip()) or "Report"
    return cleaned[:31]
