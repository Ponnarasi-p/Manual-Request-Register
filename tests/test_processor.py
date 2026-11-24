import io
from pathlib import Path

import pytest
from openpyxl import Workbook, load_workbook

import processor as P

HEADERS = ["Day", "Work centre", "Item", "Day Plan Qty", "CTN", "ENV", "TAG", "BLEND", "PROMO", "Contractor"]


def make_xlsx(rows, headers=HEADERS) -> bytes:
    wb = Workbook()
    ws = wb.active
    ws.append(headers)
    for r in rows:
        ws.append(r)
    buf = io.BytesIO()
    wb.save(buf)
    return buf.getvalue()


def row(wc, env=None, tag=None, qty="\xa0\xa0 1,000", item="ITEM"):
    return ["Monday", wc, item, qty, 10, env, tag, 1, None, "Bipin"]


def test_columns_mapping_and_kg():
    data = make_xlsx([row("C250-A", "\xa0\xa0 3,00,000", 300000), row("MAISA ENV", 230400, 230400)])
    res = P.process(data)
    assert not res.issues
    a, b = res.rows
    assert a.day_plan_qty == 1000 and a.env == 300000
    assert a.env_kg == pytest.approx(300000 / 1740 * 1.05)
    assert a.tag_kg == pytest.approx(300000 / 10780 * 1.05)
    assert b.env_factor == 1410 and b.tag_factor == 9080


@pytest.mark.parametrize(
    "wc,env,tag",
    [
        ("CONSTANTA ENV BT 25", 1525, 10780),
        ("CONSTANTA ENV HF", 1525, 10780),
        ("C250 AF", 1740, 10780),
        ("C250", 1740, 10780),
        ("MD20 ENV GARANT", 1700, 12600),
        ("MD20 4GM TAG", 1700, 12600),
        ("PERFECTA ENV", 1520, 10780),
    ],
)
def test_work_centre_rules(wc, env, tag):
    cfg = P.load_config()
    rule = P.match_rule(wc, cfg)
    assert cfg["factors"][rule["env"]] == env
    assert cfg["factors"][rule["tag"]] == tag


def test_unknown_work_centre_is_flagged_not_calculated():
    res = P.process(make_xlsx([row("FFS POUCH", 1000), row("FFS POUCH")]))
    assert len(res.rows) == 2  # row without quantities still kept
    assert res.rows[0].env_kg is None and res.rows[0].env_flag
    assert [i.row for i in res.issues if i.level == "error"] == [2]


def test_non_numeric_and_empty_rows():
    data = make_xlsx([row("C250", "abc", 100), [None] * 10, row("C250", None, None)])
    res = P.process(data)
    assert len(res.rows) == 2
    assert res.rows[0].env == "abc" and res.rows[0].env_kg is None
    assert res.rows[0].tag_kg is not None
    assert res.issues[0].row == 2 and res.issues[0].level == "error"


def test_missing_column():
    headers = [h for h in HEADERS if h != "Contractor"]
    with pytest.raises(P.InputError, match="Contractor"):
        P.process(make_xlsx([["Monday", "C250", "X", 1, 1, 1, 1, 1, 1]], headers))


def test_invalid_file():
    with pytest.raises(P.InputError):
        P.process(b"not an excel file")


def test_output_workbook_structure():
    res = P.process(make_xlsx([row("C250", 1000, 2000), row("PERFECTA ENV", 500)]))
    ws = load_workbook(io.BytesIO(P.build_workbook(res, "23-11-2026", "23-11-2026"))).active
    assert ws.title == "23-11-2026"
    assert ws["A1"].value == "MANUAL REQUEST REGISTER - 23-11-2026"
    assert "A1:L1" in {str(m) for m in ws.merged_cells.ranges}
    assert [c.value for c in ws[2]] == P.OUTPUT_HEADERS and "Day" not in P.OUTPUT_HEADERS
    assert ws["A3"].value == "C250"
    assert ws["G3"].value == "=(E3/1740)*105%"
    assert ws["H3"].value == "=(F3/10780)*105%"
    assert ws["H4"].value is None
    for r in (3, 4):
        assert all(ws.cell(r, c).value is None for c in range(9, 13))
    assert ws["C5"].value == "TOTAL" and ws["G5"].value == "=SUM(G3:G4)" and ws["I5"].value == "=G5+H5"


SAMPLE = Path(__file__).resolve().parents[2] / "Manual Request Register Oct26.xlsx"


@pytest.mark.skipif(not SAMPLE.exists(), reason="sample workbook not available")
def test_sample_sheet4():
    res = P.process(SAMPLE.read_bytes(), sheet_name="Sheet4")
    assert len(res.rows) == 19 and not res.issues
