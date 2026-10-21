"""Tests for numbers_tool.py, run with: .venv/Scripts/python py/test_numbers_tool.py"""

import os
import sys
import tempfile

from openpyxl import load_workbook

sys.path.insert(0, os.path.dirname(__file__))
import numbers_tool as nt  # noqa: E402

FIXTURES = os.path.join(os.path.dirname(__file__), "..", "test", "fixtures")
OUT = tempfile.mkdtemp(prefix="numbers-tool-")


def fixture(name):
    return os.path.join(FIXTURES, name + ".numbers")


def convert(name):
    out = os.path.join(OUT, name + ".xlsx")
    result = nt.convert(fixture(name), out)
    return result, load_workbook(out)


class FakeBook:
    by_table = {("Main", "Prices"): "Main - Prices", (None, "Prices"): "Main - Prices", (None, "Totals"): "Totals"}
    by_sheet = {"Totals": "Totals"}


def test_date_patterns():
    cases = {
        "MMM d, y": "mmm d, yyyy",
        "EEEE, MMMM d, y h:mm a": "dddd, mmmm d, yyyy h:mm AM/PM",
        "yyyy-MM-dd": "yyyy-mm-dd",
        "HH:mm:ss.SSS": "hh:mm:ss.000",
        "d 'de' MMMM": 'd "de" mmmm',
        "'Day' D 'of' yyyy": "m/d/yyyy",  # day-of-year has no Excel equivalent
        "K:mm": "m/d/yyyy h:mm",
        "": "m/d/yyyy",
    }
    for pattern, expected in cases.items():
        assert nt.icu_to_excel_date(pattern) == expected, (pattern, nt.icu_to_excel_date(pattern))


def test_formula_translation():
    t = lambda f: nt.translate_formula(f, FakeBook, "Main")  # noqa: E731
    assert t("SUM(A1:B2)") == "=SUM(A1:B2)"
    assert t("A1÷B1×2") == "=A1/B1*2"
    assert t("IF(A1≥2,\"big\",\"small\")") == '=IF(A1>=2,"big","small")'
    assert t("SUM(Prices::B2:B9)") == "=SUM('Main - Prices'!B2:B9)"
    assert t("Main::Prices::C3*2") == "='Main - Prices'!C3*2"
    assert t("SUM((A1+B1),C1)") == "=SUM((A1+B1),C1)"
    assert t("TODAY()") == "=TODAY()"
    assert t("IF(A1>0,1)") == "=IF(A1>0,1)"
    # Things Excel would reject or compute differently keep their values instead.
    assert t("IF(A1)") is None  # too few arguments
    assert t("PI(1)") is None  # too many arguments
    assert t("DURATION(1,2,3)") is None  # same name, different meaning in Excel
    assert t("SUM(A:A)") is None  # whole-column reference
    assert t("AVERAGEIFS({1,2},{1,0},\">0\")") is None  # array constant
    assert t("COUNTA(fruit)") is None  # header-name reference
    assert t("SUM(Unknown::A1)") is None
    assert t("STOCK(\"AAPL\",1)") is None
    assert t("GCD(5,(A1,A2))") is None  # list syntax
    assert t("DATEDIF(\"15/05/2001\",\"15/07/2002\",\"D\")") is None  # region-dependent text dates


def test_styles_are_kept():
    result, wb = convert("test-styles")
    assert result["sheets"] == 4
    ws = wb["Styles"]
    title = ws["A1"]
    assert title.value == 'name="Title",bold,font_size=24.0'
    assert title.font.b and title.font.sz == 24
    red = next(c for row in ws.iter_rows() for c in row if c.value and "Heading Red" in str(c.value))
    assert red.font.color.rgb.endswith("EE220C")
    strike = next(c for row in ws.iter_rows() for c in row if c.value and "strikethrough" in str(c.value) and "font" not in str(c.value))
    assert strike.font.strike


def test_fills_and_borders_are_kept():
    _, wb = convert("test-bgcolour")
    fills = {c.fill.fgColor.rgb for ws in wb.worksheets for row in ws.iter_rows() for c in row if c.fill.fill_type == "solid"}
    assert len(fills) > 1
    _, wb = convert("test-border-widths")
    borders = [c for ws in wb.worksheets for row in ws.iter_rows() for c in row if c.border.top.style or c.border.left.style]
    assert borders


def test_formulas_are_kept():
    result, wb = convert("create-formulas")
    assert result["formulas"] > 30
    ws = wb["Main Sheet - Formula Tests"]
    assert ws["A7"].value == "=SUM(A2:B2)"


def test_preview():
    data = nt.preview(fixture("test-styles"))
    sheet = data["sheets"][0]
    assert sheet["name"] == "Styles"
    assert sheet["colWidths"] and sheet["rowHeights"]
    first = sheet["rows"][0][0]
    text, style = first
    assert text.startswith('name="Title"')
    assert sheet["styles"][style]["b"] == 1 and sheet["styles"][style]["size"] == 24


def test_charts_are_reported():
    data = nt.preview(fixture("issue-9"))
    assert data["unsupported"]["charts"] == 11


def test_custom_formats_convert():
    convert("test-custom-formats")


def test_password_protected():
    try:
        nt.load(fixture("encrypted"))
    except nt.UserError as err:
        assert "password" in str(err)
    else:
        raise AssertionError("expected a UserError")


if __name__ == "__main__":
    tests = [(name, fn) for name, fn in globals().items() if name.startswith("test_")]
    for name, fn in tests:
        fn()
        print("ok", name)
    print(f"All {len(tests)} helper tests passed.")
