"""Tests for numbers_tool.py, run with: .venv/Scripts/python py/test_numbers_tool.py"""

import datetime as dt
import os
import sys
import tempfile

from numbers_parser import Document
from numbers_parser.formula import FUNCTION_NAME_TO_ID
from numbers_parser.generated import TSCEArchives_pb2 as TSCE
from openpyxl import load_workbook

sys.path.insert(0, os.path.dirname(__file__))
import numbers_tool as nt  # noqa: E402
import pivots  # noqa: E402

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
    assert t('"Row "&CHAR(64+COLUMN())&ROW()') == '="Row "&CHAR(64+COLUMN())&ROW()'
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
    assert result["formulas"] == 41 and result["formulasAsValues"] == 0
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


def test_theme_styles_follow_table_regions():
    # Cells styled only by the table theme must take the header/body/header-column style of
    # where they are, not whichever theme-styled cell happened to be read first.
    sheet = nt.preview(fixture("test-pivot"))["sheets"][0]
    style = lambda r, c: sheet["styles"][sheet["rows"][r][c][1]] if isinstance(sheet["rows"][r][c], list) else {}  # noqa: E731
    assert style(0, 1).get("bg") == "#BEC0BF"  # header row
    assert style(1, 0).get("bg") == "#DCDCDC"  # header column
    assert "bg" not in style(1, 1)  # body


def test_text_is_never_invisible():
    # Light text only on a dark background, in every fixture.
    for name in [f[:-8] for f in os.listdir(FIXTURES) if f.endswith(".numbers") and f != "encrypted.numbers"]:
        for sheet in nt.preview(fixture(name))["sheets"]:
            for s in sheet["styles"]:
                if "color" in s and nt.is_light(s["color"]):
                    assert not nt.is_light(s.get("bg", "#FFFFFF")), (name, sheet["name"], s)


def test_pivot_values_match_numbers():
    # This file was saved by a Numbers version that also stored the pivot results, so the
    # recalculation can be checked against Numbers' own numbers.
    doc = Document(fixture("test-pivot"))
    checked = 0
    for sheet in doc.sheets:
        for table in sheet.tables:
            pivot = pivots.pivot_for(table, recompute_stored=True)
            if pivot is None:
                continue
            assert pivot.missing == 0
            for (r, c), (value, _) in pivot.values.items():
                assert abs(table.cell(r, c).value - value) < 1e-6, (sheet.name, r, c)
                checked += 1
    assert checked == 28


def test_pivot_tables_convert():
    result, wb = convert("test-pivot")
    assert result["unsupported"]["pivots"] == 0
    formulas = [c.value for ws in wb.worksheets for row in ws.iter_rows() for c in row if isinstance(c.value, str) and c.value.startswith("=")]
    assert "=" not in formulas  # undecodable pivot formulas must not become empty formulas
    assert wb["Sales - Sales Pivot"]["B3"].value == 18400


def rpn(*tokens):
    """Build a grouping-formula node list: numbers, strings, 'x', 'FUNC/n', operators, '{' and '}' for thunks."""
    nodes = []
    for tok in tokens:
        n = TSCE.ASTNodeArrayArchive.ASTNodeArchive()
        if isinstance(tok, (int, float)):
            n.AST_node_type, n.AST_number_node_number = pivots.NODE.Value("NUMBER_NODE"), tok
        elif tok == "x":
            n.AST_node_type, n.AST_symbol = pivots.NODE.Value("VAR_NODE"), 0
        elif tok in ("{", "}"):
            n.AST_node_type = pivots.NODE.Value("BEGIN_THUNK_NODE" if tok == "{" else "END_THUNK_NODE")
        elif "/" in tok and tok.split("/")[0].isupper():
            name, argc = tok.split("/")
            n.AST_node_type = pivots.NODE.Value("FUNCTION_NODE")
            n.AST_function_node_index = FUNCTION_NAME_TO_ID[name]
            n.AST_function_node_numArgs = int(argc)
        elif tok in ("+", "-", "*", "/", "="):
            names = {"+": "ADDITION", "-": "SUBTRACTION", "*": "MULTIPLICATION", "/": "DIVISION", "=": "EQUAL_TO"}
            n.AST_node_type = pivots.NODE.Value(names[tok] + "_NODE")
        else:
            n.AST_node_type, n.AST_string_node_string = pivots.NODE.Value("STRING_NODE"), tok
        nodes.append(n)
    return nodes


def test_pivot_grouping_formulas():
    d = dt.datetime(2021, 3, 17, 14, 30)
    # Numbers' own "by month" rule: IF(AND(SUM(x)=x, NOT(ISBLANK(x))), DATE(2020, MONTH(x), 1), 1/0)
    month = rpn("x", "SUM/1", "x", "=", "x", "ISBLANK/1", "NOT/1", "AND/2",
                "{", 2020, "x", "MONTH/1", 1, "DATE/3", "}", "{", 1, 0, "/", "}", "IF/3")
    assert pivots.evaluate(month, d) == dt.datetime(2020, 3, 1)
    for bad in (None, "text"):
        try:
            pivots.evaluate(month, bad)
        except pivots.GroupError:
            pass
        else:
            raise AssertionError(bad)
    quarter = rpn("Q", "x", "MONTH/1", 2, "+", 3, "/", "INT/1", "CONCATENATE/2")
    assert pivots.evaluate(quarter, d) == "Q1"
    assert pivots.evaluate(rpn("x", "YEAR/1"), d) == 2021
    assert pivots.evaluate(rpn("x", "WEEKDAY/1"), d) == 4  # Wednesday, Sunday = 1
    assert pivots.evaluate(rpn("x", 2, "WEEKDAY/2"), d) == 3  # Monday = 1
    assert pivots.evaluate(rpn("x", "WEEKNUM/1"), dt.datetime(2021, 1, 2)) == 1
    assert pivots.evaluate(rpn("x", "WEEKNUM/1"), dt.datetime(2021, 1, 3)) == 2  # first Sunday starts week 2
    assert pivots.evaluate(rpn("x", 10, "FLOOR/2"), 37) == 30
    try:
        pivots.evaluate(rpn("x", "STOCK/1"), d)
    except pivots.Unsupported:
        pass
    else:
        raise AssertionError("unknown functions must not be guessed")


def test_pivot_aggregates():
    values = [3, None, 5.5, "n/a", 1.5]
    assert pivots.aggregate("sum", values) == 10
    assert pivots.aggregate("average", values) == 10 / 3
    assert pivots.aggregate("count", values) == 4
    assert pivots.aggregate("min", values) == 1.5 and pivots.aggregate("max", values) == 5.5
    assert pivots.aggregate("sum", [None, "x"]) is None


def test_format_number():
    assert nt.format_number(1234.5, '"$"#,##0') == "$1,235"
    assert nt.format_number(-5, '"$"#,##0.00') == "-$5.00"
    assert nt.format_number(0.256, "0.0%") == "25.6%"
    assert nt.format_number(476.0, None) == "476"


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
