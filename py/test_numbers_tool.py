"""Tests for numbers_tool.py, run with: .venv/Scripts/python py/test_numbers_tool.py"""

import datetime as dt
import os
import sys
import tempfile
import zipfile
from types import SimpleNamespace

from numbers_parser import Document
from numbers_parser.constants import DurationStyle, DurationUnits, FormatType
from numbers_parser.formula import FUNCTION_NAME_TO_ID
from numbers_parser.generated import TSCEArchives_pb2 as TSCE
from openpyxl import load_workbook

sys.path.insert(0, os.path.dirname(__file__))
import charts  # noqa: E402
import layout  # noqa: E402
import numbers_tool as nt  # noqa: E402
import pivots  # noqa: E402

FIXTURES = os.path.join(os.path.dirname(__file__), "..", "test", "fixtures")
OUT = tempfile.mkdtemp(prefix="numbers-tool-")


def fixture(name):
    return os.path.join(FIXTURES, name + ".numbers")


def convert(name, options=None):
    out = os.path.join(OUT, name + (f" ({options['default']})" if options else "") + ".xlsx")
    result = nt.convert(fixture(name), out, layout_options=options)
    return result, load_workbook(out)


def placed(name, sheet, options=None, table=0):
    """Where a table sits on a converted worksheet: its first, n-th, or the one with this name."""
    ws = next(ws for ws in nt.load(fixture(name)).layout(options).sheets if ws.name == sheet)
    return ws.tables[table] if isinstance(table, int) else next(p for p in ws.tables if p.table.name == table)


def at(name, ws, ref, options=None, table=0):
    """The cell at table position `ref` (e.g. "A7") on a converted sheet, allowing for the title row
    above tables whose name Numbers shows, and for where the table sits on the sheet."""
    return ws[shifted(name, ws.title, ref, options, table)]


def shifted(name, sheet, ref, options=None, table=0):
    p = placed(name, sheet, options, table)
    return nt.shift_rows(ref, p.row, p.col)


def grid_cell(sheet, p, r, c):
    """A table's cell (r, c) in a previewed worksheet's grid."""
    return sheet["rows"][p.row + r][p.col + c]


# Formulas naming a table "Prices" on sheet "Main", and a table "Totals" alone on its sheet.
FakeLayout = SimpleNamespace(
    by_table={("Main", "Prices"): 1, (None, "Prices"): 1, (None, "Totals"): 2}, by_sheet={"Totals": 2},
    place={1: layout.Placed(None, "Main - Prices", 0, 0, None), 2: layout.Placed(None, "Totals", 0, 0, None)})


def test_date_patterns():
    cases = {
        "MMM d, y": "mmm d, yyyy",
        "EEEE, MMMM d, y h:mm a": "dddd, mmmm d, yyyy h:mm AM/PM",
        "yyyy-MM-dd": "yyyy-mm-dd",
        "HH:mm:ss.SSS": "hh:mm:ss.000",
        "d 'de' MMMM": 'd "de" mmmm',
        "'Day' D 'of' yyyy": "m/d/yyyy",  # day-of-year has no Excel equivalent
        "K:mm": "m/d/yyyy h:mm",
        "hhmmssSSS a": "m/d/yyyy h:mm",  # Excel only accepts fractional seconds after "s."
        "": "m/d/yyyy",
    }
    for pattern, expected in cases.items():
        assert nt.icu_to_excel_date(pattern) == expected, (pattern, nt.icu_to_excel_date(pattern))


def test_formula_translation():
    t = lambda f: nt.translate_formula(f, FakeLayout, "Main")  # noqa: E731
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
    assert t("COUNTA(fruit)") is None  # header name nothing defines
    assert t("SUM(Unknown::A1)") is None
    assert t("STOCK(\"AAPL\",1)") is None
    assert t("GCD(5,(A1,A2))") is None  # list syntax
    assert t("DATEDIF(\"15/05/2001\",\"15/07/2002\",\"D\")") is None  # region-dependent text dates


def test_header_names_and_titles_in_formulas():
    # In Numbers a header name means the column's body cells inside a function call
    # (COUNTA('Column X & Y') counts 7 cells) and the same row's cell elsewhere. The table also
    # moves down a row in Excel under its title, which references and ROW() must allow for.
    book = nt.load(fixture("test-styles"))
    sheet_name, table = book.tables[1]
    lay = book.layout()
    assert lay.place[table._table_id].row == 1
    cell = next(c for row in nt.table_rows(table) for c in row if c.is_formula)
    assert cell.formula == '"Row Header "&CHAR(64+COLUMN())&ROW()'
    assert nt.translate_formula(cell.formula, lay, sheet_name, cell) == '="Row Header "&CHAR(64+COLUMN())&(ROW()-1)'

    _, wb = convert("create-formulas")
    ws = wb["Main Sheet - Formula Tests"]
    assert at("create-formulas", ws, "A7").value == "=SUM(" + shifted("create-formulas", ws.title, "A2:B2") + ")"


def test_approximate_lookups_need_sorted_data():
    # Excel's approximate MATCH/LOOKUP/VLOOKUP assume sorted data; Numbers doesn't. Unsorted -> value.
    book = SimpleNamespace(table_by_id={})
    rows = {"sorted": [1, 2, 5, 9, None], "unsorted": ["Pears", "Apples", "Leeks"]}

    def check(name, mode, values):
        table = SimpleNamespace(num_rows=len(values) + 1, num_cols=1, _table_id=7,
                                _model=SimpleNamespace(objects={7: SimpleNamespace(
                                    number_of_header_rows=1, number_of_footer_rows=0, number_of_header_columns=0,
                                    HasField=lambda f: False)}))
        book.table_by_id[7] = table
        original = nt.table_rows
        nt.table_rows = lambda t: [[SimpleNamespace(value="Header")]] + [[SimpleNamespace(value=v)] for v in values]
        try:
            args = [[("str", '"x"')], [("exact", "⟦7|||0||0⟧")]] + ([[mode]] if mode else [])  # all of column A
            return nt.lookup_is_safe(name, args, book)
        finally:
            nt.table_rows = original

    assert check("MATCH", None, rows["sorted"])  # approximate on sorted data, trailing blank ignored
    assert not check("MATCH", None, rows["unsorted"])
    assert check("MATCH", ("num", "0"), rows["unsorted"])  # exact match is always fine
    assert not check("LOOKUP", None, rows["unsorted"])


def test_styles_are_kept():
    result, wb = convert("test-styles")
    assert result["sheets"] == 4
    assert result["formulas"] == 41 and result["formulasAsValues"] == 0
    ws = wb["Styles"]
    title = at("test-styles", ws, "A1")
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
    assert at("create-formulas", ws, "A7").value == "=SUM(" + shifted("create-formulas", ws.title, "A2:B2") + ")"


def test_preview():
    data = nt.preview(fixture("test-styles"))
    sheet = data["sheets"][0]
    assert sheet["name"] == "Styles"
    assert sheet["colWidths"] and sheet["rowHeights"]
    first = grid_cell(sheet, placed("test-styles", "Styles"), 0, 0)
    text, style = first
    assert text.startswith('name="Title"')
    assert sheet["styles"][style]["b"] == 1 and sheet["styles"][style]["size"] == 24


def test_theme_styles_follow_table_regions():
    # Cells styled only by the table theme must take the header/body/header-column style of
    # where they are, not whichever theme-styled cell happened to be read first.
    sheet = nt.preview(fixture("test-pivot"))["sheets"][0]
    p = placed("test-pivot", sheet["name"])
    cell = lambda r, c: grid_cell(sheet, p, r, c)  # noqa: E731
    style = lambda r, c: sheet["styles"][cell(r, c)[1]] if isinstance(cell(r, c), list) else {}  # noqa: E731
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
    assert at("test-pivot", wb["Sales - Sales Pivot"], "B3").value == 18400


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


def test_duration_formats():
    fmt = lambda style, largest, smallest, auto=False: SimpleNamespace(  # noqa: E731
        duration_style=style, duration_unit_largest=largest, duration_unit_smallest=smallest,
        use_automatic_duration_units=auto)
    S, L, C = DurationStyle.SHORT, DurationStyle.LONG, DurationStyle.COMPACT
    U = DurationUnits
    two_days = 2 * 86400 + 4 * 3600 + 20 * 60 + 2
    assert nt.duration_format(fmt(S, U.DAY, U.SECOND), two_days) == 'd"d" h"h" m"m" s"s"'
    assert nt.duration_format(fmt(C, U.HOUR, U.SECOND), two_days) == "[h]:mm:ss"
    assert nt.duration_format(fmt(L, U.HOUR, U.MINUTE), 3600 + 60) == '[h]" hour" m" minute"'
    assert nt.duration_format(fmt(S, U.DAY, U.DAY), 60 * 86400) == '0"d"'  # any number of days
    assert nt.duration_format(fmt(S, U.DAY, U.HOUR), 40 * 86400) == '[h]"h"'  # past 31 days: hours
    assert nt.duration_format(nt.DEFAULT_DURATION_FORMAT, 10 * 86400) == '0"d"'


def test_number_formats():
    fmt = lambda **kw: SimpleNamespace(**{"decimal_places": 253, "show_thousands_separator": False,  # noqa: E731
                                          "negative_style": 0, "currency_code": "USD", "use_accounting_style": False,
                                          "fraction_accuracy": 0, **kw})
    assert nt.builtin_number_format(fmt(format_type=FormatType.PERCENT), 6.43) == "0%"
    assert nt.builtin_number_format(fmt(format_type=FormatType.PERCENT), 0.0533) == "0.00%"
    assert nt.builtin_number_format(fmt(format_type=FormatType.DECIMAL, show_thousands_separator=True), 1234.56) == "#,##0.00"
    assert nt.builtin_number_format(fmt(format_type=FormatType.FRACTION, fraction_accuracy=0xFFFFFFFE), 12.34) == "# ??/??"
    assert nt.builtin_number_format(fmt(format_type=FormatType.CURRENCY, decimal_places=2, use_accounting_style=True), 5) == '"$"* #,##0.00'


def test_format_number():
    assert nt.format_number(1234.5, '"$"#,##0') == "$1,235"
    assert nt.format_number(-5, '"$"#,##0.00') == "-$5.00"
    assert nt.format_number(0.256, "0.0%") == "25.6%"
    assert nt.format_number(476.0, None) == "476"


def chart_on(ws, kind=None):
    """The charts on a converted sheet, left to right, optionally only those of one kind."""
    found = sorted(ws._charts, key=lambda c: (c.anchor._from.col, c.anchor._from.colOff))
    return [c for c in found if kind is None or c.tagname == kind]


def test_charts_convert():
    # issue-9 has a chart of each kind Numbers offers. Excel has no bubble charts, so that one is reported.
    result, wb = convert("issue-9")
    assert result["charts"] == 10 and result["unsupported"]["charts"] == 1
    # Charts read the table cells, so they stay live: here B2:B6 of the table, below its title row.
    ws = wb["Column and Bar Charts"]
    column, bar = chart_on(ws, "barChart")
    top = placed("issue-9", ws.title).row
    assert bar.barDir == "bar" and column.barDir == "col"
    assert bar.series[0].val.numRef.f == f"'Column and Bar Charts'!$B${2 + top}:$B${6 + top}"
    assert bar.series[0].cat.strRef.f == f"'Column and Bar Charts'!$A${2 + top}:$A${6 + top}"
    assert [ws[f"B{r + top}"].value for r in range(2, 7)] == [11, 15, 9, 14, 21]
    assert bar.x_axis.scaling.orientation == "maxMin"  # Numbers lists bars from the top
    # One data set, one bar per series: drawn as one series with a colour per bar, named underneath.
    assert len(column.series) == 1 and len(column.series[0].dPt) == 5
    assert column.series[0].dLbls.showVal  # Numbers shows the values on the bars

    pie, = chart_on(wb["Pie and Donut Charts"], "pieChart")
    donut, = chart_on(wb["Pie and Donut Charts"], "doughnutChart")
    assert len(pie.series) == 1 and len(pie.series[0].dPt) == 5 and pie.series[0].dLbls.showPercent
    assert donut.holeSize == 50

    combined, = chart_on(wb["2-Axis Chart"])
    assert {c.tagname for c in combined._charts} == {"barChart", "lineChart"}
    scatter, = chart_on(wb["Scatter Chart"])
    assert scatter.series[0].xVal.numRef.f.startswith("'Scatter Chart'!$A$")

    # A chart above its table gets rows of its own height above it; the table and every
    # reference to it move down to match.
    ws = wb["Interactive Chart"]
    interactive, = chart_on(ws)
    p = placed("issue-9", ws.title)
    assert p.title_row == 1 and p.row == 2  # a row for the chart, then the title, then the table
    assert interactive.anchor._from.row == 0
    assert at("issue-9", ws, "A2").value == "Region 1"
    assert interactive.series[0].val.numRef.f == f"'Interactive Chart'!$B${2 + p.row}:$E${2 + p.row}"

    sheets = {s["name"]: s for s in nt.preview(fixture("issue-9"))["sheets"]}
    assert [c["kind"] for c in sheets["Column and Bar Charts"]["charts"]] == ["column", "bar"]
    assert sheets["Interactive Chart"]["rowHeights"][0] > 100  # the row making room for the chart


def test_chart_links():
    # This chart's data was stored as a range of row IDs plus a reference to cells since deleted,
    # which Numbers ignores.
    result, wb = convert("issue-66-collab")
    assert result["charts"] == 1 and result["unsupported"]["charts"] == 0
    chart, = chart_on(wb["Credit"])
    assert chart.grouping == "stacked"
    assert [s.val.numRef.f for s in chart.series] == ["Credit!$B$10:$B$15", "Credit!$C$10:$C$15", "Credit!$D$10:$D$15"]

    # Cells in a run become a range; scattered cells on one sheet a union of ranges, as in Apple's
    # export; cells on several sheets, or a chart's own copy of values, go on a hidden sheet.
    columns = []

    class Data:
        def column(self, values, formats=None):
            columns.append(values)
            return ["Chart Data", 0, len(columns) - 1, len(values) - 1, len(columns) - 1]

    place = lambda cell: (cell[0], cell[1], cell[2])  # noqa: E731  (sheet, row, col) stand in for table cells
    info = lambda cell: (7, None)  # noqa: E731
    assert charts.reference([("S", 1, 1), ("S", 2, 1), ("S", 3, 1)], [], Data(), place, info) == ["S", 1, 1, 3, 1]
    assert charts.reference([("S", 1, 1), ("S", 1, 3), ("S", 1, 5)], [], Data(), place, info) == "=(S!$B$2,S!$D$2,S!$F$2)"
    assert charts.reference([("S", 1, 1), ("T", 1, 1)], [], Data(), place, info) == ["Chart Data", 0, 0, 1, 0]
    assert columns[0] == [("=S!$B$2", 7), ("=T!$B$2", 7)]
    assert charts.reference(None, [1.5, 2.5], Data(), place, info) == ["Chart Data", 0, 1, 1, 1]


def test_layout_arrangement():
    # Laying out tables as they sit on a Numbers sheet, along one axis: tables one below the other
    # get their own rows, with a row the size of the space between; tables beside each other share
    # rows, lining up with the nearest row edge.
    starts, gaps = layout.arrange([(0, [20] * 5), (178, [20] * 4)], 0, 409)
    assert starts == [0, 6] and gaps == {5: 78}
    starts, gaps = layout.arrange([(0, [20] * 5), (0, [20] * 3)], 0, 409)
    assert starts == [0, 0] and gaps == {}
    starts, _ = layout.arrange([(0, [20] * 5), (41, [20] * 3)], 0, 409)
    assert starts == [0, 2]
    # Tables sharing columns can't share rows: the second moves below the first, after a gap row.
    starts, gaps = layout.arrange([(0, [20] * 5), (50, [20] * 3)], 0, 409, lambda i, j: True)
    assert starts == [0, 6] and gaps == {5: layout.GAP}
    # A big space is split into rows no taller than Excel allows.
    starts, gaps = layout.arrange([(500, [20])], 0, 409)
    assert starts == [2] and gaps == {0: 250, 1: 250}


def test_one_sheet_layout():
    # Laid out as in Numbers, a Numbers sheet's tables share one worksheet, each where it sits on
    # the Numbers sheet, and formulas reading another table point at it there.
    options = {"default": "sheet"}
    result, wb = convert("create-formulas", options)
    assert result["sheets"] == 2 and wb.sheetnames == ["Main Sheet", "Powers Sheet"]
    ws = wb["Main Sheet"]
    plan = nt.load(fixture("create-formulas")).layout(options).sheets[0]
    tables = {p.table.name: p for p in plan.tables}
    tests, food, animals = tables["Formula Tests"], tables["Food Table"], tables["Animal Table"]
    assert food.row == tests.row and food.col > tests.col  # side by side, sharing rows
    assert animals.col == food.col and animals.row > food.row  # one below the other, sharing columns
    formula = at("create-formulas", ws, "A7", options, "Formula Tests").value
    assert formula == "=SUM(" + shifted("create-formulas", ws.title, "A2:B2", options, "Formula Tests") + ")"
    textjoin = next(c.value for row in ws.iter_rows() for c in row if str(c.value).startswith("=_xlfn.TEXTJOIN("))
    assert textjoin == '=_xlfn.TEXTJOIN(",",FALSE,' + nt.shift_rows("B1:D1", animals.row, animals.col) + ")"
    powers = next(c.value for row in ws.iter_rows() for c in row if str(c.value).startswith("=MEDIAN("))
    assert "'Powers Sheet'!" in powers  # a table on another Numbers sheet is on another worksheet

    sheets = nt.preview(fixture("create-formulas"), layout_options=options)["sheets"]
    assert [(s["name"], s["tableCount"], s["layout"]) for s in sheets] == [("Main Sheet", 9, "sheet"), ("Powers Sheet", 5, "sheet")]
    # Each Numbers sheet can be laid out its own way.
    mixed = {"default": "sheet", "sheets": {"Powers Sheet": "tables"}}
    names = [s["name"] for s in nt.preview(fixture("create-formulas"), layout_options=mixed)["sheets"]]
    assert names[0] == "Main Sheet" and len(names) == 6 and all(n.startswith("Powers Sheet - ") for n in names[1:])


def test_font_names():
    assert nt.font_name("Helvetica Neue") == "Arial" and nt.font_name("HelveticaNeue-Medium") == "Arial"
    assert nt.font_name("AvenirNext-DemiBold") == "Segoe UI" and nt.font_is_bold("AvenirNext-DemiBold")
    assert nt.font_name("Georgia-Bold") == "Georgia" and not nt.font_is_bold("Georgia")


def test_custom_formats_convert():
    convert("test-custom-formats")


def test_password_protected():
    for password, expected in ((None, "password-protected"), ("wrong", "isn't right")):
        try:
            nt.load(fixture("encrypted"), password)
        except nt.UserError as err:
            assert expected in str(err) and err.extra["needsPassword"]
            assert err.extra["hint"] == "s3cr3t"  # shown from the first prompt; this sample's hint is its password
        else:
            raise AssertionError("expected a UserError")
    assert nt.preview(fixture("encrypted"), "s3cr3t")["sheets"]
    convert_with = os.path.join(OUT, "encrypted.xlsx")
    nt.convert(fixture("encrypted"), convert_with, "s3cr3t")
    assert load_workbook(convert_with).worksheets


def test_folder_package():
    # Numbers on a Mac can save a file as a folder; those open directly.
    assert os.path.isdir(fixture("folder-package"))
    result, _ = convert("folder-package")
    assert result["sheets"] >= 1


def test_form_controls():
    _, wb = convert("test-actions")
    ws = wb.worksheets[0]
    rules = {str(dv.sqref): dv for dv in ws.data_validations.dataValidation}
    rule = lambda ref: rules[shifted("test-actions", ws.title, ref)]  # noqa: E731
    assert rule("A5").type == "list" and rule("A5").formula1 == '"Item 1,Item 2,Item 3"'  # pop-up menu
    assert rule("B2").type == "whole" and rule("B2").formula2 == "5"  # star rating 0-5
    assert rule("A3").type == "whole" and rule("A3").formula2 == "100"  # slider
    assert at("test-actions", ws, "A1").value is False and at("test-actions", ws, "B1").value is True  # checkboxes


def test_comments_links_bullets_and_rich_text():
    _, wb = convert("test-comments")
    note = at("test-comments", wb.worksheets[0], "A1").comment
    assert note.text.startswith("Comment #1") and note.author == "Jon Connell"

    _, wb = convert("test-hlinks")
    ws = wb.worksheets[0]
    assert at("test-hlinks", ws, "A1").hyperlink.target == "http://news.bbc.co.uk/"
    assert "http://google.co.uk/" in at("test-hlinks", ws, "A1").comment.text  # one link per cell; the rest are noted
    assert at("test-hlinks", ws, "A3").value == "this is a list:\n• first item\n• second item\n• third item"

    _, wb = convert("test-bullets")
    ws = wb.worksheets[0]
    assert at("test-bullets", ws, "B1").value == "• this\n• is some\n• bulleted\n• text"
    # B2 is read by =LEFT(B2,1); bullets are display-only in Numbers, so its text stays as it was.
    assert at("test-bullets", ws, "B2").value == "this\nis some\nnumbered\ntext"

    out = os.path.join(OUT, "rich.xlsx")
    nt.convert(fixture("test-formats"), out)
    ws = load_workbook(out, rich_text=True).worksheets[0]
    rich = at("test-formats", ws, "C7").value
    runs = [(str(part.text), bool(part.font.b), bool(part.font.i)) for part in rich if hasattr(part, "font")]
    assert ("the", True, False) in runs and ("quick", False, True) in runs

    preview = nt.preview(fixture("test-formats"))["sheets"][0]
    text, _, info = grid_cell(preview, placed("test-formats", preview["name"]), 6, 2)
    assert text == "the quick brown fox jumped" and len(info["runs"]) == 9


def test_hidden_rows_and_columns():
    # No sample file has hidden rows, so mark some hidden in the loaded file (as Numbers stores them).
    path = fixture("test-bgcolour")
    book = nt.load(path)
    table = book.tables[0][1]
    model = table._model
    bds = model.objects[table._table_id].base_data_store
    for bucket, index in ((bds.rowHeaders.buckets[0].identifier, 1), (bds.columnHeaders.identifier, 2)):
        headers = model.objects[bucket].headers
        header = next((h for h in headers if h.index == index), None) or headers.add()
        header.index, header.hidingState = index, 1
    book.layouts.clear()  # laid out before the change
    sheet = nt.preview(path)["sheets"][0]
    p = book.layout().place[table._table_id]
    assert sheet["hiddenRows"] == [p.row + 1] and sheet["hiddenCols"] == [p.col + 2]
    out = os.path.join(OUT, "hidden.xlsx")
    nt.convert(path, out)
    ws = load_workbook(out).worksheets[0]
    assert ws.row_dimensions[p.row + 2].hidden and ws.column_dimensions["C"].hidden
    nt._cache.clear()


def test_formulas_say_why_they_keep_values():
    def why(formula):
        reasons = []
        assert nt.translate_formula(formula, FakeLayout, "Main", why=reasons) is None
        return reasons[0]
    assert why('STOCK("AAPL",1)') == "function:STOCK"
    assert why("IF(A1)") == "arguments:IF"
    assert why("GCD(5,(A1,A2))") == "list"
    assert why('DATEDIF("15/05/2001","15/07/2002","D")') == "date-text"
    assert why('AVERAGEIFS({1,2},{1,0},">0")') == "array"
    assert why("SUM(Unknown::A1)") == "missing"
    assert why("SUM(#REF!)") == "broken"


def test_change_report():
    def report(name):
        return nt.preview(fixture(name))["report"]

    formulas = report("create-formulas")
    (kept,) = [c for c in formulas["changes"] if c["kind"] == "formulas"]
    assert kept["title"] == "1 formula keeps its last value" and "#REF!" in kept["detail"]
    cell = kept["cells"][0]
    sheet = next(s for s in nt.preview(fixture("create-formulas"))["sheets"] if s["name"] == cell["sheet"])
    shown = sheet["rows"][cell["row"]][cell["col"]]  # plain text, or [text, style, ...]
    assert (shown if isinstance(shown, str) else shown[0]) == "#REF!" and cell["ref"] == f"A{cell['row'] + 1}"
    assert "109 formulas that keep calculating in Excel" in formulas["kept"]

    charts_file = report("issue-9")
    kinds = {c["kind"]: c for c in charts_file["changes"]}
    assert kinds["shapes"]["title"].startswith("7 text boxes")  # text boxes were missed before
    assert "Bubble charts" in kinds["charts"]["detail"]
    assert any(n["kind"] == "interactive" for n in charts_file["notes"])

    assert any(c["kind"] == "highlighting" for c in report("issue-66-collab")["changes"])
    custom = {c["kind"]: c for c in report("test-custom-formats")["changes"]}
    assert "DURATION" in custom["formulas"]["detail"] and custom["formats"]["title"].endswith("custom number format")
    assert any(n["kind"] == "pivot-values" for n in report("test-pivot")["notes"])
    fonts = next(n for n in report("test-styles")["notes"] if n["kind"] == "fonts")
    assert "Helvetica Neue → Arial" in fonts["detail"]
    assert not report("test-styles")["changes"]


def test_unreadable_files_show_their_picture():
    bad = os.path.join(OUT, "Damaged.numbers")
    with zipfile.ZipFile(fixture("test-styles")) as zin, zipfile.ZipFile(bad, "w") as zout:
        for item in zin.infolist():
            zout.writestr(item, b"\x00broken" * 40 if item.filename.endswith("Document.iwa") else zin.read(item.filename))
    try:
        nt.load(bad)
    except nt.UserError as err:
        assert err.extra["reason"] == "unreadable" and err.extra["picture"].startswith("data:image/jpeg;base64,")
        assert "Damaged" not in err.extra["details"] and OUT not in err.extra["details"]
    else:
        raise AssertionError("expected a UserError")
    assert nt.technical(OSError(f"can't open {bad}"), bad) == "OSError: can't open <file>"


def test_contents_sheet():
    out = os.path.join(OUT, "contents.xlsx")
    result = nt.convert(fixture("create-formulas"), out, layout_options={"default": "sheet"}, contents=True)
    wb = load_workbook(out)
    assert result["contents"] and wb.sheetnames == ["Contents", "Main Sheet", "Powers Sheet"]
    ws = wb["Contents"]
    links = [row[2] for row in ws.iter_rows(min_row=5) if row[2].hyperlink]
    assert len(links) == 14 and all(link.hyperlink.location.startswith("'") for link in links)
    target = links[0].hyperlink.location  # e.g. 'Main Sheet'!H42: the table's top-left cell
    sheet, cell = target.split("!")
    assert wb[sheet.strip("'")][cell].value is not None
    plain = nt.convert(fixture("create-formulas"), os.path.join(OUT, "no-contents.xlsx"))
    assert not plain["contents"] and "Contents" not in load_workbook(os.path.join(OUT, "no-contents.xlsx")).sheetnames


if __name__ == "__main__":
    tests = [(name, fn) for name, fn in globals().items() if name.startswith("test_")]
    for name, fn in tests:
        fn()
        print("ok", name)
    print(f"All {len(tests)} helper tests passed.")
