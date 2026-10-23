"""
Numbers charts -> native Excel charts.

How Numbers stores a chart (a ChartDrawableArchive with a ChartArchive extension):
  * geometry: the chart body's position and size on the sheet canvas, in points (the title,
    legend and axis labels sit around it);
  * chart_type, and a data grid: row_name = series names, column_name = category names, grid_row =
    the values, a snapshot taken when the file was saved;
  * a mediator with one formula per series saying which table cells the series reads: a range,
    a header cell (meaning that column's or row's body cells), a single cell, or a list of cells;
  * style objects whose settings live in "generated" extensions with readable names
    (tschchartinfodefaulttitle, tschchartseriescolumnfill, ...), resolved through a series' own
    style, its theme style and their parents.

read_charts turns each chart into a Chart description, layout.py decides which worksheet each chart
goes on and where, and write_charts builds the Excel charts. Charts read the converted table cells,
so they stay live in Excel, as in Apple's own export. Data Excel can't reference in place (cells on
several sheets, or values with no table behind them) goes on a hidden "Chart Data" sheet, as
formulas pointing at the table cells where there are any.
"""

import math
import re

from numbers_parser.generated import TSCEArchives_pb2 as TSCE
from numbers_parser.generated import TSCHArchives_GEN_pb2 as GEN
from numbers_parser.generated import TSCHArchives_pb2 as TSCH
from numbers_parser.numbers_uuid import NumbersUUID
from xlsxwriter.utility import quote_sheetname, xl_range_abs, xl_rowcol_to_cell

import extras
from layout import overlap

NODE = TSCE.ASTNodeArrayArchive.ASTNodeType
CHART = TSCH.ChartArchive.unity
CHART_TYPES = TSCH.ChartArchive.DESCRIPTOR.fields_by_name["chart_type"].enum_type
SHARED_X = TSCH.ChartArchive.DESCRIPTOR.fields_by_name["scatter_format"].enum_type.values_by_name[
    "scatter_format_shared_x"].number
GENERATED = {name: getattr(GEN, name).current for name in (
    "ChartStyleArchive", "ChartNonStyleArchive", "ChartSeriesStyleArchive", "ChartSeriesNonStyleArchive",
    "ChartAxisStyleArchive", "ChartAxisNonStyleArchive", "LegendStyleArchive")}

# Numbers chart type -> (Excel chart type, subtype). 3D charts become their 2D equivalents, and
# interactive ("multi-data") charts show all their data sets at once.
TYPES = {
    "columnChartType2D": ("column", None), "columnChartType3D": ("column", None),
    "multiDataColumnChartType2D": ("column", None),
    "barChartType2D": ("bar", None), "barChartType3D": ("bar", None), "multiDataBarChartType2D": ("bar", None),
    "stackedColumnChartType2D": ("column", "stacked"), "stackedColumnChartType3D": ("column", "stacked"),
    "stackedBarChartType2D": ("bar", "stacked"), "stackedBarChartType3D": ("bar", "stacked"),
    "lineChartType2D": ("line", None), "lineChartType3D": ("line", None),
    "areaChartType2D": ("area", None), "areaChartType3D": ("area", None),
    "stackedAreaChartType2D": ("area", "stacked"), "stackedAreaChartType3D": ("area", "stacked"),
    "pieChartType2D": ("pie", None), "pieChartType3D": ("pie", None),
    "donutChartType2D": ("doughnut", None), "donutChartType3D": ("doughnut", None),
    "scatterChartType2D": ("scatter", None), "multiDataScatterChartType2D": ("scatter", None),
    "mixedChartType2D": ("mixed", None), "twoAxisChartType2D": ("mixed", None),
    "radarChartType2D": ("radar", None),
}

# Which properties style each kind of series: tschchartseries<part><property>.
PARTS = {
    "column": {"fill": "column", "stroke": "bar", "labels": "bar", "position": "bar", "text": "bar"},
    "bar": {"fill": "bar", "stroke": "bar", "labels": "bar", "position": "bar", "text": "bar"},
    "line": {"stroke": "line", "symbol": "line", "labels": "line", "position": "line", "text": "line"},
    "area": {"fill": "area", "stroke": "area", "symbol": "area", "labels": "area", "position": "area", "text": "area"},
    "pie": {"fill": "pie", "stroke": "pie", "labels": "pie", "text": "pie"},
    "doughnut": {"fill": "pie", "stroke": "pie", "labels": "pie", "text": "donut"},
    "scatter": {"stroke": "scatter", "symbol": "scatter", "labels": "scatter", "position": "scatter", "text": "default"},
    "radar": {"fill": "radararea", "stroke": "radararea", "symbol": "radararea", "labels": "radar",
              "position": "radararea", "text": "radar"},
    "mixed column": {"fill": "mixedcolumn", "stroke": "mixedcolumn", "labels": "mixed", "position": "mixedcolumn",
                     "text": "mixed"},
    "mixed line": {"stroke": "mixedline", "symbol": "mixedline", "labels": "mixed", "position": "mixedline",
                   "text": "mixed"},
    "mixed area": {"fill": "mixedarea", "stroke": "mixedarea", "symbol": "mixedarea", "labels": "mixed",
                   "position": "mixedarea", "text": "mixed"},
}
MIXED_KINDS = {0: "column", 1: "area", 2: "line"}  # tschchartseriesdefaultseriestype
LABEL_POSITIONS = {4: "inside_end", 3: "center", 2: "inside_base", 1: "outside_end"}  # column and bar charts
PT_TO_PX = 4 / 3


class Unsupported(Exception):
    pass


# ---------------------------------------------------------------- styles


def _generated(obj):
    ext = GENERATED.get(type(obj).__name__)
    try:
        return obj.Extensions[ext] if ext is not None and obj.HasExtension(ext) else None
    except Exception:
        return None


def _chain(model, ids):
    for rid in ids:
        seen = 0
        while rid and rid in model.objects and seen < 12:
            obj = model.objects[rid]
            yield obj
            rid = obj.super.parent.identifier if obj.HasField("super") and obj.super.HasField("parent") else 0
            seen += 1


def prop(model, ids, name, default=None):
    """A chart setting, from the most specific style that sets it."""
    return prop_and_owner(model, ids, name, default)[0]


def prop_and_owner(model, ids, name, default=None):
    """A chart setting and the style object it comes from (None when it's the default)."""
    for obj in _chain(model, ids):
        g = _generated(obj)
        if g is not None and name in g.DESCRIPTOR.fields_by_name and g.HasField(name):
            return getattr(g, name), obj
    return default, None


def paint(fill_or_stroke):
    """(#RRGGBB, opacity) of a fill (its first gradient stop) or a visible stroke, or None."""
    if fill_or_stroke is None:
        return None
    fields = fill_or_stroke.DESCRIPTOR.fields_by_name
    if "pattern" in fields:  # a stroke
        if fill_or_stroke.width <= 0 or (fill_or_stroke.HasField("pattern")
                                         and fill_or_stroke.pattern.type == fill_or_stroke.pattern.TSDEmptyPattern):
            return None
    if fill_or_stroke.HasField("color"):
        c = fill_or_stroke.color
    elif "gradient" in fields and fill_or_stroke.HasField("gradient") and fill_or_stroke.gradient.stops:
        c = fill_or_stroke.gradient.stops[0].color
    else:
        return None
    return "#{:02X}{:02X}{:02X}".format(round(c.r * 255), round(c.g * 255), round(c.b * 255)), c.a


def text_style(model, chart, index, tools, default_size=10.0):
    """{font, size, bold, italic, color} of one of the chart's paragraph styles."""
    try:
        style = model.objects[chart.paragraph_styles[int(index)].identifier]
        colour = model.char_property(style, "font_color")
        name = model.char_property(style, "font_name")
        return {"font": tools.font(name) or "Arial", "size": round(model.char_property(style, "font_size"), 1),
                "bold": bool(model.char_property(style, "bold")) or tools.bold(name),
                "italic": bool(model.char_property(style, "italic")),
                "color": "#{:02X}{:02X}{:02X}".format(round(colour.r * 255), round(colour.g * 255), round(colour.b * 255))}
    except Exception:
        return {"font": "Arial", "size": default_size, "bold": False, "italic": False, "color": "#000000"}


def lighten(colour, amount):
    """A colour mixed with white: Numbers' theme colours come round again a little lighter."""
    if not colour or amount <= 0:
        return colour
    rgb = [int(colour[0][i:i + 2], 16) for i in (1, 3, 5)]
    return "#{:02X}{:02X}{:02X}".format(*(round(v + (255 - v) * amount) for v in rgb)), colour[1]


# ---------------------------------------------------------------- data links


def decode_link(model, formula):
    """[(table_id, row1, col1, row2, col2)] blocks of table cells a series formula reads; None for a
    whole row or column."""
    stack = []
    for node in formula.AST_node_array.AST_node:
        kind = NODE.Name(node.AST_node_type)
        if kind in ("CELL_REFERENCE_NODE", "COLON_TRACT_NODE"):
            if not node.HasField("AST_cross_table_reference_extra_info"):
                raise Unsupported("chart reads cells without a table")
            table_id = model.table_uuids_to_id(NumbersUUID(node.AST_cross_table_reference_extra_info.table_id).hex)
            if table_id is None:
                raise Unsupported("chart reads a table that isn't there")
            if kind == "CELL_REFERENCE_NODE":
                r, c = node.AST_row.row, node.AST_column.column
                stack.append([(table_id, r, c, r, c)])
            else:  # a block given by row and column ranges (blank = all of them)
                ref = model.node_to_ref(table_id, 0, 0, node)
                if ref._ref_error:
                    raise Unsupported("broken chart reference")
                stack.append([(table_id, ref.row_start, ref.col_start,
                               ref.row_end if ref.row_end is not None else ref.row_start,
                               ref.col_end if ref.col_end is not None else ref.col_start)])
        elif kind in ("COLON_NODE", "COLON_NODE_WITH_UIDS"):
            b, a = stack.pop(), stack.pop()
            if len(a) != 1 or len(b) != 1 or a[0][0] != b[0][0] or None in a[0] or None in b[0]:
                raise Unsupported("unusual chart range")
            (t, r1, c1, _, _), (_, r2, c2, _, _) = a[0], b[0]
            stack.append([(t, min(r1, r2), min(c1, c2), max(r1, r2), max(c1, c2))])
        elif kind == "REFERENCE_ERROR_WITH_UIDS":
            stack.append([])  # cells since deleted: they no longer add anything to the series
        elif kind == "FUNCTION_NODE":  # the mediator lists a series' references in a function
            count = node.AST_function_node_numArgs
            args = stack[len(stack) - count:] if count else []
            del stack[len(stack) - len(args):]
            stack.append([block for arg in args for block in arg])
        else:
            raise Unsupported(f"chart link uses {kind}")
    if len(stack) != 1 or not stack[0]:
        raise Unsupported("unexpected chart link")
    return stack[0]


def expand(book, block):
    """The cells a block covers, as [(table_id, row, col)]. A header cell on its own stands for its
    column's (or row's) body cells, as in Numbers, and so does a whole column (or row)."""
    table_id, r1, c1, r2, c2 = block
    table = book.table_by_id.get(table_id)
    if table is None:
        raise Unsupported("chart reads a table we didn't convert")
    tm = table._model.objects[table_id]
    if tm.HasField("pivot_owner"):
        raise Unsupported("chart reads a pivot table")  # its rows are rewritten; keep the chart's own copy
    hr, hc, fr = tm.number_of_header_rows, tm.number_of_header_columns, tm.number_of_footer_rows
    if r1 is None or ((r1, c1) == (r2, c2) and r1 < hr):
        r1, r2 = hr, table.num_rows - fr - 1
    if c1 is None or ((r1, c1) == (r2, c2) and c1 < hc):
        c1, c2 = hc, table.num_cols - 1
    if r2 >= table.num_rows or c2 >= table.num_cols or r1 > r2 or c1 > c2:
        raise Unsupported("chart reads cells outside its table")
    return [(table_id, r, c) for r in range(r1, r2 + 1) for c in range(c1, c2 + 1)]


# ---------------------------------------------------------------- reading


class Series:
    def __init__(self, name):
        self.name = name             # text
        self.name_cell = None        # (table_id, row, col) holding the name, when there is one
        self.cells = None            # [(table_id, row, col)] the values come from, or None
        self.values = []             # the chart's own copy of the values
        self.x_cells = None          # scatter charts: where the x values come from
        self.x_values = None
        self.kind = None             # column/line/area in mixed charts, else the chart's kind
        self.colour = None           # (#RRGGBB, opacity) fill
        self.line = None             # (#RRGGBB, opacity, width) line or outline
        self.marker = None           # {size, colour, outline}
        self.joined = True           # scatter charts: whether a line joins the points
        self.point_colours = None    # a fill per point: pie slices, or bars of a folded chart
        self.labels = None           # {"value" or "percentage": True, "position", "font"}
        self.y2 = False


class Chart:
    def __init__(self, drawable_id):
        self.id = drawable_id
        self.kind = None
        self.subtype = None
        self.series = []
        self.categories = []         # text
        self.category_cells = None   # [(table_id, row, col)] when the table holds the categories
        self.title = None
        self.title_style = None
        self.legend = None           # (x, y, w, h) on the canvas, when shown
        self.legend_style = None
        self.body = None             # (x, y, w, h) on the canvas
        self.sheet_id = None
        self.gap = None
        self.overlap = None
        self.hole = None
        self.rotation = None
        self.outside_labels = False
        self.label_width = None      # characters in the widest value label, for sizing the frame
        self.value_format = None     # how values read in the table, for the preview: see value_format()
        self.value_axes = []         # y, then y2: {labels, line, gridlines, minor, title, min, max, font, ...}
        self.category_axis = {}
        self.skip_hidden = True
        self.tables = []             # ids of tables the chart reads
        self.frame = None            # (x, y, w, h) on the canvas of everything it draws: body, title, labels, legend
        self.plot = None             # where its plot goes: its body, or a little less for some pies
        self.size = None             # (width, height) in pixels
        self.flat = False            # a 3D chart, drawn flat in Excel
        self.interactive = False     # an interactive chart, which shows every data set at once in Excel


def read_charts(book, tools):
    """[Chart] for every chart in the document, plus why any couldn't be read: [Numbers chart type
    name, or "unreadable"]. `tools` gives a cell's text as Numbers shows it (text) and its Excel
    number format (number_format), and the Windows font for a Mac font (font, bold)."""
    model = book.tables[0][1]._model
    found, problems = [], []
    for oid, obj in list(model.objects._objects.items()):
        if type(obj).__name__ != "ChartDrawableArchive":
            continue
        try:
            found.append(read_chart(book, model, oid, obj, tools))
        except Unsupported as err:
            problems.append(str(err))
        except Exception:  # a chart stored in a way we don't recognise
            problems.append("unreadable")
    found.sort(key=lambda c: (c.body[1], c.body[0]))
    return found, problems


def read_chart(book, model, oid, drawable, tools):
    c = drawable.Extensions[CHART]
    type_name = CHART_TYPES.values_by_number[c.chart_type].name
    if type_name not in TYPES:
        raise Unsupported(type_name)  # bubble charts: Excel's take their data in a different shape
    chart = Chart(oid)
    chart.kind, chart.subtype = TYPES[type_name]
    chart.flat = type_name.endswith("3D")
    chart.interactive = type_name.startswith("multiData")
    style, nonstyle = [c.chart_style.identifier], [c.chart_non_style.identifier]
    g = drawable.super.geometry
    chart.body = (g.position.x, g.position.y, g.size.width, g.size.height)
    if chart.body[2] <= 0 or chart.body[3] <= 0:
        raise Unsupported("chart has no size")
    chart.sheet_id = sheet_of(model, drawable)

    if prop(model, nonstyle, "tschchartinfodefaultshowtitle", True) and not drawable.super.title_hidden:
        chart.title = prop(model, nonstyle, "tschchartinfodefaulttitle") or None
        chart.title_style = text_style(model, c, prop(model, style, "tschchartinfodefaulttitleparagraphstyleindex", 0),
                                       tools, 12.0)
    if prop(model, nonstyle, "tschchartinfodefaultshowlegend", True) and c.HasField("legend_frame"):
        lf = c.legend_frame
        cx, cy = g.position.x + g.size.width / 2, g.position.y + g.size.height / 2
        chart.legend = (cx + lf.origin.x, cy + lf.origin.y, lf.size.width, lf.size.height)
        index = prop(model, [c.legend_style.identifier], "tschlegendmodeldefaultlabelparagraphstyleindex", 0)
        chart.legend_style = text_style(model, c, index, tools)
    gap = prop(model, style, "tschchartinfodefaultintersetgap")  # percent of a bar's width, as in Excel
    chart.gap = round(gap) if gap is not None else None
    bar_gap = prop(model, style, "tschchartinfodefaultinterbargap")
    chart.overlap = -round(bar_gap) if bar_gap is not None else None
    if chart.kind in ("pie", "doughnut"):
        start = prop(model, nonstyle, "tschchartinfodefaultpiestartangle", 0.0)
        chart.rotation = round(360 - start) % 360
        inner = prop(model, nonstyle, "tschchartinfodefaultinnerradius")
        chart.hole = max(10, min(90, round(inner * 100))) if inner else 50
        # Labels outside the pie, joined to their slices by lines (else they sit on the slices).
        chart.outside_labels = bool(prop(model, nonstyle + style, "tschchartinfopiecalloutlinetype", 0))
    chart.skip_hidden = bool(prop(model, nonstyle, "tschchartinfodefaultskiphiddendata", True))

    for s, n in zip(c.value_axis_styles, c.value_axis_nonstyles):
        chart.value_axes.append(read_axis(model, c, [s.identifier, n.identifier], "value", tools))
    if c.category_axis_styles and c.category_axis_nonstyles:
        ids = [c.category_axis_styles[0].identifier, c.category_axis_nonstyles[0].identifier]
        chart.category_axis = read_axis(model, c, ids, "category", tools)
        chart.category_axis["series_labels"] = bool(prop(model, ids, "tschchartaxiscategoryshowserieslabels", False))

    read_series(book, model, c, chart, tools)
    chart.frame = frame(chart)
    chart.plot = plot_area(chart)
    chart.size = (max(40, round(chart.frame[2] * PT_TO_PX)), max(40, round(chart.frame[3] * PT_TO_PX)))
    return chart


def read_axis(model, c, ids, which, tools):
    def get(name, default=None):
        return prop(model, ids, f"tschchartaxis{which}{name}", prop(model, ids, f"tschchartaxisdefault{name}", default))

    axis = {
        "labels": bool(get("showlabels", True)),
        "line": bool(get("showaxis", which == "category")),
        "gridlines": paint(get("majorgridlinestroke")) if get("showmajorgridlines", which == "value") else None,
        # Numbers counts minor gridlines as steps between major ones: 1 step draws none.
        "minor": paint(get("minorgridlinestroke")) if get("showminorgridlines", False)
        and (get("numberofminorgridlines", 1) or 1) > 1 else None,
        "title": (get("title") or None) if get("showtitle", False) else None,
        "font": text_style(model, c, get("labelparagraphstyleindex", 0), tools),
        "title_font": text_style(model, c, get("titleparagraphstyleindex", 0), tools),
        "angle": get("labelsorientation", 0.0) or 0.0,
    }
    for key in ("min", "max"):
        number = prop(model, ids, f"tschchartaxisdefaultuser{key}")
        if number is not None and number.HasField("number_archive") and math.isfinite(number.number_archive):
            axis[key] = number.number_archive
    return axis


def sheet_of(model, drawable):
    rid, seen = drawable.super.parent.identifier, 0
    while rid and rid in model.objects and seen < 8:
        obj = model.objects[rid]
        if type(obj).__name__ == "SheetArchive":
            return rid
        rid = obj.super.parent.identifier if hasattr(obj, "super") and obj.super.HasField("parent") else 0
        seen += 1
    return None


def series_ids(c, index):
    """Where series `index` looks its settings up: its own style, its theme style, its non-style settings."""
    private = {e.index: e.reference.identifier for e in c.series_private_styles.entries}
    themes = [r.identifier for r in c.series_theme_styles]
    style = ([private[index]] if index in private else []) + ([themes[index % len(themes)]] if themes else [])
    return style + [e.reference.identifier for e in c.series_non_styles.entries if e.index == index]


def style_series(model, c, chart, s, index, tools):
    ids = series_ids(c, index)
    kind = f"mixed {s.kind}" if chart.kind == "mixed" else chart.kind
    parts = PARTS[kind]
    # Series past the theme's colours reuse them, lighter each time round, unless given their own.
    themes = len(c.series_theme_styles) or 1
    private = {e.index: e.reference.identifier for e in c.series_private_styles.entries}

    def get(name, part, default=None):
        return prop(model, ids, f"tschchartseries{part}{name}", default)

    def colour(name, part):
        value, owner = prop_and_owner(model, ids, f"tschchartseries{part}{name}")
        own = owner is not None and index in private and model.objects[private[index]] is owner
        return lighten(paint(value), 0 if own else 0.125 * (index // themes)), value

    if "fill" in parts:
        s.colour = colour("fill", parts["fill"])[0] or colour("fill", "default")[0]
    line, stroke = colour("stroke", parts["stroke"])
    if line:
        s.line = (*line, stroke.width)
    if ("symbol" in parts and get("showsymbol", parts["symbol"], chart.kind == "scatter")
            and get("symboltype", parts["symbol"], 1) != 0):
        size = get("symbolsize", parts["symbol"], 7.0)
        outline = colour("symbolstroke", parts["symbol"])[0]
        s.marker = {"size": round(size) if size and 2 <= size <= 72 else 7,  # "automatic" is stored as a huge size
                    "colour": colour("symbolfill", parts["symbol"])[0] or s.colour or line,
                    "outline": outline or line or s.colour}
    if chart.kind == "scatter":
        s.joined = bool(get("showline", "scatter", False))
    names = kind in ("pie", "doughnut") and get("showserieslabels", "pie", False)
    if get("showvaluelabels", parts["labels"], False) or names:
        labels = {}
        if get("showvaluelabels", parts["labels"], False):
            percent = kind in ("pie", "doughnut") and get("numberformattype", "pie", 0) == 4
            labels["percentage" if percent else "value"] = True
        if names:
            labels["category"] = True  # a slice's name is its category once slices are folded into one series
        position = parts.get("position")
        if chart.subtype == "stacked" and position in ("bar", "area"):
            position = "stacked" + position
        text = parts["text"] + ("outside" if chart.outside_labels else "")
        style_index = get("labelparagraphstyleindex", text, get("labelparagraphstyleindex", "default", 0))
        labels["font"] = text_style(model, c, style_index, tools, 12.0)
        labels["position"] = get("valuelabelposition", position) if position else None
        s.labels = labels


def read_series(book, model, c, chart, tools):
    grid = c.grid
    names = list(grid.row_name)
    chart.categories = list(grid.column_name)
    snapshot = [[v.numeric_value if v.HasField("numeric_value") else None for v in row.value] for row in grid.grid_row]
    if not names or len(snapshot) != len(names):
        raise Unsupported("chart has no data")
    mediator = model.objects[c.mediator.identifier] if c.HasField("mediator") else None
    formulas = list(mediator.formulas.data_formulae) if mediator is not None else []
    links = None
    if len(formulas) == len(names):
        try:
            links = [[cell for block in decode_link(model, f) for cell in expand(book, block)] for f in formulas]
        except Exception:
            links = None
    if links and any(not matches(book, cells, snapshot[i], chart.skip_hidden) for i, cells in enumerate(links)):
        links = None  # the table changed shape since the chart's snapshot; use the snapshot
    if links:
        chart.tables = list(dict.fromkeys(cell[0] for cells in links for cell in cells))

    for i, name in enumerate(names):
        s = Series(name)
        s.values = snapshot[i]
        s.cells = links[i] if links else None
        s.kind = chart.kind
        if chart.kind == "mixed":
            ids = series_ids(c, i)
            s.kind = MIXED_KINDS.get(prop(model, ids, "tschchartseriesdefaultseriestype", 0), "column")
            s.y2 = prop(model, ids, "tschchartseriesdefaultyaxisordinal", 0) == 1
        style_series(model, c, chart, s, i, tools)
        chart.series.append(s)

    if chart.kind == "scatter":
        # Each series is plotted against x values: the first series' for all (shared x), or in pairs.
        xs = chart.series
        pairs = [(xs[0], y) for y in xs[1:]] if c.scatter_format == SHARED_X else list(zip(xs[0::2], xs[1::2]))
        for x, y in pairs:
            y.x_cells, y.x_values = x.cells, x.values
        chart.series = [y for _, y in pairs]
        if not chart.series:
            raise Unsupported("scatter chart without pairs")
    elif chart.kind in ("pie", "doughnut"):
        # Numbers keeps each slice as its own series; Excel draws one series with a colour per slice.
        fold(chart, c.multidataset_index if 0 <= c.multidataset_index < len(chart.categories) else 0)
    elif (chart.kind in ("column", "bar") and not chart.subtype and len(chart.categories) == 1
          and len(chart.series) > 1 and chart.legend is None):
        # One data set: Numbers draws a bar per series and names each one underneath. One series
        # with a colour per bar looks the same in Excel.
        fold(chart, 0)
        chart.category_axis["labels"] = chart.category_axis.get("labels") or chart.category_axis.get("series_labels")
    find_labels(book, chart, tools)
    widths = [len(cell_text(book, cell, tools) or "") for s in chart.series for cell in (s.cells or [])]
    widths += [len(f"{v:,.0f}") for s in chart.series if not s.cells for v in s.values if isinstance(v, (int, float))]
    chart.label_width = max(widths, default=3)
    chart.value_format = value_format(book, chart, tools)


def fold(chart, index):
    slices = chart.series
    one = Series(chart.categories[index] if chart.categories else "")
    one.kind = chart.kind
    one.cells = [s.cells[index] for s in slices] if all(s.cells and index < len(s.cells) for s in slices) else None
    one.values = [s.values[index] if index < len(s.values) else None for s in slices]
    one.point_colours = [s.colour for s in slices]
    one.line = slices[0].line
    one.labels = slices[0].labels
    chart.categories = [s.name for s in slices]
    chart.series = [one]


def matches(book, cells, values, skip_hidden):
    """Whether linked cells still line up with the chart's snapshot of them (which leaves out
    hidden rows and columns when the chart skips hidden data)."""
    if len(cells) == len(values):
        return True
    if not skip_hidden:
        return False
    hidden, visible = {}, 0
    for table_id, r, c in cells:
        if table_id not in hidden:
            table = book.table_by_id[table_id]
            hidden[table_id] = (extras.hidden(table, columns=False), extras.hidden(table, columns=True))
        rows, cols = hidden[table_id]
        visible += r not in rows and c not in cols
    return visible == len(values)


def cell_text(book, cell, tools):
    table_id, r, c = cell
    try:
        return tools.text(book.table_by_id[table_id]._data[r][c])
    except (IndexError, KeyError):
        return None


def value_format(book, chart, tools):
    """How the chart's values read in their table: {prefix, suffix, decimals, thousands, percent}.
    Excel shows the axis and value labels in the cells' own format; the preview copies it."""
    for s in chart.series:
        for cell in (s.cells or [])[:1]:
            text = cell_text(book, cell, tools) or ""
            digits = [m.start() for m in re.finditer(r"\d", text)]
            if not digits:
                continue
            table_id, r, c = cell
            code = tools.number_format(book.table_by_id[table_id]._data[r][c]) or ""
            head = code.split(";")[0]
            places = re.search(r"\.([0#?]+)", head)
            suffix = text[digits[-1] + 1:].strip(" )")
            return {"prefix": re.sub(r"[-()\s]", "", text[:digits[0]]), "suffix": suffix,
                    "decimals": (len(places.group(1)) if places else 0) if code else None,
                    "thousands": "," in head, "percent": suffix.startswith("%")}
    return None


def find_labels(book, chart, tools):
    """Point series names and categories at the table cells holding the same text, so they
    follow the table in Excel too."""
    def text(cell):
        return cell_text(book, cell, tools)

    def model_of(table_id):
        return book.table_by_id[table_id]._model.objects[table_id]

    for s in chart.series:
        if not s.cells or s.name in (None, ""):
            continue
        table_id = s.cells[0][0]
        tm = model_of(table_id)
        rows, cols = sorted({cell[1] for cell in s.cells}), sorted({cell[2] for cell in s.cells})
        if len(cols) == 1:    # a column of values: its name is in a header row, or just above
            candidates = [(table_id, r, cols[0]) for r in (*range(tm.number_of_header_rows), rows[0] - 1) if 0 <= r < rows[0]]
        elif len(rows) == 1:  # a row of values: its name is in a header column, or just left of them
            candidates = [(table_id, rows[0], c) for c in (*range(tm.number_of_header_columns), cols[0] - 1)
                          if 0 <= c < cols[0]]
        else:
            candidates = []
        s.name_cell = next((cell for cell in reversed(candidates) if text(cell) == str(s.name).strip()), None)

    first = next((s for s in chart.series if s.cells and len(s.cells) == len(chart.categories)), None)
    if first is None or not chart.categories:
        return
    table_id = first.cells[0][0]
    if any(cell[0] != table_id for cell in first.cells):
        return
    tm = model_of(table_id)
    rows, cols = [cell[1] for cell in first.cells], [cell[2] for cell in first.cells]
    if len(set(cols)) == 1:    # categories in a column alongside: a header column, or one of the nearest
        nearby = (*range(tm.number_of_header_columns), *range(cols[0] - 1, max(cols[0] - 5, -1), -1))
        options = [[(table_id, r, col) for r in rows] for col in dict.fromkeys(nearby) if col != cols[0]]
    elif len(set(rows)) == 1:  # categories in a row above: a header row, or one of the nearest
        nearby = (*range(tm.number_of_header_rows), *range(rows[0] - 1, max(rows[0] - 5, -1), -1))
        options = [[(table_id, row, c) for c in cols] for row in dict.fromkeys(nearby) if row != rows[0]]
    else:
        options = []
    wanted = [str(label).strip() for label in chart.categories]
    chart.category_cells = next((cells for cells in options if [text(cell) for cell in cells] == wanted), None)


# ---------------------------------------------------------------- layout


def frame(chart):
    """The chart's whole frame on the canvas, (x, y, w, h): its body plus room for the title, axis
    labels and legend, much as Apple's own export sizes it."""
    x, y, w, h = chart.body
    left, top, right, bottom = x, y, x + w, y + h
    if chart.kind in ("pie", "doughnut") and chart.outside_labels:
        size = (chart.series[0].labels or {}).get("font", {}).get("size", 12) if chart.series else 12
        margin = size * 3.75  # labels and their lines around the pie
        left, top, right, bottom = left - margin, top - margin, right + margin, bottom + margin
    if chart.kind == "radar":  # category names sit around the web
        size = chart.category_axis.get("font", {}).get("size", 10)
        widest = max((len(str(name)) for name in chart.categories), default=4)
        left, right = left - 0.55 * size * widest - 6, right + 0.55 * size * widest + 6
        top, bottom = top - size * 1.8, bottom + size * 1.8
    if chart.title:
        # Title band: the line height of its font plus spacing, as in Apple's export.
        band = chart.title_style["size"] * (2.5 if chart.title_style["font"] == "Segoe UI" else 2.12)
        top = min(top, y - band) if chart.kind != "radar" else top - band
    if chart.kind not in ("pie", "doughnut", "radar"):
        category = chart.category_axis
        size = category.get("font", {}).get("size", 10)
        names = [str(name) for name in chart.categories] or [""]
        widest = max(len(name) for name in names)
        digits = (chart.label_width or 3) + 1
        if chart.kind == "bar":  # categories down the left, values along the bottom
            left -= 0.55 * size * widest + 10
            bottom += size * 2
        else:
            left -= 0.55 * size * digits + 8
            if category.get("labels", True):
                fits = len(names) * (0.55 * size * widest + 6) <= w and not category.get("angle")
                bottom += size * 2 if fits else 0.55 * size * widest + 12
        if any(s.y2 for s in chart.series):
            right += 0.55 * size * digits + 8
            if len(chart.value_axes) > 1 and chart.value_axes[1].get("title"):
                right += chart.value_axes[1]["title_font"]["size"] * 2.4
        if chart.value_axes and chart.value_axes[0].get("title"):
            left -= chart.value_axes[0]["title_font"]["size"] * 2.4
        if category.get("title"):
            bottom += category["title_font"]["size"] * 2.4
    if chart.legend:
        lx, ly, lw, lh = chart.legend
        left, top, right, bottom = min(left, lx), min(top, ly), max(right, lx + lw), max(bottom, ly + lh)
    return left, top, right - left, bottom - top


def plot_area(chart):
    """Where the plot goes: the chart's body, except that a pie stays clear of a legend below it
    (Numbers draws the pie inside its box with room to spare; Excel fills the box)."""
    if chart.kind in ("pie", "doughnut") and chart.legend and overlap(chart.body, chart.legend) > 0:
        bx, by, bw, bh = chart.body
        ly = chart.legend[1]
        if ly > by + bh / 2:
            side = max(bh / 2, ly - 4 - by)
            return bx + (bw - side) / 2, by, side, side
    return chart.body


# ---------------------------------------------------------------- writing


class ChartData:
    """A hidden sheet for chart data Excel can't reference in place."""

    def __init__(self, workbook, name):
        self.workbook, self.name, self.ws, self.col = workbook, name, None, 0

    def column(self, values, formats=None):
        """Writes a column of values, each ("=formula", cached value) or a plain value; returns its range."""
        if self.ws is None:
            self.ws = self.workbook.add_worksheet(self.name)
            self.ws.hide()
        for r, value in enumerate(values):
            fmt = formats[r] if formats else None
            if isinstance(value, tuple):
                self.ws.write_formula(r, self.col, value[0], fmt, value[1])
            elif value is None:
                self.ws.write_blank(r, self.col, None, fmt)
            elif isinstance(value, (int, float)):
                self.ws.write_number(r, self.col, value, fmt)
            else:
                self.ws.write_string(r, self.col, str(value), fmt)
        ref = [self.name, 0, self.col, max(len(values) - 1, 0), self.col]
        self.col += 1
        return ref


def reference(cells, values, data, place, cell_info):
    """An Excel reference for a series' cells: a range when they're one straight run, a union of
    ranges when they're on one sheet (as in Apple's export), otherwise a Chart Data column of
    formulas pointing at them; or the chart's own copy of the values when there are no cells."""
    if not cells:
        return data.column(list(values)) if values else None
    placed = [place(cell) for cell in cells]
    sheet = placed[0][0]
    if all(p[0] == sheet for p in placed):
        runs = []
        for _, r, c in placed:
            last = runs[-1] if runs else None
            if last and last[2] == last[4] == c and r == last[3] + 1:
                last[3] = r
            elif last and last[1] == last[3] == r and c == last[4] + 1:
                last[4] = c
            else:
                runs.append([sheet, r, c, r, c])
        if len(runs) == 1:
            return runs[0]
        union = "=(" + ",".join(f"{quote_sheetname(sheet)}!{xl_range_abs(*run[1:])}" for run in runs) + ")"
        if len(union) <= 250:  # Excel limits the length of a series reference
            return union
    formulas, formats = [], []
    for cell, (sheet, r, c) in zip(cells, placed):
        cached, fmt = cell_info(cell)
        formulas.append((f"={quote_sheetname(sheet)}!{xl_rowcol_to_cell(r, c, True, True)}", cached))
        formats.append(fmt)
    return data.column(formulas, formats)


def font(style, default_size=10, bold=None):
    style = style or {}
    return {"name": style.get("font") or "Arial", "size": style.get("size") or default_size,
            "bold": bool(style.get("bold")) if bold is None else bold, "italic": bool(style.get("italic")),
            "color": style.get("color") or "#000000"}


def fill(colour):
    if not colour:
        return {"none": True}
    out = {"color": colour[0]}
    if colour[1] < 0.99:
        out["transparency"] = max(0, min(100, round((1 - colour[1]) * 100)))
    return out


def line(spec, default_width=1.0):
    if not spec:
        return {"none": True}
    out = {"color": spec[0], "width": max(0.25, round(spec[2] or default_width, 2))}
    if spec[1] < 0.99:
        out["transparency"] = max(0, min(100, round((1 - spec[1]) * 100)))
    return out


def series_options(chart, s, categories, data, place, cell_info):
    options = {"values": reference(s.cells, s.values, data, place, cell_info)}
    if options["values"] is None:
        return None
    if chart.kind == "scatter":
        options["categories"] = reference(s.x_cells, s.x_values, data, place, cell_info)
    elif categories:
        options["categories"] = categories
    if s.name_cell:
        sheet, r, c = place(s.name_cell)
        options["name"] = [sheet, r, c]
    elif s.name not in (None, ""):
        options["name"] = str(s.name)

    kind = s.kind
    if kind in ("line", "scatter"):
        options["line"] = line(s.line or (s.colour and (*s.colour, 2.0)), 2.0)
        if kind == "scatter" and not s.joined:
            options["line"] = {"none": True}
    else:
        options["fill"] = fill(s.colour)
        options["border"] = line(s.line) if kind in ("area", "radar") else {"none": True}
    if kind in ("line", "scatter", "area", "radar"):
        m = s.marker
        options["marker"] = ({"type": "circle", "size": m["size"], "fill": fill(m["colour"]),
                              "border": fill(m["outline"] or m["colour"])} if m else {"type": "none"})
    if s.point_colours:
        separator = {"color": "#FFFFFF", "width": 0.75} if kind in ("pie", "doughnut") else {"none": True}
        options["points"] = [{"fill": fill(colour), "border": separator} for colour in s.point_colours]
    if kind in ("column", "bar"):
        if chart.gap is not None:
            options["gap"] = max(0, min(500, chart.gap))
        if chart.subtype == "stacked":
            options["overlap"] = 100
        elif chart.overlap is not None:
            options["overlap"] = max(-100, min(100, chart.overlap))
    if s.labels:
        labels = {k: True for k in ("category", "value", "percentage") if s.labels.get(k)}
        labels["font"] = font(s.labels.get("font"), 12)
        if kind in ("column", "bar"):
            position = LABEL_POSITIONS.get(s.labels.get("position"))
            if position and not (chart.subtype == "stacked" and position == "outside_end"):
                labels["position"] = position
        elif kind == "pie":
            labels["position"] = "outside_end" if chart.outside_labels else "center"
        if chart.outside_labels:
            labels["leader_lines"] = True
        if "percentage" in labels:
            labels["num_format"] = "0%"
        if "category" in labels and len(labels) > 2:
            labels["separator"] = "\n"
        options["data_labels"] = labels
    if s.y2:
        options["y2_axis"] = True
    return options


def axis_options(axis, kind):
    axis = axis or {}
    options = {"num_font": font(axis.get("font"), bold=False), "major_tick_mark": "none", "minor_tick_mark": "none"}
    if axis.get("angle"):
        options["num_font"]["rotation"] = max(-90, min(90, -round(axis["angle"])))
    if axis.get("title"):
        options["name"] = axis["title"]
        options["name_font"] = font(axis.get("title_font"), bold=False)
    for key in ("min", "max"):
        if key in axis:
            options[key] = axis[key]
    if kind == "value":
        gridlines, minor = axis.get("gridlines"), axis.get("minor")
        options["major_gridlines"] = {"visible": True, "line": line((*gridlines, 0.5))} if gridlines else {"visible": False}
        if minor:
            options["minor_gridlines"] = {"visible": True, "line": line((*minor, 0.5))}
        options["line"] = {"color": "#000000", "width": 1} if axis.get("line") else {"none": True}
        options["label_position"] = "next_to"
    else:
        options["line"] = {"color": "#000000", "width": 1} if axis.get("line", True) else {"none": True}
        options["label_position"] = "low"
    if not axis.get("labels", True):
        options["label_position"] = "none"
    return options


def write_chart(workbook, chart, data, place, cell_info):
    kind = chart.kind
    secondary = None
    if kind == "mixed":
        primary_kind = chart.series[0].kind
        excel = workbook.add_chart({"type": primary_kind})
        other = next((s.kind for s in chart.series if s.kind != primary_kind), None)
        if other:
            secondary = workbook.add_chart({"type": other})
    else:
        options = {"type": kind}
        if chart.subtype:
            options["subtype"] = chart.subtype
        if kind == "scatter":
            chart.subtype = options["subtype"] = ("straight_with_markers" if any(s.joined for s in chart.series)
                                                  else "marker_only")
        if kind == "radar":
            options["subtype"] = "filled" if any(s.colour for s in chart.series) else "with_markers"
        excel = workbook.add_chart(options)

    categories = None
    if kind != "scatter" and (chart.category_cells or chart.categories):
        categories = reference(chart.category_cells, chart.categories, data, place, cell_info)
    added = 0
    for s in chart.series:
        options = series_options(chart, s, categories, data, place, cell_info)
        if options is None:
            continue
        (secondary if kind == "mixed" and s.kind != chart.series[0].kind else excel).add_series(options)
        added += 1
    if not added:
        return None
    if secondary is not None:
        excel.combine(secondary)

    if chart.title:
        excel.set_title({"name": chart.title, "name_font": font(chart.title_style, 12), "overlay": True})
    else:
        excel.set_title({"none": True})

    fx, fy, fw, fh = chart.frame

    def relative(x, y, w, h):
        clamp = lambda v: max(0.0, min(1.0, round(v, 4)))  # noqa: E731
        return {"x": clamp((x - fx) / fw), "y": clamp((y - fy) / fh), "width": clamp(w / fw), "height": clamp(h / fh)}

    if chart.legend:
        # A tall legend lists its entries down, a wide one across.
        excel.set_legend({"position": "right" if chart.legend[2] < chart.legend[3] * 2 else "bottom",
                          "font": font(chart.legend_style, bold=False), "layout": relative(*chart.legend)})
    else:
        excel.set_legend({"none": True})
    excel.set_plotarea({"layout": relative(*chart.plot), "fill": {"none": True}, "border": {"none": True}})
    excel.set_chartarea({"border": {"none": True}, "fill": {"none": True}})

    if kind in ("pie", "doughnut"):
        excel.set_rotation(chart.rotation or 0)
        if kind == "doughnut":
            excel.set_hole_size(chart.hole or 50)
    else:
        values = axis_options(chart.value_axes[0] if chart.value_axes else None, "value")
        category = axis_options(chart.category_axis, "category")
        if kind == "bar":
            # Numbers lists bars from the top; Excel starts at the bottom unless the axis is reversed.
            category.update(reverse=True, crossing="max")
            excel.set_x_axis(values)
            excel.set_y_axis(category)
        elif kind == "scatter":
            excel.set_x_axis({**axis_options(chart.category_axis, "value"), "major_gridlines": {"visible": False},
                              "line": {"color": "#000000", "width": 1}})
            excel.set_y_axis(values)
        else:
            excel.set_x_axis(category)
            excel.set_y_axis(values)
        if any(s.y2 for s in chart.series):
            y2 = axis_options(chart.value_axes[1] if len(chart.value_axes) > 1 else None, "value")
            y2["major_gridlines"] = {"visible": False}
            target = secondary if secondary is not None and any(s.y2 and s.kind != chart.series[0].kind
                                                                for s in chart.series) else excel
            target.set_y2_axis(y2)
    if not chart.skip_hidden:
        excel.show_hidden_data()
    excel.set_size({"width": chart.size[0], "height": chart.size[1]})
    return excel


def write_charts(workbook, layout, worksheets, cell_info, unique_name):
    """Adds every chart to the worksheet the layout puts it on; returns how many were written.
    cell_info(cell) gives a table cell's (value, format) for copies on the Chart Data sheet."""
    data = ChartData(workbook, unique_name("Chart Data"))

    def place(cell):
        table_id, r, c = cell
        placed = layout.place[table_id]
        return (placed.sheet, *placed.cell(r, c))

    written = 0
    for plan in layout.sheets:
        for chart, (x, y) in plan.charts:
            try:
                excel = write_chart(workbook, chart, data, place, cell_info)
            except Exception:
                excel = None
            if excel is None:
                layout.book.unsupported["charts"] += 1
                continue
            worksheets[plan.name].insert_chart(0, 0, excel, {"x_offset": x, "y_offset": y})
            written += 1
    return written


# ---------------------------------------------------------------- preview


def describe(chart, position):
    """What the app's preview needs to draw a chart as it will look in Excel: its place (`position`)
    and size in pixels from the sheet's top-left cell, where the plot and legend sit in it
    (fractions), and its data and colours."""
    fx, fy, fw, fh = chart.frame

    def relative(rect):
        x, y, w, h = rect
        return [round((x - fx) / fw, 4), round((y - fy) / fh, 4), round(w / fw, 4), round(h / fh, 4)]

    def numbers(values):
        return [v if isinstance(v, (int, float)) and math.isfinite(v) else None for v in values or []]

    def hex_of(colour):
        return colour[0] if colour else None

    series = []
    for s in chart.series:
        labels = s.labels or {}
        series.append({
            "name": "" if s.name is None else str(s.name),
            "kind": s.kind,
            "values": numbers(s.values),
            "x": numbers(s.x_values) if chart.kind == "scatter" else None,
            "color": hex_of(s.colour),
            "line": hex_of(s.line),
            "lineWidth": s.line[2] if s.line else None,
            "points": [hex_of(colour) for colour in s.point_colours] if s.point_colours else None,
            "marker": {"fill": hex_of(s.marker["colour"]), "outline": hex_of(s.marker["outline"]),
                       "size": s.marker["size"]} if s.marker else None,
            "joined": s.joined,
            "labels": "percentage" if labels.get("percentage") else "value" if labels.get("value") else None,
            "labelNames": bool(labels.get("category")),
            "labelStyle": labels.get("font"),
            "labelPosition": LABEL_POSITIONS.get(labels.get("position")),
            "y2": s.y2,
        })
    axis = chart.value_axes[0] if chart.value_axes else {}
    return {
        "kind": chart.kind, "stacked": chart.subtype == "stacked",
        "x": position[0], "y": position[1], "w": chart.size[0], "h": chart.size[1],
        "title": chart.title, "titleStyle": chart.title_style,
        "plot": relative(chart.plot), "legend": relative(chart.legend) if chart.legend else None,
        "legendStyle": chart.legend_style,
        "categories": [str(c) for c in chart.categories], "series": series,
        "gap": chart.gap, "overlap": chart.overlap, "hole": chart.hole, "rotation": chart.rotation,
        "outsideLabels": chart.outside_labels,
        "gridlines": hex_of(axis.get("gridlines")), "valueLabels": axis.get("labels", True),
        "axisStyle": axis.get("font"), "categoryLabels": chart.category_axis.get("labels", True),
        "categoryAngle": chart.category_axis.get("angle", 0), "categoryLine": chart.category_axis.get("line", True),
        "format": chart.value_format,
        "valueTitle": axis.get("title"), "categoryTitle": chart.category_axis.get("title"),
        "y2Title": chart.value_axes[1].get("title") if len(chart.value_axes) > 1 and any(s.y2 for s in chart.series) else None,
        "titleFont": axis.get("title_font"),
    }
