"""
Reads Apple Numbers files and writes formatted Excel workbooks.

Runs as a long-lived helper for the Electron app: one JSON request per line on stdin,
one JSON response per line on stdout.

  {"id": 1, "cmd": "preview", "path": "C:/x.numbers"}
  {"id": 2, "cmd": "convert", "path": "C:/x.numbers", "out": "C:/x.xlsx"}
"""

import datetime as dt
import json
import os
import re
import sys
import warnings
from collections import OrderedDict
from types import SimpleNamespace
from decimal import ROUND_HALF_UP, Decimal

import xlsxwriter
from numbers_parser import Document
from xlsxwriter.utility import xl_range_abs

import charts
import extras
import pivots
import strokes
from layout import SHEET, TABLES, Footprint, Layout
from numbers_parser.cell import (
    RGB,
    BoolCell,
    DateCell,
    DurationCell,
    EmptyCell,
    ErrorCell,
    MergedCell,
    NumberCell,
)
from numbers_parser.cell import _auto_units
from numbers_parser.constants import DATETIME_FIELD_MAP, DurationStyle, DurationUnits, FormatType
from numbers_parser.containers import ObjectStore
from numbers_parser.xrefs import CellRange
from numbers_parser.numbers_uuid import NumbersUUID

warnings.simplefilter("ignore")

# ObjectStore.find_refs scans every object in the document on each call, and formula decoding calls it
# hundreds of thousands of times. We never modify documents, so the results can be cached per store.
_find_refs = ObjectStore.find_refs


def _cached_find_refs(self, ref_name):
    cache = self.__dict__.setdefault("_refs_cache", {})
    if ref_name not in cache:
        cache[ref_name] = _find_refs(self, ref_name)
    return cache[ref_name]


ObjectStore.find_refs = _cached_find_refs


# numbers-parser prints references the way Numbers displays them: header names, and table names
# only where Numbers thinks they're needed, which loses which table a name points to. Print every
# reference as an exact code instead, ⟦table|row start|row end|column start|column end|absolute
# flags⟧ (blank = whole row or column), which translate_formula turns into Excel references.
def _exact_reference(self):
    if self._ref_error:
        return "#REF!"
    table = self.to_table_id if self.to_table_id is not None else self.from_table_id
    flags = (int(bool(self.row_start_is_abs)) | int(bool(self.col_start_is_abs)) << 1
             | int(bool(self.row_end_is_abs)) << 2 | int(bool(self.col_end_is_abs)) << 3)
    part = lambda v: "" if v is None else str(v)  # noqa: E731
    return f"⟦{table}|{part(self.row_start)}|{part(self.row_end)}|{part(self.col_start)}|{part(self.col_end)}|{flags}⟧"


def _light_setup(self):
    # The exact codes don't use names, so skip numbers-parser's per-reference name lookups (slow).
    self._ref_error = (self.row_start is not None and self.row_start < 0) or (
        self.col_start is not None and self.col_start < 0)


CellRange.__str__ = _exact_reference
CellRange.__post_init__ = _light_setup

# numbers-parser formats dates with glibc's "%-d" (no leading zero), which Windows rejects; Windows spells it "%#d".
if os.name == "nt":
    for _field, _code in DATETIME_FIELD_MAP.items():
        if isinstance(_code, str):
            DATETIME_FIELD_MAP[_field] = _code.replace("%-", "%#")

PREVIEW_MAX_ROWS = 1000
PREVIEW_MAX_COLS = 50
DECIMAL_PLACES_AUTO = 253

# Fonts that ship with macOS but not Windows; Excel would otherwise fall back unpredictably.
FONT_SUBSTITUTES = {
    "HelveticaNeue": "Arial",
    "Helvetica Neue": "Arial",
    "Helvetica": "Arial",
    "Avenir": "Segoe UI",
    "Avenir Next": "Segoe UI",
    "SF Pro": "Segoe UI",
    ".SFNS": "Segoe UI",
    "Menlo": "Consolas",
    "Monaco": "Consolas",
    "Lucida Grande": "Segoe UI",
}

CURRENCY_SYMBOLS = {
    "USD": "$", "EUR": "€", "GBP": "£", "JPY": "¥", "CNY": "¥", "INR": "₹", "KRW": "₩",
    "AUD": "A$", "CAD": "CA$", "NZD": "NZ$", "CHF": "CHF ", "ZAR": "R", "BRL": "R$",
    "MXN": "MX$", "SEK": "kr ", "NOK": "kr ", "DKK": "kr ", "RUB": "₽", "PHP": "₱",
}

# Functions that behave the same in Numbers and Excel, with the argument counts Excel accepts
# (None = any number). A formula using anything else, or the wrong number of arguments, keeps its
# last value instead: Excel refuses to open a file containing a formula it can't parse.
N = None
EXCEL_FUNCTIONS = {
    # math
    "SUM": (1, N), "PRODUCT": (1, N), "ABS": (1, 1), "ROUND": (2, 2), "ROUNDUP": (2, 2), "ROUNDDOWN": (2, 2),
    "INT": (1, 1), "TRUNC": (1, 2), "MOD": (2, 2), "POWER": (2, 2), "SQRT": (1, 1), "EXP": (1, 1), "LN": (1, 1),
    "LOG": (1, 2), "LOG10": (1, 1), "PI": (0, 0), "SIGN": (1, 1), "CEILING": (2, 2), "FLOOR": (2, 2),
    "MROUND": (2, 2), "EVEN": (1, 1), "ODD": (1, 1), "QUOTIENT": (2, 2), "GCD": (1, N), "LCM": (1, N),
    "FACT": (1, 1), "COMBIN": (2, 2), "PERMUT": (2, 2), "RAND": (0, 0), "RANDBETWEEN": (2, 2),
    "SUMPRODUCT": (1, N), "SUMSQ": (1, N), "DEGREES": (1, 1), "RADIANS": (1, 1), "SIN": (1, 1), "COS": (1, 1),
    "TAN": (1, 1), "ASIN": (1, 1), "ACOS": (1, 1), "ATAN": (1, 1), "ATAN2": (2, 2),
    # statistics
    "AVERAGE": (1, N), "AVERAGEA": (1, N), "MIN": (1, N), "MAX": (1, N), "MINA": (1, N), "MAXA": (1, N),
    "COUNT": (1, N), "COUNTA": (1, N), "COUNTBLANK": (1, 1), "COUNTIF": (2, 2), "COUNTIFS": (2, N),
    "SUMIF": (2, 3), "SUMIFS": (3, N), "AVERAGEIF": (2, 3), "AVERAGEIFS": (3, N), "MEDIAN": (1, N),
    "MODE": (1, N), "STDEV": (1, N), "STDEVP": (1, N), "VAR": (1, N), "VARP": (1, N), "LARGE": (2, 2),
    "SMALL": (2, 2), "RANK": (2, 3), "PERCENTILE": (2, 2), "QUARTILE": (2, 2),
    # logic and information
    "IF": (2, 3), "IFS": (2, N), "IFERROR": (2, 2), "AND": (1, N), "OR": (1, N), "NOT": (1, 1), "XOR": (1, N),
    "TRUE": (0, 0), "FALSE": (0, 0), "ISBLANK": (1, 1), "ISERROR": (1, 1), "ISNUMBER": (1, 1),
    "ISTEXT": (1, 1), "ISEVEN": (1, 1), "ISODD": (1, 1),
    # text
    "CONCATENATE": (1, N), "TEXTJOIN": (3, N), "LEFT": (1, 2), "RIGHT": (1, 2), "MID": (3, 3), "LEN": (1, 1),
    "LOWER": (1, 1), "UPPER": (1, 1), "PROPER": (1, 1), "TRIM": (1, 1), "CLEAN": (1, 1),
    "SUBSTITUTE": (3, 4), "REPLACE": (4, 4), "FIND": (2, 3), "SEARCH": (2, 3), "REPT": (2, 2),
    "EXACT": (2, 2), "VALUE": (1, 1), "CHAR": (1, 1), "CODE": (1, 1), "T": (1, 1),
    # dates
    "DATE": (3, 3), "TIME": (3, 3), "TODAY": (0, 0), "NOW": (0, 0), "YEAR": (1, 1), "MONTH": (1, 1),
    "DAY": (1, 1), "HOUR": (1, 1), "MINUTE": (1, 1), "SECOND": (1, 1), "WEEKDAY": (1, 2), "WEEKNUM": (1, 2),
    "EDATE": (2, 2), "EOMONTH": (2, 2), "DAYS360": (2, 3), "NETWORKDAYS": (2, 3), "WORKDAY": (2, 3),
    "DATEDIF": (3, 3), "YEARFRAC": (2, 3), "DATEVALUE": (1, 1), "TIMEVALUE": (1, 1), "DAYS": (2, 2),
    # lookup
    "VLOOKUP": (3, 4), "HLOOKUP": (3, 4), "LOOKUP": (2, 3), "MATCH": (2, 3), "INDEX": (2, 4),
    "CHOOSE": (2, N), "ROW": (0, 1), "COLUMN": (0, 1), "ROWS": (1, 1), "COLUMNS": (1, 1), "XLOOKUP": (3, 6),
    # finance
    "PMT": (3, 5), "FV": (3, 5), "PV": (3, 5), "NPER": (3, 5), "RATE": (3, 6), "NPV": (2, N), "IRR": (1, 2),
    "IPMT": (4, 6), "PPMT": (4, 6),
}


class UserError(Exception):
    """An error whose message can be shown to the user as-is. `extra` is sent along with it."""

    def __init__(self, message, **extra):
        super().__init__(message)
        self.extra = extra


# ---------------------------------------------------------------- loading

_cache = OrderedDict()


def load(path, password=None):
    """Parse a Numbers file (or a Mac folder package) into table models, cached by path + mtime."""
    if not path.lower().rstrip("\\/").endswith(".numbers"):
        raise UserError("Not a .numbers file.")
    key = (path, os.path.getmtime(path), password)
    if key in _cache:
        _cache.move_to_end(key)
        return _cache[key]

    try:
        doc = Document(path, password=password or None)
    except Exception as err:  # numbers-parser raises a variety of errors for bad files
        text = str(err)
        if "No password provided" in text:
            raise UserError("This file is password-protected.", needsPassword=True) from None
        if "Invalid password" in text:
            hint = re.search(r"Hint is '(.*)'", text)
            raise UserError("That password isn't right.", needsPassword=True,
                            hint=hint.group(1) if hint and hint.group(1) else None) from None
        raise UserError("Could not read this Numbers file (it may be damaged or not a real Numbers file).") from None
    finally:
        release_file()

    book = Book(doc)
    _cache[key] = book
    while len(_cache) > 4:
        _cache.popitem(last=False)
    return book


def release_file():
    """numbers-parser reads the whole archive when opening but leaves the zip file open, which on
    Windows stops the user moving or deleting the .numbers file. Close any zip it left open."""
    import gc
    import zipfile

    for obj in gc.get_objects():
        if isinstance(obj, zipfile.ZipFile) and obj.fp is not None:
            try:
                obj.close()
            except Exception:
                pass


class Book:
    """A Numbers document as read: its sheets and tables, table titles, charts and pivot tables.
    Where everything goes in the workbook depends on the layout chosen (see layout())."""

    def __init__(self, doc):
        self.doc = doc
        self.sheets = []  # [(sheet id, sheet name, [tables])] in Numbers' order
        self.tables = []  # [(sheet name, table)]
        for sheet in doc.sheets:
            tables = list(sheet.tables)
            background = sheet_background(doc._model, sheet)
            for table in tables:
                if background:
                    SHEET_BACKGROUNDS[(id(table._model), table._table_id)] = background
                self.tables.append((sheet.name, table))
            self.sheets.append((sheet._sheet_id, sheet.name, tables))
        if not self.tables:
            raise UserError("No tables found in this file.")
        self.table_by_id = {table._table_id: table for _, table in self.tables}
        self.sheet_name_of = {table._table_id: sheet_name for sheet_name, table in self.tables}
        self.titles = {table._table_id: table_title(table) for _, table in self.tables}
        self.unsupported = count_drawables(doc)
        try:
            self.charts, self.unsupported["charts"] = charts.read_charts(self, CHART_TOOLS)
        except Exception:  # never let a chart stop the tables converting
            self.charts = []
        # Pivot tables whose summary values have to be recalculated: {table_id: {(row, col): (value, source cell)}}
        self.pivot_values = {}
        for _, table in self.tables:
            try:
                pivot = pivots.pivot_for(table)
            except Exception:  # pivots.Unsupported, or a pivot laid out in a way we don't recognise
                self.unsupported["pivots"] += 1
                continue
            if pivot is not None:
                self.pivot_values[table._table_id] = pivot.values
                if pivot.missing:
                    self.unsupported["pivots"] += 1
        self.layouts = {}

    def layout(self, options=None):
        """Where everything goes for a set of layout choices (see layout.Layout), worked out once."""
        key = json.dumps(options or {}, sort_keys=True)
        if key not in self.layouts:
            self.layouts[key] = Layout(self, options, lambda table: footprint(self, table), unique_sheet_name)
        return self.layouts[key]


# What charts.py needs from here: a cell's text as Numbers shows it and its Excel number format, and
# the Windows font for a Mac font. (Lambdas, as the functions are defined further down.)
CHART_TOOLS = SimpleNamespace(text=lambda cell: "" if isinstance(cell, (EmptyCell, MergedCell)) else formatted(cell).strip(),
                              font=lambda name: font_name(name), bold=lambda name: font_is_bold(name),
                              number_format=lambda cell: excel_number_format(cell))


def footprint(book, table):
    """A table as the layout sees it: its position on its Numbers sheet, and its rows (its title row
    first, if Numbers shows its name) and columns as they're written out."""
    title = book.titles[table._table_id]
    hidden_rows, hidden_cols = extras.hidden(table, columns=False), extras.hidden(table, columns=True)
    rows = [(height, r in hidden_rows) for r, height in enumerate(header_sizes(table, columns=False))]
    cols = [(max(6.0, width), c in hidden_cols) for c, width in enumerate(header_sizes(table, columns=True))]
    if title:
        rows.insert(0, (title_height(title), False))
    try:
        g = pivots.find_table_info(table._model, table._table_id).super.geometry
        position = (g.position.x, g.position.y)
    except Exception:
        position = (0.0, 0.0)
    return Footprint(table, position, rows, cols, title)


def table_title(table):
    """{text, font, size, bold, italic, color} for a table whose name Numbers shows above it, else None."""
    try:
        model = table._model
        tm = model.objects[table._table_id]
        if not tm.table_name_enabled or not table.name:
            return None
        style = model.objects[tm.table_name_style.identifier]
        colour = rgb_hex(rgb_from_proto(model.char_property(style, "font_color")))
        font = model.char_property(style, "font_name")
        title = {"text": table.name, "font": font_name(font) or "Arial",
                 "size": round(model.char_property(style, "font_size"), 1),
                 "bold": bool(model.char_property(style, "bold")) or font_is_bold(font),
                 "italic": bool(model.char_property(style, "italic")),
                 "color": None if colour == "#000000" else colour}
        title["height"] = title_height(title)
        return title
    except Exception:
        return None


# {(model id, table id): "#RRGGBB"} for tables on sheets with a coloured page background.
SHEET_BACKGROUNDS = {}


def sheet_background(model, sheet):
    """The sheet's page colour (templates often use one), or None for plain white."""
    try:
        style = model.objects[model.objects[sheet._sheet_id].style.identifier]
        fill = style.sheet_properties.fill
        if fill.HasField("color"):
            colour = rgb_hex(rgb_from_proto(fill.color))
            return None if colour == "#FFFFFF" else colour
    except Exception:
        pass
    return None


def stored_sizes(table, columns):
    """Row heights (or column widths) in points that Numbers stores explicitly; other rows are
    sized automatically to fit their content. Table.row_height rebuilds its lookup on every call,
    which is quadratic on big tables, so the stored sizes are read once."""
    model, table_id = table._model, table._table_id
    bds = model.objects[table_id].base_data_store
    bucket_id = bds.columnHeaders.identifier if columns else bds.rowHeaders.buckets[0].identifier
    return {h.index: round(h.size) for h in model.objects[bucket_id].headers if h.size}


def header_sizes(table, columns):
    """All row heights (or column widths) in points. Rows Numbers sizes automatically get the height
    Numbers last fitted them to (kept in the table's layout cache), and never less than the default."""
    count = table.num_cols if columns else table.num_rows
    try:
        sizes = stored_sizes(table, columns)
        table_model = table._model.objects[table._table_id]
        default = table_model.default_column_width if columns else table_model.default_row_height
        fitted = {} if columns else fitted_row_heights(table)
        return [round(sizes.get(i, max(default, fitted.get(i, 0))), 1) for i in range(count)]
    except Exception:
        size = table.col_width if columns else table.row_height
        return [size(i) for i in range(count)]


def fitted_row_heights(table):
    """{row: height} Numbers computed to fit each row's content, from the table's layout cache."""
    try:
        info = pivots.find_table_info(table._model, table._table_id)
        entries = info.layout_engine.width_height_cache.rows_fitting_entries
        return {row: entry.fitting_size for row, entry in enumerate(entries) if entry.fitting_size > 0}
    except Exception:
        return {}


def unique_sheet_name(name, used):
    name = re.sub(r"[\[\]:*?/\\]", "-", name).strip("' ") or "Sheet"
    base = name[:31]
    candidate, n = base, 2
    while candidate.lower() in used:
        suffix = f" ({n})"
        candidate = base[: 31 - len(suffix)] + suffix
        n += 1
    used.add(candidate.lower())
    return candidate


def count_drawables(doc):
    counts = {"charts": 0, "images": 0, "shapes": 0, "pivots": 0}
    kinds = {"ChartDrawableArchive": "charts", "ImageArchive": "images", "MovieArchive": "images", "ShapeArchive": "shapes"}
    try:
        for obj in doc._model.objects._objects.values():
            kind = kinds.get(type(obj).__name__)
            if kind:
                counts[kind] += 1
    except Exception:
        pass
    return counts


# ---------------------------------------------------------------- cell helpers


def error_text(cell):
    """Excel's name for a Numbers formula error. numbers-parser doesn't say which error it is, so
    a broken reference shows as #REF! and anything else as #N/A; both spread through other formulas
    the way Numbers' errors do."""
    try:
        return "#REF!" if "#REF!" in (cell.formula or "") else "#N/A"
    except Exception:
        return "#N/A"


def formatted(cell):
    if isinstance(cell, ErrorCell):
        return error_text(cell)
    bullets = extras.bullet_text(cell)
    if bullets:
        return bullets
    try:
        if (isinstance(cell, NumberCell) and isinstance(cell.value, float) and cell._control_id is None
                and cell._num_format_id is None and cell._currency_format_id is None):
            return f"{cell.value:.15g}"  # numbers-parser prints float noise (0.9999000000000001); Numbers doesn't
        text = cell.formatted_value
        # numbers-parser can print scientific notation with hundreds of digits; show what Excel will.
        if len(text) > 40 and re.fullmatch(r"-?\d\.\d+E[+-]\d+", text) and isinstance(cell.value, (int, float)):
            code = excel_number_format(cell) or "0.00E+00"
            places = len(re.search(r"\.(0*)", code).group(1)) if "." in code else 0
            return f"{cell.value:.{places}E}"
        return text
    except Exception:  # some Numbers date formats can't be rendered by Windows' strftime
        value = cell.value
        if isinstance(value, dt.datetime):
            return f"{value.month}/{value.day}/{value.year}" + (
                f" {value:%H:%M}" if (value.hour or value.minute) else ""
            )
        return "" if value is None else str(value)


def rgb_hex(color):
    if isinstance(color, (list, tuple)) and not hasattr(color, "r"):  # gradient fill: use its first colour
        color = color[0] if color else None
    if color is None:
        return None
    return "#{:02X}{:02X}{:02X}".format(color.r, color.g, color.b)


def font_name(name):
    """The font Excel should use: a Windows stand-in for Mac-only fonts. Numbers sometimes names
    fonts by their PostScript name (AvenirNext-DemiBold), whose family Excel knows as "Avenir Next"."""
    if not name:
        return None
    if name in FONT_SUBSTITUTES:
        return FONT_SUBSTITUTES[name]
    if "-" not in name:
        return name
    family = name.split("-")[0]
    spaced = re.sub(r"(?<=[a-z])(?=[A-Z])", " ", family)
    return FONT_SUBSTITUTES.get(family, FONT_SUBSTITUTES.get(spaced, spaced))


def font_is_bold(name):
    """Whether a PostScript font name is a bold weight (AvenirNext-DemiBold, HelveticaNeue-Bold)."""
    return bool(name) and "-" in name and bool(re.search(r"Bold|Heavy|Black|Demi", name.split("-", 1)[1]))


_style_cache = {}


def table_region(cell):
    """Which part of the table's theme styles a cell: header row, footer row, header column or
    body (in that order, so a "Total" cell in the header column takes the footer's look, as in
    Numbers), plus whether it's a shaded band in a banded table."""
    tm = cell._model.objects[cell._table_id]
    if cell.row < tm.number_of_header_rows:
        return "header_row", False
    if tm.number_of_footer_rows and cell.row >= tm.number_of_rows - tm.number_of_footer_rows:
        return "footer_row", False
    if cell.col < tm.number_of_header_columns:
        return "header_column", False
    return "body", (cell.row - tm.number_of_header_rows) % 2 == 1


def theme_fill(cell, region, banded):
    """Background colour a cell gets from its table's theme rather than from its own style."""
    try:
        model = cell._model
        tm = model.objects[cell._table_id]
        if region == "body" and banded:
            style = model.objects[tm.table_style.identifier]
            fill = strokes.style_property(model, style, "banded_fill")
            if strokes.style_property(model, style, "banded_rows") and fill is not None and fill.HasField("color"):
                return rgb_hex(rgb_from_proto(fill.color))
        fill = model.objects[getattr(tm, f"{region}_style" if region != "body" else "body_cell_style").identifier]
        fill = fill.cell_properties.cell_fill
        if fill.HasField("color"):
            return rgb_hex(rgb_from_proto(fill.color))
    except Exception:
        pass
    return None


def is_light(hex_color):
    r, g, b = (int(hex_color[i:i + 2], 16) for i in (1, 3, 5))
    return 0.299 * r + 0.587 * g + 0.114 * b > 200


def rgb_from_proto(color):
    return RGB(round(color.r * 255), round(color.g * 255), round(color.b * 255))


def cell_style(cell):
    """Style dict shared by the preview (as CSS-ish values) and the Excel writer."""
    # Cells without their own styles take them from the table theme, which differs by region.
    region, banded = table_region(cell)
    key = (id(cell._model), cell._table_id, cell._cell_style_id, cell._text_style_id, region, banded)
    s = _style_cache.get(key)
    if s is None:
        s = text_and_fill_style(cell)
        if "bg" not in s:
            bg = theme_fill(cell, region, banded)
            if bg and bg != "#FFFFFF":
                s["bg"] = bg
        if "bg" not in s and (id(cell._model), cell._table_id) in SHEET_BACKGROUNDS:
            s["bg"] = SHEET_BACKGROUNDS[(id(cell._model), cell._table_id)]  # unfilled cells show the page colour
        if "color" in s and is_light(s["color"]) and is_light(s.get("bg", "#FFFFFF")):
            del s["color"]  # light text we couldn't find the dark background for: keep it readable
        _style_cache[key] = s
    s = dict(s)
    s.update(border_style(cell))
    return s


def text_and_fill_style(cell):
    s = {}
    try:
        st = cell.style
    except Exception:
        st = None
    if st is not None:
        if st.bold:
            s["b"] = 1
        if st.italic:
            s["i"] = 1
        if st.underline:
            s["u"] = 1
        if st.strikethrough:
            s["st"] = 1
        color = rgb_hex(st.font_color)
        if color and color != "#000000":
            s["color"] = color
        bg = rgb_hex(st.bg_color)
        if bg:
            s["bg"] = bg
        if st.font_size and abs(st.font_size - 10) > 0.01:
            s["size"] = round(st.font_size, 1)
        name = font_name(st.font_name)
        if name and name != "Arial":
            s["font"] = name
        if st.text_wrap:
            s["wrap"] = 1
        try:
            h = st.alignment.horizontal.name.lower()
            v = st.alignment.vertical.name.lower()
            if h in ("left", "right", "center", "justify", "justified"):
                s["ha"] = "justify" if h.startswith("justif") else h
            if v in ("top", "middle", "bottom"):
                s["va"] = v
        except Exception:
            pass
    corner = footer_corner_text(cell)
    if corner:
        for key, value in corner.items():
            if value:
                s[key] = value
            else:
                s.pop(key, None)
    return s


def footer_corner_text(cell):
    """numbers-parser styles a footer-row cell in the header column like the header column; Numbers
    uses the footer row's text style there ("Total" labels). Returns {"b", "i", "color"} or None."""
    if cell._text_style_id is not None:
        return None
    model = cell._model
    tm = model.objects[cell._table_id]
    in_footer = tm.number_of_footer_rows and cell.row >= tm.number_of_rows - tm.number_of_footer_rows
    if not (in_footer and cell.col < tm.number_of_header_columns and cell.row >= tm.number_of_header_rows):
        return None
    try:
        style = model.objects[tm.footer_row_text_style.identifier]
        colour = rgb_hex(rgb_from_proto(model.char_property(style, "font_color")))
        return {"b": 1 if model.char_property(style, "bold") else 0,
                "i": 1 if model.char_property(style, "italic") else 0,
                "color": colour if colour != "#000000" else None}
    except Exception:
        return None


def border_style(cell):
    """The cell's lines: the table style's lines with any set on individual cells drawn over them."""
    try:
        return strokes.cell_lines(cell)
    except Exception:
        pass
    s = {}  # fall back to the cell-level lines numbers-parser reads
    try:
        border = cell.border
    except Exception:
        border = None
    if border is not None:
        for side, key in (("top", "bt"), ("right", "br"), ("bottom", "bb"), ("left", "bl")):
            seg = getattr(border, side, None)
            if seg is not None and seg.style is not None and seg.style.name != "NONE" and seg.width > 0:
                s[key] = [round(seg.width, 2), rgb_hex(seg.color) or "#000000", seg.style.name.lower()]
    return s


def table_rows(table):
    """The table's rows in the order they're written out. numbers-parser lists rows of a table with
    categories (grouped rows) in grouped display order, but formula references, merges, row sizes
    and hidden rows all use the original row numbers, so such tables are written in original order.
    Pivot tables keep display order: their layout is defined by it and nothing refers to their rows."""
    rows = list(table.iter_rows())
    if table._model.objects[table._table_id].HasField("pivot_owner"):
        return rows
    return sorted(rows, key=lambda row: row[0].row if row else 0)


def is_numeric(cell):
    return isinstance(cell, (NumberCell, DateCell, DurationCell))


# ---------------------------------------------------------------- preview


def preview(path, password=None, layout_options=None):
    """Each worksheet the file becomes, laid out as it will be saved (see layout.Layout): its cells
    in a grid, with every table placed where it goes and its title in a merged cell above it.
    Cells are sent as plain text, [text, style index], or [text, style index, extras] where extras
    may hold "runs" ([[text, style index], ...]), "link", "note" and "control"; None marks a cell
    covered by a merged one. Each sheet also says which Numbers sheet it shows ("numbersSheet"),
    how many tables that has ("tableCount") and how it's laid out ("layout"), and lists its charts
    (see charts.describe)."""
    book = load(path, password)
    layout = book.layout(layout_options)
    return {"sheets": [preview_sheet(book, ws) for ws in layout.sheets], "unsupported": book.unsupported}


def preview_sheet(book, ws):
    total_rows, total_cols = ws.size()
    rows_shown, cols_shown = min(total_rows, PREVIEW_MAX_ROWS), min(total_cols, PREVIEW_MAX_COLS)
    grid = [[""] * cols_shown for _ in range(rows_shown)]
    styles, style_ids = [{}], {"{}": 0}  # index 0 = unstyled; such cells are sent as plain strings
    merges = []

    def style_index(st):
        key = json.dumps(st, sort_keys=True)
        if key not in style_ids:
            style_ids[key] = len(styles)
            styles.append(st)
        return style_ids[key]

    def merge(r1, c1, r2, c2):
        if r1 < rows_shown and c1 < cols_shown:
            r2, c2 = min(r2, rows_shown - 1), min(c2, cols_shown - 1)
            merges.append([r1, c1, r2, c2])
            for r in range(r1, r2 + 1):
                for c in range(c1, c2 + 1):
                    if (r, c) != (r1, c1):
                        grid[r][c] = None

    for placed in ws.tables:
        table = placed.table
        title = book.titles[table._table_id]
        if title and placed.title_row < rows_shown and placed.col < cols_shown:
            background = SHEET_BACKGROUNDS.get((id(table._model), table._table_id))
            st = {"font": title["font"], "size": title["size"], "b": 1 if title["bold"] else 0,
                  "i": 1 if title["italic"] else 0, "color": title["color"], "bg": background, "ha": "center", "va": "middle"}
            grid[placed.title_row][placed.col] = [title["text"], style_index({k: v for k, v in st.items() if v})]
            merge(placed.title_row, placed.col, placed.title_row, placed.col + table.num_cols - 1)
        computed = book.pivot_values.get(table._table_id, {})
        for r, source_row in enumerate(table_rows(table)):
            row = placed.row + r
            if row >= rows_shown:
                break
            for c, cell in enumerate(source_row):
                col = placed.col + c
                if col >= cols_shown:
                    break
                if isinstance(cell, MergedCell):
                    grid[row][col] = None
                    continue
                st = cell_style(cell)
                if (is_numeric(cell) or (r, c) in computed) and "ha" not in st:
                    st["ha"] = "right"
                index = style_index(st)
                if (r, c) in computed:
                    value, source = computed[(r, c)]
                    text = format_number(value, source and excel_number_format(source))
                else:
                    text = "" if isinstance(cell, EmptyCell) else formatted(cell)
                info = cell_extras(cell)
                if "runs" in info:
                    info["runs"] = [[piece, style_index(run_style)] for piece, run_style in info["runs"]]
                grid[row][col] = [text, index, info] if info else ([text, index] if index else text)
        for rng in table.merge_ranges:
            r1, c1, r2, c2 = parse_range(rng)
            merge(placed.row + r1, placed.col + c1, placed.row + r2, placed.col + c2)

    return {
        "name": ws.name,
        "numbersSheet": ws.numbers_sheet,
        "tableCount": ws.table_count,
        "layout": ws.mode,
        "charts": [charts.describe(chart, position) for chart, position in ws.charts],
        "rows": grid,
        "styles": styles,
        "merges": merges,
        "colWidths": [ws.cols.get(c, 48.0) for c in range(cols_shown)],
        "rowHeights": [ws.rows.get(r, 15.0) for r in range(rows_shown)],
        "hiddenRows": sorted(r for r in ws.hidden_rows if r < rows_shown),
        "hiddenCols": sorted(c for c in ws.hidden_cols if c < cols_shown),
        "totalRows": total_rows,
        "totalCols": total_cols,
    }


def cell_extras(cell):
    """Preview extras for a cell (see preview); run styles are style dicts until indexed."""
    info = {}
    runs = extras.rich_runs(cell)
    if runs:
        base = {k: v for k, v in cell_style(cell).items() if k in FONT_KEYS}
        info["runs"] = [(piece, {**base, **style}) for piece, style in runs]
    links = extras.links(cell)
    if links:
        info["link"] = links[0][1]
    note = extras.comment(cell)
    if note:
        info["note"] = note[0] + (f"\n— {note[1]}" if note[1] else "")
    control = extras.control(cell)
    if control and control["kind"] == "popup":
        info["control"] = "popup"
    return info


FONT_KEYS = ("b", "i", "u", "st", "color", "size", "font")


def format_number(value, code):
    """Preview text for a computed number using the simple Excel format codes this tool writes."""
    if value is None:
        return ""
    head = (code or "").split(";")[0]  # positive-number section
    if "%" in head:
        value *= 100
    sep = "," if "," in head else ""
    decimals = re.search(r"\.(0+)", head)
    if head:
        places = len(decimals.group(1)) if decimals else 0
        # Excel rounds halves away from zero; Python's format() rounds them to even.
        rounded = Decimal(repr(abs(value))).quantize(Decimal(1).scaleb(-places), ROUND_HALF_UP)
        digits = f"{rounded:{sep}f}"
    else:  # General
        digits = f"{abs(value):{sep}.10g}"
    symbol = re.search(r'"([^"]*)"', head)
    return ("-" if value < 0 else "") + (symbol.group(1) if symbol else "") + digits + ("%" if "%" in head else "")


def parse_range(rng):
    a, _, b = rng.partition(":")
    r1, c1 = parse_ref(a)
    r2, c2 = parse_ref(b or a)
    return r1, c1, r2, c2


def parse_ref(ref):
    m = re.fullmatch(r"\$?([A-Z]+)\$?(\d+)", ref.strip())
    col = 0
    for ch in m.group(1):
        col = col * 26 + ord(ch) - 64
    return int(m.group(2)) - 1, col - 1


# ---------------------------------------------------------------- number formats


def excel_number_format(cell):
    """Best-effort Excel format code for a cell's Numbers data format, or None for General."""
    model, table_id = cell._model, cell._table_id
    try:
        if isinstance(cell, DateCell) and cell._date_format_id is not None:
            fmt = model.table_format(table_id, cell._date_format_id)
            pattern = fmt.date_time_format
            if fmt.HasField("custom_uid"):
                custom = model.custom_format_map()[NumbersUUID(fmt.custom_uid).hex].default_format
                pattern = custom.custom_format_string
            return icu_to_excel_date(pattern)
        if isinstance(cell, DateCell):
            return "m/d/yyyy"
        if isinstance(cell, DurationCell):
            fmt = DEFAULT_DURATION_FORMAT
            if cell._duration_format_id is not None:
                fmt = model.table_format(table_id, cell._duration_format_id)
            return duration_format(fmt, cell.value.total_seconds())
        if not isinstance(cell, NumberCell):
            return None
        fmt_id = cell._currency_format_id if cell._currency_format_id is not None else cell._num_format_id
        if fmt_id is None:
            return None
        fmt = model.table_format(table_id, fmt_id)
        if fmt.HasField("custom_uid"):
            custom = model.custom_format_map()[NumbersUUID(fmt.custom_uid).hex].default_format
            code = custom.custom_format_string
            # Only pass through simple numeric patterns Excel is sure to accept.
            if code and re.fullmatch(r"[#0?,.%E+\- ()$€£¥]+", code):
                # "0.#" shows "12570." in Excel for whole numbers, where Numbers shows "12570".
                if float(cell.value).is_integer() and re.search(r"\.#+(?![0#])", code):
                    code = re.sub(r"\.#+(?![0#])", "", code)
                return code
            return None
        return builtin_number_format(fmt, cell.value)
    except Exception:
        return None


# How Numbers shows a duration that has no format of its own: "1w 3d", "2h 30m", ...
DEFAULT_DURATION_FORMAT = SimpleNamespace(
    duration_style=DurationStyle.SHORT, use_automatic_duration_units=True,
    duration_unit_largest=DurationUnits.WEEK, duration_unit_smallest=DurationUnits.SECOND,
)

DURATION_UNITS = [  # Numbers unit, seconds, short label, long label
    (DurationUnits.DAY, 86400, "d", "day"),
    (DurationUnits.HOUR, 3600, "h", "hour"),
    (DurationUnits.MINUTE, 60, "m", "minute"),
    (DurationUnits.SECOND, 1, "s", "second"),
]


def duration_format(fmt, seconds):
    """Excel format code showing a duration the way Numbers does for this value: the same units,
    as 2:04:20 (compact), 2d 4h 20m (short) or 2 days 4 hours 20 minutes (long). Weeks, and days
    beyond 31, are folded into hours since Excel can't show more days than that."""
    largest, smallest = fmt.duration_unit_largest, fmt.duration_unit_smallest
    if fmt.use_automatic_duration_units:
        smallest, largest = _auto_units(seconds, fmt)
    millis = smallest >= DurationUnits.MILLISECOND
    if largest <= DurationUnits.DAY and smallest <= DurationUnits.DAY:  # whole days only: a plain day count
        if fmt.duration_style == DurationStyle.SHORT:
            return '0"d"'
        if fmt.duration_style == DurationStyle.LONG:
            return '0" day"' if round(seconds) == 86400 else '0" days"'
    if largest <= DurationUnits.DAY:  # Excel has no weeks, and its day count stops at 31
        largest = DurationUnits.DAY if abs(seconds) < 32 * 86400 else DurationUnits.HOUR
    smallest = max(min(smallest, DurationUnits.SECOND), largest)
    units = [u for u in DURATION_UNITS if largest <= u[0] <= smallest]
    if not units:
        return "[h]:mm:ss"
    tokens = {DurationUnits.DAY: "d", DurationUnits.HOUR: "h", DurationUnits.MINUTE: "m", DurationUnits.SECOND: "s"}
    remaining = abs(seconds)
    parts = []
    for i, (unit, size, short, long) in enumerate(units):
        amount = int(remaining // size)
        remaining -= amount * size
        token = tokens[unit]
        if i == 0 and unit != DurationUnits.DAY:
            token = f"[{token}]"  # total hours/minutes/seconds, not the remainder
        elif fmt.duration_style == DurationStyle.COMPACT and i > 0:
            token = token * 2
        if unit == DurationUnits.SECOND and millis:
            token += ".000"
        if fmt.duration_style == DurationStyle.COMPACT:
            parts.append(token)
        elif fmt.duration_style == DurationStyle.SHORT:
            parts.append(f'{token}"{short}"')
        else:
            parts.append(f'{token}" {long}{"" if amount == 1 else "s"}"')
    return (":" if fmt.duration_style == DurationStyle.COMPACT else " ").join(parts)


def auto_places(value):
    """Decimal places Numbers' "automatic" setting shows for this value (Excel has no equivalent
    once thousands separators, currency or percent are involved)."""
    if not isinstance(value, (int, float)) or value == 0:
        return 0
    return max(0, -Decimal(f"{abs(value):.15g}").normalize().as_tuple().exponent)


# Numbers' fraction accuracy -> Excel fraction format. Positive values are fixed denominators;
# 0xFFFFFFFF/FE/FD mean "up to 1/2/3 digits".
FRACTION_FORMATS = {2: "# ?/2", 4: "# ?/4", 8: "# ?/8", 16: "# ??/16", 10: "# ?/10", 100: "# ??/100",
                    0xFFFFFFFF: "# ?/?", 0xFFFFFFFE: "# ??/??", 0xFFFFFFFD: "# ???/???"}


def builtin_number_format(fmt, value=None):
    t = fmt.format_type
    places = fmt.decimal_places
    if places == DECIMAL_PLACES_AUTO:
        if t == FormatType.SCIENTIFIC and isinstance(value, (int, float)) and value:
            places = auto_places(float(f"{abs(value):.15e}".split("e")[0]))  # digits of the mantissa
        else:
            places = auto_places(value * 100 if t == FormatType.PERCENT and isinstance(value, (int, float)) else value)
    decimals = "." + "0" * min(places, 15) if places else ""
    if t == FormatType.PERCENT:
        return f"0{decimals}%"
    if t == FormatType.SCIENTIFIC:
        return f"0{decimals}E+00"
    if t == FormatType.FRACTION:
        return FRACTION_FORMATS.get(fmt.fraction_accuracy, "# ?/?")
    if t not in (FormatType.DECIMAL, FormatType.CURRENCY):
        return None
    if t == FormatType.DECIMAL and fmt.decimal_places == DECIMAL_PLACES_AUTO and not fmt.show_thousands_separator:
        return None  # General already shows exactly the digits needed
    body = ("#,##0" if fmt.show_thousands_separator or t == FormatType.CURRENCY else "0") + decimals
    if t == FormatType.CURRENCY:
        symbol = CURRENCY_SYMBOLS.get(fmt.currency_code, fmt.currency_code + " ")
        # Accounting style lines the symbol up at the left edge of the cell, as Numbers does.
        body = f'"{symbol}"* {body}' if fmt.use_accounting_style else f'"{symbol}"{body}'
    negative = {0: f"-{body}", 1: f"[Red]{body}", 2: f"({body})", 3: f"[Red]({body})"}.get(fmt.negative_style, f"-{body}")
    if fmt.negative_style in (2, 3):
        body += "_)"  # leave room for the ")" so positive and negative numbers line up, as Numbers does
    return f"{body};{negative}" if fmt.negative_style else body


ICU_TO_EXCEL = [
    ("EEEE", "dddd"), ("EEE", "ddd"), ("E", "ddd"),
    ("MMMMM", "mmmmm"), ("MMMM", "mmmm"), ("MMM", "mmm"), ("MM", "mm"), ("M", "m"),
    ("LLLL", "mmmm"), ("LLL", "mmm"), ("LL", "mm"), ("L", "m"),
    ("dd", "dd"), ("d", "d"),
    ("yyyy", "yyyy"), ("yy", "yy"), ("y", "yyyy"),
    ("HH", "hh"), ("H", "h"), ("hh", "hh"), ("h", "h"), ("kk", "hh"), ("k", "h"),
    ("mm", "mm"), ("m", "m"), ("ss", "ss"), ("s", "s"),
    ("a", "AM/PM"),
]


def icu_to_excel_date(pattern):
    """Numbers (ICU) date pattern -> Excel format code. Patterns using fields Excel can't show
    (day of year, week numbers, eras, ...) fall back to a plain date/time format."""
    fallback = "m/d/yyyy h:mm" if re.search(r"[HhKkm]", pattern or "") else "m/d/yyyy"
    if not pattern:
        return fallback
    out, literal, i = [], "", 0

    def flush():
        nonlocal literal
        literal = literal.replace('"', "")  # Excel can't escape quotes inside a quoted literal
        if literal:
            out.append(f'"{literal}"')
        literal = ""

    while i < len(pattern):
        ch = pattern[i]
        if ch == "'":
            end = pattern.find("'", i + 1)
            end = len(pattern) if end == -1 else end
            literal += pattern[i + 1:end] if end > i + 1 else "'"
            i = end + 1
        elif ch.isalpha():
            if ch == "S":  # fractional seconds; Excel only allows them straight after "s."
                run = len(re.match(r"S+", pattern[i:]).group())
                if not (out and out[-1] == "." and len(out) > 1 and out[-2] in ("s", "ss")) or literal:
                    return fallback
                flush()
                out.append("0" * min(run, 3))
                i += run
                continue
            for icu, excel in ICU_TO_EXCEL:
                if pattern.startswith(icu, i):
                    flush()
                    out.append(excel)
                    i += len(icu)
                    break
            else:
                return fallback
        elif ch in " ,./-:":
            flush()
            out.append(ch)
            i += 1
        else:
            literal += ch
            i += 1
    flush()
    return "".join(out)


# ---------------------------------------------------------------- formulas

# Whole-row/column references (A:A, 3:3) are left out on purpose: Numbers and Excel count header
# cells in them differently, so they'd silently give different results.
CELLS = r"""(?:\$?[A-Z]{1,3}\$?\d+(?::\$?[A-Z]{1,3}\$?\d+)?)"""
NAME = r"""(?:'[^']+'|[^\s:(),'"=<>+\-*/^&{}×÷≠≤≥][^:(),'"=<>+\-*/^&{}×÷≠≤≥]*?)"""
EXACT = re.compile(r"⟦(\d+)\|(\d*)\|(\d*)\|(\d*)\|(\d*)\|(\d+)⟧")
TOKEN = re.compile(
    rf"""
    (?P<str>"(?:[^"]|"")*")
  | (?P<exact>⟦[^⟧]*⟧)
  | (?P<xref>(?:{NAME}::){{1,2}}{CELLS}(?![A-Za-z0-9_:]))
  | (?P<func>[A-Z][A-Z0-9.]*)\s*\(
  | (?P<ref>{CELLS})(?![A-Za-z0-9_(:])
  | (?P<bool>TRUE|FALSE)\b
  | (?P<num>\d+(?:\.\d+)?(?:[Ee][+-]?\d+)?%?)
  | (?P<op><>|<=|>=|[-+*/^&=<>(),%])
  | (?P<ws>\s+)
  | (?P<other>.)
    """,
    re.VERBOSE,
)

OPERATOR_MAP = {"÷": "/", "×": "*", "≠": "<>", "≤": "<=", "≥": ">="}


def referenced_cells(layout):
    """{worksheet name: [(r1, c1, r2, c2)]} of every range a formula in the file reads, in sheet
    coordinates (0-based)."""
    refs = {}
    for sheet_name, table in layout.book.tables:
        excel_name = layout.place[table._table_id].sheet
        for row in table.iter_rows():
            for cell in row:
                if not cell.is_formula:
                    continue
                excel = excel_formula(layout, sheet_name, cell)
                for target, rng in re.findall(r"(?:'((?:[^']|'')+)'!)?(\$?[A-Z]{1,3}\$?\d+(?::\$?[A-Z]{1,3}\$?\d+)?)", excel or ""):
                    name = target.replace("''", "'") if target else excel_name
                    refs.setdefault(name, []).append(parse_range(rng.replace("$", "")))
    return refs


def is_referenced(refs, r, c):
    return any(r1 <= r <= r2 and c1 <= c <= c2 for r1, c1, r2, c2 in refs)


def excel_formula(layout, sheet_name, cell):
    """The cell's formula in Excel form (None if it has to stay a value), translated once per cell
    for each layout."""
    key = (cell._table_id, cell.row, cell.col)
    if key not in layout.formulas:
        try:
            # Numbers returns durations from date maths where Excel returns plain numbers (e.g.
            # NETWORKDAYS), which the duration format would then show as hours, so those keep values.
            layout.formulas[key] = (None if isinstance(cell, DurationCell)
                                    else translate_formula(cell.formula, layout, sheet_name, cell))
        except Exception:
            layout.formulas[key] = None
    return layout.formulas[key]


def shift_rows(cells, offset, columns=0):
    """Move every cell in a reference like B2 or $A$3:C9 down by `offset` rows and right by `columns`."""
    if not offset and not columns:
        return cells

    def move(m):
        dollar, letters, row_dollar, row = m.groups()
        col = 0
        for ch in letters:
            col = col * 26 + ord(ch) - 64
        return f"{dollar}{column_letters(col - 1 + columns)}{row_dollar}{int(row) + offset}"
    return re.sub(r"(\$?)([A-Z]{1,3})(\$?)(\d+)", move, cells)


def column_letters(index):
    letters = ""
    index += 1
    while index:
        index, rem = divmod(index - 1, 26)
        letters = chr(65 + rem) + letters
    return letters


def merge_exact_ranges(formula):
    """numbers-parser joins the two ends of a range with ":" (⟦a⟧:⟦b⟧); fold them into one code."""
    def merge(m):
        a, b = EXACT.fullmatch(m.group(1)).groups(), EXACT.fullmatch(m.group(2)).groups()
        if a[0] != b[0]:
            return m.group(0)  # a range across tables isn't something Excel can express
        rows = sorted(int(v) for v in (a[1], b[1]) if v != "")
        cols = sorted(int(v) for v in (a[3], b[3]) if v != "")
        if (a[1] == "") != (b[1] == "") or (a[3] == "") != (b[3] == ""):
            return m.group(0)
        flags = (int(a[5]) & 3) | ((int(b[5]) & 3) << 2)
        part = lambda values, i: str(values[i]) if values else ""  # noqa: E731
        return f"⟦{a[0]}|{part(rows, 0)}|{part(rows, -1)}|{part(cols, 0)}|{part(cols, -1)}|{flags}⟧"
    return re.sub(r"(⟦[^⟧]*⟧):(⟦[^⟧]*⟧)", merge, formula)


def exact_reference(code, layout, own_table, cell, as_range):
    """Excel reference for an exact reference code, or None. Whole columns (header names like
    Amount, or column letters) mean the column's body cells, without header and footer rows, when
    they're a whole function argument (SUM(Amount)); anywhere else Numbers uses the cell in the
    formula's own row (Price×Quantity), which only makes sense within the formula's own table.
    Whole rows work the same way across the body columns. Cells are placed where their table sits
    on its worksheet, which is named when it isn't the formula's own."""
    table_id, r1, r2, c1, c2, flags = EXACT.fullmatch(code).groups()
    table_id = int(table_id)
    placed = layout.place.get(table_id)
    if placed is None:
        return None
    table = placed.table
    tm = table._model.objects[table_id]
    if tm.HasField("pivot_owner"):
        return None  # pivot cells are laid out in display order, not the order references use
    own = layout.place.get(own_table)
    top, left = placed.row, placed.col
    flags = int(flags)
    prefix = "" if own is not None and own.sheet == placed.sheet else "'" + placed.sheet.replace("'", "''") + "'!"
    row_abs, col_abs, row2_abs, col2_abs = ("$" if flags >> i & 1 else "" for i in range(4))
    r1, r2, c1, c2 = (int(v) if v != "" else None for v in (r1, r2, c1, c2))
    r2 = r1 if r2 is None else r2
    c2 = c1 if c2 is None else c2

    def letters(c):
        return column_letters(c + left)

    if r1 is not None and c1 is not None:
        start = f"{col_abs}{letters(c1)}{row_abs}{r1 + top + 1}"
        end = f"{col2_abs}{letters(c2)}{row2_abs}{r2 + top + 1}"
        return prefix + (start if (r1, c1) == (r2, c2) else f"{start}:{end}")
    if c1 is not None:  # whole column(s)
        first, last = tm.number_of_header_rows, table.num_rows - tm.number_of_footer_rows - 1
        if as_range and last >= first:
            return (f"{prefix}{col_abs}{letters(c1)}{first + top + 1}:"
                    f"{col2_abs or col_abs}{letters(c2)}{last + top + 1}")
        if not as_range and c1 == c2 and table_id == own_table and cell is not None and first <= cell.row <= last:
            return f"{col_abs}{letters(c1)}{cell.row + top + 1}"
        return None
    if r1 is not None:  # whole row(s)
        first, last = tm.number_of_header_columns, table.num_cols - 1
        if as_range and last >= first:
            return f"{prefix}{letters(first)}{row_abs}{r1 + top + 1}:{letters(last)}{row2_abs or row_abs}{r2 + top + 1}"
        if not as_range and r1 == r2 and table_id == own_table and cell is not None and first <= cell.col <= last:
            return f"{letters(cell.col)}{row_abs}{r1 + top + 1}"
    return None


# Lookups that can do an approximate match: (argument with the values searched, argument choosing
# the match type, or None if always approximate).
LOOKUP_ARGUMENTS = {"MATCH": (1, 2), "LOOKUP": (1, None), "VLOOKUP": (1, 3), "HLOOKUP": (1, 3)}


def lookup_is_safe(name, args, book):
    """An approximate-match lookup gives the same answer in Excel only if the values it searches are
    sorted: Numbers finds the right value regardless, Excel's search assumes sorted data. Exact
    matches are always safe. Anything we can't check counts as unsafe."""
    values_arg, mode_arg = LOOKUP_ARGUMENTS[name]
    meaningful = lambda tokens: [t for t in tokens if t[0] != "ws"]  # noqa: E731
    mode = meaningful(args[mode_arg]) if mode_arg is not None and mode_arg < len(args) else []
    mode_text = "".join(t[1] for t in mode).upper()
    if mode_text in ("0", "FALSE", "FALSE()"):
        return True
    if mode_text not in ("", "1", "TRUE", "TRUE()", "-1"):
        return False
    searched = meaningful(args[values_arg]) if values_arg < len(args) else []
    if len(searched) != 1 or searched[0][0] != "exact":
        return False
    values = searched_values(searched[0][1], book, first_column=name == "VLOOKUP", first_row=name == "HLOOKUP")
    if values is None:
        return False
    while values and values[-1] is None:  # trailing blank rows are harmless
        values.pop()
    if any(v is None for v in values):
        return False
    keys = [excel_sort_key(v) for v in values]
    if mode_text == "-1":
        return all(a >= b for a, b in zip(keys, keys[1:]))
    return all(a <= b for a, b in zip(keys, keys[1:]))


def searched_values(code, book, first_column=False, first_row=False):
    """Cell values an exact reference covers (whole columns/rows meaning their body cells)."""
    table_id, r1, r2, c1, c2, _ = EXACT.fullmatch(code).groups()
    table = book.table_by_id.get(int(table_id))
    if table is None:
        return None
    tm = table._model.objects[table._table_id]
    rows = table_rows(table)
    if r1 == "":  # whole column(s): the body rows
        r1, r2 = tm.number_of_header_rows, table.num_rows - tm.number_of_footer_rows - 1
    else:
        r1 = int(r1)
        r2 = int(r2) if r2 else r1
    if c1 == "":  # whole row(s): the body columns
        c1, c2 = tm.number_of_header_columns, table.num_cols - 1
    else:
        c1 = int(c1)
        c2 = int(c2) if c2 else c1
    if first_column:
        c2 = c1
    if first_row:
        r2 = r1
    if r1 != r2 and c1 != c2:
        return None  # a block, not a single row or column
    try:
        return [rows[r][c].value for r in range(r1, r2 + 1) for c in range(c1, c2 + 1)]
    except IndexError:
        return None


def excel_sort_key(value):
    """How Excel orders values when it searches: numbers and dates, then text, then TRUE/FALSE."""
    if isinstance(value, bool):
        return (2, value)
    if isinstance(value, (int, float)):
        return (0, float(value))
    if isinstance(value, (dt.datetime, dt.timedelta)):
        return (0, cached_value_of(value))
    return (1, str(value).casefold())


def cached_value_of(value):
    if isinstance(value, dt.timedelta):
        return value.total_seconds() / 86400
    delta = naive(value) - EXCEL_EPOCH
    return delta.days + delta.seconds / 86400


def translate_formula(formula, layout, sheet_name, cell=None):
    """Numbers formula text -> Excel formula, or None if it can't be represented safely. `cell` is
    the formula's own cell; references are placed where their tables sit (layout.place)."""
    if not formula or not formula.strip():
        return None  # e.g. pivot table summaries, which numbers-parser can't decode
    place = getattr(layout, "place", {})
    own_table = cell._table_id if cell is not None else None

    def offsets(table_id):
        placed = place.get(table_id)
        return (placed.row, placed.col) if placed else (0, 0)

    out = []
    # One frame per open parenthesis: [function name or None, commas seen, any argument text seen,
    # the tokens of each argument, where the function's name is in `out`].
    stack = []
    tokens = [(m.lastgroup, m.group()) for m in TOKEN.finditer(merge_exact_ranges(formula))]
    significant = [i for i, (kind, _) in enumerate(tokens) if kind != "ws"]
    for i, (kind, text) in enumerate(tokens):
        if stack and kind != "ws" and text not in (",", ")"):
            stack[-1][2] = True
        if stack and text not in (",", ")"):
            stack[-1][3][-1].append((kind, text))
        if kind == "op" and text == "(":
            stack.append([None, 0, False, [[]], None])
        elif kind == "op" and text == ",":
            if not stack:
                return None
            stack[-1][1] += 1
            stack[-1][3].append([])
        elif kind == "op" and text == ")":
            if not stack:
                return None
            name, commas, seen, args, at = stack.pop()
            if not name and commas:
                return None  # Numbers' (A1,B2) list syntax; Excel reads it as a union and gives #VALUE!
            if name:
                count = commas + 1 if (commas or seen) else 0
                low, high = EXCEL_FUNCTIONS[name]
                if count < low or (high is not None and count > high):
                    return None
                if name in LOOKUP_ARGUMENTS and not lookup_is_safe(name, args, layout.book):
                    return None
                if name in ("ROW", "COLUMN"):
                    # Numbers counts rows (columns) within a table, Excel across the sheet, where
                    # the table may start further down (right): its title row, other tables.
                    target = next((int(EXACT.fullmatch(t).group(1)) for k, t in args[0] if k == "exact"), own_table)
                    shift = offsets(target)[0 if name == "ROW" else 1]
                    if shift:
                        out[at] = "(" + out[at]
                        out.append(f")-{shift})")
                        continue
        if kind == "str":
            if re.search(r"\d{1,4}[/.-]\d{1,2}[/.-]\d{1,4}", text):
                return None  # dates written as text are read using the computer's region settings
            out.append(text)
        elif kind == "xref":
            *names, cells = [p.strip("'") for p in text.split("::")]
            if len(names) == 2:
                target = layout.by_table.get((names[0], names[1]))
            else:
                target = (
                    layout.by_table.get((sheet_name, names[0]))
                    or layout.by_table.get((None, names[0]))
                    or layout.by_sheet.get(names[0])
                )
            if target is None or target not in place:
                return None
            out.append("'" + place[target].sheet.replace("'", "''") + "'!" + shift_rows(cells, *offsets(target)))
        elif kind == "ref":
            out.append(shift_rows(text, *offsets(own_table)))
        elif kind == "func":
            name = text.rstrip("( \t")
            if name not in EXCEL_FUNCTIONS:
                return None
            stack.append([name, 0, False, [[]], len(out)])
            out.append(name + "(")
        elif kind in ("bool", "num", "op", "ws"):
            out.append(text)
        elif text in OPERATOR_MAP:
            out.append(OPERATOR_MAP[text])
        elif kind == "exact":
            position = significant.index(i)
            before = tokens[significant[position - 1]] if position > 0 else (None, None)
            after = tokens[significant[position + 1]][1] if position + 1 < len(significant) else None
            opens_argument = before[0] == "func" or before[1] in ("(", ",")  # "SUM(" arrives as one token
            as_range = bool(stack and stack[-1][0]) and opens_argument and after in (",", ")")
            excel = exact_reference(text, layout, own_table, cell, as_range) if place else None
            if excel is None:
                return None
            out.append(excel)
        else:
            return None  # array constants, ranges across tables, unknown syntax, etc.
    if stack:
        return None
    return "=" + "".join(out)


# ---------------------------------------------------------------- convert


class FormatCache:
    def __init__(self, workbook):
        self.workbook = workbook
        self.cache = {}

    def get(self, style, num_format):
        key = (json.dumps(style, sort_keys=True), num_format)
        if key in self.cache:
            return self.cache[key]
        props = {"font_name": style.get("font", "Arial"), "font_size": style.get("size", 10), "valign": "vcenter"}
        if style.get("b"):
            props["bold"] = True
        if style.get("i"):
            props["italic"] = True
        if style.get("u"):
            props["underline"] = 1
        if style.get("st"):
            props["font_strikeout"] = True
        if style.get("color"):
            props["font_color"] = style["color"]
        if style.get("bg"):
            props["bg_color"] = style["bg"]
            props["pattern"] = 1
        if style.get("wrap"):
            props["text_wrap"] = True
        if style.get("ha"):
            props["align"] = style["ha"]
        if style.get("va"):
            props["valign"] = {"top": "top", "middle": "vcenter", "bottom": "bottom"}[style["va"]]
        for key_, side in (("bt", "top"), ("br", "right"), ("bb", "bottom"), ("bl", "left")):
            if key_ in style:
                width, color, kind = style[key_]
                props[side] = excel_border(width, kind)
                props[f"{side}_color"] = color
        if num_format:
            props["num_format"] = num_format
        fmt = self.workbook.add_format(props)
        self.cache[key] = fmt
        return fmt


def excel_border(width, kind):
    if kind == "dashes":
        return 8 if width >= 1.5 else 3  # medium dashed / dashed
    if kind == "dots":
        return 4
    if width >= 2.5:
        return 5  # thick
    if width >= 1.5:
        return 2  # medium
    return 1  # thin


def naive(value):
    return value.replace(tzinfo=None) if value.tzinfo else value


class OffsetSheet:
    """A worksheet seen from a table placed `top` rows down and `left` columns across it: calls that
    take a cell or a range are shifted, so the table can be written in its own coordinates."""

    CELL_ARGS = {name: ((0,), (1,)) for name in (
        "write_blank", "write_string", "write_number", "write_boolean", "write_datetime", "write_formula",
        "write_url", "write_rich_string", "write_comment", "insert_checkbox")}
    CELL_ARGS.update({name: ((0, 2), (1, 3)) for name in ("merge_range", "data_validation")})

    def __init__(self, ws, top, left=0):
        self.ws, self.top, self.left = ws, top, left

    def __getattr__(self, name):
        method = getattr(self.ws, name)
        positions = self.CELL_ARGS.get(name)
        if not positions or not (self.top or self.left):
            return method

        def shifted(*args, **kwargs):
            args = list(args)
            for indexes, offset in zip(positions, (self.top, self.left)):
                for p in indexes:
                    if p < len(args) and isinstance(args[p], int):
                        args[p] += offset
            return method(*args, **kwargs)
        return shifted


def title_height(title):
    """Height of a table's title row: the line height of its font plus Numbers' spacing (Apple's
    export: 27.5pt for 12pt Helvetica, 30pt for 12pt Avenir)."""
    return round(title["size"] * (2.5 if title["font"] == "Segoe UI" else 2.3) * 2) / 2


def write_title(ws, row, col, title, table, formats, background):
    """The table's name in a row above it, merged across the table, as Numbers shows it."""
    style = {"font": title["font"], "size": title["size"], "ha": "center", "va": "middle",
             "b": 1 if title["bold"] else 0, "i": 1 if title["italic"] else 0,
             "color": title["color"], "bg": background}
    fmt = formats.get({k: v for k, v in style.items() if v}, None)
    if table.num_cols > 1:
        ws.merge_range(row, col, row, col + table.num_cols - 1, title["text"], fmt)
    else:
        ws.write_string(row, col, title["text"], fmt)


def convert(path, out, password=None, layout_options=None):
    book = load(path, password)
    layout = book.layout(layout_options)
    workbook = xlsxwriter.Workbook(out, {"use_future_functions": True, "strings_to_numbers": False,
                                         "strings_to_formulas": False, "strings_to_urls": False})
    formats = FormatCache(workbook)
    lists = ListSheet(workbook)
    stats = {"formulas": 0, "formulasAsValues": 0}
    refs = referenced_cells(layout)
    worksheets = {}

    for plan in layout.sheets:
        sheet = worksheets[plan.name] = workbook.add_worksheet(plan.name)
        sheet.hide_gridlines(2)  # Numbers shows no grid outside a table's own lines, which are drawn as borders
        # Sizes stay in points, like the fonts, so text fits its cells as it does in Numbers.
        for c, width in sorted(plan.cols.items()):
            sheet.set_column_pixels(c, c, max(8, round(width * 4 / 3)), None, {"hidden": c in plan.hidden_cols})
        for r, height in sorted(plan.rows.items()):
            sheet.set_row(r, height, None, {"hidden": r in plan.hidden_rows})
        if plan.freeze:
            sheet.freeze_panes(*plan.freeze)
        if plan.autofilter:
            sheet.autofilter(*plan.autofilter)
        for placed in plan.tables:
            write_table(sheet, placed, layout, formats, lists, stats, refs.get(plan.name, []))

    def cell_info(cell):  # a table cell's value and number format, for chart data copied to another sheet
        table_id, r, c = cell
        source = book.table_by_id[table_id]._data[r][c]
        value = cached_value(source) if not isinstance(source, (EmptyCell, MergedCell, ErrorCell)) else None
        return value, formats.get({}, excel_number_format(source))

    stats["charts"] = charts.write_charts(workbook, layout, worksheets, cell_info,
                                          lambda name: unique_sheet_name(name, set(layout.used_names)))
    workbook.close()
    return {"outputPath": out, "sheets": len(layout.sheets), "unsupported": book.unsupported, **stats}


def write_table(sheet, placed, layout, formats, lists, stats, sheet_refs):
    table = placed.table
    book = layout.book
    sheet_name = book.sheet_name_of[table._table_id]
    title = book.titles[table._table_id]
    if title:
        write_title(sheet, placed.title_row, placed.col, title, table, formats,
                    SHEET_BACKGROUNDS.get((id(table._model), table._table_id)))
    ws = OffsetSheet(sheet, placed.row, placed.col)
    top, left = placed.row, placed.col

    merged_anchors = {}
    for rng in table.merge_ranges:
        r1, c1, r2, c2 = parse_range(rng)
        merged_anchors[(r1, c1)] = (r2, c2)

    computed = book.pivot_values.get(table._table_id, {})
    for r, row in enumerate(table_rows(table)):
        for c, cell in enumerate(row):
            if isinstance(cell, MergedCell):
                continue
            style = cell_style(cell)
            if (r, c) in computed:  # recalculated pivot summary
                value, source = computed[(r, c)]
                fmt = formats.get(style, source and excel_number_format(source))
                if value is None:
                    ws.write_blank(r, c, None, fmt)
                else:
                    ws.write_number(r, c, value, fmt)
                continue
            num_format = excel_number_format(cell)
            fmt = formats.get(style, num_format)
            if (r, c) in merged_anchors:
                r2, c2 = merged_anchors[(r, c)]
                ws.merge_range(r, c, r2, c2, None, fmt)
            if not write_special(ws, r, c, cell, style, fmt, formats,
                                 lambda r=r, c=c: is_referenced(sheet_refs, r + top, c + left)):
                write_cell(ws, r, c, cell, fmt, layout, sheet_name, stats)
            add_cell_extras(ws, r, c, cell, lists)


def write_special(ws, r, c, cell, style, fmt, formats, referenced=lambda: False):
    """Write cells that need more than a plain value: checkboxes, links and mixed formatting.
    Returns False to let write_cell handle the cell normally. `referenced()` says whether a
    formula reads the cell (only worked out when it matters, as it's slow on big files)."""
    if cell.is_formula:
        return False
    bullets = extras.bullet_text(cell)
    # Bullets are only display formatting in Numbers, so a cell a formula reads keeps its plain text.
    if bullets and not extras.links(cell) and not referenced():
        ws.write_string(r, c, bullets, fmt)
        return True
    control = extras.control(cell)
    if control and control["kind"] == "checkbox" and isinstance(cell.value, bool):
        ws.insert_checkbox(r, c, cell.value, fmt)  # a real checkbox in Excel 365, TRUE/FALSE in older Excel
        return True
    links = extras.links(cell)
    if links:  # Excel allows one link per cell; any others are listed in the cell's note
        ws.write_url(r, c, links[0][1], formats.get({**style, "color": style.get("color", "#0563C1"), "u": 1}, None),
                     string=str(cell.value or links[0][0]))
        return True
    runs = extras.rich_runs(cell)
    if runs and len(runs) > 1:
        base = {k: v for k, v in style.items() if k in FONT_KEYS}
        fragments = []
        for piece, run_style in runs:
            fragments += [formats.get({**base, **run_style}, None), piece]
        try:
            ws.write_rich_string(r, c, *fragments, fmt)
            return True
        except Exception:
            return False
    return False


def add_cell_extras(ws, r, c, cell, lists):
    """Comments, extra links and input rules (pop-up menus, sliders, steppers, ratings)."""
    note = extras.comment(cell)
    links = extras.links(cell)
    text = note[0] if note else ""
    if len(links) > 1:
        text += ("\n\n" if text else "") + "Links:\n" + "\n".join(f"{label}: {url}" for label, url in links)
    if text:
        ws.write_comment(r, c, text, {"author": note[1] if note and note[1] else "", "x_scale": 1.6, "y_scale": 1.4})

    control = extras.control(cell)
    if not control or control["kind"] == "checkbox":
        return
    if control["kind"] == "popup":
        items = [str(i) for i in control["items"]]
        if not items:
            return
        inline = ",".join(items)
        source = items if len(inline) <= 250 and not any("," in i or '"' in i for i in items) else lists.add(items)
        ws.data_validation(r, c, r, c, {"validate": "list", "source": source})
    else:  # stepper, slider, rating: keep the same allowed range
        low, high, step = (float(control[k]) for k in ("min", "max", "step"))
        whole = all(v.is_integer() for v in (low, high, step))
        ws.data_validation(r, c, r, c, {"validate": "integer" if whole else "decimal", "criteria": "between",
                                        "minimum": int(low) if whole else low, "maximum": int(high) if whole else high})


class ListSheet:
    """A hidden sheet holding pop-up menu items too long to write inline in a dropdown rule."""

    def __init__(self, workbook):
        self.workbook = workbook
        self.ws = None
        self.col = 0

    def add(self, items):
        if self.ws is None:
            self.ws = self.workbook.add_worksheet("Pop-up Lists")
            self.ws.hide()
        for i, item in enumerate(items):
            self.ws.write_string(i, self.col, item)
        ref = xl_range_abs(0, self.col, len(items) - 1, self.col)
        self.col += 1
        return f"='Pop-up Lists'!{ref}"


EXCEL_EPOCH = dt.datetime(1899, 12, 30)


def cached_value(cell):
    value = cell.value
    if isinstance(cell, DateCell):
        delta = naive(value) - EXCEL_EPOCH
        return delta.days + delta.seconds / 86400
    if isinstance(cell, DurationCell):
        return value.total_seconds() / 86400
    if isinstance(value, (bool, int, float)):
        return value
    return "" if value is None else str(value)


def write_cell(ws, r, c, cell, fmt, layout, sheet_name, stats):
    if cell.is_formula:
        excel = excel_formula(layout, sheet_name, cell)
        try:
            formula = cell.formula
        except Exception:
            formula = None
        if excel:
            stats["formulas"] += 1
            ws.write_formula(r, c, excel, fmt, cached_value(cell))
            return
        if formula and formula.strip():  # pivot summaries have no readable formula or value; they're reported separately
            stats["formulasAsValues"] += 1

    value = cell.value
    if isinstance(cell, ErrorCell):
        error = error_text(cell)
        ws.write_formula(r, c, "=" + error, fmt, error)
    elif isinstance(cell, EmptyCell) or value is None:
        ws.write_blank(r, c, None, fmt)
    elif isinstance(cell, BoolCell) or isinstance(value, bool):
        ws.write_boolean(r, c, bool(value), fmt)
    elif isinstance(cell, DateCell) and value.year < 1900:
        ws.write_string(r, c, formatted(cell), fmt)  # Excel can't store dates before 1900
    elif isinstance(cell, DateCell):
        ws.write_datetime(r, c, naive(value), fmt)
    elif isinstance(cell, DurationCell) and value.total_seconds() < 0:
        ws.write_string(r, c, formatted(cell), fmt)  # Excel shows negative times as #####
    elif isinstance(cell, DurationCell):
        ws.write_number(r, c, value.total_seconds() / 86400, fmt)
    elif isinstance(value, (int, float)):
        ws.write_number(r, c, value, fmt)
    else:
        ws.write_string(r, c, str(value), fmt)


# ---------------------------------------------------------------- server


def handle(request):
    cmd = request.get("cmd")
    if cmd == "preview":
        return preview(request["path"], request.get("password"), request.get("layout"))
    if cmd == "convert":
        return convert(request["path"], request["out"], request.get("password"), request.get("layout"))
    if cmd == "ping":
        return {}
    raise UserError(f"Unknown command: {cmd}")


def main():
    stdout = open(sys.stdout.fileno(), "w", encoding="utf-8", newline="\n", closefd=False)
    for line in sys.stdin.buffer:
        line = line.strip()
        if not line:
            continue
        req_id = None
        try:
            request = json.loads(line)
            req_id = request.get("id")
            response = {"id": req_id, "ok": True, **handle(request)}
        except UserError as err:
            response = {"id": req_id, "ok": False, "error": str(err), **err.extra}
        except Exception as err:  # never let one bad file kill the helper
            response = {"id": req_id, "ok": False, "error": f"Conversion failed ({type(err).__name__}: {err})."}
        stdout.write(json.dumps(response, ensure_ascii=False) + "\n")
        stdout.flush()


if __name__ == "__main__":
    main()
