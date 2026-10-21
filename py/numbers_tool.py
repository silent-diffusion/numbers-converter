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

import extras
import pivots
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
    def __init__(self, doc):
        self.doc = doc
        self.tables = []  # (excel_name, sheet_name, table)
        used = set()
        for sheet in doc.sheets:
            tables = list(sheet.tables)
            for table in tables:
                label = sheet.name if len(tables) == 1 else f"{sheet.name} - {table.name}"
                self.tables.append((unique_sheet_name(label, used), sheet.name, table))
        if not self.tables:
            raise UserError("No tables found in this file.")
        # Table name lookup for cross-table formula references.
        self.by_table, self.by_sheet = {}, {}
        for excel_name, sheet_name, table in self.tables:
            self.by_table[(sheet_name, table.name)] = excel_name
            self.by_table.setdefault((None, table.name), excel_name)
            self.by_sheet.setdefault(sheet_name, excel_name)
        self.unsupported = count_drawables(doc)
        # Pivot tables whose summary values have to be recalculated: {table_id: {(row, col): (value, source cell)}}
        self.pivot_values = {}
        for _, _, table in self.tables:
            try:
                pivot = pivots.pivot_for(table)
            except Exception:  # pivots.Unsupported, or a pivot laid out in a way we don't recognise
                self.unsupported["pivots"] += 1
                continue
            if pivot is not None:
                self.pivot_values[table._table_id] = pivot.values
                if pivot.missing:
                    self.unsupported["pivots"] += 1


def stored_sizes(table, columns):
    """Row heights (or column widths) in points that Numbers stores explicitly; other rows are
    sized automatically to fit their content. Table.row_height rebuilds its lookup on every call,
    which is quadratic on big tables, so the stored sizes are read once."""
    model, table_id = table._model, table._table_id
    bds = model.objects[table_id].base_data_store
    bucket_id = bds.columnHeaders.identifier if columns else bds.rowHeaders.buckets[0].identifier
    return {h.index: round(h.size) for h in model.objects[bucket_id].headers if h.size}


def header_sizes(table, columns):
    """All row heights (or column widths) in points, with automatic ones at the table default."""
    count = table.num_cols if columns else table.num_rows
    try:
        sizes = stored_sizes(table, columns)
        table_model = table._model.objects[table._table_id]
        default = round(table_model.default_column_width if columns else table_model.default_row_height)
        return [sizes.get(i, default) for i in range(count)]
    except Exception:
        size = table.col_width if columns else table.row_height
        return [size(i) for i in range(count)]


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


def formatted(cell):
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
    if not name:
        return None
    return FONT_SUBSTITUTES.get(name, FONT_SUBSTITUTES.get(name.split("-")[0], name))


_style_cache = {}


def table_region(cell):
    """Which part of the table's theme styles a cell: header row, header column, footer or body
    (same precedence as numbers-parser), plus whether it's a shaded band in a banded table."""
    tm = cell._model.objects[cell._table_id]
    if cell.row < tm.number_of_header_rows:
        return "header_row", False
    if cell.col < tm.number_of_header_columns:
        return "header_column", False
    if tm.number_of_footer_rows and cell.row >= tm.number_of_rows - tm.number_of_footer_rows:
        return "footer_row", False
    return "body", (cell.row - tm.number_of_header_rows) % 2 == 1


def theme_fill(cell, region, banded):
    """Background colour a cell gets from its table's theme rather than from its own style."""
    try:
        model = cell._model
        tm = model.objects[cell._table_id]
        if region == "body" and banded:
            props = model.objects[tm.table_style.identifier].table_properties
            if props.banded_rows and props.banded_fill.HasField("color"):
                return rgb_hex(rgb_from_proto(props.banded_fill.color))
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
    return s


def border_style(cell):
    s = {}
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


def preview(path, password=None):
    """Cells are sent as plain text, [text, style index], or [text, style index, extras] where
    extras may hold "runs" ([[text, style index], ...]), "link", "note" and "control"."""
    book = load(path, password)
    sheets = []
    for excel_name, _sheet, table in book.tables:
        rows_shown = min(table.num_rows, PREVIEW_MAX_ROWS)
        cols_shown = min(table.num_cols, PREVIEW_MAX_COLS)
        styles, style_ids = [{}], {"{}": 0}  # index 0 = unstyled; such cells are sent as plain strings

        def style_index(st):
            key = json.dumps(st, sort_keys=True)
            if key not in style_ids:
                style_ids[key] = len(styles)
                styles.append(st)
            return style_ids[key]

        computed = book.pivot_values.get(table._table_id, {})
        source_rows = table_rows(table)
        rows = []
        for r in range(rows_shown):
            row = []
            for c in range(cols_shown):
                cell = source_rows[r][c]
                if isinstance(cell, MergedCell):
                    row.append(None)
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
                if info:
                    row.append([text, index, info])
                else:
                    row.append([text, index] if index else text)
            rows.append(row)

        merges = []
        for rng in table.merge_ranges:
            r1, c1, r2, c2 = parse_range(rng)
            if r1 < rows_shown and c1 < cols_shown:
                merges.append([r1, c1, min(r2, rows_shown - 1), min(c2, cols_shown - 1)])

        sheets.append({
            "name": excel_name,
            "rows": rows,
            "styles": styles,
            "merges": merges,
            "colWidths": header_sizes(table, columns=True)[:cols_shown],
            "rowHeights": header_sizes(table, columns=False)[:rows_shown],
            "headerRows": table.num_header_rows,
            "headerCols": table.num_header_cols,
            "hiddenRows": sorted(i for i in extras.hidden(table, columns=False) if i < rows_shown),
            "hiddenCols": sorted(i for i in extras.hidden(table, columns=True) if i < cols_shown),
            "totalRows": table.num_rows,
            "totalCols": table.num_cols,
        })
    return {"sheets": sheets, "unsupported": book.unsupported}


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
NAME = r"""(?:'[^']+'|[^\s:(),'"=<>+\-*/^&{}][^:(),'"=<>+\-*/^&{}]*?)"""
TOKEN = re.compile(
    rf"""
    (?P<str>"(?:[^"]|"")*")
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


def referenced_cells(book):
    """{excel sheet name: [(r1, c1, r2, c2)]} of every range a formula in the file reads."""
    refs = {}
    for excel_name, sheet_name, table in book.tables:
        for row in table.iter_rows():
            for cell in row:
                if not cell.is_formula:
                    continue
                try:
                    excel = translate_formula(cell.formula, book, sheet_name)
                except Exception:
                    continue
                for target, rng in re.findall(r"(?:'((?:[^']|'')+)'!)?(\$?[A-Z]{1,3}\$?\d+(?::\$?[A-Z]{1,3}\$?\d+)?)", excel or ""):
                    name = target.replace("''", "'") if target else excel_name
                    refs.setdefault(name, []).append(parse_range(rng.replace("$", "")))
    return refs


def is_referenced(refs, r, c):
    return any(r1 <= r <= r2 and c1 <= c <= c2 for r1, c1, r2, c2 in refs)


def translate_formula(formula, book, sheet_name):
    """Numbers formula text -> Excel formula, or None if it can't be represented safely."""
    if not formula or not formula.strip():
        return None  # e.g. pivot table summaries, which numbers-parser can't decode
    out = []
    # One frame per open parenthesis: [function name or None, commas seen, any argument text seen].
    stack = []
    for m in TOKEN.finditer(formula):
        kind, text = m.lastgroup, m.group()
        if stack and kind != "ws" and text not in (",", ")"):
            stack[-1][2] = True
        if kind == "op" and text == "(":
            stack.append([None, 0, False])
        elif kind == "op" and text == ",":
            if not stack:
                return None
            stack[-1][1] += 1
        elif kind == "op" and text == ")":
            if not stack:
                return None
            name, commas, seen = stack.pop()
            if not name and commas:
                return None  # Numbers' (A1,B2) list syntax; Excel reads it as a union and gives #VALUE!
            if name:
                count = commas + 1 if (commas or seen) else 0
                low, high = EXCEL_FUNCTIONS[name]
                if count < low or (high is not None and count > high):
                    return None
        if kind == "str":
            if re.search(r"\d{1,4}[/.-]\d{1,2}[/.-]\d{1,4}", text):
                return None  # dates written as text are read using the computer's region settings
            out.append(text)
        elif kind == "xref":
            *names, cells = [p.strip("'") for p in text.split("::")]
            if len(names) == 2:
                target = book.by_table.get((names[0], names[1]))
            else:
                target = (
                    book.by_table.get((sheet_name, names[0]))
                    or book.by_table.get((None, names[0]))
                    or book.by_sheet.get(names[0])
                )
            if target is None:
                return None
            out.append("'" + target.replace("'", "''") + "'!" + cells)
        elif kind == "ref":
            out.append(text)
        elif kind == "func":
            name = m.group("func")
            if name not in EXCEL_FUNCTIONS:
                return None
            stack.append([name, 0, False])
            out.append(name + "(")
        elif kind in ("bool", "num", "op", "ws"):
            out.append(text)
        elif text in OPERATOR_MAP:
            out.append(OPERATOR_MAP[text])
        else:
            return None  # column-name references, array constants, unknown syntax, etc.
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


def convert(path, out, password=None):
    book = load(path, password)
    workbook = xlsxwriter.Workbook(out, {"use_future_functions": True, "strings_to_numbers": False,
                                         "strings_to_formulas": False, "strings_to_urls": False})
    formats = FormatCache(workbook)
    lists = ListSheet(workbook)
    stats = {"formulas": 0, "formulasAsValues": 0}
    refs = referenced_cells(book)

    for excel_name, sheet_name, table in book.tables:
        ws = workbook.add_worksheet(excel_name)
        sheet_refs = refs.get(excel_name, [])
        hidden_cols = extras.hidden(table, columns=True)
        for c, width in enumerate(header_sizes(table, columns=True)):
            ws.set_column_pixels(c, c, max(8, width), None, {"hidden": c in hidden_cols})
        try:
            heights = stored_sizes(table, columns=False)
        except Exception:
            heights = {}
        hidden_rows = extras.hidden(table, columns=False)
        for r in sorted(set(heights) | hidden_rows):  # rows Numbers sizes automatically are left for Excel to fit
            if r >= table.num_rows:
                continue
            if r in heights:
                ws.set_row_pixels(r, heights[r], None, {"hidden": r in hidden_rows})
            else:
                ws.set_row(r, None, None, {"hidden": True})
        if table.num_header_rows or table.num_header_cols:
            ws.freeze_panes(table.num_header_rows, table.num_header_cols)
        if extras.filter_enabled(table) and table.num_rows > 1:
            header = max(table.num_header_rows - 1, 0)
            ws.autofilter(header, 0, table.num_rows - 1, table.num_cols - 1)

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
                if not write_special(ws, r, c, cell, style, fmt, formats, is_referenced(sheet_refs, r, c)):
                    write_cell(ws, r, c, cell, fmt, book, sheet_name, stats)
                add_cell_extras(ws, r, c, cell, lists)

    workbook.close()
    return {"outputPath": out, "sheets": len(book.tables), "unsupported": book.unsupported, **stats}


def write_special(ws, r, c, cell, style, fmt, formats, referenced=False):
    """Write cells that need more than a plain value: checkboxes, links and mixed formatting.
    Returns False to let write_cell handle the cell normally."""
    if cell.is_formula:
        return False
    bullets = extras.bullet_text(cell)
    # Bullets are only display formatting in Numbers, so a cell a formula reads keeps its plain text.
    if bullets and not referenced and not extras.links(cell):
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


def write_cell(ws, r, c, cell, fmt, book, sheet_name, stats):
    if cell.is_formula:
        excel = formula = None
        try:
            formula = cell.formula
            # Numbers returns durations from date maths where Excel returns plain numbers (e.g. NETWORKDAYS),
            # which the duration format would then show as hours, so those keep their values.
            if not isinstance(cell, DurationCell):
                excel = translate_formula(formula, book, sheet_name)
        except Exception:
            pass
        if excel:
            stats["formulas"] += 1
            ws.write_formula(r, c, excel, fmt, cached_value(cell))
            return
        if formula and formula.strip():  # pivot summaries have no readable formula or value; they're reported separately
            stats["formulasAsValues"] += 1

    value = cell.value
    if isinstance(cell, EmptyCell) or value is None:
        ws.write_blank(r, c, None, fmt)
    elif isinstance(cell, ErrorCell):
        ws.write_string(r, c, formatted(cell) or "#ERROR", fmt)
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
        return preview(request["path"], request.get("password"))
    if cmd == "convert":
        return convert(request["path"], request["out"], request.get("password"))
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
