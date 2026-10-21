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

import xlsxwriter
from numbers_parser import Document
from numbers_parser.cell import (
    BoolCell,
    DateCell,
    DurationCell,
    EmptyCell,
    ErrorCell,
    MergedCell,
    NumberCell,
)
from numbers_parser.constants import DATETIME_FIELD_MAP, FormatType
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
    "CHOOSE": (2, N), "ROWS": (1, 1), "COLUMNS": (1, 1), "XLOOKUP": (3, 6),
    # finance
    "PMT": (3, 5), "FV": (3, 5), "PV": (3, 5), "NPER": (3, 5), "RATE": (3, 6), "NPV": (2, N), "IRR": (1, 2),
    "IPMT": (4, 6), "PPMT": (4, 6),
}


class UserError(Exception):
    """An error whose message can be shown to the user as-is."""


# ---------------------------------------------------------------- loading

_cache = OrderedDict()


def load(path):
    """Parse a Numbers file into a list of table models (cached by path + mtime)."""
    if os.path.isdir(path):
        raise UserError(
            "This file was saved as a Numbers package. Re-save it in Numbers as a single file, or zip it first."
        )
    if not path.lower().endswith(".numbers"):
        raise UserError("Not a .numbers file.")
    key = (path, os.path.getmtime(path))
    if key in _cache:
        _cache.move_to_end(key)
        return _cache[key]

    try:
        doc = Document(path)
    except Exception as err:  # numbers-parser raises a variety of errors for bad files
        text = str(err).lower()
        if "encrypt" in text or "password" in text:
            raise UserError("This file is password-protected. Remove the password in Numbers and try again.")
        raise UserError("Could not read this Numbers file (it may be damaged or not a real Numbers file).")
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
    counts = {"charts": 0, "images": 0, "shapes": 0}
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
    try:
        return cell.formatted_value
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


def cell_style(cell):
    """Style dict shared by the preview (as CSS-ish values) and the Excel writer."""
    key = (id(cell._model), cell._table_id, cell._cell_style_id, cell._text_style_id)
    s = _style_cache.get(key)
    if s is None:
        s = _style_cache[key] = text_and_fill_style(cell)
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


def is_numeric(cell):
    return isinstance(cell, (NumberCell, DateCell, DurationCell))


# ---------------------------------------------------------------- preview


def preview(path):
    book = load(path)
    sheets = []
    for excel_name, _sheet, table in book.tables:
        rows_shown = min(table.num_rows, PREVIEW_MAX_ROWS)
        cols_shown = min(table.num_cols, PREVIEW_MAX_COLS)
        styles, style_ids = [{}], {"{}": 0}  # index 0 = unstyled; such cells are sent as plain strings
        rows = []
        for r in range(rows_shown):
            row = []
            for c in range(cols_shown):
                cell = table.cell(r, c)
                if isinstance(cell, MergedCell):
                    row.append(None)
                    continue
                st = cell_style(cell)
                if is_numeric(cell) and "ha" not in st:
                    st["ha"] = "right"
                key = json.dumps(st, sort_keys=True)
                if key not in style_ids:
                    style_ids[key] = len(styles)
                    styles.append(st)
                text = "" if isinstance(cell, EmptyCell) else formatted(cell)
                row.append([text, style_ids[key]] if style_ids[key] else text)
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
            "totalRows": table.num_rows,
            "totalCols": table.num_cols,
        })
    return {"sheets": sheets, "unsupported": book.unsupported}


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
            return "[h]:mm:ss"
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
                return code
            return None
        return builtin_number_format(fmt)
    except Exception:
        return None


def builtin_number_format(fmt):
    t = fmt.format_type
    places = fmt.decimal_places
    auto = places == DECIMAL_PLACES_AUTO
    decimals = "" if auto or places == 0 else "." + "0" * min(places, 15)
    if t == FormatType.PERCENT:
        return f"0{decimals or ('.##' if auto else '')}%"
    if t == FormatType.SCIENTIFIC:
        return f"0{decimals or '.00'}E+00"
    if t == FormatType.FRACTION:
        return "# ?/?"
    if t not in (FormatType.DECIMAL, FormatType.CURRENCY):
        return None
    if t == FormatType.DECIMAL and auto and not fmt.show_thousands_separator:
        return None
    body = ("#,##0" if fmt.show_thousands_separator or t == FormatType.CURRENCY else "0") + (
        "" if auto and t == FormatType.DECIMAL else decimals
    )
    if t == FormatType.CURRENCY:
        symbol = CURRENCY_SYMBOLS.get(fmt.currency_code, fmt.currency_code + " ")
        body = f'"{symbol}"{body}'
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
            if ch == "S":  # fractional seconds
                run = len(re.match(r"S+", pattern[i:]).group())
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


def translate_formula(formula, book, sheet_name):
    """Numbers formula text -> Excel formula, or None if it can't be represented safely."""
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


def convert(path, out):
    book = load(path)
    workbook = xlsxwriter.Workbook(out, {"use_future_functions": True, "strings_to_numbers": False,
                                         "strings_to_formulas": False, "strings_to_urls": False})
    formats = FormatCache(workbook)
    stats = {"formulas": 0, "formulasAsValues": 0}

    for excel_name, sheet_name, table in book.tables:
        ws = workbook.add_worksheet(excel_name)
        for c, width in enumerate(header_sizes(table, columns=True)):
            ws.set_column_pixels(c, c, max(8, width))
        try:
            heights = stored_sizes(table, columns=False)
        except Exception:
            heights = {}
        for r, height in heights.items():  # rows Numbers sizes automatically are left for Excel to fit
            if r < table.num_rows:
                ws.set_row_pixels(r, height)
        if table.num_header_rows or table.num_header_cols:
            ws.freeze_panes(table.num_header_rows, table.num_header_cols)

        merged_anchors = {}
        for rng in table.merge_ranges:
            r1, c1, r2, c2 = parse_range(rng)
            merged_anchors[(r1, c1)] = (r2, c2)

        for r, row in enumerate(table.iter_rows()):
            for c, cell in enumerate(row):
                if isinstance(cell, MergedCell):
                    continue
                style = cell_style(cell)
                num_format = excel_number_format(cell)
                fmt = formats.get(style, num_format)
                if (r, c) in merged_anchors:
                    r2, c2 = merged_anchors[(r, c)]
                    ws.merge_range(r, c, r2, c2, None, fmt)
                write_cell(ws, r, c, cell, fmt, book, sheet_name, stats)

    workbook.close()
    return {"outputPath": out, "sheets": len(book.tables), "unsupported": book.unsupported, **stats}


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
        excel = None
        # Numbers returns durations from date maths where Excel returns plain numbers (e.g. NETWORKDAYS),
        # which the duration format would then show as hours, so those keep their values.
        if not isinstance(cell, DurationCell):
            try:
                excel = translate_formula(cell.formula, book, sheet_name)
            except Exception:
                pass
        if excel:
            stats["formulas"] += 1
            ws.write_formula(r, c, excel, fmt, cached_value(cell))
            return
        stats["formulasAsValues"] += 1

    value = cell.value
    if isinstance(cell, EmptyCell) or value is None:
        ws.write_blank(r, c, None, fmt)
    elif isinstance(cell, ErrorCell):
        ws.write_string(r, c, formatted(cell) or "#ERROR", fmt)
    elif isinstance(cell, BoolCell) or isinstance(value, bool):
        ws.write_boolean(r, c, bool(value), fmt)
    elif isinstance(cell, DateCell):
        ws.write_datetime(r, c, naive(value), fmt)
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
        return preview(request["path"])
    if cmd == "convert":
        return convert(request["path"], request["out"])
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
            response = {"id": req_id, "ok": False, "error": str(err)}
        except Exception as err:  # never let one bad file kill the helper
            response = {"id": req_id, "ok": False, "error": f"Conversion failed ({type(err).__name__}: {err})."}
        stdout.write(json.dumps(response, ensure_ascii=False) + "\n")
        stdout.flush()


if __name__ == "__main__":
    main()
