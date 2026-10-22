"""
Where everything goes in the workbook.

A Numbers sheet that holds several tables can be laid out two ways:
  * TABLES: each table on a worksheet of its own ("Sheet - Table"), starting in column A, with rows
    above it making room for charts (and text boxes, shapes and pictures) that sat above it;
  * SHEET: the whole Numbers sheet on one worksheet, with each table where it sits in Numbers.
    Tables side by side share rows and tables one above the other share columns (each shared row
    or column as tall or wide as the largest of them needs, so text still fits), and the space
    between them becomes rows and columns of that size.

The file is read once (numbers_tool.Book); a Layout is worked out from it for each set of choices:
a default, and a choice per Numbers sheet.
"""

import math

import extras

TABLES, SHEET = "tables", "sheet"
EXCEL_MAX_ROW_HEIGHT = 409    # points
EXCEL_MAX_COLUMN_WIDTH = 1300  # points (Excel allows 255 characters)
GAP = 20                       # points between tables that would otherwise overlap
MARGIN = 6                     # points left above or beside charts at a sheet's edge
FROZEN_ROOM = 150              # points of room above a table that may be frozen along with its header rows
PT_TO_PX = 4 / 3


class Placed:
    """A table on a worksheet: the worksheet's name, the table's first row and column there
    (0-based), and the row holding its title when Numbers shows its name."""

    def __init__(self, table, sheet, row, col, title_row):
        self.table, self.sheet, self.row, self.col, self.title_row = table, sheet, row, col, title_row


class Worksheet:
    def __init__(self, name, numbers_sheet, mode, table_count):
        self.name = name
        self.numbers_sheet = numbers_sheet  # the Numbers sheet it comes from
        self.mode = mode                    # how that Numbers sheet is laid out
        self.table_count = table_count      # tables on that Numbers sheet
        self.tables = []                    # [Placed]
        self.rows, self.cols = {}, {}       # {index: height or width in points}
        self.hidden_rows, self.hidden_cols = set(), set()
        self.used_rows, self.used_cols = set(), set()  # rows and columns holding a table
        self.charts = []                    # [(chart, (x, y) in pixels from the top-left cell)]
        self.drawings = []                  # [(text box, shape or picture, (x, y) in pixels)], in stacking order
        self.freeze = None                  # (rows, columns) kept in view, counted from the top-left
        self.autofilter = None              # (first row, first column, last row, last column)

    def size(self):
        return max(self.rows, default=-1) + 1, max(self.cols, default=-1) + 1

    def offset(self, columns, index):
        """Points from the sheet's top (or left) edge to row (or column) `index`."""
        sizes, hidden = (self.cols, self.hidden_cols) if columns else (self.rows, self.hidden_rows)
        return sum(size for i, size in sizes.items() if i < index and i not in hidden)

    def rect(self, placed):
        """(x, y, width, height) in points of a table on this sheet, title row included."""
        x, y = self.offset(True, placed.col), self.offset(False, placed.row if placed.title_row is None else placed.title_row)
        return (x, y, self.offset(True, placed.col + placed.table.num_cols) - x,
                self.offset(False, placed.row + placed.table.num_rows) - y)


class Footprint:
    """A table on its Numbers sheet: its position there, and its rows (title row first, if shown)
    and columns as they're written out: [(size in points, hidden)]."""

    def __init__(self, table, position, rows, cols, title):
        self.table, self.position, self.rows, self.cols, self.title = table, position, rows, cols, title

    def spans(self, columns):
        """Sizes along one axis as the table takes up space on the Numbers sheet (hidden: none)."""
        return [0 if hidden else size for size, hidden in (self.cols if columns else self.rows)]

    def rect(self):
        return (*self.position, sum(self.spans(True)), sum(self.spans(False)))


class Layout:
    """`options`: {"default": TABLES or SHEET, "sheets": {Numbers sheet name: TABLES or SHEET}}.
    `measure(table)` gives a table's Footprint, `unique_name(name, used)` an unused worksheet name."""

    def __init__(self, book, options, measure, unique_name):
        options = options or {}
        default = options.get("default") if options.get("default") in (TABLES, SHEET) else TABLES
        chosen = {name: mode for name, mode in (options.get("sheets") or {}).items() if mode in (TABLES, SHEET)}
        self.book = book
        self.sheets = []       # [Worksheet] in workbook order
        self.place = {}        # {table id: Placed}
        self.used_names = set()
        self.formulas = {}     # formula translations for this layout (see numbers_tool.excel_formula)
        self.formula_reasons = {}  # why formulas that couldn't be translated keep their values
        self.highlights = {}   # tables' highlighting rules as Excel's (see numbers_tool.table_highlights)
        self.report = None     # what will be different in Excel (numbers_tool.change_report), once worked out
        for sheet_id, sheet_name, tables in book.sheets:
            mode = chosen.get(sheet_name, default)
            footprints = [measure(table) for table in tables]
            sheet_charts = [chart for chart in book.charts if chart.sheet_id == sheet_id]
            sheet_drawings = [d for d in book.drawings if d.sheet_id == sheet_id]
            if not tables:
                if sheet_charts or sheet_drawings:
                    self.chart_sheet(sheet_name, mode, sheet_charts, sheet_drawings, unique_name)
            elif mode == SHEET:
                self.whole_sheet(sheet_name, footprints, sheet_charts, sheet_drawings, unique_name)
            else:
                self.table_sheets(sheet_name, footprints, sheet_charts, sheet_drawings, unique_name)
        # Table names, for formulas that name a table (Table::A1) rather than point at it exactly.
        self.by_table, self.by_sheet = {}, {}
        for _, sheet_name, tables in book.sheets:
            for table in tables:
                self.by_table[(sheet_name, table.name)] = table._table_id
                self.by_table.setdefault((None, table.name), table._table_id)
                self.by_sheet.setdefault(sheet_name, table._table_id)

    # ------------------------------------------------------------ one worksheet per table

    def table_sheets(self, sheet_name, footprints, sheet_charts, sheet_drawings, unique_name):
        hosted = {id(f): [] for f in footprints}
        for chart in sheet_charts:
            own = [f for f in footprints if f.table._table_id in chart.tables]
            host = own[0] if own else min(footprints, key=lambda f: distance(chart.frame, f.rect()))
            hosted[id(host)].append(chart)
        drawn = {id(f): [] for f in footprints}  # text boxes, shapes and pictures go with the nearest table
        for d in sheet_drawings:
            drawn[id(min(footprints, key=lambda f: host_distance(d.frame, f.rect())))].append(d)
        for f in footprints:
            label = sheet_name if len(footprints) == 1 else f"{sheet_name} - {f.table.name}"
            ws = Worksheet(unique_name(label, self.used_names), sheet_name, TABLES, len(footprints))
            charts = hosted[id(f)]
            x, y = f.position
            # Rows above the table make room for charts over it. The table starts in column A, so
            # with charts to its left it moves left of where it was, and the charts keep their places.
            origin = corner([f], charts, drawn[id(f)])
            starts, gaps = arrange([(y, f.spans(False))], origin[1], EXCEL_MAX_ROW_HEIGHT)
            self.put(ws, f, starts[0], 0)
            fill(ws, gaps, {})
            # Header rows stay frozen under a little room made for a title or note (frozen with them).
            for_drawings = not arrange([(y, f.spans(False))], corner([f], charts)[1], EXCEL_MAX_ROW_HEIGHT)[1]
            freeze(ws, {} if for_drawings and sum(gaps.values()) <= FROZEN_ROOM else gaps)
            self.position_charts(ws, charts, origin)
            self.position_drawings(ws, drawn[id(f)], origin)
            self.sheets.append(ws)

    # ------------------------------------------------------------ the whole Numbers sheet on one worksheet

    def whole_sheet(self, sheet_name, footprints, sheet_charts, sheet_drawings, unique_name):
        ws = Worksheet(unique_name(sheet_name, self.used_names), sheet_name, SHEET, len(footprints))
        origin = corner(footprints, sheet_charts, sheet_drawings)
        cols, col_gaps = arrange([(f.position[0], f.spans(True)) for f in footprints], origin[0], EXCEL_MAX_COLUMN_WIDTH)
        spans = [(start, start + len(f.cols)) for start, f in zip(cols, footprints)]

        def share_columns(i, j):
            return spans[i][0] < spans[j][1] and spans[j][0] < spans[i][1]

        rows, row_gaps = arrange([(f.position[1], f.spans(False)) for f in footprints], origin[1],
                                 EXCEL_MAX_ROW_HEIGHT, share_columns)
        for f, row, col in zip(footprints, rows, cols):
            self.put(ws, f, row, col)
        fill(ws, row_gaps, col_gaps)
        if len(footprints) == 1 and not col_gaps:  # a lone table in column A freezes as on a sheet of its own
            freeze(ws, row_gaps)
        self.position_charts(ws, sheet_charts, origin, footprints)
        self.position_drawings(ws, sheet_drawings, origin, footprints)
        self.sheets.append(ws)

    def chart_sheet(self, sheet_name, mode, sheet_charts, sheet_drawings, unique_name):
        """A Numbers sheet with charts, text boxes, shapes or pictures but no tables."""
        ws = Worksheet(unique_name(sheet_name or "Charts", self.used_names), sheet_name, mode, 0)
        origin = corner([], sheet_charts, sheet_drawings)
        self.position_charts(ws, sheet_charts, origin)
        self.position_drawings(ws, sheet_drawings, origin)
        self.sheets.append(ws)

    # ------------------------------------------------------------ shared

    def put(self, ws, f, start_row, start_col):
        """Places a table (its title row first) at a row and column, recording the sizes of its rows
        and columns. A row (column) shared with another table is as large as the larger of them
        needs, and hidden only if both hide it."""
        title = 1 if f.title else 0
        placed = Placed(f.table, ws.name, start_row + title, start_col, start_row if title else None)
        ws.tables.append(placed)
        self.place[f.table._table_id] = placed
        for sizes, hidden, used, items, start in ((ws.rows, ws.hidden_rows, ws.used_rows, f.rows, start_row),
                                                  (ws.cols, ws.hidden_cols, ws.used_cols, f.cols, start_col)):
            for k, (size, is_hidden) in enumerate(items):
                i = start + k
                if i in used:
                    sizes[i] = max(sizes[i], size)
                    if not is_hidden:
                        hidden.discard(i)
                else:
                    used.add(i)
                    sizes[i] = size
                    if is_hidden:
                        hidden.add(i)
        table = f.table
        if ws.autofilter is None and extras.filter_enabled(table) and table.num_rows > 1:
            header = max(table.num_header_rows - 1, 0)  # Excel puts the filter buttons on the last header row
            ws.autofilter = (placed.row + header, placed.col, placed.row + table.num_rows - 1,
                             placed.col + table.num_cols - 1)

    def position_charts(self, ws, sheet_charts, origin, footprints=None):
        """Charts keep their place on the Numbers sheet, measured from the sheet's top-left corner
        (`origin`, on the Numbers sheet) or, on a sheet laid out as in Numbers, from the nearest
        table, so a chart moves along with its table when rows and columns around it grow. A chart
        that would cover a table's cells goes just right of it instead."""
        for chart in sheet_charts:
            ws.charts.append((chart, spot(ws, chart.frame, origin, footprints)))

    def position_drawings(self, ws, sheet_drawings, origin, footprints=None):
        """Text boxes, shapes and pictures keep their place the same way."""
        for d in sheet_drawings:
            ws.drawings.append((d, spot(ws, d.frame, origin, footprints)))


def spot(ws, frame, origin, footprints):
    """Where something with this frame on the Numbers sheet goes on a worksheet, in pixels from its
    top-left cell (see Layout.position_charts)."""
    rects = [(p, ws.rect(p)) for p in ws.tables]
    fx, fy, fw, fh = frame
    if footprints:
        f = min(footprints, key=lambda f: distance(frame, f.rect()))
        tx, ty, _, _ = next(rect for p, rect in rects if p.table is f.table)
        x, y = tx + fx - f.position[0], ty + fy - f.position[1]
    else:
        x, y = fx - origin[0], fy - origin[1]
    x, y = max(0.0, x), max(0.0, y)
    for _, rect in rects:
        if overlap((x, y, fw, fh), rect) > 10:
            x = rect[0] + rect[2] + 12
    return round(x * PT_TO_PX), round(y * PT_TO_PX)


def corner(footprints, charts, drawings=()):
    """The point on the Numbers sheet that becomes the worksheet's top-left corner: the top-left
    table's corner, or a little above and left of any chart further up or left (but not past the
    sheet's edge). Text boxes, shapes and pictures move it only where they'd otherwise be cut off."""
    def edge(axis):
        tables = [f.position[axis] for f in footprints]
        drawn = [max(0.0, min(c.frame[axis] for c in charts) - MARGIN)] if charts else []
        start = min(tables + drawn) if tables or drawn else None
        others = [d.frame[axis] for d in drawings]
        if others and (start is None or min(others) < start - MARGIN):
            moved = max(0.0, min(others) - MARGIN)
            start = moved if start is None else min(start, moved)
        return start if start is not None else 0.0
    return edge(0), edge(1)


def fill(ws, row_gaps, col_gaps):
    """Rows and columns that only make space take their size, unless a table ended up in them."""
    for gaps, sizes, used in ((row_gaps, ws.rows, ws.used_rows), (col_gaps, ws.cols, ws.used_cols)):
        for index, size in gaps.items():
            if index not in used:
                sizes[index] = size


def arrange(blocks, origin, limit, share=None):
    """Lays blocks out along one axis of the worksheet grid. Each block is (where it starts on the
    Numbers sheet, [sizes of its rows (or columns) there]). A block starting within another's span
    lines up with that block's nearest row edge and shares its rows from there; a block starting
    beyond all of them gets new rows, after rows sized to the space in between (none if they
    touch), at most `limit` each. On the row axis, `share(i, j)` says whether tables i and j share
    columns: those can't share rows too, so a table that would overlap one placed before it moves
    below it, after a GAP-high row. Returns each block's first index, and {index: size} of the rows
    that only make space."""
    order = sorted(range(len(blocks)), key=lambda i: blocks[i][0])
    starts, edges, gaps = [None] * len(blocks), {}, {}
    next_index, free = 0, origin   # the first unused index, and where unused space starts on the sheet

    def add_gap(size):
        nonlocal next_index
        pieces = max(1, math.ceil(size / limit))
        for _ in range(pieces):
            gaps[next_index] = round(size / pieces, 2)
            next_index += 1

    for i in order:
        position, sizes = blocks[i]
        edges[i] = [position]
        for size in sizes:
            edges[i].append(edges[i][-1] + size)
        holders = [j for j in edges if j != i and edges[j][0] - 1 <= position < edges[j][-1] - 1]
        if holders:
            def nearest(j):
                return min(range(len(edges[j])), key=lambda k: abs(edges[j][k] - position))
            j = min(holders, key=lambda j: (abs(edges[j][nearest(j)] - position), -len(edges[j])))
            start = starts[j] + nearest(j)
        else:
            if position - free >= 2:
                add_gap(position - free)
            start = next_index
        while share:
            hits = [j for j in edges if j != i and share(i, j)
                    and start < starts[j] + len(blocks[j][1]) and starts[j] < start + len(sizes)]
            if not hits:
                break
            start = max(starts[j] + len(blocks[j][1]) for j in hits)
            if start >= next_index:
                next_index = start
                add_gap(GAP)
                start = next_index
        starts[i] = start
        next_index = max(next_index, start + len(sizes))
        free = max(free, edges[i][-1])
    return starts, gaps


def freeze(ws, row_gaps):
    """Keeps a worksheet's only table's header rows and columns in view where Numbers freezes them;
    not its rows when rows above it make room for charts, which would be frozen too."""
    table = ws.tables[0].table
    tm = table._model.objects[table._table_id]
    rows = table.num_header_rows if tm.header_rows_frozen else 0
    cols = table.num_header_cols if tm.header_columns_frozen else 0
    if rows and not row_gaps:
        ws.freeze = (ws.tables[0].row + rows, cols)
    elif cols:
        ws.freeze = (0, cols)


def distance(a, b):
    ax, ay, aw, ah = a
    bx, by, bw, bh = b
    return math.hypot(max(bx - (ax + aw), ax - (bx + bw), 0), max(by - (ay + ah), ay - (by + bh), 0))


def host_distance(drawing, table):
    """How near a text box, shape or picture is to a table, for choosing the table it goes with: a
    table it sits beside (a note next to it) comes before one it's just above or below (which would
    need room made above it), unless that one is far away."""
    _, dy, _, dh = drawing
    _, ty, _, th = table
    shared = min(dy + dh, ty + th) - max(dy, ty)
    beside = shared >= max(dh, 1.0) / 2 and distance(drawing, table) <= 100
    return (0 if beside else 1, distance(drawing, table))


def overlap(a, b):
    """How far two rectangles overlap: the smaller of the overlap's width and height (0 if apart)."""
    w = min(a[0] + a[2], b[0] + b[2]) - max(a[0], b[0])
    h = min(a[1] + a[3], b[1] + b[3]) - max(a[1], b[1])
    return max(0.0, min(w, h))
