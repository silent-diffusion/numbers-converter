"""
Table lines ("strokes") as Numbers draws them.

A table's style defines default lines per region: an outer border, the separators under header
rows / beside header columns / above footer rows, and the inner lines of each region. Lines set on
individual cells are stored separately in the table's stroke sidecar and drawn on top, in order.
numbers-parser only reads the sidecar, so tables that use their style's lines (most of them) came
out with no lines at all.

table_lines(model, table_id) returns two grids of lines, each None or (width, "#RRGGBB", "solid"/"dashes"/"dots"):
  horizontal[y][c]  the line along the top of row y (y == num_rows is the bottom of the table)
  vertical[r][x]    the line along the left of column x (x == num_cols is the right edge)
"""

from numbers_parser.generated import TSDArchives_pb2 as TSD

PATTERN = TSD.StrokePatternArchive.StrokePatternType
_cache = {}

VISIBILITY = {
    "hr_separator_visible": ["header_row_separator_stroke"],
    "hc_separator_visible": ["header_column_separator_stroke"],
    "footer_separator_visible": ["footer_row_separator_stroke"],
    "table_hr_divider_visible": ["header_row_vertical_stroke"],
    "table_hc_divider_visible": ["header_column_horizontal_stroke"],
    "table_footer_divider_visible": ["footer_row_vertical_stroke"],
    "table_header_border_visible": ["header_row_border_stroke", "header_column_border_stroke"],
    "table_border_visible": ["header_row_border_stroke", "header_column_border_stroke", "footer_row_border_stroke",
                             "table_body_horizontal_border_stroke", "table_body_vertical_border_stroke"],
}


def stroke_line(stroke):
    """(width, colour, style) for a StrokeArchive, or None if it draws nothing."""
    if stroke is None or not stroke.HasField("pattern"):
        return None
    kind = stroke.pattern.type
    if kind == PATTERN.TSDEmptyPattern or stroke.width <= 0:
        return None
    if kind == PATTERN.TSDPattern:
        style = "dots" if stroke.pattern.pattern and stroke.pattern.pattern[0] < 1.0 else "dashes"
    else:
        style = "solid"
    c = stroke.color
    # Excel's thinnest line is still heavier than Numbers' hairlines, so lines under 1pt are faded
    # toward white in proportion to their width, as Apple's own export does (0.35pt black -> #A5A5A5).
    fade = min(1.0, stroke.width)
    rgb = [round(255 - (255 - channel * 255) * fade) for channel in (c.r, c.g, c.b)]
    colour = "#{:02X}{:02X}{:02X}".format(*rgb)
    return (round(stroke.width, 2), colour, style)


def style_property(model, style, name):
    """A table style property, following the style's parents when it isn't set directly."""
    seen = 0
    while style is not None and seen < 10:
        props = style.table_properties
        if props.HasField(name):
            return getattr(props, name)
        parent = style.super.parent.identifier if style.HasField("super") and style.super.HasField("parent") else 0
        style = model.objects[parent] if parent and parent in model.objects else None
        seen += 1
    return None


def table_lines(model, table_id):
    key = (id(model), table_id)
    if key not in _cache:
        _cache[key] = compute(model, table_id)
    return _cache[key]


def compute(model, table_id):
    tm = model.objects[table_id]
    rows, cols = tm.number_of_rows, tm.number_of_columns
    hr, hc, fr = tm.number_of_header_rows, tm.number_of_header_columns, tm.number_of_footer_rows
    style = model.objects[tm.table_style.identifier] if tm.HasField("table_style") else None

    def line(name):
        return stroke_line(style_property(model, style, name)) if style is not None else None

    s = {name: line(name) for name in (
        "header_row_border_stroke", "header_row_separator_stroke", "header_row_horizontal_stroke",
        "header_row_vertical_stroke", "header_column_border_stroke", "header_column_separator_stroke",
        "header_column_horizontal_stroke", "header_column_vertical_stroke", "footer_row_border_stroke",
        "footer_row_separator_stroke", "footer_row_horizontal_stroke", "footer_row_vertical_stroke",
        "table_body_horizontal_border_stroke", "table_body_vertical_border_stroke",
        "table_body_horizontal_stroke", "table_body_vertical_stroke",
    )}
    # Switches in the style that hide some of those lines (unset means shown).
    for switch, names in VISIBILITY.items():
        if style is not None and style_property(model, style, switch) is False:
            for name in names:
                s[name] = None

    def row_region(r):
        if r < hr:
            return "header_row"
        if fr and r >= rows - fr:
            return "footer_row"
        return "body"

    horizontal = [[None] * cols for _ in range(rows + 1)]
    for y in range(rows + 1):
        for c in range(cols):
            in_header_col = c < hc
            if y == 0 or y == rows:  # outer border
                region = row_region(0 if y == 0 else rows - 1)
                if region == "header_row":
                    horizontal[y][c] = s["header_row_border_stroke"]
                elif region == "footer_row":
                    horizontal[y][c] = s["footer_row_border_stroke"]
                elif in_header_col:
                    horizontal[y][c] = s["header_column_border_stroke"]
                else:
                    horizontal[y][c] = s["table_body_horizontal_border_stroke"]
            elif hr and y == hr:
                horizontal[y][c] = s["header_row_separator_stroke"]
            elif fr and y == rows - fr:
                horizontal[y][c] = s["footer_row_separator_stroke"]
            else:
                region = row_region(y)
                if region == "header_row":
                    horizontal[y][c] = s["header_row_horizontal_stroke"]
                elif region == "footer_row":
                    horizontal[y][c] = s["footer_row_horizontal_stroke"]
                elif in_header_col:
                    horizontal[y][c] = s["header_column_horizontal_stroke"]
                else:
                    horizontal[y][c] = s["table_body_horizontal_stroke"]

    vertical = [[None] * (cols + 1) for _ in range(rows)]
    for r in range(rows):
        region = row_region(r)
        for x in range(cols + 1):
            if x == 0 or x == cols:  # outer border
                if region == "header_row":
                    vertical[r][x] = s["header_row_border_stroke"]
                elif region == "footer_row":
                    vertical[r][x] = s["footer_row_border_stroke"]
                elif x == 0 and hc:
                    vertical[r][x] = s["header_column_border_stroke"]
                else:
                    vertical[r][x] = s["table_body_vertical_border_stroke"]
            elif hc and x == hc and region == "body":
                vertical[r][x] = s["header_column_separator_stroke"]
            elif region == "header_row":
                vertical[r][x] = s["header_row_vertical_stroke"]
            elif region == "footer_row":
                vertical[r][x] = s["footer_row_vertical_stroke"]
            elif x < hc:
                vertical[r][x] = s["header_column_vertical_stroke"]
            else:
                vertical[r][x] = s["table_body_vertical_stroke"]

    apply_sidecar(model, tm, horizontal, vertical, rows, cols)
    return horizontal, vertical


def apply_sidecar(model, tm, horizontal, vertical, rows, cols):
    """Lines set on individual cells, drawn over the style's lines in their stored order."""
    if not tm.HasField("stroke_sidecar") or not tm.stroke_sidecar.identifier:
        return
    sidecar = model.objects[tm.stroke_sidecar.identifier]
    runs = []
    for side, layers in (("top", sidecar.top_row_stroke_layers), ("bottom", sidecar.bottom_row_stroke_layers),
                         ("left", sidecar.left_column_stroke_layers), ("right", sidecar.right_column_stroke_layers)):
        for ref in layers:
            layer = model.objects[ref.identifier]
            for run in layer.stroke_runs:
                runs.append((run.order, side, layer.row_column_index, run))
    for _, side, index, run in sorted(runs, key=lambda item: item[0]):
        value = stroke_line(run.stroke)
        for k in range(run.origin, run.origin + run.length):
            if side in ("top", "bottom"):
                y = index + (1 if side == "bottom" else 0)
                if 0 <= y <= rows and 0 <= k < cols:
                    horizontal[y][k] = value
            else:
                x = index + (1 if side == "right" else 0)
                if 0 <= x <= cols and 0 <= k < rows:
                    vertical[k][x] = value


def cell_lines(cell):
    """{"bt"/"bb"/"bl"/"br": [width, colour, style]} for one cell (missing key = no line)."""
    horizontal, vertical = table_lines(cell._model, cell._table_id)
    row, col = cell.row, cell.col
    if row >= len(vertical) or col + 1 >= len(vertical[0]):
        return {}
    lines = {}
    for key, value in (("bt", horizontal[row][col]), ("bb", horizontal[row + 1][col]),
                       ("bl", vertical[row][col]), ("br", vertical[row][col + 1])):
        if value is not None:
            lines[key] = list(value)
    return lines
