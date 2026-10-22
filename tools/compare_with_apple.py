"""
Measures how close our conversions are to Apple's own Excel exports.

    .venv/Scripts/python tools/compare_with_apple.py reference/raw [--ours reference/ours] [--rows 120]

For every X.numbers in the folder that has an Apple export next to it (X.xlsx, and X 2.xlsx etc. for
Numbers' other export options), the file is converted with py/numbers_tool.py, laid out the way the
export was (an Excel sheet per table, or one per Numbers sheet), and both workbooks are opened in
Excel (so values, text, fonts, fills, borders, alignment and number formats are read exactly as Excel
shows them). Each Numbers table is located in both workbooks and compared cell by cell, and so is
each chart. Prints a score per attribute, per file, and example differences. Needs Windows with
Excel installed (pywin32).
"""

import argparse
import collections
import glob
import os
import sys
import warnings

warnings.simplefilter("ignore")
HERE = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, os.path.join(HERE, "..", "py"))
import numbers_tool as nt  # noqa: E402
import win32com.client as win32  # noqa: E402

EDGES = {"top": 8, "bottom": 9, "left": 7, "right": 10}  # xlEdgeTop/Bottom/Left/Right
NO_LINE = (-4142, None)  # xlNone
WHITE = 16777215
# Apple writes a font's weight into its name ("Avenir Next Demi Bold"); we use the family plus bold.
WEIGHTS = ("ultralight", "light", "thin", "regular", "medium", "demi bold", "demibold", "semibold",
           "bold", "heavy", "black", "italic", "condensed")
HEAVY = ("demi bold", "demibold", "semibold", "bold", "heavy", "black")
# We deliberately swap Mac fonts that Windows lacks for close Windows equivalents.
FONT_EQUIVALENTS = {k.lower(): v.lower() for k, v in nt.FONT_SUBSTITUTES.items()}


def font_key(name):
    name = (name or "").strip().lower()
    changed = True
    while changed:
        changed = False
        for word in WEIGHTS:
            if name.endswith(" " + word):
                name, changed = name[: -len(word) - 1].strip(), True
    return FONT_EQUIVALENTS.get(name, name)


def heavy(name):
    name = (name or "").lower()
    return any(name.endswith(" " + w) or f" {w} " in name for w in HEAVY)


def cell_props(ws, r, c):
    """Everything we compare for one cell (1-based r, c)."""
    cell = ws.Cells(r, c)
    font = cell.Font
    props = {
        "text": (cell.Text or "").strip(),
        "formula": bool(cell.HasFormula),
        "font": font_key(font.Name),
        "size": round(font.Size or 0, 1),
        "bold": bool(font.Bold) or heavy(font.Name),
        "italic": bool(font.Italic),
        "underline": font.Underline not in (-4142, None, 0, False),
        "strike": bool(font.Strikethrough),
        "color": font.Color,
        "fill": None if cell.Interior.Pattern in (-4142, None) or cell.Interior.Color == WHITE else cell.Interior.Color,
        "halign": cell.HorizontalAlignment,
        "valign": cell.VerticalAlignment,
        "wrap": bool(cell.WrapText),
        "numfmt": cell.NumberFormat,
    }
    for side, index in EDGES.items():
        b = cell.Borders(index)
        props["edge_" + side] = NO_LINE if b.LineStyle in (-4142, None) else (b.LineStyle, b.Weight, b.Color)
    return props


def locate(ws, expected):
    """(row, column) offset where a table starts on a sheet: where the text in its first rows best
    matches ours (matching on words, since numbers are stored raw but shown formatted)."""
    used = ws.UsedRange
    values = used.Value
    if not isinstance(values, tuple):
        values = ((values,),)
    texts = [[v.strip() if isinstance(v, str) else None for v in row] for row in values]
    words = [(r, c, text) for r, row in enumerate(expected) for c, text in enumerate(row)
             if text and any(ch.isalpha() for ch in text)]
    best, best_score = None, 0
    for r0 in range(len(texts)):
        for c0 in range(len(texts[0])):
            score = sum(1 for r, c, text in words
                        if r0 + r < len(texts) and c0 + c < len(texts[0]) and texts[r0 + r][c0 + c] == text)
            if score > best_score:
                best, best_score = (used.Row - 1 + r0, used.Column - 1 + c0), score
    return best


def layout_of(book, apple):
    """How an Apple export laid out Numbers sheets holding several tables: a sheet per table
    ("Sheet - Table") or one per Numbers sheet."""
    names = {ws.Name for ws in apple.Worksheets}
    for _, sheet_name, tables in book.sheets:
        if len(tables) > 1:
            return "tables" if any(f"{sheet_name} - {t.name}"[:31] in names for t in tables) else "sheet"
    return "tables"


def same(key, a, b):
    if key == "color" or key == "fill":
        return colour_close(a, b)
    if key.startswith("edge_"):
        if (a == NO_LINE) != (b == NO_LINE):
            return False
        return a == NO_LINE or (a[0] == b[0] and a[1] == b[1] and colour_close(a[2], b[2]))
    if key == "halign":  # General (1) and Left (-4131) look the same for text; compare what is visible
        return a == b or {a, b} <= {1, -4131}
    if key == "numfmt":
        return a == b or {a, b} <= {"General", "@"}
    return a == b


def colour_close(a, b, tolerance=24):
    if a is None or b is None:
        return a == b
    a, b = int(a), int(b)
    return all(abs(((a >> s) & 255) - ((b >> s) & 255)) <= tolerance for s in (0, 8, 16))


def effective_edges(grid):
    """A line between two cells shows if either cell draws it; fold each edge into the cell on its
    top/left so both workbooks are compared the same way."""
    rows, cols = len(grid), len(grid[0])
    for r in range(rows):
        for c in range(cols):
            p = grid[r][c]
            if r + 1 < rows and p["edge_bottom"] == NO_LINE:
                p["edge_bottom"] = grid[r + 1][c]["edge_top"]
            if c + 1 < cols and p["edge_right"] == NO_LINE:
                p["edge_right"] = grid[r][c + 1]["edge_left"]
    for r in range(rows):
        for c in range(cols):
            if r > 0:
                grid[r][c]["edge_top"] = grid[r - 1][c]["edge_bottom"]
            if c > 0:
                grid[r][c]["edge_left"] = grid[r][c - 1]["edge_right"]


def find_sheet(wb, candidates):
    names = {ws.Name: ws for ws in wb.Worksheets}
    for name in candidates:
        if name[:31] in names:
            return names[name[:31]]
    return None


def compare_file(excel, numbers_path, apple_path, ours_path, max_rows, totals, examples):
    book = nt.load(numbers_path)
    apple = excel.Workbooks.Open(os.path.abspath(apple_path), UpdateLinks=0, ReadOnly=True)
    options = {"default": layout_of(book, apple)}
    layout = book.layout(options)
    nt.convert(numbers_path, ours_path, layout_options=options)
    ours = excel.Workbooks.Open(os.path.abspath(ours_path), UpdateLinks=0, ReadOnly=True)
    file_scores = collections.Counter()
    file_totals = collections.Counter()
    name = f"{os.path.basename(apple_path)} ({options['default']})"
    widened = set()
    try:
        for sheet_name, table in book.tables:
            if table._model.objects[table._table_id].HasField("pivot_owner"):
                continue  # Apple writes a real Excel pivot table, laid out Excel's way; ours mirrors Numbers
            placed = layout.place[table._table_id]
            candidates = [f"{sheet_name} - {table.name}", sheet_name, placed.sheet]
            a_ws, o_ws = find_sheet(apple, candidates), find_sheet(ours, [placed.sheet])
            if a_ws is None or o_ws is None:
                print(f"  {name}: no matching sheet for {candidates[0]!r}")
                continue
            rows = min(table.num_rows, max_rows)
            cols = table.num_cols
            expected = [[(nt.formatted(cell) or "").strip() for cell in row[:cols]] for row in nt.table_rows(table)[:min(rows, 6)]]
            a_off = locate(a_ws, expected)
            if a_off is None:
                a_off = (1, 0)  # empty table: Apple puts the title in row 1
            o_off = (placed.row, placed.col)  # we know where we put it
            # Read sizes before widening columns for text comparison.
            for c in range(cols):
                a_w, o_w = a_ws.Columns(1 + a_off[1] + c).Width, o_ws.Columns(1 + o_off[1] + c).Width
                record(totals, file_scores, file_totals, examples, "column width", abs(a_w - o_w) <= max(2, 0.06 * a_w),
                       f"{name} [{a_ws.Name}] col {c + 1}: apple {a_w:.0f}pt ours {o_w:.0f}pt")
            for r in range(rows):
                a_h, o_h = a_ws.Rows(1 + a_off[0] + r).RowHeight, o_ws.Rows(1 + o_off[0] + r).RowHeight
                record(totals, file_scores, file_totals, examples, "row height", abs(a_h - o_h) <= max(2, 0.08 * a_h),
                       f"{name} [{a_ws.Name}] row {r + 1}: apple {a_h:.1f}pt ours {o_h:.1f}pt")
            for ws in (a_ws, o_ws):
                if (ws.Parent.Name, ws.Name) not in widened:  # so cell texts aren't cut to "###"
                    ws.Columns.ColumnWidth = 80
                    widened.add((ws.Parent.Name, ws.Name))
            a_grid = [[cell_props(a_ws, 1 + a_off[0] + r, 1 + a_off[1] + c) for c in range(cols)] for r in range(rows)]
            o_grid = [[cell_props(o_ws, 1 + o_off[0] + r, 1 + o_off[1] + c) for c in range(cols)] for r in range(rows)]
            effective_edges(a_grid)
            effective_edges(o_grid)
            for r in range(rows):
                for c in range(cols):
                    a, o = a_grid[r][c], o_grid[r][c]
                    for key in a:
                        if key in ("color",) and not a["text"]:
                            continue  # font colour of an empty cell isn't visible
                        ok = same(key, a[key], o[key])
                        if key == "numfmt" and not ok and a["text"] == o["text"]:
                            ok = True  # different code, same result on screen (e.g. "#,##0%" and "0%")
                        record(totals, file_scores, file_totals, examples, key, ok,
                               f"{name} [{a_ws.Name}] {col_name(c)}{r + 1}: apple {a[key]!r} ours {o[key]!r}"
                               + (f"  text={a['text'][:30]!r}" if key not in ("text",) else ""))
        # Charts: matched by title (else in order), comparing type and the data Excel reads for each series.
        mine = chart_list(ours)
        for a in chart_list(apple):
            o = next((c for c in mine if c["title"] == a["title"]), mine[0] if mine else None)
            label = f"{name} chart {a['title']!r}"
            record(totals, file_scores, file_totals, examples, "chart", o is not None, f"{label}: missing from ours")
            if o is None:
                continue
            mine.remove(o)
            record(totals, file_scores, file_totals, examples, "chart type", a["type"] == o["type"],
                   f"{label}: apple type {a['type']} ours {o['type']}")
            record(totals, file_scores, file_totals, examples, "chart data", a["values"] == o["values"],
                   f"{label}: apple {a['values']} ours {o['values']}")
            record(totals, file_scores, file_totals, examples, "chart labels", a["labels"] == o["labels"],
                   f"{label}: apple {a['labels']} ours {o['labels']}")
        # Text boxes: matched by their text, comparing the font of their first letter and, where both
        # workbooks put a Numbers sheet on one worksheet, where they sit. Then lines and other shapes.
        mine = shape_list(ours)
        for a in shape_list(apple):
            if not a["text"]:
                continue
            o = next((s for s in mine if s["text"] == a["text"]), None)
            label = f"{name} text box {a['text'][:30]!r}"
            record(totals, file_scores, file_totals, examples, "text box", o is not None, f"{label}: missing from ours")
            if o is None:
                continue
            mine.remove(o)
            record(totals, file_scores, file_totals, examples, "text box font", (a["size"], a["bold"]) == (o["size"], o["bold"]),
                   f"{label}: apple {a['size']}pt bold={a['bold']} ours {o['size']}pt bold={o['bold']}")
            if options["default"] == "sheet" and a["sheet"] == o["sheet"]:
                near = abs(a["left"] - o["left"]) <= 12 and abs(a["top"] - o["top"]) <= 12
                record(totals, file_scores, file_totals, examples, "text box place", near,
                       f"{label}: apple at ({a['left']:.0f}, {a['top']:.0f})pt ours ({o['left']:.0f}, {o['top']:.0f})pt")
        others = lambda shapes: sorted(s["kind"] for s in shapes if not s["text"])  # noqa: E731
        record(totals, file_scores, file_totals, examples, "shapes", others(shape_list(apple)) == others(shape_list(ours)),
               f"{name}: apple {others(shape_list(apple))} ours {others(shape_list(ours))}")
    finally:
        apple.Close(False)
        ours.Close(False)
    scored = {k: f"{100 * file_scores[k] / file_totals[k]:.0f}%" for k in file_totals}
    print(f"  {name}: " + ", ".join(f"{k} {v}" for k, v in scored.items()))


def chart_list(wb):
    """Every chart in a workbook: its title, Excel chart type, and the values and category labels
    Excel reads for each series (so live links are checked, not just copied numbers)."""
    found = []
    for ws in wb.Worksheets:
        for i in range(1, ws.ChartObjects().Count + 1):
            chart = ws.ChartObjects(i).Chart
            values, labels = [], []
            for j in range(1, chart.SeriesCollection().Count + 1):
                series = chart.SeriesCollection(j)
                values.append([round(v, 6) if isinstance(v, float) else v for v in (series.Values or ())])
                labels.append([str(x) for x in (series.XValues or ())])
            found.append({"title": chart.ChartTitle.Text if chart.HasTitle else None, "type": chart.ChartType,
                          "values": values, "labels": labels})
    return found


def shape_list(wb):
    """Every text box, shape and picture in a workbook (not charts): its worksheet, text (spaces
    and line breaks folded), place in points, the font size and boldness of its first letter, and
    what kind of shape it is."""
    found = []
    for ws in wb.Worksheets:
        for i in range(1, ws.Shapes.Count + 1):
            shape = ws.Shapes(i)
            if shape.HasChart:
                continue
            frame = shape.TextFrame2 if shape.Type != 13 else None  # msoPicture has no text
            text = " ".join(str(frame.TextRange.Text).split()) if frame is not None and frame.HasText else ""
            first = shape.TextFrame.Characters(1, 1).Font if text else None
            found.append({"sheet": ws.Name, "text": text, "left": shape.Left, "top": shape.Top,
                          "size": round(first.Size, 1) if first else None, "bold": bool(first.Bold) if first else None,
                          "kind": "picture" if shape.Type == 13 else "line" if shape.Height < 1 or shape.Width < 1 else "shape"})
    return found


def record(totals, file_scores, file_totals, examples, key, ok, example):
    totals[key][0] += ok
    totals[key][1] += 1
    file_scores[key] += ok
    file_totals[key] += 1
    if not ok and len(examples[key]) < 6:
        examples[key].append(example)


def col_name(i):
    s = ""
    for i in (i + 1,):
        while i:
            i, rem = divmod(i - 1, 26)
            s = chr(65 + rem) + s
    return s


def main():
    parser = argparse.ArgumentParser()
    parser.add_argument("folder")
    parser.add_argument("--ours", default=None)
    parser.add_argument("--rows", type=int, default=120)
    parser.add_argument("--only", default=None)
    args = parser.parse_args()
    ours_dir = args.ours or os.path.join(args.folder, "..", "ours")
    os.makedirs(ours_dir, exist_ok=True)

    totals = collections.defaultdict(lambda: [0, 0])
    examples = collections.defaultdict(list)
    excel = win32.DispatchEx("Excel.Application")
    excel.Visible = False
    excel.DisplayAlerts = False
    try:
        for numbers_path in sorted(glob.glob(os.path.join(args.folder, "*.numbers"))):
            base = os.path.splitext(numbers_path)[0]
            if args.only and args.only not in base:
                continue
            exports = [base + ".xlsx"] + sorted(glob.glob(glob.escape(base) + " [0-9].xlsx"))
            for apple_path in filter(os.path.exists, exports):
                # Excel can't open two workbooks with the same file name at once.
                stem = os.path.splitext(os.path.basename(apple_path))[0]
                ours_path = os.path.join(ours_dir, stem + " (ours).xlsx")
                compare_file(excel, numbers_path, apple_path, ours_path, args.rows, totals, examples)
    finally:
        excel.Quit()

    print("\nOverall (share of cells/rows/columns that match Apple's export):")
    for key, (ok, n) in sorted(totals.items(), key=lambda kv: kv[1][0] / kv[1][1]):
        print(f"  {key:14} {100 * ok / n:5.1f}%  ({n - ok} of {n} differ)")
    print("\nExamples:")
    for key, (ok, n) in sorted(totals.items(), key=lambda kv: kv[1][0] / kv[1][1]):
        if ok < n:
            print(f"  {key}:")
            for e in examples[key]:
                print("     ", e)


if __name__ == "__main__":
    main()
