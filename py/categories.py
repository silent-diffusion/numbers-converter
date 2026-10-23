"""
Tables grouped into categories, as Numbers shows them.

Numbers keeps a categorized table's rows in their original order and only shows them grouped. A
category column on the left holds the disclosure triangles and the groups' names. Each group has a
row with its name and its summaries (the sum, average and so on of a column's values in the
group), above its own rows and any groups nested in it. When a level of groups has summaries, a
label row above its first group names them ("Name:" … "Sum:"). Apple's export writes this view out
as cells, with formulas for the summaries, and so does this app; each group also becomes an Excel
outline group, so it collapses as in Numbers.

read(table, tools) gives the view of a table with categories switched on (a Grouping), or None.
"""

from datetime import timedelta

from numbers_parser.cell import Cell, _decode_date_format
from numbers_parser.constants import EPOCH
from numbers_parser.numbers_uuid import NumbersUUID


class Unsupported(Exception):
    """Categories stored in a way this module doesn't understand: the table keeps its row order."""


# Summary functions: {name: (label row text, Excel function)}. Numbers stores a code for each; only
# the sum's is confirmed against Numbers files, so other codes are recognised from the results
# Numbers saved for each group (see settle).
FUNCTIONS = {
    "sum": ("Sum", "SUM"), "average": ("Average", "AVERAGE"), "min": ("Minimum", "MIN"),
    "max": ("Maximum", "MAX"), "count": ("Count", "COUNTA"), "count numbers": ("Count", "COUNT"),
}
FUNCTION_BY_CODE = {2: "sum"}
UNKNOWN = "?"  # a summary whose code isn't in FUNCTION_BY_CODE


class Level:
    """One level of grouping: the column whose values group the rows (`grouping` 0 by value, else
    by a part of a date), its header text, its summaries {column: function name, or None when only
    the results Numbers saved can be kept}, whether a label row names them, and the heights Numbers
    gives its group and label rows (None: the table's default)."""

    def __init__(self, number, column, grouping, header):
        self.number, self.column, self.grouping, self.header = number, column, grouping, header
        self.summaries = {}
        self.label = False
        self.height = self.label_height = None


class Group:
    """A group of rows: its level (1 = outermost), name, the group it's in, the table rows in it (its
    own and those of the groups inside it, in display order), whether Numbers shows it collapsed,
    and its summaries {column: (function name or None, our result, (Numbers' result, its number
    format) or None)}."""

    def __init__(self, uid, level, parent, node):
        self.uid, self.level, self.parent, self.node = uid, level, parent, node
        self.kids = {NumbersUUID(child.group_uid): child for child in node.child}
        self.name = ""
        self.rows = []
        self.collapsed = False
        self.summaries = {}

    def hidden(self):
        """Whether a group it's inside is collapsed, which hides its row too."""
        parent = self.parent
        while parent is not None:
            if parent.collapsed:
                return True
            parent = parent.parent
        return False


class Grouping:
    """A table's grouped view.
    rows:   [(kind, item)] in display order: ("row", table row), ("group", Group) or ("label", Level)
    at:     {table row: index in rows}
    depth:  [outline level of each of rows]: how many groups it's inside
    levels: [Level], outermost first;  groups: [Group] in display order
    width:  the category column's width in points"""

    def __init__(self, rows, depth, levels, groups, width):
        self.rows, self.depth, self.levels, self.groups, self.width = rows, depth, levels, groups, width
        self.at = {item: i for i, (kind, item) in enumerate(rows) if kind == "row"}


def group_by(table):
    """The table's GroupByArchive when it has categories switched on, else None. Pivot tables are
    left out: Numbers lays their rows out itself."""
    model = table._model
    tm = model.objects[table._table_id]
    if tm.HasField("pivot_owner") or not tm.category_owner.identifier:
        return None
    owner = model.objects[tm.category_owner.identifier]
    if not owner.group_by:
        return None
    archive = model.objects[owner.group_by[0].identifier]
    return archive if archive.is_enabled and archive.group_column else None


def enabled(table):
    try:
        return group_by(table) is not None
    except Exception:
        return False


def read(table, tools):
    """The table's grouped view, or None without categories. `tools` gives a cell's text as Numbers
    shows it (text), its value as Excel stores a number (number, None for text and blanks) and its
    Excel number format (number_format), whether it holds anything (filled), and a table's hidden
    rows (hidden). Raises Unsupported for categories it can't follow."""
    archive = group_by(table)
    if archive is None:
        return None
    model, table_id = table._model, table._table_id
    tm = model.objects[table_id]
    info = model.objects[model.table_info_id(table_id)]
    summary_model = model.objects[info.summary_model.identifier] if info.summary_model.identifier else None
    data, num_rows = table._data, table.num_rows
    header_rows, footer_rows = tm.number_of_header_rows, tm.number_of_footer_rows
    uid_map = model.objects[tm.base_column_row_uids.identifier]
    column_of = {NumbersUUID(uid_map.sorted_column_uids[i]): c for c, i in enumerate(uid_map.column_uid_for_index)}

    levels = []
    for number, grouping in enumerate(archive.group_column, start=1):
        column = column_of.get(NumbersUUID(grouping.column_uid))
        if column is None:
            raise Unsupported("grouping column")
        levels.append(Level(number, column, grouping.grouping_type, tools.text(data[0][column]) if header_rows else ""))
    for agg in archive.column_agg_type:
        column = column_of.get(NumbersUUID(agg.column_uid))
        if column is not None and agg.agg_type and 1 <= agg.level <= len(levels):
            # "Show as" summaries (shares of a total and the like) keep Numbers' results
            levels[agg.level - 1].summaries[column] = None if agg.show_as_type else FUNCTION_BY_CODE.get(agg.agg_type, UNKNOWN)
    if summary_model is not None:
        for level in levels:
            level.label = bool(level.summaries) and per_level(summary_model.label_row_visibility_list, level.number) == 0
            level.height = per_level(summary_model.summary_row_height_list, level.number) or None
            level.label_height = per_level(summary_model.label_row_height_list, level.number) or None

    # The display order lists table rows and groups (by uid). A group's uid can repeat elsewhere in
    # the tree, so each is looked up among the children of the groups open at that point.
    order = model.objects[model.objects[info.category_order.identifier].uid_map.identifier]
    view = [NumbersUUID(order.sorted_row_uids[i]) for i in order.row_uid_for_index]
    row_of = {NumbersUUID(uid): r for r, uid in enumerate(archive.row_uid_lookup.uuids)}
    top = {NumbersUUID(node.group_uid): node for node in archive.group_node_root.child}
    rows, depth, groups, stack = [], [], [], []
    for uid in view:
        r = row_of.get(uid)
        if r is not None:
            if r < header_rows or r >= num_rows - footer_rows:
                if r >= header_rows:
                    stack = []  # footer rows are in no group
                rows.append(("row", r))
                depth.append(0)
                continue
            rows.append(("row", r))
            depth.append(len(stack))
            for group in stack:
                group.rows.append(r)
            continue
        while stack and uid not in stack[-1].kids:
            stack.pop()
        node = (stack[-1].kids if stack else top).get(uid)
        if node is None or len(stack) >= len(levels):
            raise Unsupported("group order")
        group = Group(uid, len(stack) + 1, stack[-1] if stack else None, node)
        rows.append(("group", group))
        depth.append(group.level - 1)
        groups.append(group)
        stack.append(group)
    if sorted(item for kind, item in rows if kind == "row") != list(range(num_rows)):
        raise Unsupported("rows missing from the display order")

    hidden = tools.hidden(table)
    for group in groups:
        group.name = group_name(group, levels[group.level - 1], data, tools)
        group.collapsed = bool(group.rows) and all(r in hidden for r in group.rows)
    summarize(table, levels, groups, summary_model, column_of, tools)

    # A level's label row goes just above its first group.
    labelled, view_rows, view_depth = set(), [], []
    for (kind, item), d in zip(rows, depth):
        if kind == "group" and item.level not in labelled and levels[item.level - 1].label:
            labelled.add(item.level)
            view_rows.append(("label", levels[item.level - 1]))
            view_depth.append(d)
        view_rows.append((kind, item))
        view_depth.append(d)
    width = summary_model.category_column_width if summary_model is not None and summary_model.category_column_width else 50.0
    return Grouping(view_rows, view_depth, levels, groups, width)


def per_level(values, number):
    """A level's entry in one of the summary model's per-level lists (index 0 isn't a level)."""
    return values[number] if number < len(values) else 0


def group_name(group, level, data, tools):
    """A group's name as Numbers shows it in its row: the value its rows share, as their cells show
    it; for groups by a part of a date (year, month...), that part; and for checkboxes, “Column”
    and Not “Column”."""
    value = group.node.group_cell_value
    if level.grouping and value.HasField("date_value"):
        when = EPOCH + timedelta(seconds=value.date_value.value)
        return _decode_date_format(value.date_value.format.date_time_format, when)
    if value.HasField("boolean_value"):
        name = f"“{level.header}”" if level.header else "Checked"
        return name if value.boolean_value.value else f"Not {name}"
    if group.rows:
        return tools.text(data[group.rows[0]][level.column])
    if value.HasField("string_value"):
        return value.string_value.value
    if value.HasField("number_value"):
        return f"{value.number_value.value:g}"
    return ""


def summarize(table, levels, groups, summary_model, column_of, tools):
    """Works out each group's summaries: the function for each level's summary columns (see settle),
    and for each group its result and the one Numbers saved."""
    if not any(level.summaries for level in levels):
        return
    saved = {}
    if summary_model is not None:
        try:
            saved = saved_results(table, summary_model, column_of, tools)
        except Exception:
            saved = {}  # without them, summaries whose code we know still work out
    data = table._data
    for level in levels:
        here = [g for g in groups if g.level == level.number]
        for column, code in list(level.summaries.items()):
            function = settle(code, here, column, saved, data, tools)
            level.summaries[column] = function
            for g in here:
                result = compute(function, [data[r][column] for r in g.rows], tools) if function else None
                g.summaries[column] = (function, result, saved.get((g.uid, column)))


def settle(code, groups, column, saved, data, tools):
    """The function a level's summaries in a column use: the one Numbers' code names, if the results
    Numbers saved agree with it; for codes we don't know, the only function that gives those
    results. None when there's no telling: the saved results are then kept as they are."""
    if code is None:
        return None
    checks = []
    for g in groups:
        value = saved.get((g.uid, column), (None, None))[0]
        if value is not None:
            checks.append(([data[r][column] for r in g.rows], value))

    def fits(name):
        return all(close(compute(name, cells, tools), value) for cells, value in checks)

    if code != UNKNOWN:
        return code if fits(code) else None
    found = [name for name in FUNCTIONS if fits(name)] if checks else []
    if found == ["count", "count numbers"]:  # every cell holds a number: Numbers' Count counts them all
        return "count"
    return found[0] if len(found) == 1 else None


def compute(name, cells, tools):
    """A summary's result over cells, as its Excel function gives it (None for Excel's #DIV/0!)."""
    if name == "count":
        return sum(1 for cell in cells if tools.filled(cell))
    numbers = [n for n in (tools.number(cell) for cell in cells) if n is not None]
    if name == "count numbers":
        return len(numbers)
    if name == "sum":
        return sum(numbers)
    if name == "average":
        return sum(numbers) / len(numbers) if numbers else None
    if name == "min":
        return min(numbers) if numbers else 0
    if name == "max":
        return max(numbers) if numbers else 0
    return None


def close(a, b):
    return a is not None and b is not None and abs(a - b) <= 1e-9 * max(1.0, abs(b))


def saved_results(table, summary_model, column_of, tools):
    """{(group uid, column): (number, number format)} of the results Numbers saved for the groups'
    summaries. They're in the summary model's own cell store, which numbers-parser reads like a
    table's once it's under a table id of its own: a copy of the table's model holding that store,
    there only while it's read (under an id no other object or table uses, as numbers-parser caches
    what it reads by id)."""
    model, tm = table._model, table._model.objects[table._table_id]
    uids = model.objects[summary_model.column_row_uids.identifier]
    group_uids = [NumbersUUID(uids.sorted_row_uids[i]) for i in uids.row_uid_for_index]
    columns = [column_of.get(NumbersUUID(uids.sorted_column_uids[i])) for i in uids.column_uid_for_index]
    stand_in = type(tm)()
    stand_in.CopyFrom(tm)
    stand_in.base_data_store.CopyFrom(summary_model.data_store)
    stand_in.number_of_rows, stand_in.number_of_columns = len(group_uids), len(columns)
    store = model.objects._objects
    stand_in_id = max(store) + 1 + table._table_id
    store[stand_in_id] = stand_in
    try:
        results = {}
        for row, buffers in model.storage_buffers(stand_in_id).items():
            for col, buffer in enumerate(buffers):
                if buffer is not None and row < len(group_uids) and col < len(columns) and columns[col] is not None:
                    cell = Cell._from_storage(stand_in_id, row, col, buffer, model)
                    results[(group_uids[row], columns[col])] = (tools.number(cell), tools.number_format(cell))
        return results
    finally:
        del store[stand_in_id]


def edges(grouping, horizontal, header_rows, footer_rows, num_rows):
    """Lines across the grouped view, from the table's own (horizontal[y][c] runs along the top of
    table row y; see strokes.table_lines): a list per column for the top of each view row, plus one
    for the bottom of the table. Header and footer rows keep theirs. A group's row, and a label row,
    gets the line under the header rows above it, as Numbers draws it and Apple's export writes it
    (none between a label row and the group under it); a group's own rows are divided by the
    body's lines."""
    rows = grouping.rows
    columns = len(horizontal[0])
    separator = horizontal[header_rows] if header_rows else horizontal[0]
    inner_row = header_rows + 1
    inner = horizontal[inner_row] if inner_row < num_rows - footer_rows else [None] * columns

    def region(i):
        kind, item = rows[i]
        if kind != "row":
            return kind
        if item < header_rows:
            return "header"
        return "footer" if item >= num_rows - footer_rows else "body"

    out = [horizontal[0]]
    for i in range(1, len(rows)):
        above, below = region(i - 1), region(i)
        if above == below and below in ("header", "footer"):
            out.append(horizontal[rows[i][1]])
        elif above == "header":
            out.append(separator)
        elif below == "footer":
            out.append(horizontal[num_rows - footer_rows])
        elif below == "group" and above == "label":
            out.append([None] * columns)
        elif below in ("group", "label"):
            out.append(separator)
        elif above in ("group", "label"):
            out.append(inner)
        else:  # two rows in a group: their own line where they're neighbours in the table too
            out.append(horizontal[rows[i][1]] if rows[i][1] == rows[i - 1][1] + 1 else inner)
    out.append(horizontal[num_rows])
    return out
