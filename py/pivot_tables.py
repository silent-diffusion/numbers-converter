"""
Numbers pivot tables as native Excel PivotTables.

Numbers keeps a pivot table as its rules (which source columns group its rows and columns, and which
are summarised how; see pivots.py), a snapshot of its source data from the last refresh, and the table
it last showed: its labels, in the order Numbers shows them, with subtotal and grand total rows and
columns where its options put them (the table's view order). Excel keeps a PivotTable as a definition
(fields on rows and columns, values, items in order, totals) and a cache of the source data, while the
worksheet's cells hold what it last showed. This module turns one into the other:

  * build() works out the PivotTable: a cache field for each column of the converted source table,
    holding the snapshot's rows; the row, column and value fields, with Numbers' groupings (by value,
    by a date part Excel can group by itself, or else by listing which values make up each group);
    items in Numbers' order with Numbers' labels; Numbers' totals and subtotals; and every cell of
    Excel's tabular layout, with the values pivots.py would calculate for it;
  * write() puts those cells on the worksheet in place of the table (numbers_tool.write_table);
  * add_parts() adds the pivot cache and PivotTable parts to the file XlsxWriter saved, with the
    workbook's list of pivot caches, the relationships and content types, and a PivotTable style made
    from the Numbers table's colours, so the table keeps its look when Excel refreshes it.

Excel opens the file showing the cells as written; they are exactly what it shows after a refresh.
Anything Excel can't show the same way raises pivots.Unsupported, and the pivot is written as values.
"""

import datetime as dt
import os
import re
import zipfile
from types import SimpleNamespace
from xml.sax.saxutils import escape

from numbers_parser.cell import (
    BoolCell,
    DateCell,
    DurationCell,
    EmptyCell,
    ErrorCell,
    MergedCell,
    NumberCell,
)
from numbers_parser.numbers_uuid import NumbersUUID

import pivots
from pivots import Unsupported

EXCEL_EPOCH = pivots.EXCEL_EPOCH
DATA = "data"  # the Values field: one entry per data field, on whichever axis lists them
GRAND_UID = (1, 0)  # Numbers' grand total row and column

MAIN_NS = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
REL_NS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
REL_TYPE = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/"
CONTENT_TYPE = "application/vnd.openxmlformats-officedocument.spreadsheetml."
X14_NS = "http://schemas.microsoft.com/office/spreadsheetml/2009/9/main"
SUBTOTALS = {"sum": None, "average": "average", "count": "count", "min": "min", "max": "max",
             "stdev": "stdDev", "var": "var"}


# ---------------------------------------------------------------- values as Excel sees them


class Error(str):
    """An error value (#N/A, #REF!) in the source data."""


def excel_value(cell, text):
    """A source cell's value as the converted cell holds it (see numbers_tool.write_cell): None for
    blank, a number, text, True/False, a date, or an Error."""
    if isinstance(cell, (EmptyCell, MergedCell)) or cell is None:
        return None
    if isinstance(cell, ErrorCell):
        return Error("#REF!" if "#REF!" in (getattr(cell, "formula", "") or "") else "#N/A")
    value = cell.value
    if value is None or value == "":
        return None
    if isinstance(cell, BoolCell) or isinstance(value, bool):
        return bool(value)
    if isinstance(cell, DateCell):
        value = value.replace(tzinfo=None)
        return text(cell) if value.year < 1900 else value
    if isinstance(cell, DurationCell):
        seconds = value.total_seconds()
        return text(cell) if seconds < 0 else seconds / 86400
    if isinstance(value, (int, float)):
        return float(value)
    return str(value)


def kind_of(value):
    if value is None:
        return "m"
    if isinstance(value, Error):
        return "e"
    if isinstance(value, bool):
        return "b"
    if isinstance(value, dt.datetime):
        return "d"
    if isinstance(value, float):
        return "n"
    return "s"


def item_key(value):
    """When Excel counts two values as the same PivotTable item (text ignores case)."""
    kind = kind_of(value)
    if kind == "s":
        return (kind, value.casefold())
    return (kind, value)


def serial(value):
    delta = value - EXCEL_EPOCH
    return delta.days + (delta.seconds + delta.microseconds / 1e6) / 86400


def sort_key(value):
    """Excel's A to Z order of PivotTable items: blanks, numbers and dates, text (ignoring case),
    TRUE/FALSE, then errors."""
    kind = kind_of(value)
    if kind == "n":
        return (0, value, "")
    if kind == "d":
        return (0, serial(value), "")
    if kind == "s":
        return (1, 0, value.casefold())
    if kind == "b":
        return (2, int(value), "")
    if kind == "e":
        return (3, 0, value)
    return (-1, 0, "")


def default_text(value):
    """How Excel labels an item it has no caption for (before any number format)."""
    kind = kind_of(value)
    if kind == "m":
        return "(blank)"
    if kind == "b":
        return "TRUE" if value else "FALSE"
    if kind == "n":
        return f"{value:.15g}"
    if kind == "d":
        time = f" {value.hour}:{value.minute:02d}" if value.hour or value.minute else ""
        return f"{value.month}/{value.day}/{value.year}{time}"
    return str(value)


# ---------------------------------------------------------------- the PivotTable


class CacheField:
    """A field of the pivot cache: a column of the source range, or a grouping of another field
    that Excel works out itself (`group`)."""

    def __init__(self, name, values=None):
        self.name = name
        self.values = values  # per record, for fields read from the source range
        self.shared = None    # distinct values, when records refer to them by index
        self.index = None     # per record: index into shared
        self.group = None     # SimpleNamespace(base, kind "range"/"discrete", by, start, end, items, mapping)
        self.parent = None    # index of the field that groups this one

    def share(self):
        if self.shared is None:
            self.shared, positions, self.index = [], {}, []
            for value in self.values:
                key = item_key(value)
                if key not in positions:
                    positions[key] = len(self.shared)
                    self.shared.append(value)
                self.index.append(positions[key])
        return self.shared


class Axis:
    """A row or column field: which cache field it shows, its items in the order written (each
    {"x": shared or group item index, "caption" if Excel's own name isn't Numbers' label, "hidden",
    "key": Numbers' group key (None for items Numbers doesn't show), "label": Numbers' label}), the
    item each record falls in (None for records Numbers leaves out), and how it sorts and totals."""

    def __init__(self, field, items, keys, caption=None):
        self.field = field
        self.items = items
        self.keys = keys
        self.caption = caption
        self.subtotals = True
        self.sort = None      # None (as listed), "ascending" or "descending"
        self.sort_by = None   # data field index when sorted by its values
        self.num_format = None  # the number format of numbers and dates among its items
        self.num_fmt_id = 0


class DataField:
    def __init__(self, field, kind, name, num_format):
        self.field, self.kind, self.name, self.num_format = field, kind, name, num_format


class Spec:
    """A Numbers pivot table as an Excel PivotTable (see build)."""

    def __init__(self):
        self.table = None      # the Numbers pivot table
        self.source = None     # the converted source table, and the part of it the cache reads
        self.fields = []       # [CacheField]
        self.records = 0
        self.rows, self.cols = [], []  # [Axis or DATA]
        self.data = []         # [DataField]
        self.grand_row = self.grand_col = True
        self.fill_down = False
        self.values = []       # per data field: per record, the value it adds up (Excluded if left out)
        self.included = []     # per record: whether Numbers counts it
        self.data_on_rows = False
        self.value_columns = {}  # per data field: the Numbers table's columns showing it
        self.grid = None       # the cells of Excel's layout (see render)
        self.heights, self.widths = [], []
        self.styles = {}       # Numbers' look for each kind of cell (see regions)


# ---------------------------------------------------------------- reading the Numbers pivot


def uid_key(uid):
    return (uid.lower, uid.upper)


def uid_index(uid_map, columns):
    """{uid: index} from a ColumnRowUIDMapArchive."""
    axis = "column" if columns else "row"
    uids = getattr(uid_map, f"sorted_{axis}_uids")
    indexes = getattr(uid_map, f"{axis}_index_for_uid")
    return {uid_key(u): indexes[i] for i, u in enumerate(uids)}


def view_order(model, table, columns):
    """The table's rows (or columns) in the order Numbers shows them: their index in the table, or
    "summary" for a subtotal row, "grand" for the grand total; with sizes stored for the extra ones."""
    tm = model.objects[table._table_id]
    info = pivots.find_table_info(model, table._table_id)
    base = uid_index(model.objects[tm.base_column_row_uids.identifier], columns)
    view = uid_index(model.objects[info.view_column_row_uids.identifier], columns)
    order = []
    for uid in sorted(view, key=view.get):
        if uid in base:
            order.append(base[uid])
        else:
            order.append("grand" if uid == GRAND_UID else "summary")
    sizes = {}
    try:
        data_store = model.objects[info.summary_model.identifier].data_store
        bucket = data_store.columnHeaders.identifier if columns else data_store.rowHeaders.buckets[0].identifier
        sizes = {h.index: h.size for h in model.objects[bucket].headers if h.size}
    except Exception:
        pass
    return order, sizes


def build(book, table, base, tools):
    """The Excel PivotTable for a Numbers pivot table, given its pivots.Pivot. `tools` has what this
    needs from numbers_tool: text(cell), number_format(cell), auto_format(fmt, value), style(cell),
    table_style(table, style ref, text style ref), lines(table, rows, cols, header rows, header
    columns, footer rows), header_sizes(table, columns). Raises Unsupported."""
    model = table._model
    tm = model.objects[table._table_id]
    owner = base.owner
    if owner.is_empty_pivot or not hasattr(base, "source"):
        raise Unsupported("empty pivot")
    spec = Spec()
    spec.table = table
    spec.name = table.name or "PivotTable"
    row_rules = list(owner.grouping_columns_for_rows.group_column)
    col_rules = list(owner.grouping_columns_for_columns.group_column)
    rules = row_rules + col_rules
    if not owner.aggregate_columns.aggregates:
        raise Unsupported("no values")
    if len({uid_key(r.column_uid) for r in rules}) < len(rules):
        raise Unsupported("a column grouped twice")

    # The converted source table, and the snapshot rows the pivot last read.
    source = book.table_by_id.get(model.table_uuids_to_id(NumbersUUID(owner.source_table_uid).hex))
    if source is None:
        raise Unsupported("source table not found")
    source_tm = model.objects[source._table_id]
    if source_tm.HasField("pivot_owner") or source_tm.number_of_header_rows < 1:
        raise Unsupported("source table has no header row")
    if source._table_id in getattr(book, "groupings", {}):
        raise Unsupported("source table grouped into categories")  # its group rows would be read as data
    header_row = source_tm.number_of_header_rows - 1
    last_row = source.num_rows - source_tm.number_of_footer_rows - 1
    live_uids = uid_index(model.objects[source_tm.base_column_row_uids.identifier], True)
    snapshot_col = {live_uids[uid]: col for uid, col in base.column_for_uid.items() if uid in live_uids}
    rule_cols = []
    for rule in rules + list(owner.aggregate_columns.aggregates):
        uid = uid_key(rule.column_uid)
        if uid not in live_uids or uid not in base.column_for_uid:
            raise Unsupported("source column not found")
        rule_cols.append(live_uids[uid])
    names = [tools.text(source.cell(header_row, c)).strip() for c in range(source.num_cols)]
    first, last = 0, source.num_cols - 1
    if not all(names):  # Excel needs a name for every column it reads: read only the columns used
        first, last = min(rule_cols), max(rule_cols)
        if not all(names[first:last + 1]):
            raise Unsupported("a source column has no name")
    spec.source = SimpleNamespace(table=source, header_row=header_row, last_row=last_row, first=first, last=last)
    used = set()
    snapshot_rows = list(base.rows)
    spec.records = len(snapshot_rows)
    for c in range(first, last + 1):
        name = unique_name(names[c], used)
        sc = snapshot_col.get(c)
        values = [None if sc is None else excel_value(base.source.cell(r, sc), tools.text) for r in snapshot_rows]
        field = CacheField(name, values)
        field.column = c
        spec.fields.append(field)
    field_of = {f.column: i for i, f in enumerate(spec.fields)}

    # What Numbers shows: labels in view order, totals where they are.
    hdr_rows, hdr_cols = tm.number_of_header_rows, tm.number_of_header_columns
    row_order, row_sizes = view_order(model, table, False)
    col_order, col_sizes = view_order(model, table, True)
    if row_order[:hdr_rows] != list(range(hdr_rows)) or col_order[:hdr_cols] != list(range(hdr_cols)):
        raise Unsupported("unexpected header layout")
    if hdr_cols != max(len(row_rules), 1) or hdr_rows != len(col_rules) + 1:
        raise Unsupported("values listed in rows")  # Numbers shows its values side by side; this layout isn't known

    def label(r, c):
        cell = table.cell(r, c)
        return None if isinstance(cell, (EmptyCell, MergedCell)) else cell

    # Values: the label over each value column names its aggregate ("Units (Sum)", which also says
    # how it's summarised); several take turns under each column group, in the order of the rules.
    aggs = list(owner.aggregate_columns.aggregates)
    body_rows, body_cols = row_order[hdr_rows:], col_order[hdr_cols:]
    agg_row, previous, agg_labels, col_data = len(col_rules), None, {}, {}
    for pos, index in enumerate(body_cols):
        if not isinstance(index, int):
            continue
        text = tools.text(label(agg_row, index)).strip()
        if not text:
            if len(aggs) > 1 or previous is None:
                raise Unsupported("a value column without a label")
            text = previous
        previous = text
        col_data[pos] = agg_labels.setdefault(text, len(agg_labels))
    if len(agg_labels) != len(aggs):
        raise Unsupported("values don't match their labels")
    for index, (rule, name) in enumerate(zip(aggs, agg_labels)):
        if rule.show_as_type != 0:
            raise Unsupported("show-as values (percentages, running totals)")
        m = re.fullmatch(r"(.*)\s+\(([^()]+)\)", name)
        header = base.source_header(base.column(rule.column_uid))
        if not m or m.group(1).strip().casefold() != header.strip().casefold():
            raise Unsupported("a value's label doesn't name its column")
        kind = pivots.AGG_LABELS.get(m.group(2).strip().casefold())
        if kind not in SUBTOTALS or pivots.AGG_BY_CODE.get(rule.agg_type, kind) != kind:
            raise Unsupported("unknown summary")
        field = field_of.get(rule_cols[len(rules) + index])
        if field is None:
            raise Unsupported("value column outside the source range")
        spec.data.append(DataField(field, kind, name, None))

    # Records: which group each source row falls in for each row and column rule. Rows whose value
    # Numbers can't group are left out of the pivot (and hidden from Excel's).
    keys, spec.included = [], [True] * spec.records
    for rule, col in zip(rules, rule_cols):
        sc = base.column(rule.column_uid)
        rule_keys = []
        for i, r in enumerate(snapshot_rows):
            try:
                rule_keys.append(pivots.group_key(rule, base.source.cell(r, sc).value))
            except pivots.GroupError:
                rule_keys.append(Excluded)
                spec.included[i] = False
        keys.append(rule_keys)

    # Each rule's groups as Numbers shows them (labels carried down / across in view order). Under
    # each column group, one column per value, in order.
    row_paths = display_paths(body_rows, len(row_rules), lambda pos, level: label(pos, level), tools)
    col_paths = display_paths(body_cols, len(col_rules), lambda pos, level: label(level, pos), tools)
    runs = []
    for pos in sorted(col_paths):
        if col_paths[pos] is None:
            continue
        if runs and runs[-1][0] == col_paths[pos] and pos == runs[-1][1][-1] + 1:
            runs[-1][1].append(pos)
        else:
            runs.append((col_paths[pos], [pos]))
    if any([col_data[pos] for pos in positions] != list(range(len(aggs))) for _, positions in runs):
        raise Unsupported("values in an unexpected order")
    for pos, d in col_data.items():
        spec.value_columns.setdefault(d, []).append(body_cols[pos])

    axes = []
    for level, (rule, col) in enumerate(zip(rules, rule_cols)):
        on_rows = level < len(row_rules)
        paths = row_paths if on_rows else col_paths
        depth = level if on_rows else level - len(row_rules)
        shown = {}  # Numbers' group key -> its label, in order of appearance
        for path in paths.values():
            if path is not None and path[depth][0] not in shown:
                shown[path[depth][0]] = path[depth][1]
        found = {k for k, ok in zip(keys[level], spec.included) if ok}
        if set(shown) != found:
            raise Unsupported("labels don't match the source data")
        name_cell = label(hdr_rows - 1, depth) if on_rows else label(depth, hdr_cols - 1)
        sc = base.column(rule.column_uid)
        sample = next((base.source.cell(r, sc) for r in snapshot_rows if base.source.cell(r, sc).value not in (None, "")), None)
        axes.append(make_axis(spec, field_of[col], rule, keys[level], shown, tools.text(name_cell).strip(),
                              tools.number_format(sample) if sample is not None else None))
    spec.rows = axes[:len(row_rules)]
    spec.cols = axes[len(row_rules):]

    # The values each record adds up.
    for data in spec.data:
        spec.values.append([value if ok else Excluded for value, ok in zip(spec.fields[data.field].values, spec.included)])

    # Numbers' order and totals, read from what it shows.
    order_and_totals(spec, spec.rows, row_paths, body_rows)
    order_and_totals(spec, spec.cols, col_paths, body_cols, per=len(aggs))
    spec.grand_row, spec.grand_col = "grand" in body_rows, "grand" in body_cols
    spec.fill_down = fills_down(table, body_rows, len(row_rules))

    # Excel won't name a value like a field.
    names = {f.name.casefold() for f in spec.fields}
    for data in spec.data:
        while data.name.casefold() in names:
            data.name += " "

    # The layout, and the number format Numbers shows the values in.
    render(spec)
    for index, data in enumerate(spec.data):
        data.num_format = value_format(spec, index, base, tools)
    spec.heights, spec.widths = sizes(spec, tools, row_order, row_sizes, col_order, col_sizes)
    spec.styles = regions(spec, tools)
    return spec


class _Excluded:
    def __repr__(self):
        return "Excluded"


Excluded = _Excluded()  # a record Numbers leaves out of the pivot


def unique_name(name, used):
    """Excel's names for repeated column names: Units, Units2, Units3."""
    candidate, n = name, 2
    while candidate.casefold() in used:
        candidate = f"{name}{n}"
        n += 1
    used.add(candidate.casefold())
    return candidate


def display_paths(order, levels, cell, tools):
    """{position in view order: [(group key, label text)] for each level} for the body rows (or
    columns) Numbers shows, None for subtotal and grand total rows. Labels are carried down from the
    row above (Numbers shows each group's label once)."""
    paths, previous = {}, [None] * levels
    for pos, index in enumerate(order):
        if not isinstance(index, int):
            paths[pos] = None
            continue
        path = []
        for level in range(levels):
            c = cell(index, level)
            if c is not None and c.value not in (None, ""):
                path.append((pivots.norm(c.value), tools.text(c).strip()))
                previous[level + 1:] = [None] * (levels - level - 1)
            elif previous[level] is not None:
                path.append(previous[level])
            else:
                raise Unsupported("a group without a label")
            previous[level] = path[-1]
        paths[pos] = path
    return paths


def make_axis(spec, base_field, rule, keys, shown, name, num_format=None):
    """The Excel field for one of Numbers' row or column rules, named as Numbers names it: the
    source column itself when the rule groups by value, or a field grouping it (by a date part
    Excel groups by itself, or else by listing which values make up each group)."""
    field = spec.fields[base_field]
    shared = field.share()
    excel_of = {}  # shared index -> Numbers' group key (Excluded where Numbers can't group the value)
    for x, key in zip(field.index, keys):
        if excel_of.setdefault(x, key) != key:
            raise Unsupported("Excel would group values Numbers keeps apart")
    if not rule.HasField("grouping_functor"):
        if len({k for k in excel_of.values() if k is not Excluded}) != len([k for k in excel_of.values() if k is not Excluded]):
            raise Unsupported("Numbers groups values Excel keeps apart")
        x_of_key = {k: x for x, k in excel_of.items() if k is not Excluded}
        items = []
        for key, text in shown.items():
            x = x_of_key[key]
            value = shared[x]
            # Text shows as Numbers labels it; numbers and dates in the column's number format.
            caption = text if kind_of(value) == "s" and text != value.strip() else None
            items.append({"x": x, "caption": caption, "hidden": False, "value": value, "key": key, "label": text})
        listed = {item["x"] for item in items}
        for x, key in excel_of.items():  # Excel lists every value: those Numbers leaves out hidden
            if x not in listed:
                items.append({"x": x, "caption": None, "hidden": key is Excluded, "value": shared[x], "key": None,
                              "label": default_text(shared[x])})
        pos_of_x = {item["x"]: i for i, item in enumerate(items)}
        axis = Axis(base_field, items, [pos_of_x[x] if ok else None for x, ok in zip(field.index, spec.included)],
                    name if name and name != field.name else None)
        if any(kind_of(it["value"]) in ("n", "d") for it in items):
            axis.num_format = num_format or "General"
        if len({(it["caption"] or default_text(it["value"])).casefold() for it in items}) < len(items):
            raise Unsupported("two items would have the same name")
        axis.value_sort = [sort_key(it["value"]) for it in items]
        return axis

    # A grouping: a date part Excel groups by itself, or the groups listed value by value.
    group_field = CacheField(unique_name(name or f"{field.name} (group)", {f.name.casefold() for f in spec.fields}))
    group_field.group = native_group(rule, shared, excel_of) or discrete_group(shared, excel_of, shown)
    group_field.group.base = base_field
    field.parent = len(spec.fields)
    spec.fields.append(group_field)
    g = group_field.group
    items, pos_of_group = [], {}
    for key, text in shown.items():
        gi = g.item_of_key[key]
        pos_of_group[gi] = len(items)
        items.append({"x": gi, "caption": None if text == g.items[gi] else text, "hidden": False, "key": key,
                      "order": g.order[gi], "label": text})
    for gi in range(len(g.items)):
        if gi not in pos_of_group:
            pos_of_group[gi] = len(items)
            items.append({"x": gi, "caption": None, "hidden": gi in g.hidden, "key": None, "order": g.order[gi],
                          "label": g.items[gi]})
    record_keys = []
    for x, ok in zip(field.index, spec.included):
        record_keys.append(pos_of_group[g.mapping[x]] if ok else None)
    axis = Axis(len(spec.fields) - 1, items, record_keys)
    names = [(it["caption"] or g.items[it["x"]]).casefold() for it in items]
    if len(set(names)) < len(names):
        raise Unsupported("two groups would have the same name")
    axis.value_sort = [it["order"] for it in items]
    return axis


# Excel's own date groupings: the part of a date each one keeps, and the names it gives its items.
MONTHS = ["Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul", "Aug", "Sep", "Oct", "Nov", "Dec"]
DATE_PARTS = {
    "years": lambda d: d.year,
    "quarters": lambda d: (d.month - 1) // 3,
    "months": lambda d: d.month,
    "hours": lambda d: d.hour,
    "minutes": lambda d: d.minute,
    "seconds": lambda d: d.second,
}


def probe_dates(values):
    """Dates to compare a Numbers grouping with Excel's on: the data's own, and every day of a leap
    year and its neighbours at a few times of day."""
    dates = [v for v in values if isinstance(v, dt.datetime)]
    start = dt.datetime(2019, 1, 1)
    for day in range(0, 365 * 4 + 1, 1):
        for hour, minute, second in ((0, 0, 0), (13, 17, 41)):
            dates.append(start + dt.timedelta(days=day, hours=hour, minutes=minute, seconds=second))
    for minute in range(60):
        dates.append(dt.datetime(2020, 2, 29, 23, minute, minute))
    return dates


def native_group(rule, shared, excel_of):
    """The Excel date grouping (years, quarters, months, hours, minutes, seconds) that splits dates
    exactly as the Numbers rule does, if there is one and the column holds only dates (and blanks,
    which Excel puts in its "before the first date" item, hidden here as Numbers leaves them out)."""
    if not any(isinstance(v, dt.datetime) for v in shared):
        return None
    if any(v is not None and not isinstance(v, dt.datetime) for v in shared):
        return None
    if any(v is None and excel_of[x] is not Excluded for x, v in enumerate(shared)):
        return None
    nodes = list(rule.grouping_functor.formula.AST_node_array.AST_node)
    probes = probe_dates(shared)
    numbers_keys = []
    for d in probes:
        try:
            numbers_keys.append(pivots.norm(pivots.evaluate(nodes, d)))
        except (pivots.GroupError, Unsupported):
            return None
    for by, part in DATE_PARTS.items():
        to_excel, to_numbers = {}, {}
        if all(to_excel.setdefault(k, part(d)) == part(d) and to_numbers.setdefault(part(d), k) == k
               for d, k in zip(probes, numbers_keys)):
            break
    else:
        return None
    dates = [v for v in shared if isinstance(v, dt.datetime)]
    start, end = min(dates), max(dates)
    if by == "years":
        names = [str(y) for y in range(start.year, end.year + 1)]
        bucket = lambda d: d.year - start.year + 1  # noqa: E731
    elif by == "quarters":
        names = [f"Qtr{q}" for q in range(1, 5)]
        bucket = lambda d: (d.month - 1) // 3 + 1  # noqa: E731
    elif by == "months":
        names = list(MONTHS)
        bucket = lambda d: d.month  # noqa: E731
    elif by == "hours":
        names = [f"{(h % 12) or 12} {'AM' if h < 12 else 'PM'}" for h in range(24)]
        bucket = lambda d: d.hour + 1  # noqa: E731
    else:
        names = [f":{m:02d}" for m in range(60)]
        bucket = (lambda d: d.minute + 1) if by == "minutes" else (lambda d: d.second + 1)
    short = lambda d: f"{d.month}/{d.day}/{d.year}"  # noqa: E731
    items = [f"<{short(start)}"] + names + [f">{short(end)}"]
    mapping = [0 if v is None else bucket(v) for v in shared]
    item_of_key = {}
    for x, key in excel_of.items():
        if key is not Excluded:
            item_of_key[key] = mapping[x]
    hidden = {0} if any(v is None for v in shared) else set()
    return SimpleNamespace(kind="range", by=by, start=start, end=end, items=items, mapping=mapping,
                           item_of_key=item_of_key, hidden=hidden, order=list(range(len(items))))


def discrete_group(shared, excel_of, shown):
    """Groups listed value by value (Excel's "Group Selection"), named with Numbers' labels. Values
    Numbers leaves out get a group of their own, hidden."""
    items, item_of_key, hidden, mapping = [], {}, set(), []
    for key, text in shown.items():
        item_of_key[key] = len(items)
        items.append(text)
    used = {t.casefold() for t in items}
    for x, value in enumerate(shared):
        key = excel_of[x]
        if key is Excluded:
            name = unique_name(default_text(value), used)
            hidden.add(len(items))
            mapping.append(len(items))
            items.append(name)
        else:
            mapping.append(item_of_key[key])
    return SimpleNamespace(kind="discrete", items=items, mapping=mapping, item_of_key=item_of_key, hidden=hidden,
                           order=[(1, 0, t.casefold()) for t in items])


def order_and_totals(spec, axes, paths, order, per=1):
    """Puts each field's items in the order Numbers shows them, and works out from where Numbers
    shows subtotals which fields have them (`per` subtotal columns for each, one per value)."""
    if not axes:
        return
    shown = [paths[pos] for pos in sorted(paths) if paths[pos] is not None]
    for level, axis in enumerate(axes):
        # The order Numbers shows each parent group's groups in at this level.
        within = {}
        for path in shown:
            keys = within.setdefault(tuple(p[0] for p in path[:level]), [])
            if path[level][0] not in keys:
                keys.append(path[level][0])
        index = {it["key"]: i for i, it in enumerate(axis.items) if it["key"] is not None}
        by_name = sorted(index, key=lambda k: axis.value_sort[index[k]])
        rank = {k: r for r, k in enumerate(by_name)}
        listed = by_name
        if all(keys == sorted(keys, key=rank.get) for keys in within.values()):
            axis.sort = "ascending"
        elif all(keys == sorted(keys, key=lambda k: -rank[k]) for keys in within.values()):
            axis.sort, listed = "descending", by_name[::-1]
        else:
            found = value_order(spec, axes, level, within, rank)
            if found:
                axis.sort, axis.sort_by = found
            else:  # as listed: one order every parent group's groups follow
                listed = merged(within.values())
        if axis.sort_by is None and any(axis.items[index[k]]["caption"] for k in listed):
            axis.sort = None  # Excel would sort these by their captions, which can differ from Numbers' order
            listed = merged(within.values())
        new_order = [index[k] for k in listed]
        new_order += [i for i in range(len(axis.items)) if i not in set(new_order)]
        remap = {old: new for new, old in enumerate(new_order)}
        axis.items = [axis.items[i] for i in new_order]
        axis.keys = [None if k is None else remap[k] for k in axis.keys]
        axis.value_sort = [axis.value_sort[i] for i in new_order]

    # Subtotals: which fields' groups end with summary rows (deepest first) at each boundary between
    # groups, and at the end, where the grand totals follow them.
    levels = len(axes)
    boundaries = []  # (levels whose groups end here, summary rows found, grand totals expected)
    count, previous = 0, None
    for pos in sorted(paths):
        path = paths[pos]
        if path is None:
            count += 1
            continue
        if previous is not None and path != previous:
            differ = next(i for i in range(levels) if i == levels - 1 or path[i][0] != previous[i][0])
            boundaries.append((set(range(differ, levels - 1)), count, 0))
        elif count:
            raise Unsupported("a summary inside a group")
        count, previous = 0, path
    boundaries.append((set(range(levels - 1)), count, per if "grand" in order else 0))
    choices = []
    for mask in range(2 ** max(levels - 1, 0)):
        on = {i for i in range(levels - 1) if mask >> i & 1}
        if all(per * len(ends & on) + grand == found for ends, found, grand in boundaries):
            choices.append(on)
    if not choices:
        raise Unsupported("subtotals Excel can't place")
    on = max(choices, key=lambda s: sorted(s, reverse=True))
    for level, axis in enumerate(axes):
        axis.subtotals = level in on or level == levels - 1


def value_order(spec, axes, level, within, rank):
    """(direction, data field) if Numbers sorts this field's groups by their totals of a value, as
    Excel's "sort by value" does (within each parent group, blank totals last, ties by name)."""
    records = {}
    for i, ok in enumerate(spec.included):
        if ok:
            path = tuple(axes[lv].items[axes[lv].keys[i]]["key"] for lv in range(level + 1))
            records.setdefault(path, []).append(i)
    for d in range(len(spec.data)):
        totals = {path: aggregate(spec, d, recs) for path, recs in records.items()}
        for direction in ("ascending", "descending"):
            def key(parent, k):
                value = totals.get(parent + (k,))
                return (1, 0, rank[k]) if value is None else (0, -value if direction == "descending" else value, rank[k])
            if all(keys == sorted(keys, key=lambda k: key(parent, k)) for parent, keys in within.items()):
                return direction, d
    return None


def merged(orders):
    """One order of all the keys that keeps each of the given orders, or Unsupported."""
    after, first = {}, {}  # first: where each key first appears, which breaks ties
    for order in orders:
        for k in order:
            first.setdefault(k, len(first))
        for a, b in zip(order, order[1:]):
            after.setdefault(a, set()).add(b)
    before = {k: 0 for k in first}
    for a in after:
        for b in after[a]:
            before[b] += 1
    result, ready = [], [k for k in first if not before[k]]
    while ready:
        k = ready.pop(0)
        result.append(k)
        for b in after.get(k, ()):
            before[b] -= 1
            if not before[b]:
                ready.append(b)
        ready.sort(key=first.get)
    if len(result) < len(first):
        raise Unsupported("groups in different orders under different parents")
    return result


def fills_down(table, body_rows, levels):
    """Whether Numbers repeats each group's label on every row ("Repeat group names")."""
    if levels < 2:
        return False
    rows = [r for r in body_rows if isinstance(r, int)]
    return len(rows) > 1 and all(not isinstance(table.cell(r, 0), (EmptyCell, MergedCell)) for r in rows)


# ---------------------------------------------------------------- Excel's tabular layout


class Grid:
    """The cells of a PivotTable in Excel's tabular form: {(row, col): Cell} relative to its top-left
    cell, and the row and column items Excel lists for it."""

    def __init__(self):
        self.cells = {}
        self.rows = self.cols = 0
        self.first_header_row = self.first_data_row = self.first_data_col = 0
        self.row_items, self.col_items = [], []


class Cell:
    """value: text, a number or None; role: caption, button, label, total-label or value; region:
    header, label, body, subtotal or grand; data: the data field of a value."""

    def __init__(self, value, role, region, data=None):
        self.value, self.role, self.region, self.data = value, role, region, data
        self.num_format = self.axis = None  # for a label holding a number or date


def axis_entries(spec, axes, totals, on_rows):
    """Excel's items for one axis: ("data", path, data index) for each row (column) of values, with
    ("default", path, data index) after a group with subtotals and ("grand", (), data index) for the
    grand total. A path holds each field's item position (the data index for the Values field).
    `totals` has the records under each combination of leading items (see index_records)."""
    entries = []
    real = [a for a in axes if a is not DATA]
    has_data = DATA in axes
    n = len(spec.data)
    records = [i for i, ok in enumerate(spec.included) if ok]

    def walk(level, path, recs):
        if level == len(axes):
            entries.append(("data", tuple(path), path_data(axes, path)))
            return
        axis = axes[level]
        if axis is DATA:
            for d in range(n):
                walk(level + 1, path + [d], recs)
            return
        groups = {}
        for i in recs:
            groups.setdefault(axis.keys[i], []).append(i)
        present = [pos for pos, item in enumerate(axis.items) if pos in groups and not item["hidden"]]
        if axis.sort_by is not None:
            parent = tuple(p for lv, p in enumerate(path) if axes[lv] is not DATA)
            present = by_value(spec, axis, present, parent, totals, on_rows)
        for pos in present:
            walk(level + 1, path + [pos], groups[pos])
            later = [a for a in axes[level + 1:] if a is not DATA]
            if axis.subtotals and later:
                if has_data and DATA in axes[level + 1:]:
                    for d in range(n):
                        entries.append(("default", tuple(path + [pos]), d))
                else:
                    entries.append(("default", tuple(path + [pos]), 0))

    if axes:
        walk(0, [], records)
    grand = spec.grand_row if on_rows else spec.grand_col
    if real and grand:
        for d in (range(n) if has_data else [0]):
            entries.append(("grand", (), d))
    return entries or [("blank", (), 0)]


def by_value(spec, axis, positions, parent, totals, on_rows):
    """Items sorted by their totals of a value (Excel's "sort by value") within their parent group:
    blank totals last, and equal totals in the items' order."""
    def total(pos):
        key = (parent + (pos,), ()) if on_rows else ((), parent + (pos,))
        value = aggregate(spec, axis.sort_by, totals.get(key, []))
        return (1, 0) if value is None else (0, -value if axis.sort == "descending" else value)
    return sorted(positions, key=total)


def path_data(axes, path):
    return path[axes.index(DATA)] if DATA in axes else 0


def leading(axes, kind, path):
    """The items, field by field, of the records a row (or column) item covers."""
    if kind in ("grand", "blank"):
        return ()
    return tuple(p for level, p in enumerate(path) if axes[level] is not DATA)


def index_records(spec, rows, cols):
    """{(leading row items, leading column items): [records]} for every combination of each
    record's leading items, so each cell's records are looked up rather than searched for."""
    real_rows, real_cols = [a for a in rows if a is not DATA], [a for a in cols if a is not DATA]
    totals = {}
    for i, ok in enumerate(spec.included):
        if not ok:
            continue
        rk, ck = tuple(a.keys[i] for a in real_rows), tuple(a.keys[i] for a in real_cols)
        for lr in range(len(rk) + 1):
            for lc in range(len(ck) + 1):
                totals.setdefault((rk[:lr], ck[:lc]), []).append(i)
    return totals


def render(spec):
    """Lays the PivotTable out as Excel does in tabular form (each row field in its own column,
    subtotals below their group), with the value of every cell."""
    n = len(spec.data)
    rows = list(spec.rows) + ([DATA] if n > 1 and spec.data_on_rows else [])
    cols = list(spec.cols) + ([DATA] if n > 1 and not spec.data_on_rows else [])
    spec.rows_axis, spec.cols_axis = rows, cols
    g = Grid()
    real_cols = [a for a in cols if a is not DATA]
    n_r, n_c = len(rows), len(cols)
    g.first_data_col = n_r if n_r else (1 if real_cols and n == 1 else 0)
    if real_cols:
        g.first_header_row, g.first_data_row = 1, 1 + n_c
    elif cols:  # only the Values field: its names are the header
        g.first_header_row, g.first_data_row = 0, 1
    else:
        g.first_header_row, g.first_data_row = 1, 1
    totals = index_records(spec, rows, cols)
    row_entries = axis_entries(spec, rows, totals, True)
    col_entries = axis_entries(spec, cols, totals, False)
    g.rows = g.first_data_row + len(row_entries)
    g.cols = g.first_data_col + len(col_entries)
    g.row_items = items_xml_data(row_entries, rows)
    g.col_items = items_xml_data(col_entries, cols)
    put = g.cells.__setitem__

    def name_of(axis):
        if axis is DATA:
            return "Values"
        return axis.caption or spec.fields[axis.field].name

    def caption(axis, pos):
        """An item's label: its text, or a number shown in the field's number format."""
        item = axis.items[pos]
        if item["caption"] is not None:
            return item["caption"]
        field = spec.fields[axis.field]
        if field.group is not None:
            return field.group.items[item["x"]]
        value = field.shared[item["x"]]
        if axis.num_format is not None and kind_of(value) in ("n", "d"):
            return serial(value) if isinstance(value, dt.datetime) else value
        return default_text(value)

    def text_of(axis, pos):
        """An item's label as text, for "X Total"."""
        label = caption(axis, pos)
        return axis.items[pos].get("label", default_text(label)) if not isinstance(label, str) else label

    def item_text(axis, pos):
        return spec.data[pos].name if axis is DATA else caption(axis, pos)

    def label_cell(axis, pos, region):
        value = item_text(axis, pos)
        cell = Cell(value, "label", region)
        if not isinstance(value, str):
            cell.num_format, cell.axis = axis.num_format, axis
        return cell

    # Captions and field buttons.
    if real_cols:
        if n == 1 and n_r:
            put((0, 0), Cell(spec.data[0].name, "caption", "header"))
        for k, axis in enumerate(cols):
            put((0, g.first_data_col + k), Cell(name_of(axis), "button", "header"))
    elif cols:
        for d in range(n):
            put((0, g.first_data_col + d), Cell(spec.data[d].name, "label", "header"))
    elif n == 1:
        put((0, g.first_data_col), Cell(spec.data[0].name, "caption", "header"))
    for k, axis in enumerate(rows):
        put((g.first_data_row - 1, k), Cell(name_of(axis), "button", "header"))

    # Column labels.
    previous = None
    for j, (kind, path, d) in enumerate(col_entries):
        c = g.first_data_col + j
        if not real_cols:
            break
        if kind == "data":
            start = 0 if previous is None else common(previous, path)
            for level in range(start, len(path)):
                put((1 + level, c), label_cell(cols[level], path[level], "header"))
            previous = path
        elif kind == "default":
            level = len(path) - 1
            text = text_of(cols[level], path[level])
            put((1 + level, c), Cell(f"{text} {spec.data[d].name}" if DATA in cols else f"{text} Total", "total-label", "header"))
        elif kind == "grand":
            put((1, c), Cell(f"Total {spec.data[d].name}" if DATA in cols else "Grand Total", "total-label", "header"))

    # Row labels.
    previous = None
    for i, (kind, path, d) in enumerate(row_entries):
        r = g.first_data_row + i
        if kind == "data":
            start = 0 if previous is None else common(previous, path)
            if spec.fill_down:
                start = 0
            for level in range(len(path)):
                if level >= start or rows[level] is DATA:
                    put((r, level), label_cell(rows[level], path[level], "label"))
            previous = path
        elif kind == "default":
            level = len(path) - 1
            text = text_of(rows[level], path[level])
            put((r, level), Cell(f"{text} {spec.data[d].name}" if DATA in rows else f"{text} Total", "total-label", "subtotal"))
        elif kind == "grand":
            put((r, 0), Cell(f"Total {spec.data[d].name}" if DATA in rows else "Grand Total", "total-label", "grand"))
        elif kind == "blank" and real_cols and n == 1 and g.first_data_col:
            put((r, 0), Cell(spec.data[0].name, "label", "label"))

    # Values.
    for i, (rkind, rpath, rd) in enumerate(row_entries):
        rkey = leading(rows, rkind, rpath)
        for j, (ckind, cpath, cd) in enumerate(col_entries):
            d = rd if DATA in rows else cd if DATA in cols else 0
            value = aggregate(spec, d, totals.get((rkey, leading(cols, ckind, cpath)), []))
            if rkind == "grand" or ckind == "grand":
                region = "grand"
            elif rkind == "default" or ckind == "default":
                region = "subtotal"
            else:
                region = "body"
            put((g.first_data_row + i, g.first_data_col + j), Cell(value, "value", region, d))
    spec.grid = g
    spec.row_entries, spec.col_entries = row_entries, col_entries


def common(a, b):
    n = 0
    while n < min(len(a), len(b)) and a[n] == b[n]:
        n += 1
    return n


def aggregate(spec, d, records):
    """A data field's total over some records, as Excel's PivotTable works it out: None (a blank
    cell) when there's nothing to total, and a count of 0 is blank too."""
    data = spec.data[d]
    values = [spec.values[d][i] for i in records]
    if not values:
        return None
    if data.kind == "count":
        count = sum(1 for v in values if v is not None)
        return count or None
    if any(isinstance(v, (Error, bool)) for v in values):
        raise Unsupported("errors or checkboxes in the values")  # Numbers and Excel total these differently
    nums = [serial(v) if isinstance(v, dt.datetime) else v for v in values if kind_of(v) in ("n", "d")]
    if not nums:
        return None
    return pivots.aggregate(data.kind, nums)


def items_xml_data(entries, axes):
    """Each entry as Excel lists it: (type, repeated leading fields, data index, [item positions])."""
    out, previous = [], None
    for kind, path, d in entries:
        if kind == "blank":
            out.append(("blank", 0, 0, []))
        elif kind == "data":
            r = 0 if previous is None else common(previous, path)
            out.append(("data", r, d, list(path[r:])))
            previous = path
        elif kind == "default":
            level = len(path) - 1
            out.append(("default", level, d, [path[level]]))
        else:
            out.append(("grand", 0, d, [0]))
    return out


# ---------------------------------------------------------------- formats, sizes and looks


def value_format(spec, d, base, tools):
    """The number format of a data field: the one Numbers shows the pivot's values in (or else its
    source column's). Where that shows as many decimal places as each value needs, which a single
    Excel format can't, it gets as many as the values need, up to two more than the source data's."""
    data = spec.data[d]
    if data.kind == "count":
        return None
    table = spec.table
    tm = table._model.objects[table._table_id]
    shown = [table.cell(r, c) for c in spec.value_columns.get(d, []) for r in range(tm.number_of_header_rows, table.num_rows)]
    source = base.format_cell(base.column(base.owner.aggregate_columns.aggregates[d].column_uid))
    for cell in [c for c in shown if isinstance(c, NumberCell)] + ([source] if source is not None else []):
        fmt_id = cell._currency_format_id if cell._currency_format_id is not None else cell._num_format_id
        if not isinstance(cell, NumberCell) or fmt_id is None:
            continue
        fmt = cell._model.table_format(cell._table_id, fmt_id)
        if fmt.HasField("custom_uid"):
            return tools.number_format(cell)
        return automatic(spec, d, tools, fmt)
    if source is not None and tools.number_format(source):
        return tools.number_format(source)
    # Numbers shows pivot values with thousands separators and the decimal places each needs.
    return automatic(spec, d, tools, tools.default_format)


def automatic(spec, d, tools, fmt):
    """The format for a data field's values (see value_format)."""
    def places(code):
        return len(re.search(r"\.(0*)", code).group(1)) if code and "." in code else 0
    values = [c.value for c in spec.grid.cells.values() if c.role == "value" and c.data == d and isinstance(c.value, (int, float))]
    sources = [v for v, ok in zip(spec.values[d], spec.included) if ok and isinstance(v, float)]
    limit = max([places(tools.auto_format(fmt, v)) for v in sources] + [0]) + 2
    best = max((tools.auto_format(fmt, v) for v in values), key=places, default=tools.auto_format(fmt, 0))
    if places(best) > limit:
        probe = 1 + 10 ** -limit
        best = tools.auto_format(fmt, probe / 100 if "%" in (best or "") else probe)
    return best


def sizes(spec, tools, row_order, row_sizes, col_order, col_sizes):
    """Row heights and column widths in points: Numbers' own, row for row and column for column
    (the view's subtotal and grand total rows and columns included), the first header row's for any
    header row Excel adds."""
    table = spec.table
    tm = table._model.objects[table._table_id]
    heights, widths = tools.header_sizes(table, False), tools.header_sizes(table, True)

    def size(order, stored, base_sizes, default, pos):
        index = order[pos]
        if isinstance(index, int):
            return base_sizes[index]
        return stored.get(pos) or default

    g = spec.grid
    hdr_rows, hdr_cols = tm.number_of_header_rows, tm.number_of_header_columns
    row_heights = []
    for r in range(g.first_data_row):
        numbers_row = hdr_rows - g.first_data_row + r
        row_heights.append(heights[max(0, numbers_row)])
    body = list(range(hdr_rows, len(row_order)))
    for i in range(g.rows - g.first_data_row):
        row_heights.append(size(row_order, row_sizes, heights, tm.default_row_height, body[i]) if i < len(body)
                           else tm.default_row_height)
    col_widths = []
    for c in range(g.first_data_col):
        col_widths.append(widths[min(c, hdr_cols - 1)])
    body = list(range(hdr_cols, len(col_order)))
    for j in range(g.cols - g.first_data_col):
        col_widths.append(size(col_order, col_sizes, widths, tm.default_column_width, body[j]) if j < len(body)
                          else tm.default_column_width)
    return [round(h, 1) for h in row_heights], [round(w, 1) for w in col_widths]


def regions(spec, tools):
    """Numbers' look for each kind of cell: {region: style dict as numbers_tool.cell_style gives}.
    Header cells look like the table's header row, row labels like its header column, values like
    its body; subtotal and grand total rows and columns use the pivot's summary styles."""
    table = spec.table
    tm = table._model.objects[table._table_id]
    hdr_rows, hdr_cols = tm.number_of_header_rows, tm.number_of_header_columns
    body_row = hdr_rows if hdr_rows < table.num_rows else None

    def from_cell(r, c):
        if r is None or r >= table.num_rows or c >= table.num_cols:
            return {}
        return {k: v for k, v in tools.style(table.cell(r, c)).items() if k not in ("bt", "bb", "bl", "br")}

    styles = {
        "header": from_cell(0, hdr_cols if hdr_cols < table.num_cols else 0),
        "label": from_cell(body_row, 0),
        "body": from_cell(body_row, min(hdr_cols, table.num_cols - 1)),
    }
    styles["subtotal"] = {**styles["body"], **tools.table_style(table, "pivot_body_summary_row_style", "category_level_1_text_style")}
    styles["subtotal-label"] = {**styles["label"], **tools.table_style(table, "pivot_header_column_summary_style", "category_level_1_text_style")}
    styles["subtotal-col"] = {**styles["body"], **tools.table_style(table, "pivot_body_summary_column_style", "category_level_1_text_style")}
    styles["grand"] = {**styles["body"], **tools.table_style(table, "footer_row_style", "footer_row_text_style")}
    grand_label = tools.table_style(table, "footer_row_style", "footer_row_text_style")
    if "bg" not in grand_label and "bg" in styles["label"]:
        grand_label["bg"] = styles["label"]["bg"]
    styles["grand-label"] = {**styles["label"], **grand_label}
    for key in ("subtotal", "subtotal-col", "grand"):
        styles[key].pop("ha", None)
    return styles


def cell_style(spec, r, c, cell):
    """The look of one cell of the layout, with the table's lines drawn as Numbers would for a table
    of this size (header rows and columns, the grand total row as its footer)."""
    g = spec.grid
    data_row = r >= g.first_data_row
    if not data_row:
        style = dict(spec.styles["header"])
    elif c < g.first_data_col:
        kind = spec.row_entries[r - g.first_data_row][0]
        style = dict(spec.styles["subtotal-label" if kind == "default" else "grand-label" if kind == "grand" else "label"])
    else:
        rkind = spec.row_entries[r - g.first_data_row][0]
        ckind = spec.col_entries[c - g.first_data_col][0]
        if rkind == "grand" or (ckind == "grand" and rkind != "default"):
            style = dict(spec.styles["grand"] if rkind == "grand" else spec.styles["subtotal-col"])
        elif rkind == "default":
            style = dict(spec.styles["subtotal"])
        elif ckind in ("default", "grand"):
            style = dict(spec.styles["subtotal-col"])
        else:
            style = dict(spec.styles["body"])
        if cell is not None and isinstance(cell.value, (int, float)) and "ha" not in style:
            style["ha"] = "right"
    lines = spec.lines
    for key, value in (("bt", lines[0][r][c]), ("bb", lines[0][r + 1][c]), ("bl", lines[1][r][c]), ("br", lines[1][r][c + 1])):
        if value is not None:
            style[key] = list(value)
    return {key: value for key, value in style.items() if value not in (None, 0, "")}


def prepare_lines(spec, tools):
    g = spec.grid
    footer = 1 if spec.grand_row and spec.rows else 0
    spec.lines = tools.lines(spec.table, g.rows, g.cols, g.first_data_row, g.first_data_col, footer)


def cells(spec, tools):
    """Every cell of the layout: (row, col, value, style dict, number format) relative to the
    PivotTable's top-left cell, the value None for an empty cell."""
    if not hasattr(spec, "lines"):
        prepare_lines(spec, tools)
    g = spec.grid
    for r in range(g.rows):
        for c in range(g.cols):
            cell = g.cells.get((r, c))
            num_format = spec.data[cell.data].num_format if cell is not None and cell.role == "value" else (
                cell.num_format if cell is not None else None)
            yield r, c, None if cell is None else cell.value, cell_style(spec, r, c, cell), num_format, cell


def write(ws, spec, formats, tools, button):
    """Writes the PivotTable's cells (ws is offset to its top-left cell); field buttons get the
    format `button(style)` gives, which add_parts gives Excel's drop-down arrow. Returns what
    add_parts needs: the formats used, {"buttons": [...], data index: a format of its values}."""
    used = {"buttons": []}
    for r, c, value, style, num_format, cell in cells(spec, tools):
        fmt = formats.get(style, num_format)
        if cell is not None and cell.role == "button":
            fmt = button(style)
            used["buttons"].append(fmt)
        elif cell is not None and cell.role == "value":
            used.setdefault(cell.data, fmt)
        elif cell is not None and cell.num_format is not None:
            used.setdefault(("axis", cell.axis.field), fmt)
        if value is None:
            ws.write_blank(r, c, None, fmt)
        elif isinstance(value, str):
            ws.write_string(r, c, value, fmt)
        else:
            ws.write_number(r, c, value, fmt)
    return used


# ---------------------------------------------------------------- the file's parts


def number(value):
    if isinstance(value, float) and value.is_integer() and abs(value) < 1e15:
        return str(int(value))
    return repr(float(value))


def iso(value):
    return value.strftime("%Y-%m-%dT%H:%M:%S")


def attr(value):
    return escape(str(value), {'"': "&quot;", "\n": "&#10;", "\r": "&#13;", "\t": "&#9;"})


def value_xml(value):
    kind = kind_of(value)
    if kind == "m":
        return "<m/>"
    if kind == "n":
        return f'<n v="{number(value)}"/>'
    if kind == "d":
        return f'<d v="{iso(value)}"/>'
    if kind == "b":
        return f'<b v="{int(value)}"/>'
    if kind == "e":
        return f'<e v="{attr(value)}"/>'
    return f'<s v="{attr(value)}"/>'


def shared_items_xml(values, listed):
    """<sharedItems> with the flags Excel writes for what the field holds, and the items when
    records refer to them by index."""
    kinds = {kind_of(v) for v in values}
    nums = [v for v in values if kind_of(v) == "n"]
    dates = [v for v in values if kind_of(v) == "d"]
    has = kinds.__contains__
    attrs = []
    if not (has("s") or has("m") or has("b") or has("e")):
        attrs.append('containsSemiMixedTypes="0"')
    if (has("d")) and not (has("s") or has("n") or has("b") or has("e")):
        attrs.append('containsNonDate="0"')
    if has("d"):
        attrs.append('containsDate="1"')
    if not (has("s") or has("b") or has("e")):
        attrs.append('containsString="0"')
    if has("m"):
        attrs.append('containsBlank="1"')
    if sum(map(has, ("s", "n", "d", "b"))) > 1:
        attrs.append('containsMixedTypes="1"')
    if nums:
        attrs.append('containsNumber="1"')
        if all(float(v).is_integer() for v in nums):
            attrs.append('containsInteger="1"')
        attrs.append(f'minValue="{number(min(nums))}" maxValue="{number(max(nums))}"')
    if dates:
        attrs.append(f'minDate="{iso(min(dates))}" maxDate="{iso(max(dates))}"')
    if any(kind_of(v) == "s" and len(v) > 255 for v in values):
        attrs.append('longText="1"')
    if listed is None:
        return f'<sharedItems {" ".join(attrs)}/>' if attrs else "<sharedItems/>"
    attrs.append(f'count="{len(listed)}"')
    return f'<sharedItems {" ".join(attrs)}>' + "".join(value_xml(v) for v in listed) + "</sharedItems>"


def cache_definition_xml(spec, sheet, ref):
    fields = []
    for field in spec.fields:
        if field.group is not None:
            g = field.group
            if g.kind == "range":
                inner = (f'<rangePr groupBy="{g.by}" startDate="{iso(g.start)}" endDate="{iso(g.end)}"/>')
            else:
                inner = f'<discretePr count="{len(g.mapping)}">' + "".join(f'<x v="{m}"/>' for m in g.mapping) + "</discretePr>"
            items = "".join(f'<s v="{attr(name)}"/>' for name in g.items)
            fields.append(f'<cacheField name="{attr(field.name)}" numFmtId="0" databaseField="0">'
                          f'<fieldGroup base="{g.base}">{inner}<groupItems count="{len(g.items)}">{items}</groupItems>'
                          "</fieldGroup></cacheField>")
            continue
        shared = shared_items_xml(field.values, field.shared)
        parent = f'<fieldGroup par="{field.parent}"/>' if field.parent is not None else ""
        fields.append(f'<cacheField name="{attr(field.name)}" numFmtId="0">{shared}{parent}</cacheField>')
    return ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
            f'<pivotCacheDefinition xmlns="{MAIN_NS}" xmlns:r="{REL_NS}" r:id="rId1" refreshOnLoad="0" '
            f'createdVersion="6" refreshedVersion="6" minRefreshableVersion="3" recordCount="{spec.records}">'
            f'<cacheSource type="worksheet"><worksheetSource ref="{ref}" sheet="{attr(sheet)}"/></cacheSource>'
            f'<cacheFields count="{len(spec.fields)}">' + "".join(fields) + "</cacheFields>"
            f'<extLst><ext uri="{{725AE2AE-9491-48be-B2B4-4EB974FC3084}}" xmlns:x14="{X14_NS}">'
            "<x14:pivotCacheDefinition/></ext></extLst></pivotCacheDefinition>")


def cache_records_xml(spec):
    rows = []
    stored = [f for f in spec.fields if f.group is None]
    for i in range(spec.records):
        parts = []
        for field in stored:
            parts.append(f'<x v="{field.index[i]}"/>' if field.shared is not None else value_xml(field.values[i]))
        rows.append("<r>" + "".join(parts) + "</r>")
    return ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
            f'<pivotCacheRecords xmlns="{MAIN_NS}" xmlns:r="{REL_NS}" count="{spec.records}">'
            + "".join(rows) + "</pivotCacheRecords>")


def items_xml(items):
    out = []
    for kind, repeat, d, positions in items:
        if kind == "blank":
            out.append("<i/>")
            continue
        attrs = ""
        if kind != "data":
            attrs += f' t="{kind}"'
        if repeat:
            attrs += f' r="{repeat}"'
        if d:
            attrs += f' i="{d}"'
        xs = "".join("<x/>" if p == 0 else f'<x v="{p}"/>' for p in positions)
        out.append(f"<i{attrs}>{xs}</i>")
    return f'count="{len(items)}">' + "".join(out)


def pivot_table_xml(spec, name, cache_id, top_left):
    g = spec.grid
    axis_of, axes = {}, {}
    for kind, axes_list in (("axisRow", spec.rows), ("axisCol", spec.cols)):
        for axis in axes_list:
            axis_of[axis.field] = kind
            axes[axis.field] = axis
    data_fields = {d.field for d in spec.data}
    fields = []
    for index, field in enumerate(spec.fields):
        attrs = ""
        axis = axes.get(index)
        if axis is not None:
            attrs += f' axis="{axis_of[index]}"'
        if index in data_fields:
            attrs += ' dataField="1"'
        if axis is not None and axis.caption:
            attrs += f' name="{attr(axis.caption)}"'
        attrs += ' compact="0"'
        if axis is not None and axis.num_fmt_id:
            attrs += f' numFmtId="{axis.num_fmt_id}"'
        attrs += ' outline="0" showAll="0"'
        if axis is not None and axis.sort:
            attrs += f' sortType="{axis.sort}"'
        if axis is not None and not axis.subtotals:
            attrs += ' defaultSubtotal="0"'
        inner = ""
        if axis is not None:
            listed = []
            for item in axis.items:
                a = ""
                if item["caption"] is not None:
                    a += f' n="{attr(item["caption"])}"'
                if item["hidden"]:
                    a += ' h="1"'
                listed.append(f'<item{a} x="{item["x"]}"/>')
            if axis.subtotals:
                listed.append('<item t="default"/>')
            inner = f'<items count="{len(listed)}">' + "".join(listed) + "</items>"
            if axis.sort_by is not None:
                inner += ('<autoSortScope><pivotArea dataOnly="0" outline="0" fieldPosition="0"><references count="1">'
                          f'<reference field="4294967294" count="1" selected="0"><x v="{axis.sort_by}"/></reference>'
                          "</references></pivotArea></autoSortScope>")
        elif field.shared is not None or field.group is not None:
            count = len(field.shared) if field.shared is not None else len(field.group.items)
            listed = [f'<item x="{x}"/>' for x in range(count)] + ['<item t="default"/>']
            inner = f'<items count="{len(listed)}">' + "".join(listed) + "</items>"
        if spec.fill_down and axis is not None and axis_of[index] == "axisRow":
            inner += (f'<extLst><ext uri="{{2946ED86-A175-432a-8AC1-64E0C546D7DE}}" xmlns:x14="{X14_NS}">'
                      '<x14:pivotField fillDownLabels="1"/></ext></extLst>')
        fields.append(f"<pivotField{attrs}>{inner}</pivotField>" if inner else f"<pivotField{attrs}/>")

    def field_list(tag, axes_list):
        if not axes_list:
            return ""
        xs = "".join(f'<field x="{-2 if a is DATA else a.field}"/>' for a in axes_list)
        return f'<{tag} count="{len(axes_list)}">{xs}</{tag}>'

    top, left = top_left
    ref = f"{cell_ref(top, left)}:{cell_ref(top + g.rows - 1, left + g.cols - 1)}"
    data = []
    for d in spec.data:
        a = f' name="{attr(d.name)}" fld="{d.field}"'
        if SUBTOTALS[d.kind]:
            a += f' subtotal="{SUBTOTALS[d.kind]}"'
        a += ' baseField="0" baseItem="0"'
        if d.num_fmt_id:
            a += f' numFmtId="{d.num_fmt_id}"'
        data.append(f"<dataField{a}/>")
    totals = ""
    if not spec.grand_row:
        totals += ' rowGrandTotals="0"'
    if not spec.grand_col:
        totals += ' colGrandTotals="0"'
    formats = (f'<formats count="1"><format dxfId="{spec.font_dxf}"><pivotArea type="all" dataOnly="0" outline="0" '
               'fieldPosition="0"/></format></formats>')
    fill_down = ' fillDownLabelsDefault="1"' if spec.fill_down else ""
    data_on_rows = ' dataOnRows="1"' if spec.data_on_rows and len(spec.data) > 1 else ""
    ext = f'<x14:pivotTableDefinition{fill_down} hideValuesRow="1"/>'
    return ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
            f'<pivotTableDefinition xmlns="{MAIN_NS}" name="{attr(name)}" cacheId="{cache_id}"'
            f'{data_on_rows} applyNumberFormats="0" '
            'applyBorderFormats="0" applyFontFormats="0" applyPatternFormats="0" applyAlignmentFormats="0" '
            'applyWidthHeightFormats="1" dataCaption="Values" updatedVersion="6" minRefreshableVersion="3" '
            f'useAutoFormatting="0"{totals} itemPrintTitles="1" createdVersion="6" indent="0" compact="0" '
            'compactData="0" multipleFieldFilters="0">'
            f'<location ref="{ref}" firstHeaderRow="{g.first_header_row}" firstDataRow="{g.first_data_row}" '
            f'firstDataCol="{g.first_data_col}"/>'
            f'<pivotFields count="{len(fields)}">' + "".join(fields) + "</pivotFields>"
            + field_list("rowFields", spec.rows_axis)
            + f"<rowItems {items_xml(g.row_items)}</rowItems>"
            + field_list("colFields", spec.cols_axis)
            + f"<colItems {items_xml(g.col_items)}</colItems>"
            + f'<dataFields count="{len(data)}">' + "".join(data) + "</dataFields>"
            + formats
            + f'<pivotTableStyleInfo name="{attr(spec.style_name)}" showRowHeaders="1" showColHeaders="1" '
              'showRowStripes="0" showColStripes="0" showLastColumn="1"/>'
            + f'<extLst><ext uri="{{962EF5D1-5CA2-4c93-8EF4-DBF5C05439D2}}" xmlns:x14="{X14_NS}">{ext}</ext></extLst>'
            + "</pivotTableDefinition>")


def cell_ref(row, col):
    letters, col = "", col + 1
    while col:
        col, rem = divmod(col - 1, 26)
        letters = chr(65 + rem) + letters
    return f"{letters}{row + 1}"


def style_elements(spec):
    """A PivotTable style with Numbers' look, so the table keeps it when Excel lays it out again
    (a refresh clears the cells' own formatting): fills, bold, italic and text colour for the
    header row, row labels, subtotals and grand totals, and the table's lines. (Fonts and sizes
    can't go in a style; add_parts keeps them with a format for the whole table.)"""
    if not hasattr(spec, "lines"):
        raise ValueError("cells() first")
    s, g = spec.styles, spec.grid
    horizontal, vertical = spec.lines
    r0, c0 = min(g.first_data_row, g.rows - 1), min(g.first_data_col, g.cols - 1)

    def at(grid, r, c):
        return grid[r][c] if 0 <= r < len(grid) and 0 <= c < len(grid[r]) else None

    elements = {
        "wholeTable": dxf(s["body"], left=at(vertical, r0, 0), right=at(vertical, r0, g.cols), top=at(horizontal, 0, c0),
                          bottom=at(horizontal, g.rows, c0), vertical=at(vertical, r0, c0 + 1),
                          horizontal=at(horizontal, r0 + 1, c0)),
        "headerRow": dxf(s["header"], bottom=at(horizontal, g.first_data_row, c0)),
        "firstColumn": dxf(s["label"], right=at(vertical, r0, g.first_data_col) if g.first_data_col else None),
    }
    if spec.rows and spec.grand_row:
        elements["totalRow"] = dxf(s["grand"], top=at(horizontal, g.rows - 1, c0))
    if spec.cols and spec.grand_col:  # Excel draws this over the header row too, so no fill: its header keeps its colour
        elements["lastColumn"] = dxf({k: v for k, v in s["subtotal-col"].items() if k != "bg"})
    for which in ("first", "second", "third"):
        elements[f"{which}SubtotalRow"] = dxf(s["subtotal"])
        elements[f"{which}SubtotalColumn"] = dxf(s["subtotal-col"])
    return elements


def dxf(style, **borders):
    """A style element: bold, italic and text colour, fill, and lines ((width, colour, kind))."""
    font = ""
    if style.get("b"):
        font += "<b/>"
    if style.get("i"):
        font += "<i/>"
    if style.get("color"):
        font += f'<color rgb="FF{style["color"].lstrip("#")}"/>'
    xml = f"<font>{font}</font>" if font else ""
    if style.get("bg"):
        rgb = style["bg"].lstrip("#")
        xml += f'<fill><patternFill patternType="solid"><fgColor rgb="FF{rgb}"/><bgColor rgb="FF{rgb}"/></patternFill></fill>'
    edges = ""
    for side in ("left", "right", "top", "bottom", "vertical", "horizontal"):
        line = borders.get(side)
        if line:
            width, colour, kind = line
            edges += f'<{side} style="{border_style(width, kind)}"><color rgb="FF{colour.lstrip("#")}"/></{side}>'
    if edges:
        xml += f"<border>{edges}</border>"
    return f"<dxf>{xml}</dxf>"


def border_style(width, kind):
    """Excel's line for a Numbers line (as numbers_tool.excel_border)."""
    if kind == "dashes":
        return "mediumDashed" if width >= 1.5 else "dashed"
    if kind == "dots":
        return "dotted"
    return "thick" if width >= 2.5 else "medium" if width >= 1.5 else "thin"


class Pending:
    """PivotTables written to a workbook, waiting for add_parts once XlsxWriter has saved it."""

    def __init__(self):
        self.tables = []  # SimpleNamespace(sheet, spec, top_left, source_sheet, source_ref, formats)

    def add(self, sheet, spec, top_left, source_sheet, source_ref, formats):
        self.tables.append(SimpleNamespace(sheet=sheet, spec=spec, top_left=top_left, source_sheet=source_sheet,
                                           source_ref=source_ref, formats=formats))

    def __bool__(self):
        return bool(self.tables)


def add_parts(path, pending):
    """Adds the pivot caches and PivotTables to the workbook XlsxWriter saved at `path`."""
    with zipfile.ZipFile(path) as z:
        parts = {info.filename: z.read(info.filename) for info in z.infolist()}
    text = lambda name: parts[name].decode("utf-8")  # noqa: E731
    workbook = text("xl/workbook.xml")
    workbook_rels = text("xl/_rels/workbook.xml.rels")
    content_types = text("[Content_Types].xml")
    styles = text("xl/styles.xml")
    targets = dict(re.findall(r'<Relationship Id="(rId\d+)" Type="[^"]+" Target="([^"]+)"', workbook_rels))
    sheet_files = {}
    for name, rid in re.findall(r'<sheet name="([^"]*)" sheetId="\d+" r:id="(rId\d+)"', workbook):
        sheet_files[unescape(name)] = "xl/" + targets[rid]
    next_rid = max(int(r[3:]) for r in targets) + 1

    # Styles: for each pivot, a PivotTable style with Numbers' colours and its font for the whole
    # table (both kept when Excel refreshes it), and drop-down arrows on its field buttons.
    dxfs = re.search(r'<dxfs count="\d+"\s*/>|<dxfs count="\d+">(.*?)</dxfs>', styles, re.S)
    new_dxfs = re.findall(r"<dxf>.*?</dxf>|<dxf/>", dxfs.group(1) or "", re.S) if dxfs else []
    style_xml, buttons = [], set()
    for n, written in enumerate(pending.tables, 1):
        spec = written.spec
        spec.style_name = f"Numbers Pivot {n}"
        body = spec.styles["body"]
        spec.font_dxf = len(new_dxfs)
        new_dxfs.append(f'<dxf><font><sz val="{number(float(body.get("size", 10)))}"/>'
                        f'<name val="{attr(body.get("font", "Arial"))}"/></font></dxf>')
        entries = []
        for element, dxf in style_elements(spec).items():
            entries.append(f'<tableStyleElement type="{element}" dxfId="{len(new_dxfs)}"/>')
            new_dxfs.append(dxf)
        style_xml.append(f'<tableStyle name="{attr(spec.style_name)}" table="0" count="{len(entries)}">'
                         + "".join(entries) + "</tableStyle>")
        for d, data in enumerate(spec.data):
            fmt = written.formats.get(d)
            data.num_fmt_id = fmt.num_format_index if fmt is not None and data.num_format else 0
        buttons.update(fmt.xf_index for fmt in written.formats["buttons"])
        for axis in spec.rows + spec.cols:
            fmt = written.formats.get(("axis", axis.field))
            axis.num_fmt_id = fmt.num_format_index if fmt is not None else 0
    dxfs_xml = f'<dxfs count="{len(new_dxfs)}">' + "".join(new_dxfs) + "</dxfs>"
    styles = styles[:dxfs.start()] + dxfs_xml + styles[dxfs.end():] if dxfs else styles.replace("</cellStyles>", "</cellStyles>" + dxfs_xml)
    table_styles = (f'<tableStyles count="{len(style_xml)}" defaultTableStyle="TableStyleMedium9" '
                    f'defaultPivotStyle="PivotStyleLight16">' + "".join(style_xml) + "</tableStyles>")
    styles = re.sub(r"<tableStyles[^>]*/>|<tableStyles[^>]*>.*?</tableStyles>", lambda m: table_styles, styles, count=1, flags=re.S)
    xfs = re.search(r'<cellXfs count="\d+">(.*?)</cellXfs>', styles, re.S)
    xf_list = re.findall(r"<xf [^>]*?/>|<xf [^>]*?>.*?</xf>", xfs.group(1), re.S)
    for index in buttons:
        if index is not None and index < len(xf_list) and "pivotButton" not in xf_list[index]:
            xf_list[index] = re.sub(r"^<xf ", '<xf pivotButton="1" ', xf_list[index])
    styles = styles[:xfs.start(1)] + "".join(xf_list) + styles[xfs.end(1):]
    parts["xl/styles.xml"] = styles.encode("utf-8")

    caches, overrides, used_names = [], [], {}
    for n, written in enumerate(pending.tables, 1):
        spec, sheet = written.spec, written.sheet
        cache_id = n
        definition = f"xl/pivotCache/pivotCacheDefinition{n}.xml"
        records = f"xl/pivotCache/pivotCacheRecords{n}.xml"
        table_part = f"xl/pivotTables/pivotTable{n}.xml"
        name = unique_name(spec.name, used_names.setdefault(sheet, set()))  # unique on its worksheet
        parts[definition] = cache_definition_xml(spec, written.source_sheet, written.source_ref).encode("utf-8")
        parts[records] = cache_records_xml(spec).encode("utf-8")
        parts[f"xl/pivotCache/_rels/pivotCacheDefinition{n}.xml.rels"] = rels_xml(
            [("rId1", "pivotCacheRecords", f"pivotCacheRecords{n}.xml")])
        parts[table_part] = pivot_table_xml(spec, name, cache_id, written.top_left).encode("utf-8")
        parts[f"xl/pivotTables/_rels/pivotTable{n}.xml.rels"] = rels_xml(
            [("rId1", "pivotCacheDefinition", f"../pivotCache/pivotCacheDefinition{n}.xml")])
        rid = f"rId{next_rid}"
        next_rid += 1
        workbook_rels = workbook_rels.replace("</Relationships>", relationship(rid, "pivotCacheDefinition",
                                                                                f"pivotCache/pivotCacheDefinition{n}.xml") + "</Relationships>")
        caches.append(f'<pivotCache cacheId="{cache_id}" r:id="{rid}"/>')
        sheet_file = sheet_files[sheet]
        sheet_rels = sheet_file.replace("worksheets/", "worksheets/_rels/") + ".rels"
        existing_rels = parts[sheet_rels].decode("utf-8") if sheet_rels in parts else rels_xml([]).decode("utf-8")
        ids = [int(i) for i in re.findall(r'Id="rId(\d+)"', existing_rels)]
        new_rel = relationship(f"rId{max(ids, default=0) + 1}", "pivotTable", f"../pivotTables/pivotTable{n}.xml")
        parts[sheet_rels] = existing_rels.replace("</Relationships>", new_rel + "</Relationships>").encode("utf-8")
        for part, kind in ((definition, "pivotCacheDefinition"), (records, "pivotCacheRecords"), (table_part, "pivotTable")):
            overrides.append(f'<Override PartName="/{part}" ContentType="{CONTENT_TYPE}{kind}+xml"/>')
    workbook = workbook.replace("</workbook>", "<pivotCaches>" + "".join(caches) + "</pivotCaches></workbook>")
    parts["xl/workbook.xml"] = workbook.encode("utf-8")
    parts["xl/_rels/workbook.xml.rels"] = workbook_rels.encode("utf-8")
    parts["[Content_Types].xml"] = content_types.replace("</Types>", "".join(overrides) + "</Types>").encode("utf-8")

    temp = path + ".tmp"
    with zipfile.ZipFile(temp, "w", zipfile.ZIP_DEFLATED) as z:
        for name, data in parts.items():
            z.writestr(name, data)
    os.replace(temp, path)


def relationship(rid, kind, target):
    return f'<Relationship Id="{rid}" Type="{REL_TYPE}{kind}" Target="{target}"/>'


def rels_xml(entries):
    body = "".join(relationship(rid, kind, target) for rid, kind, target in entries)
    return ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
            '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
            + body + "</Relationships>").encode("utf-8")


def unescape(text):
    return text.replace("&lt;", "<").replace("&gt;", ">").replace("&quot;", '"').replace("&apos;", "'").replace("&amp;", "&")
