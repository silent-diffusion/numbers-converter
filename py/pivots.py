"""
Recomputes Numbers pivot table values.

Recent versions of Numbers don't save a pivot table's summary values; it recalculates them from
the pivot's rules every time the file is opened. numbers-parser therefore reads those cells as
empty. This module does the same recalculation:

  * the pivot table's model points to its rules (PivotOwnerArchive): which source columns group
    the rows and columns, and which columns are aggregated how;
  * its TableInfoArchive points to a snapshot of the source table taken at the last refresh
    (pivot_data_model), whose column UID map turns the rules' column UIDs into column indexes;
  * each grouping rule carries a small formula (the "functor") that turns a source value into its
    group, e.g. DATE(2020, MONTH(x), 1) for "by month". Evaluating it covers every grouping type
    without needing to know what Numbers' grouping type codes mean;
  * the pivot's own label cells hold the group values, so each empty summary cell is matched to
    its row and column groups exactly.

Anything not understood (an unsupported function in a grouping formula, an unknown aggregate, a
"show as" percentage) leaves the affected cells empty rather than guessing.
"""

import datetime as dt
import math
import re
import statistics

from numbers_parser.document import Table
from numbers_parser.generated import TSCEArchives_pb2 as TSCE
from numbers_parser.generated.functionmap import FUNCTION_MAP
from numbers_parser.model import _NumbersModel

NODE = TSCE.ASTNodeArrayArchive.ASTNodeType
EXCEL_EPOCH = dt.datetime(1899, 12, 30)

# A pivot's source snapshot has no TableInfoArchive, which numbers-parser needs for grouped
# (categorised) tables. Treat such tables as ungrouped instead of failing.
_calculate_table_categories = _NumbersModel.calculate_table_categories


def _safe_categories(self, table_id):
    try:
        return _calculate_table_categories(self, table_id)
    except IndexError:
        self._table_categories_data[table_id] = None
        self._table_categories_row_mapper[table_id] = None
        return None


_NumbersModel.calculate_table_categories = _safe_categories


class Unsupported(Exception):
    """The pivot uses something we can't reproduce exactly."""


class GroupError(Exception):
    """A source value doesn't belong to any group (Numbers' grouping formula returns an error)."""


# ---------------------------------------------------------------- values


def norm(value):
    """Comparable form of a cell or group value."""
    if value is None or value == "":
        return None
    if isinstance(value, bool):
        return ("b", value)
    if isinstance(value, (int, float)):
        return ("n", round(float(value), 9))
    if isinstance(value, dt.datetime):
        return ("d", value.replace(tzinfo=None))
    if isinstance(value, dt.timedelta):
        return ("t", value.total_seconds())
    return ("s", str(value).strip().casefold())


def serial(value):
    if isinstance(value, dt.datetime):
        delta = value.replace(tzinfo=None) - EXCEL_EPOCH
        return delta.days + delta.seconds / 86400
    if isinstance(value, dt.timedelta):
        return value.total_seconds() / 86400
    if isinstance(value, bool):
        return float(value)
    if isinstance(value, (int, float)):
        return float(value)
    raise GroupError


def as_date(value):
    if isinstance(value, dt.datetime):
        return value.replace(tzinfo=None)
    if isinstance(value, (int, float)) and not isinstance(value, bool):
        return EXCEL_EPOCH + dt.timedelta(days=value)
    raise GroupError


def truthy(value):
    if isinstance(value, str):
        raise GroupError
    return bool(value)


# ---------------------------------------------------------------- grouping formulas


def _date(y, m, d):
    y, m, d = int(y), int(m), int(d)
    y += (m - 1) // 12
    m = (m - 1) % 12 + 1
    return dt.datetime(y, m, 1) + dt.timedelta(days=d - 1)


def _weekday(d, kind=1):
    wd = as_date(d).isoweekday()  # Mon=1..Sun=7
    kind = int(kind)
    if kind == 1:
        return wd % 7 + 1  # Sun=1
    if kind == 2:
        return wd  # Mon=1
    if kind == 3:
        return wd - 1  # Mon=0
    raise Unsupported(f"WEEKDAY type {kind}")


def _weeknum(d, kind=1):
    """Week of the year; week 1 contains 1 January and weeks start on Sunday (1) or Monday (2)."""
    d = as_date(d)
    starts = {1: 6, 2: 0}  # Python weekday() of the first day of the week
    if int(kind) not in starts:
        raise Unsupported(f"WEEKNUM type {kind}")
    jan1 = dt.datetime(d.year, 1, 1)
    offset = (jan1.weekday() - starts[int(kind)]) % 7
    return ((d - jan1).days + offset) // 7 + 1


def _sum(*args):
    if len(args) == 1 and isinstance(args[0], dt.datetime):
        return args[0]
    total = 0.0
    for a in args:
        if a is None:
            continue
        if isinstance(a, str):
            raise GroupError
        total += serial(a)
    return total


FUNCTIONS = {
    "SUM": _sum,
    "AND": lambda *a: all(truthy(x) for x in a),
    "OR": lambda *a: any(truthy(x) for x in a),
    "NOT": lambda a: not truthy(a),
    "ISBLANK": lambda a: a is None or a == "",
    "ISNUMBER": lambda a: isinstance(a, (int, float)) and not isinstance(a, bool),
    "ISTEXT": lambda a: isinstance(a, str),
    "DATE": _date,
    "TIME": lambda h, m, s: (int(h) * 3600 + int(m) * 60 + int(s)) / 86400,
    "YEAR": lambda d: as_date(d).year,
    "MONTH": lambda d: as_date(d).month,
    "DAY": lambda d: as_date(d).day,
    "HOUR": lambda d: as_date(d).hour,
    "MINUTE": lambda d: as_date(d).minute,
    "SECOND": lambda d: as_date(d).second,
    "WEEKDAY": _weekday,
    "WEEKNUM": _weeknum,
    "ISOWEEKNUM": lambda d: as_date(d).isocalendar()[1],
    "INT": lambda x: math.floor(serial(x)),
    "TRUNC": lambda x, n=0: math.trunc(serial(x) * 10 ** int(n)) / 10 ** int(n),
    "FLOOR": lambda x, f=1: math.floor(serial(x) / serial(f)) * serial(f),
    "CEILING": lambda x, f=1: math.ceil(serial(x) / serial(f)) * serial(f),
    "ROUND": lambda x, n=0: round(serial(x), int(n)),
    "ROUNDDOWN": lambda x, n=0: math.trunc(serial(x) * 10 ** int(n)) / 10 ** int(n),
    "MOD": lambda a, b: serial(a) % serial(b),
    "QUOTIENT": lambda a, b: math.trunc(serial(a) / serial(b)),
    "ABS": lambda x: abs(serial(x)),
    "CONCATENATE": lambda *a: "".join(text(x) for x in a),
    "LEFT": lambda s, n=1: text(s)[: int(n)],
    "RIGHT": lambda s, n=1: text(s)[-int(n):] if int(n) else "",
    "UPPER": lambda s: text(s).upper(),
    "LOWER": lambda s: text(s).lower(),
    "LEN": lambda s: len(text(s)),
}


def text(value):
    if value is None:
        return ""
    if isinstance(value, float) and value.is_integer():
        return str(int(value))
    return str(value)


class Thunk:
    def __init__(self, nodes):
        self.nodes = nodes


def force(value, x):
    return evaluate(value.nodes, x) if isinstance(value, Thunk) else value


def evaluate(nodes, x):
    """Evaluate an RPN formula from a grouping functor with VAR_NODE bound to x."""
    stack = []
    i = 0
    while i < len(nodes):
        node = nodes[i]
        kind = NODE.Name(node.AST_node_type)
        if kind == "BEGIN_THUNK_NODE":
            depth, j = 1, i + 1
            while depth:
                k = NODE.Name(nodes[j].AST_node_type)
                depth += k == "BEGIN_THUNK_NODE"
                depth -= k == "END_THUNK_NODE"
                j += 1
            stack.append(Thunk(nodes[i + 1:j - 1]))
            i = j
            continue
        if kind == "VAR_NODE":
            stack.append(x)
        elif kind == "NUMBER_NODE":
            stack.append(node.AST_number_node_number)
        elif kind == "STRING_NODE":
            stack.append(node.AST_string_node_string)
        elif kind == "BOOLEAN_NODE":
            stack.append(node.AST_boolean_node_boolean)
        elif kind == "FUNCTION_NODE":
            name = FUNCTION_MAP.get(node.AST_function_node_index)
            argc = node.AST_function_node_numArgs
            args = stack[len(stack) - argc:] if argc else []
            del stack[len(stack) - argc:]
            if name == "IF":
                cond = truthy(force(args[0], x))
                branch = args[1] if cond else (args[2] if len(args) > 2 else False)
                stack.append(force(branch, x))
            elif name == "IFERROR":
                try:
                    stack.append(force(args[0], x))
                except GroupError:
                    stack.append(force(args[1], x))
            elif name in FUNCTIONS:
                try:
                    stack.append(FUNCTIONS[name](*[force(a, x) for a in args]))
                except (TypeError, ValueError, OverflowError, ZeroDivisionError) as err:
                    raise GroupError from err
            else:
                raise Unsupported(f"function {name}")
        elif kind in BINARY:
            b, a = force(stack.pop(), x), force(stack.pop(), x)
            stack.append(BINARY[kind](a, b))
        elif kind == "NEGATION_NODE":
            stack.append(-serial(force(stack.pop(), x)))
        elif kind == "PLUS_SIGN_NODE":
            pass
        elif kind == "PERCENT_NODE":
            stack.append(serial(force(stack.pop(), x)) / 100)
        else:
            raise Unsupported(f"formula node {kind}")
        i += 1
    if len(stack) != 1:
        raise Unsupported("malformed grouping formula")
    return force(stack[0], x)


def _divide(a, b):
    if serial(b) == 0:
        raise GroupError
    return serial(a) / serial(b)


BINARY = {
    "ADDITION_NODE": lambda a, b: serial(a) + serial(b),
    "SUBTRACTION_NODE": lambda a, b: serial(a) - serial(b),
    "MULTIPLICATION_NODE": lambda a, b: serial(a) * serial(b),
    "DIVISION_NODE": _divide,
    "POWER_NODE": lambda a, b: serial(a) ** serial(b),
    "CONCATENATION_NODE": lambda a, b: text(a) + text(b),
    "EQUAL_TO_NODE": lambda a, b: norm(a) == norm(b),
    "NOT_EQUAL_TO_NODE": lambda a, b: norm(a) != norm(b),
    "GREATER_THAN_NODE": lambda a, b: serial(a) > serial(b),
    "GREATER_THAN_OR_EQUAL_TO_NODE": lambda a, b: serial(a) >= serial(b),
    "LESS_THAN_NODE": lambda a, b: serial(a) < serial(b),
    "LESS_THAN_OR_EQUAL_TO_NODE": lambda a, b: serial(a) <= serial(b),
}


def group_key(rule, value):
    if not rule.HasField("grouping_functor"):
        return norm(value)
    result = evaluate(list(rule.grouping_functor.formula.AST_node_array.AST_node), value)
    if isinstance(result, float) and result.is_integer():
        result = int(result)
    return norm(result)


# ---------------------------------------------------------------- aggregates

AGG_BY_CODE = {2: "sum"}  # confirmed against Numbers files; other types are recognised by label
AGG_LABELS = {
    "sum": "sum", "total": "sum", "average": "average", "mean": "average", "count": "count",
    "minimum": "min", "min": "min", "maximum": "max", "max": "max",
    "standard deviation": "stdev", "std dev": "stdev", "variance": "var",
}


def aggregate(kind, values):
    if kind == "count":
        return len([v for v in values if v is not None and v != ""])
    nums = [serial(v) for v in values if v is not None and v != "" and not isinstance(v, str)]
    if not nums:
        return None
    if kind == "sum":
        return sum(nums)
    if kind == "average":
        return sum(nums) / len(nums)
    if kind == "min":
        return min(nums)
    if kind == "max":
        return max(nums)
    if kind == "stdev":
        return statistics.stdev(nums) if len(nums) > 1 else None
    if kind == "var":
        return statistics.variance(nums) if len(nums) > 1 else None
    raise Unsupported(kind)


# ---------------------------------------------------------------- pivot tables


def find_table_info(model, table_id):
    for ref in model.objects.find_refs("TableInfoArchive"):
        info = model.objects[ref]
        if info.tableModel.identifier == table_id:
            return info
    return None


def uid_key(uid):
    return (uid.lower, uid.upper)


class Pivot:
    """Computed values for one pivot table: {(row, col): (value, source cell for formatting)}."""

    def __init__(self, table, recompute_stored=False):
        self.table = table
        self.recompute_stored = recompute_stored  # also compute cells whose value Numbers saved (for testing)
        self.values = {}
        self.missing = 0  # summary cells we couldn't compute
        self._format_cells = {}
        model = table._model
        tm = model.objects[table._table_id]
        self.owner = model.objects[tm.pivot_owner.identifier]
        info = find_table_info(model, table._table_id)
        if self.owner.is_empty_pivot or info is None or not info.HasField("pivot_data_model"):
            return
        self.source_id = info.pivot_data_model.identifier
        self.source = Table(model, self.source_id)
        source_tm = model.objects[self.source_id]
        uid_map = model.objects[source_tm.base_column_row_uids.identifier]
        self.column_for_uid = {
            uid_key(uid): uid_map.column_index_for_uid[i] for i, uid in enumerate(uid_map.sorted_column_uids)
        }
        first = source_tm.number_of_header_rows
        last = source_tm.number_of_rows - source_tm.number_of_footer_rows
        self.rows = range(first, last)
        self.compute(tm)

    def column(self, uid):
        try:
            return self.column_for_uid[uid_key(uid)]
        except KeyError:
            raise Unsupported("unknown source column") from None

    def source_header(self, col):
        names = [self.source.cell(r, col).value for r in range(self.rows.start)]
        return next((str(n) for n in reversed(names) if n not in (None, "")), "")

    def aggregates(self):
        """[(kind, source column, label text)] for each aggregate rule."""
        result = []
        for rule in self.owner.aggregate_columns.aggregates:
            if rule.show_as_type != 0:
                raise Unsupported("show-as values (percentages, running totals)")
            col = self.column(rule.column_uid)
            result.append([AGG_BY_CODE.get(rule.agg_type), col, self.source_header(col)])
        return result

    def compute(self, tm):
        row_rules = list(self.owner.grouping_columns_for_rows.group_column)
        col_rules = list(self.owner.grouping_columns_for_columns.group_column)
        aggs = self.aggregates()
        row_cols = [self.column(g.column_uid) for g in row_rules]
        col_cols = [self.column(g.column_uid) for g in col_rules]

        # Group every source row.
        records = []
        for r in self.rows:
            try:
                rkey = tuple(group_key(g, self.source.cell(r, c).value) for g, c in zip(row_rules, row_cols))
                ckey = tuple(group_key(g, self.source.cell(r, c).value) for g, c in zip(col_rules, col_cols))
            except GroupError:
                continue  # Numbers leaves such rows out of the pivot too
            records.append((rkey, ckey, [self.source.cell(r, a[1]).value for a in aggs]))
        row_groups = [{k[i] for k, _, _ in records} for i in range(len(row_rules))]
        col_groups = [{k[i] for _, k, _ in records} for i in range(len(col_rules))]

        t = self.table
        header_rows, header_cols = tm.number_of_header_rows, tm.number_of_header_columns
        agg_rows = range(len(col_rules), header_rows)  # rows naming the aggregate (columns layout)
        agg_cols = range(len(row_rules), header_cols)

        def label(r, c):
            return t.cell(r, c).value

        def path(levels, groups, get):
            """Group keys for each level; None = all (a Total row/column)."""
            keys = []
            for level in range(levels):
                value = get(level)
                key = norm(value)
                if key in groups[level]:
                    keys.append(key)
                elif isinstance(value, str) and "total" in value.casefold():
                    keys.extend([None] * (levels - level))
                    break
                else:
                    raise LookupError
            return keys

        def carried(get_cell, positions):
            """Last non-empty label at or before the position (labels aren't repeated)."""
            for p in reversed(positions):
                value = get_cell(p)
                if value not in (None, ""):
                    return value
            return None

        for r in range(header_rows, t.num_rows):
            for c in range(header_cols, t.num_cols):
                cell = t.cell(r, c)
                if not cell.is_formula or (cell.value is not None and not self.recompute_stored):
                    continue
                try:
                    rpath = path(len(row_rules), row_groups,
                                 lambda lv: carried(lambda rr: label(rr, lv), range(header_rows, r + 1)))
                    cpath = path(len(col_rules), col_groups,
                                 lambda lv: carried(lambda cc: label(lv, cc), range(header_cols, c + 1)))
                    agg = self.pick_aggregate(aggs, r, c, agg_rows, agg_cols, label, carried, header_rows, header_cols)
                except LookupError:
                    self.missing += 1
                    continue
                index, (kind, col, _) = agg
                picked = [
                    vals[index] for rk, ck, vals in records
                    if all(k is None or k == g for k, g in zip(rpath, rk)) and all(k is None or k == g for k, g in zip(cpath, ck))
                ]
                value = aggregate(kind, picked)
                self.values[(r, c)] = (value, None if kind == "count" else self.format_cell(col))

    def format_cell(self, col):
        """A source cell whose number format the column's totals should use."""
        if col not in self._format_cells:
            self._format_cells[col] = next(
                (self.source.cell(r, col) for r in self.rows if self.source.cell(r, col).value is not None), None)
        return self._format_cells[col]

    def pick_aggregate(self, aggs, r, c, agg_rows, agg_cols, label, carried, header_rows, header_cols):
        names = []
        for row in agg_rows:
            names.append(carried(lambda cc: label(row, cc), range(header_cols, c + 1)))
        for col in agg_cols:
            names.append(carried(lambda rr: label(rr, col), range(header_rows, r + 1)))
        names = [n for n in names if isinstance(n, str)]
        for index, agg in enumerate(aggs):
            for name in names:
                m = re.fullmatch(r"(.*)\s+\(([^()]+)\)", name.strip())
                if not m or m.group(1).strip().casefold() != agg[2].strip().casefold():
                    continue
                kind = AGG_LABELS.get(m.group(2).strip().casefold())
                if agg[0] is None:
                    agg[0] = kind
                if agg[0] is not None and (kind is None or kind == agg[0]):
                    return index, agg
        if len(aggs) == 1 and aggs[0][0] is not None:
            return 0, aggs[0]
        raise LookupError


def pivot_for(table, recompute_stored=False):
    """A Pivot with computed values, or None if the table isn't a pivot. Raises Unsupported."""
    tm = table._model.objects[table._table_id]
    if not tm.HasField("pivot_owner"):
        return None
    return Pivot(table, recompute_stored)
