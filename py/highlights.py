"""
Conditional highlighting -> Excel conditional formatting.

How Numbers stores it: a cell with highlighting rules has a key (cell._cond_style_id) into its
table's conditional style table (the base data store's conditionalstyletable), whose entries are
ConditionalStyleSetArchives: the rules in order, each a predicate and the cell and text styles used
when it's true. The predicate is kept as a formula Numbers evaluates for each cell, with a
LINKED_CELL_REF_NODE standing for the cell itself (alongside a structured copy: a type and values).

Each rule becomes an Excel rule on the same cells: "Cell Value < 0" style rules where the formula
compares the cell with a value, and a formula rule (written for the top-left cell of each block
of cells, as Excel expects) for anything else Excel can calculate. Numbers uses the first rule a
cell matches and Excel every one, so where a cell has several rules each stops the ones after it.
"""

import re
from types import SimpleNamespace

from numbers_parser.formula import NODE_FUNCTION_MAP, Formula
from numbers_parser.generated import TSCEArchives_pb2 as TSCE

NODE = TSCE.ASTNodeArrayArchive.ASTNodeType
CONDITIONAL_STYLE = 9  # TST.TableDataList.ListType
DURATION_SECONDS = {1: 604800, 2: 86400, 3: 3600, 4: 60, 5: 1, 6: 0.001}  # weeks ... milliseconds
COMPARISONS = {"<": "<", ">": ">", "<=": "<=", ">=": ">=", "=": "==", "<>": "!="}


class Unsupported(Exception):
    pass


class Rules:
    """A table's conditional highlighting: for each key, the cells it covers and its rules
    ([{"nodes": the predicate's formula nodes, "style": format}], in Numbers' order)."""

    def __init__(self, table):
        self.table = table
        self.cells = {}      # {key: [(row, col)]}
        self.sets = {}       # {key: [rule]}
        self.unreadable = 0  # rules (or keys, when their set can't be found) that couldn't be read


def unreadable_rules(table):
    """For a table whose rules couldn't be read at all: its highlighted cells, with no rules."""
    rules = Rules(table)
    try:
        rules.cells = highlighted_cells(table)
    except Exception:
        return None
    rules.unreadable = len(rules.cells)
    return rules if rules.cells else None


def highlighted_cells(table):
    cells = {}
    for row in table.iter_rows():
        for cell in row:
            key = getattr(cell, "_cond_style_id", None)
            if key is not None:
                cells.setdefault(key, []).append((cell.row, cell.col))
    return cells


def read_rules(table):
    """The table's highlighting rules (a Rules), or None if it has none."""
    model, table_id = table._model, table._table_id
    if model.objects[table_id].HasField("pivot_owner"):
        return None  # a pivot table's cells are rewritten in display order; its rules aren't carried over
    store = model.objects[table_id].base_data_store
    entries = {}
    if store.HasField("conditionalstyletable") and store.conditionalstyletable.identifier in model.objects:
        found = model.objects[store.conditionalstyletable.identifier]
        if found.listType == CONDITIONAL_STYLE:
            entries = {e.key: e.reference.identifier for e in found.entries}
        if not found.entries:
            return None  # no rules anywhere in the table (saves reading every cell)
    cells = highlighted_cells(table)
    if not cells:
        return None
    rules = Rules(table)
    rules.cells = cells
    for key in cells:
        style_set = model.objects[entries[key]] if entries.get(key) in model.objects else None
        if style_set is None:
            rules.unreadable += 1
            continue
        stored = list(style_set.rules.rule) if style_set.HasField("rules") else list(style_set.rules_prepivot)
        rules.sets[key] = []
        for rule in stored:
            if not rule.predicate.HasField("formula"):
                rules.unreadable += 1
                continue
            rules.sets[key].append({"nodes": list(rule.predicate.formula.AST_node_array.AST_node),
                                    "style": rule_format(model, rule)})
    return rules


def rule_format(model, rule):
    """What a rule changes, as numbers_tool cell style keys: {color, b, i, u, st, bg}. Only what the
    rule's own styles set counts (they override the cell's look)."""
    style = {}
    text = model.objects[rule.text_style.identifier] if rule.HasField("text_style") and rule.text_style.identifier in model.objects else None
    if text is not None and "char_properties" in text.DESCRIPTOR.fields_by_name:
        props = text.char_properties
        if props.HasField("font_color"):
            c = props.font_color
            style["color"] = "#{:02X}{:02X}{:02X}".format(round(c.r * 255), round(c.g * 255), round(c.b * 255))
        for name, key in (("bold", "b"), ("italic", "i")):
            if props.HasField(name):
                style[key] = 1 if getattr(props, name) else 0
        if props.HasField("underline"):
            style["u"] = 1 if props.underline else 0
        if props.HasField("strikethru"):
            style["st"] = 1 if props.strikethru else 0
    cell = model.objects[rule.cell_style.identifier] if rule.HasField("cell_style") and rule.cell_style.identifier in model.objects else None
    if cell is not None and cell.HasField("cell_properties") and cell.cell_properties.HasField("cell_fill"):
        fill = cell.cell_properties.cell_fill
        colour = fill.color if fill.HasField("color") else (fill.gradient.stops[0].color if fill.HasField("gradient") and fill.gradient.stops else None)
        if colour is not None:
            style["bg"] = "#{:02X}{:02X}{:02X}".format(round(colour.r * 255), round(colour.g * 255), round(colour.b * 255))
    return style


def blocks(cells):
    """Cells as few rectangles as possible, (first row, first column, last row, last column), by
    joining runs of cells across each row, then runs the same width down."""
    rows = {}
    for r, c in sorted(cells):
        runs = rows.setdefault(r, [])
        if runs and runs[-1][1] == c - 1:
            runs[-1][1] = c
        else:
            runs.append([c, c])
    found, open_ = [], {}
    for r in sorted(rows):
        current = {}
        for c1, c2 in rows[r]:
            block = open_.get((c1, c2))
            if block is not None and block[2] == r - 1:
                block[2] = r
            else:
                block = [r, c1, r, c2]
                found.append(block)
            current[(c1, c2)] = block
        open_ = current
    return [tuple(b) for b in found]


def predicate_text(model, table_id, row, col, nodes):
    """A predicate as Numbers formula text for the cell at (row, col): the cell itself as an exact
    reference code (see numbers_tool), other references relative to it, as in cell formulas."""
    formula = Formula(model, table_id, row, col)
    for node in nodes:
        kind = NODE.Name(node.AST_node_type)
        if kind == "LINKED_CELL_REF_NODE":
            formula.push(f"⟦{table_id}|{row}|{row}|{col}|{col}|0⟧")
        elif kind == "DURATION_NODE":
            seconds = DURATION_SECONDS.get(node.AST_duration_node_unit)
            if seconds is None:
                raise Unsupported(kind)
            formula.push(repr(node.AST_duration_node_unitNum * seconds / 86400))
        elif kind == "REFERENCE_ERROR_WITH_UIDS":
            raise Unsupported(kind)
        elif kind in NODE_FUNCTION_MAP:
            if NODE_FUNCTION_MAP[kind] is not None:  # else whitespace and thunk markers
                getattr(formula, NODE_FUNCTION_MAP[kind])(row, col, node)
        else:
            raise Unsupported(kind)
    if len(formula._stack) != 1:
        raise Unsupported("unexpected predicate")
    return str(formula)


def excel_rules(rules, translate, place):
    """The Excel rules for a table's highlighting, [{"key", "areas" (on the worksheet), "options"
    (conditional_format's, but the format), "style"}] in the order they apply, and how many of
    Numbers' rules Excel can't express. translate(text, row, col) gives the Excel formula for
    Numbers formula text in the cell at (row, col) of the table, or None; place(row, col) a table
    cell's (row, col) on its worksheet."""
    model, table_id = rules.table._model, rules.table._table_id
    out, missed = [], rules.unreadable
    for key, rule_list in rules.sets.items():
        found = blocks(rules.cells[key])
        stop = len(rule_list) > 1
        for rule in rule_list:
            formulas = []
            for r1, c1, r2, c2 in found:
                try:
                    excel = translate(predicate_text(model, table_id, r1, c1, rule["nodes"]), r1, c1)
                except Exception:
                    excel = None
                if not excel:
                    break
                formulas.append(excel)
            if len(formulas) != len(found):
                missed += 1
                continue
            areas = [(*place(r1, c1), *place(r2, c2)) for r1, c1, r2, c2 in found]
            compared = [cell_comparison(excel, area) for excel, area in zip(formulas, areas)]
            if all(compared) and len({c[1:] for c in compared}) == 1 and not has_relative_reference(compared[0][2]):
                # One "Cell Value" rule covers every block when what it compares with is fixed.
                _, criteria, value = compared[0]
                out.append({"key": key, "areas": areas, "style": rule["style"],
                            "options": {"type": "cell", "criteria": criteria, "value": value, "stop_if_true": stop}})
                continue
            for excel, area, comparison in zip(formulas, areas, compared):
                options = ({"type": "cell", "criteria": comparison[1], "value": comparison[2]} if comparison
                           else {"type": "formula", "criteria": excel})
                out.append({"key": key, "areas": [area], "style": rule["style"], "options": {**options, "stop_if_true": stop}})
    return out, missed


def cell_comparison(excel, area):
    """(cell, criteria, value) when an Excel formula only compares the block's top-left cell with
    something not involving it ("=B3<0", '=B3="Done"'), else None."""
    top_left = cell_name(area[0], area[1])
    match = re.fullmatch(r"=(\$?[A-Z]{1,3}\$?\d+)(<=|>=|<>|<|>|=)(.+)", excel)
    if not match or match.group(1) != top_left:
        return None
    value = match.group(3)
    bare = re.sub(r'"(?:[^"]|"")*"', '""', value)  # outside text in quotes
    if re.search(r"[<>=]", bare) or re.search(rf"(?<![A-Z$]){re.escape(top_left)}(?!\d)", bare):
        return None
    return top_left, COMPARISONS[match.group(2)], value


def has_relative_reference(value):
    bare = re.sub(r'"(?:[^"]|"")*"', '""', value)
    return any(not (a and b) for a, b in re.findall(r"(?<![A-Za-z_])(\$?)[A-Z]{1,3}(\$?)\d+(?![\d(])", bare))


def cell_name(row, col):
    letters, col = "", col + 1
    while col:
        col, rem = divmod(col - 1, 26)
        letters = chr(65 + rem) + letters
    return f"{letters}{row + 1}"


def write(ws, rules, formats):
    """Adds the rules from excel_rules to a worksheet. formats(style) gives the Excel format for a
    rule's style dict (or None for a rule that changes nothing, which still stops later rules)."""
    for rule in rules:
        r1, c1, r2, c2 = rule["areas"][0]
        options = dict(rule["options"], format=formats(rule["style"]))
        if len(rule["areas"]) > 1:
            options["multi_range"] = " ".join(area_name(a) for a in rule["areas"])
        ws.conditional_format(r1, c1, r2, c2, options)


def area_name(area):
    r1, c1, r2, c2 = area
    return cell_name(r1, c1) if (r1, c1) == (r2, c2) else f"{cell_name(r1, c1)}:{cell_name(r2, c2)}"


def host(table_id, row, col):
    """Stands in for the table cell a formula belongs to (numbers_tool.translate_formula's `cell`)."""
    return SimpleNamespace(_table_id=table_id, row=row, col=col)


# ---------------------------------------------------------------- preview


def preview_style(rules, value):
    """The style a cell gets from its rules (excel_rules' for its key, in order) in the preview: the
    first matching "Cell Value" rule's, or None. Formula rules aren't worked out here, so the
    preview misses what they'd highlight (Excel calculates them when the file opens)."""
    for rule in rules:
        options = rule["options"]
        if options["type"] == "cell" and matches(options["criteria"], value, options["value"]):
            return rule["style"]
    return None


def matches(criteria, value, wanted):
    """Whether a cell's value passes an Excel "Cell Value" comparison, as Excel works it out: a
    blank cell counts as 0 (or "" against text), numbers sort before text and text before TRUE and
    FALSE, and text compares ignoring case."""
    wanted = wanted.strip()
    if re.fullmatch(r"-?\d+(\.\d+)?([eE][-+]?\d+)?", wanted):
        target = float(wanted)
    elif re.fullmatch(r'"(?:[^"]|"")*"', wanted):
        target = wanted[1:-1].replace('""', '"').casefold()
    elif wanted in ("TRUE", "FALSE"):
        target = wanted == "TRUE"
    else:
        return False  # a reference or calculation: not worked out for the preview
    if value is None or value == "":
        value = 0.0 if isinstance(target, float) else ("" if isinstance(target, str) else False)
    if isinstance(value, str):
        value = value.casefold()
    rank = lambda v: 2 if isinstance(v, bool) else 1 if isinstance(v, str) else 0  # noqa: E731
    a, b = (rank(value), value), (rank(target), target)
    if a[0] != b[0]:
        a, b = a[0], b[0]
    return {"<": a < b, ">": a > b, "<=": a <= b, ">=": a >= b, "==": a == b, "!=": a != b}[criteria]
