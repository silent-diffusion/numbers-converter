"""
Cell and table features beyond values and styles: form controls (checkboxes, ratings, sliders,
steppers, pop-up menus), comments, links, mixed formatting within a cell, hidden rows/columns and
filters. Everything here is best-effort: if a structure isn't what we expect, the feature is
skipped and the cell is converted as plain content.
"""

import re

# Numbers' cell interaction types (TST.CellSpecArchive.interaction_type), from sample files.
CONTROL_KINDS = {4: "stepper", 5: "slider", 6: "rating", 7: "popup", 8: "checkbox"}


def control(cell):
    """{'kind': ..., 'min'/'max'/'step' or 'items'} for a form-control cell, else None."""
    if cell._control_id is None:
        return None
    try:
        model = cell._model
        spec = model._control_specs.lookup_value(cell._table_id, cell._control_id).cell_spec
        kind = CONTROL_KINDS.get(spec.interaction_type)
        if kind in ("stepper", "slider", "rating"):
            return {"kind": kind, "min": spec.range_control_min, "max": spec.range_control_max,
                    "step": spec.range_control_inc or 1}
        if kind == "popup":
            menu = model.objects[spec.chooser_control_popup_model.identifier]
            return {"kind": kind, "items": [v for v in (popup_item(i) for i in menu.tsce_item) if v is not None]}
        return {"kind": kind} if kind else None
    except Exception:
        return None


def popup_item(item):
    if item.HasField("string_value"):
        return item.string_value.value
    if item.HasField("number_value"):
        value = item.number_value.value
        return int(value) if float(value).is_integer() else value
    return None  # the blank "None" entry Numbers puts first


def comment(cell):
    """(text, author) of the cell's comment, or None."""
    try:
        comments = cell.comments
    except Exception:
        return None
    if not comments:
        return None
    parts = [c.text.strip() for c in comments if c.text and c.text.strip()]
    if not parts:
        return None
    author = next((c.author for c in comments if c.author), None)
    return "\n\n".join(parts), author


def links(cell):
    """[(text, url)] for links inside a text cell."""
    try:
        return [(text, url) for text, url in (getattr(cell, "hyperlinks", None) or []) if url]
    except Exception:
        return []


def bullet_text(cell):
    """The cell's text with its bullet or numbering characters, or None if it has none.
    numbers-parser miscounts numbered lists that follow an unbulleted first line, so numbered
    lists are only used when every line is part of the list."""
    try:
        if not cell.is_bulleted:
            return None
        lines = cell.formatted_bullets
        plain = str(cell.value).split("\n")
    except Exception:
        return None
    if not lines:
        return None
    every_line = len(lines) == len(plain) and all(line != p for line, p in zip(lines, plain))
    numbered = any(re.match(r"\(?[0-9A-Za-z]+[.)] ", line) for line in lines)
    return "\n".join(lines) if every_line or not numbered else None


def rich_runs(cell):
    """Runs of differently formatted text within a cell: [(text, {style keys})], or None when the
    whole cell has one format. Style keys match numbers_tool.cell_style (b, i, u, st, color, size, font)."""
    if getattr(cell, "_rich_id", None) is None or bullet_text(cell):
        return None  # bulleted text is written with its bullets instead
    try:
        model = cell._model
        bds = model.objects[cell._table_id].base_data_store
        table = model.objects[bds.rich_text_table.identifier]
        entry = next(e for e in table.entries if e.key == cell._rich_id)
        storage = model.objects[model.objects[entry.rich_text_payload.identifier].storage.identifier]
        text = storage.text[0] if storage.text else ""
        runs = list(storage.table_char_style.entries)
    except Exception:
        return None
    if text.endswith("\n"):
        text = text[:-1]
    if len(runs) < 2 and not (runs and runs[0].object.identifier):
        return None

    pieces = []
    for i, run in enumerate(runs):
        start = run.character_index
        end = runs[i + 1].character_index if i + 1 < len(runs) else len(text)
        piece = text[start:end]
        if not piece:
            continue
        style = {}
        if run.object.identifier and run.object.identifier in model.objects:
            props = model.objects[run.object.identifier].char_properties
            style = char_style(props)
        pieces.append((piece, style))
    if len(pieces) < 2 and not any(s for _, s in pieces):
        return None
    return pieces


def char_style(props):
    style = {}
    if props.HasField("bold") and props.bold:
        style["b"] = 1
    if props.HasField("italic") and props.italic:
        style["i"] = 1
    if props.HasField("underline") and props.underline:
        style["u"] = 1
    if props.HasField("strikethru") and props.strikethru:
        style["st"] = 1
    if props.HasField("font_color"):
        c = props.font_color
        style["color"] = "#{:02X}{:02X}{:02X}".format(round(c.r * 255), round(c.g * 255), round(c.b * 255))
    if props.HasField("font_size"):
        style["size"] = round(props.font_size, 1)
    return style


def hidden(table, columns):
    """Indexes of rows (or columns) hidden in Numbers, by the user or by a filter."""
    try:
        model = table._model
        bds = model.objects[table._table_id].base_data_store
        bucket = bds.columnHeaders.identifier if columns else bds.rowHeaders.buckets[0].identifier
        return {h.index for h in model.objects[bucket].headers if h.hidingState}
    except Exception:
        return set()


def filter_enabled(table):
    """True if the table has filters switched on in Numbers."""
    try:
        model = table._model
        tm = model.objects[table._table_id]
        for state in tm.hidden_states_owner.hidden_states:
            ref = state.row_hidden_state_extent.filter_set
            if ref.identifier and ref.identifier in model.objects:
                fs = model.objects[ref.identifier]
                if fs.filter_rules and (not fs.HasField("is_enabled") or fs.is_enabled):
                    return True
    except Exception:
        pass
    return False
