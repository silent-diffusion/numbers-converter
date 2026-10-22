"""
Text boxes, shapes and images -> Excel drawings.

How Numbers stores what's placed on a sheet besides its tables and charts (listed, in stacking
order, in the sheet's drawable_infos):
  * a text box or shape is a ShapeInfoArchive: a geometry (position and size on the sheet canvas in
    points, rotation, flips, and whether its width and height are fixed or follow its text), a path
    giving its outline (a rectangle for text boxes; a line; a rounded rectangle, polygon, star,
    arrow or speech bubble described by a few numbers; any bezier outline), a style (fill, stroke,
    opacity, and for text its padding and vertical alignment) and its text: paragraphs, with
    paragraph, list and character styles by position;
  * an image is an ImageArchive: a geometry, a reference to the picture's file in the package's Data
    folder (listed in the package metadata), an optional mask cropping it, and a style (border,
    opacity); a movie's poster frame is shown the same way;
  * objects grouped together sit in a GroupArchive, their geometry measured within the group.

read_drawings turns each into a Drawing, layout.py decides which worksheet each goes on and where
(as for charts), and write_drawings adds them to the worksheets. XlsxWriter places them and adds the
pictures to the file, but its text boxes have one font and a rectangular outline and its pictures
can't be cropped, so the XML for each is written here instead (see _write_shape, _write_picture).
"""

import base64
import io
import math
import re

from xlsxwriter.drawing import Drawing as SheetDrawing
from xlsxwriter.image import Image
from xlsxwriter.shape import Shape
from xlsxwriter.worksheet import Worksheet

PT_TO_PX = 4 / 3
EMU = 12700        # per point
PACKAGE_ID = 2     # the package metadata, which lists the files in the Data folder
KAPPA = 0.5523     # how far a cubic curve's control points sit to draw a quarter circle
KINDS = {"ShapeInfoArchive": "shapes", "ShapeArchive": "shapes", "ImageArchive": "images", "MovieArchive": "images"}
PATH_COMMANDS = {1: "M", 2: "L", 3: "Q", 4: "C", 5: "Z"}  # TSP.Path element types
ALIGNMENTS = {0: "l", 1: "r", 2: "ctr", 3: "just"}         # 4 is "natural": left, for left-to-right text
ANCHORS = {0: "t", 1: "ctr", 2: "b"}                       # vertical alignment of text in its box
# Numbers' numbered list styles (TSWP.ListStyleArchive.NumberType) as Excel's: 1. (1) 1) for each kind.
NUMBERING = {0: "arabicPeriod", 1: "arabicParenBoth", 2: "arabicParenR", 3: "romanUcPeriod", 4: "romanUcParenBoth",
             5: "romanUcParenR", 6: "romanLcPeriod", 7: "romanLcParenBoth", 8: "romanLcParenR", 9: "alphaUcPeriod",
             10: "alphaUcParenBoth", 11: "alphaUcParenR", 12: "alphaLcPeriod", 13: "alphaLcParenBoth",
             14: "alphaLcParenR"}
# Line heights of the Windows fonts text ends up in, as a multiple of the font size.
LINE_HEIGHTS = {"Arial": 1.15, "Segoe UI": 1.33, "Calibri": 1.22, "Consolas": 1.17, "Times New Roman": 1.15,
                "Georgia": 1.14, "Verdana": 1.22, "Tahoma": 1.21}
NARROW, WIDE = set("fijlrtI!|.,:;'`()[]{}\"/"), set("mwMW@%&")


class Drawing:
    def __init__(self, oid, kind):
        self.id = oid
        self.kind = kind              # "text" (a text box), "shape", "line" or "image"
        self.sheet_id = None
        self.frame = None             # (x, y, w, h) on the canvas, in points, before rotation
        self.rotation = 0.0           # degrees clockwise
        self.flip = (False, False)    # horizontal, vertical
        self.outline = "rect"         # "rect", "line" or "path"
        self.path = None              # the outline when "path": [(command, [(x, y)])], M L Q C Z ...
        self.path_size = None         # ... drawn in a box this size, stretched to the frame
        self.fill = None              # ("solid", #RRGGBB, alpha) or ("gradient", [(position, #RRGGBB, alpha)], angle, radial)
        self.line = None              # {"color", "alpha", "width", "dash", "head", "tail"}
        self.opacity = 1.0
        self.paragraphs = []          # see read_text
        self.padding = (4.0, 4.0, 4.0, 4.0)  # left, top, right, bottom, in points
        self.anchor = "t"             # where text sits in its box: t, ctr or b
        self.autofit = False          # a text box whose height follows its text
        self.image = None             # the picture's bytes, in a format Excel shows
        self.image_name = "image.png"
        self.pixels = None            # the picture's (width, height, horizontal dpi, vertical dpi)
        self.crop = None              # fractions of the picture cut off: (left, top, right, bottom)
        self.name = ""                # how Excel lists it
        self.link = None              # a web or mail link it opens when clicked
        self.size = None              # (w, h) in pixels

    @property
    def counts_as(self):
        return "images" if self.kind == "image" else "shapes"

    @property
    def text(self):
        return "\n".join("".join(text for text, _ in p["runs"]) for p in self.paragraphs).replace("\u2028", "\n")


# ---------------------------------------------------------------- reading


def read_drawings(book, tools):
    """[Drawing] for every text box, shape and image on the file's sheets, in stacking order. `tools`
    gives the Windows font for a Mac font (font, bold) and notes Mac fonts for the change report
    (note). Anything that can't be read is left out (and stays counted as not converted)."""
    model = book.doc._model
    found = []
    for sheet_id, _, _ in book.sheets:
        for oid, obj, place in sheet_drawables(model, sheet_id):
            kind = KINDS.get(type(obj).__name__)
            if kind is None:
                continue
            try:
                drawing = read_image(model, oid, obj, place) if kind == "images" else read_shape(model, oid, obj, place, tools)
            except Exception:  # stored in a way we don't recognise
                drawing = None
            if drawing is not None:
                drawing.sheet_id = sheet_id
                found.append(drawing)
    return found


def sheet_drawables(model, sheet_id):
    """(id, archive, place) for everything on a sheet, groups opened up, in stacking order. `place`
    turns a geometry where the object sits (on the sheet or in its group) into (its frame on the
    sheet, its rotation clockwise, its flips)."""
    def walk(ids, place, depth):
        for oid in ids:
            if oid not in model.objects:
                continue
            obj = model.objects[oid]
            if type(obj).__name__ == "GroupArchive":
                if depth < 8:
                    yield from walk([child.identifier for child in obj.children], group_place(model, obj, place), depth + 1)
            else:
                yield oid, obj, place
    yield from walk([ref.identifier for ref in model.objects[sheet_id].drawable_infos], sheet_place, 0)


def geometry_of(obj):
    """(x, y, w, h, angle, flags) of a drawable's geometry, from the DrawableArchive it extends."""
    for _ in range(4):
        if obj.DESCRIPTOR.full_name == "TSD.DrawableArchive":
            g = obj.geometry
            return g.position.x, g.position.y, g.size.width, g.size.height, g.angle, g.flags
        if "super" not in obj.DESCRIPTOR.fields_by_name:
            return None
        obj = obj.super
    return None


def sheet_place(geometry):
    """Where something on the sheet itself sits. Numbers turns things counter-clockwise, Excel
    clockwise; flags 4 and 8 flip it."""
    x, y, w, h, angle, flags = geometry
    return (x, y, w, h), (360 - angle) % 360, (bool(flags & 4), bool(flags & 8))


def group_place(model, group, outer):
    """How geometries inside a group map onto the sheet: the box around the group's children becomes
    the group's frame (which works whether children are measured from the group's corner or the
    sheet's, and for a group resized since), turned with the group."""
    boxes = [geometry_of(model.objects[ref.identifier]) for ref in group.children if ref.identifier in model.objects]
    boxes = [b for b in boxes if b is not None]
    (gx, gy, gw, gh), turned, _ = outer(geometry_of(group))
    if not boxes:
        return outer
    x0, y0 = min(b[0] for b in boxes), min(b[1] for b in boxes)
    x1, y1 = max(b[0] + b[2] for b in boxes), max(b[1] + b[3] for b in boxes)
    sx = gw / (x1 - x0) if x1 - x0 > 0.01 and gw > 0 else 1.0
    sy = gh / (y1 - y0) if y1 - y0 > 0.01 and gh > 0 else 1.0

    def place(geometry):
        x, y, w, h, angle, flags = geometry
        x, y, w, h = gx + (x - x0) * sx, gy + (y - y0) * sy, w * sx, h * sy
        if turned:  # turn the child's centre about the group's
            cx, cy, a = gx + gw / 2, gy + gh / 2, math.radians(turned)
            dx, dy = x + w / 2 - cx, y + h / 2 - cy
            x = cx + dx * math.cos(a) - dy * math.sin(a) - w / 2
            y = cy + dx * math.sin(a) + dy * math.cos(a) - h / 2
        return (x, y, w, h), ((360 - angle) + turned) % 360, (bool(flags & 4), bool(flags & 8))
    return place


def read_shape(model, oid, obj, place, tools):
    info = obj.DESCRIPTOR.full_name == "TSWP.ShapeInfoArchive"
    shape = obj.super if info else obj
    geometry = geometry_of(shape)
    (x, y, w, h), rotation, flips = place(geometry)
    styles = styles_of(model, shape.style.identifier)
    source = shape.pathsource
    d = Drawing(oid, "text" if info and obj.is_text_box else "shape")
    d.rotation = rotation
    d.flip = (flips[0] != source.horizontalFlip, flips[1] != source.verticalFlip)
    natural = read_outline(d, source)
    if d.outline == "line":
        d.kind = "line"
        (x, y, w, h) = line_frame(d, (x, y, w, h))
    stroke = first(styles, shape_value("stroke"))
    d.line = read_stroke(stroke)
    if d.line is not None:
        d.line["head"] = line_end(shape.head_line_end if shape.HasField("head_line_end") else first(styles, shape_value("head_line_end")))
        d.line["tail"] = line_end(shape.tail_line_end if shape.HasField("tail_line_end") else first(styles, shape_value("tail_line_end")))
    fill = first(styles, shape_value("fill"))
    d.fill = read_fill(fill) if fill is not None and d.kind != "line" else None
    opacity = first(styles, shape_value("opacity"))
    d.opacity = max(0.0, min(1.0, opacity)) if opacity is not None else 1.0

    padding = first(styles, text_value_of("padding"))
    if padding is not None:
        d.padding = (padding.left, padding.top, padding.right, padding.bottom)
    d.anchor = ANCHORS.get(first(styles, text_value_of("vertical_alignment")), "t")
    default = first(styles, text_value_of("paragraph_style"))
    if info and obj.HasField("owned_storage") and obj.owned_storage.identifier in model.objects:
        d.paragraphs = read_text(model, model.objects[obj.owned_storage.identifier], default.identifier if default else 0, tools)
    if not d.text.strip():
        d.paragraphs = [] if d.kind != "text" else d.paragraphs
    if d.kind == "shape" and d.paragraphs and d.fill is None and d.line is None and d.outline == "rect":
        d.kind = "text"  # a shape that only holds text, like Numbers' own "Text" shape
    # A text box's width, height or both can follow its text (geometry flags 1 and 2 say which are
    # fixed): its natural size is then the size Numbers last laid it out at.
    flags = geometry[5]
    if d.kind != "line":
        if not flags & 1 or w <= 0:
            w = natural[0] if natural[0] > 0 else (text_width_needed(d) if d.paragraphs else w)
        if not flags & 2 or h <= 0:
            d.autofit = bool(d.paragraphs)
            if natural[1] > 0 and not (flags & 2 and h > 0):
                h = natural[1]
            elif h <= 0 and d.paragraphs:
                h = text_height(d.paragraphs, w, d.padding)
    d.frame = (x, y, max(w, 0.0), max(h, 0.0))
    d.name = one_line(d.text) or ("Line" if d.kind == "line" else "Shape")
    d.link = web_link(obj)
    least = 0 if d.kind == "line" else 1  # a straight line across or down has no height or width
    d.size = (max(least, round(d.frame[2] * PT_TO_PX)), max(least, round(d.frame[3] * PT_TO_PX)))
    return d


def one_line(text, limit=80):
    text = re.sub(r"\s+", " ", text or "").strip()
    return text[:limit - 1] + "…" if len(text) > limit else text


# ---------------------------------------------------------------- outlines


def read_outline(d, source):
    """The outline of a shape: d.outline "rect", "line" (d.path holding its two ends) or "path"
    (d.path in a box d.path_size, stretched to the frame). Returns the shape's natural size: its
    size in Numbers, including any it takes from its text."""
    if source.HasField("bezier_path_source") or source.HasField("connection_line_path_source"):
        bezier = source.bezier_path_source if source.HasField("bezier_path_source") else source.connection_line_path_source.super
        path = [(PATH_COMMANDS.get(e.type), [(p.x, p.y) for p in e.points]) for e in bezier.path.elements]
        path = [step for step in path if step[0]]
        while path and path[-1][0] == "M":  # a move with nothing drawn after it
            path.pop()
        d.path, d.path_size = fitted(path)
        if [cmd for cmd, _ in d.path] == ["M", "L"]:
            d.outline = "line"
        elif is_box(d.path, d.path_size):
            d.outline, d.path = "rect", None
        else:
            d.outline = "path"
        return bezier.naturalSize.width, bezier.naturalSize.height
    if source.HasField("editable_bezier_path_source"):
        editable = source.editable_bezier_path_source
        path = []
        for sub in editable.subpaths:
            nodes = list(sub.nodes)
            if not nodes:
                continue
            path.append(("M", [(nodes[0].nodePoint.x, nodes[0].nodePoint.y)]))
            pairs = list(zip(nodes, nodes[1:])) + ([(nodes[-1], nodes[0])] if sub.closed else [])
            for a, b in pairs:
                path.append(("C", [(a.outControlPoint.x, a.outControlPoint.y), (b.inControlPoint.x, b.inControlPoint.y),
                                   (b.nodePoint.x, b.nodePoint.y)]))
            if sub.closed:
                path.append(("Z", []))
        d.outline, (d.path, d.path_size) = "path", fitted(path)
        return editable.naturalSize.width, editable.naturalSize.height
    for name, build in (("scalar_path_source", scalar_path), ("point_path_source", point_path),
                        ("callout_path_source", callout_path)):
        if source.HasField(name):
            built = build(getattr(source, name))
            if built is not None:
                d.outline, (d.path, d.path_size) = "path", built
            size = getattr(source, name).naturalSize if name != "callout_path_source" else source.callout_path_source.natural_size
            return size.width, size.height
    return 0.0, 0.0


def fitted(path):
    """A bezier path moved to start at (0, 0), and the size of the box around it: Numbers stretches a
    shape's path to fill its natural size, whatever size the path was drawn at."""
    points = [p for _, ps in path for p in ps]
    if not points:
        return path, (0.0, 0.0)
    x0, y0 = min(x for x, _ in points), min(y for _, y in points)
    x1, y1 = max(x for x, _ in points), max(y for _, y in points)
    return [(cmd, [(x - x0, y - y0) for x, y in ps]) for cmd, ps in path], (x1 - x0, y1 - y0)


def is_box(path, size):
    """Whether a path is just the rectangle filling its box."""
    w, h = size
    corners = {(round(x, 1), round(y, 1)) for cmd, points in path for x, y in points}
    drawn = [cmd for cmd, _ in path]
    return drawn in (["M", "L", "L", "L", "Z"], ["M", "L", "L", "L", "L", "Z"]) and corners == {
        (0.0, 0.0), (round(w, 1), 0.0), (round(w, 1), round(h, 1)), (0.0, round(h, 1))}


def line_frame(d, frame):
    """A straight line's frame: the box between its ends, flipped so Excel draws it the same way
    (Excel's line runs from the box's top-left corner to its bottom-right one)."""
    (x0, y0), (x1, y1) = d.path[0][1][0], d.path[1][1][0]
    x, y, w, h = frame
    nw, nh = d.path_size
    sx, sy = (w / nw if nw else 1.0), (h / nh if nh else 1.0)
    ax, ay, bx, by = x + x0 * sx, y + y0 * sy, x + x1 * sx, y + y1 * sy
    d.flip = (d.flip[0] != (ax > bx), d.flip[1] != (ay > by))
    d.path = None
    return min(ax, bx), min(ay, by), abs(bx - ax), abs(by - ay)


def rounded_box(w, h, r, path=None):
    """A rectangle with corners of radius r, clockwise from the top-left."""
    r = max(0.0, min(r, w / 2, h / 2))
    k = r * (1 - KAPPA)
    path = path if path is not None else []
    path += [("M", [(r, 0)]), ("L", [(w - r, 0)])]
    if r:
        path.append(("C", [(w - k, 0), (w, k), (w, r)]))
    path.append(("L", [(w, h - r)]))
    if r:
        path.append(("C", [(w, h - k), (w - k, h), (w - r, h)]))
    path.append(("L", [(r, h)]))
    if r:
        path.append(("C", [(k, h), (0, h - k), (0, h - r)]))
    path.append(("L", [(0, r)]))
    if r:
        path.append(("C", [(0, k), (k, 0), (r, 0)]))
    path.append(("Z", []))
    return path


def polygon(points):
    return [("M", [points[0]])] + [("L", [p]) for p in points[1:]] + [("Z", [])]


def scalar_path(source):
    """Rounded rectangles (scalar: corner radius), regular polygons (scalar: sides) and chevrons."""
    w, h = source.naturalSize.width, source.naturalSize.height
    if w <= 0 or h <= 0:
        return None
    if source.type == 0:
        return rounded_box(w, h, source.scalar), (w, h)
    if source.type == 1:
        sides = max(3, min(64, round(source.scalar or 5)))
        return polygon([(w / 2 + w / 2 * math.cos(math.radians(-90 + 360 * i / sides)),
                         h / 2 + h / 2 * math.sin(math.radians(-90 + 360 * i / sides))) for i in range(sides)]), (w, h)
    if source.type == 2:
        depth = source.scalar if 0 < source.scalar < w else (source.scalar * w if 0 < source.scalar < 1 else h / 2)
        return polygon([(0, 0), (w - depth, 0), (w, h / 2), (w - depth, h), (0, h), (depth, h / 2)]), (w, h)
    return None


def point_path(source):
    """Arrows (point: where the head starts and the shaft's edge), stars (point: number of points
    and inner radius) and plus signs (point: the inner corner)."""
    w, h = source.naturalSize.width, source.naturalSize.height
    px, py = source.point.x, source.point.y
    if w <= 0 or h <= 0:
        return None
    if source.type in (0, 1, 10):
        head = px if 0 < px < w else max(0.0, w - h / 2)
        shaft = py * h if 0 < py < 1 else (py if 1 <= py < h / 2 else h / 4)
        if source.type == 1:
            points = [(0, shaft), (head, shaft), (head, 0), (w, h / 2), (head, h), (head, h - shaft), (0, h - shaft)]
        elif source.type == 0:
            points = [(w, shaft), (w - head, shaft), (w - head, 0), (0, h / 2), (w - head, h), (w - head, h - shaft),
                      (w, h - shaft)]
        else:
            head = min(head, w / 2) if head < w / 2 else w - head
            points = [(0, h / 2), (head, 0), (head, shaft), (w - head, shaft), (w - head, 0), (w, h / 2), (w - head, h),
                      (w - head, h - shaft), (head, h - shaft), (head, h)]
        return polygon(points), (w, h)
    if source.type == 100:
        count = round(px) if 3 <= px <= 64 else 5
        inner = py if 0 < py < 1 else 0.38
        points = []
        for i in range(count * 2):
            radius = 1.0 if i % 2 == 0 else inner
            angle = math.radians(-90 + 180 * i / count)
            points.append((w / 2 + w / 2 * radius * math.cos(angle), h / 2 + h / 2 * radius * math.sin(angle)))
        return polygon(points), (w, h)
    if source.type == 200:
        ix = px if 0 < px < w / 2 else (px * w if 0 < px < 0.5 else w / 3)
        iy = py if 0 < py < h / 2 else (py * h if 0 < py < 0.5 else h / 3)
        return polygon([(ix, 0), (w - ix, 0), (w - ix, iy), (w, iy), (w, h - iy), (w - ix, h - iy), (w - ix, h), (ix, h),
                        (ix, h - iy), (0, h - iy), (0, iy), (ix, iy)]), (w, h)
    return None


def callout_path(source):
    """A speech bubble: a rounded box with a tail pointing at tail_position."""
    w, h = source.natural_size.width, source.natural_size.height
    if w <= 0 or h <= 0:
        return None
    tx, ty = source.tail_position.x, source.tail_position.y
    r = max(0.0, min(source.corner_radius, w / 2, h / 2))
    path = rounded_box(w, h, r)
    if 0 <= tx <= w and 0 <= ty <= h:
        return path, (w, h)  # the tail is inside the box
    half = max(2.0, (source.tail_size or min(w, h) / 4) / 2)
    along = lambda tip, size: size / 2 if source.center_tail else max(r + half, min(size - r - half, tip))  # noqa: E731
    # The tail leaves the side facing its tip, opposite the tip (or from the middle), drawn into that
    # side as the outline goes round clockwise: before the line that ends the side.
    if ty > h:
        c = along(tx, w)
        tail, side_end = [(c + half, h), (tx, ty), (c - half, h)], (r, h)
    elif ty < 0:
        c = along(tx, w)
        tail, side_end = [(c - half, 0), (tx, ty), (c + half, 0)], (w - r, 0)
    elif tx > w:
        c = along(ty, h)
        tail, side_end = [(w, c - half), (tx, ty), (w, c + half)], (w, h - r)
    else:
        c = along(ty, h)
        tail, side_end = [(0, c + half), (tx, ty), (0, c - half)], (0, r)
    at = next(i for i, (cmd, points) in enumerate(path) if cmd == "L" and points[0] == side_end)
    path[at:at] = [("L", [p]) for p in tail]
    return path, (w, h)


# ---------------------------------------------------------------- styles


def styles_of(model, style_id):
    """A style and those it inherits from, most specific first."""
    chain = []
    while style_id and style_id in model.objects and len(chain) < 12:
        style = model.objects[style_id]
        chain.append(style)
        base = style
        while "parent" not in base.DESCRIPTOR.fields_by_name and "super" in base.DESCRIPTOR.fields_by_name:
            base = base.super
        style_id = base.parent.identifier if "parent" in base.DESCRIPTOR.fields_by_name and base.HasField("parent") else 0
    return chain


def first(styles, get):
    """The first value `get` finds along a style chain (None when no style sets it)."""
    for style in styles:
        value = get(style)
        if value is not None:
            return value
    return None


def shape_value(name):
    """Gets a drawing property (fill, stroke, opacity...) a shape or picture style sets itself."""
    def get(style):
        if style.DESCRIPTOR.full_name == "TSWP.ShapeStyleArchive":
            style = style.super
        props = None
        if style.DESCRIPTOR.full_name == "TSD.ShapeStyleArchive" and style.HasField("shape_properties"):
            props = style.shape_properties
        elif style.DESCRIPTOR.full_name == "TSD.MediaStyleArchive" and style.HasField("media_properties"):
            props = style.media_properties
        return getattr(props, name) if props is not None and name in props.DESCRIPTOR.fields_by_name and props.HasField(name) else None
    return get


def text_value_of(name):
    """Gets a text property (padding, vertical alignment, paragraph style) a shape style sets itself."""
    def get(style):
        if style.DESCRIPTOR.full_name != "TSWP.ShapeStyleArchive" or not style.HasField("shape_properties"):
            return None
        props = style.shape_properties
        return getattr(props, name) if props.HasField(name) else None
    return get


def colour_of(colour):
    """(#RRGGBB, alpha) of a TSP colour."""
    return ("#{:02X}{:02X}{:02X}".format(*(max(0, min(255, round(v * 255))) for v in (colour.r, colour.g, colour.b))),
            colour.a if colour.HasField("a") else 1.0)


def read_fill(fill):
    if fill.HasField("color"):
        return ("solid", *colour_of(fill.color))
    if fill.HasField("gradient") and fill.gradient.stops:
        g = fill.gradient
        stops = [(max(0.0, min(1.0, s.fraction)), *colour_of(s.color)) for s in g.stops]
        angle = g.anglegradient.gradientangle if g.HasField("anglegradient") else 0.0
        return ("gradient", stops, angle, g.type == 1)
    return None  # no fill, or a picture fill


def read_stroke(stroke):
    """{color, alpha, width, dash} of a line that's drawn, else None."""
    if stroke is None or not stroke.HasField("color") or stroke.width <= 0:
        return None
    pattern = stroke.pattern
    if pattern.HasField("type") and pattern.type == 2:  # TSDEmptyPattern: no line
        return None
    dash = None
    if pattern.type == 0 and pattern.count:  # dashes, as lengths in line widths
        length = pattern.pattern[0]
        dash = "sysDot" if length <= 1.5 else "dash" if length <= 4 else "lgDash"
    colour, alpha = colour_of(stroke.color)
    return {"color": colour, "alpha": alpha, "width": stroke.width, "dash": dash, "head": None, "tail": None}


def line_end(end):
    """The DrawingML arrowhead for a Numbers line end, or None."""
    name = (end.identifier if end is not None and end.HasField("identifier") else "").lower()
    if not name or name == "none":
        return None
    if "arrow" in name:
        return "arrow" if "open" in name or "simple" in name else "triangle"
    if "circle" in name or "dot" in name:
        return "oval"
    if "diamond" in name or "square" in name:
        return "diamond"
    return None


# ---------------------------------------------------------------- text


def text_value(model, style_id, group, name):
    """A character or paragraph property (group "char_properties" or "para_properties") from a text
    style or those it inherits from, else None."""
    for style in styles_of(model, style_id):
        if group in style.DESCRIPTOR.fields_by_name:
            props = getattr(style, group)
            if props.HasField(name):
                return getattr(props, name)
    return None


def read_text(model, storage, default_style, tools):
    """A text storage's paragraphs: [{"align", "before", "after", "spacing" (kind, amount), "left",
    "right", "first" (indents in points), "bullet" (None, {"char"} or {"auto"}), "runs": [(text,
    style)], "end" (the style of an empty paragraph)}]. Line breaks within a paragraph stay as
    U+2028; objects anchored in the text are left out."""
    text = storage.text[0] if storage.text else ""
    entries = lambda table: [(e.character_index, e.object.identifier if e.HasField("object") else 0)  # noqa: E731
                             for e in table.entries]
    para_styles, char_styles, list_styles = (entries(storage.table_para_style), entries(storage.table_char_style),
                                             entries(storage.table_list_style))
    levels = [(e.character_index, e.first) for e in storage.table_para_data.entries]

    def at(table, index, default=0):
        found = default
        for start, value in table:
            if start > index:
                break
            found = value
        return found

    paragraphs, start, style_id = [], 0, default_style
    for end in [m.start() for m in re.finditer("[\n\u2029]", text)] + [len(text)]:
        style_id = at(para_styles, start) or style_id
        p = paragraph_format(model, style_id)
        runs = []
        bounds = sorted({start, end} | {i for i, _ in char_styles if start < i < end})
        for a, b in zip(bounds, bounds[1:]):
            piece = clean(text[a:b])
            if piece:
                style = run_style(model, at(char_styles, a), style_id, tools)
                if runs and runs[-1][1] == style:
                    runs[-1] = (runs[-1][0] + piece, style)
                else:
                    runs.append((piece, style))
        p["runs"] = runs
        p["end"] = runs[-1][1] if runs else run_style(model, at(char_styles, start), style_id, tools)
        p["bullet"] = bullet(model, at(list_styles, start), max(0, min(8, at(levels, start))))
        paragraphs.append(p)
        start = end + 1
    while len(paragraphs) > 1 and not paragraphs[-1]["runs"] and not paragraphs[-1]["bullet"]:
        paragraphs.pop()  # Numbers ends its text with a paragraph mark
    return paragraphs


def clean(text):
    """Text as Excel can hold it: no objects anchored in it, no control characters but tabs."""
    return re.sub(r"[\x00-\x08\x0b-\x1f\ufffc\ufffe\uffff]", "", text.replace("\x0b", "\u2028"))


def paragraph_format(model, style_id):
    get = lambda name: text_value(model, style_id, "para_properties", name)  # noqa: E731
    spacing, amount = "pct", 1.0
    line = get("line_spacing")
    if line is not None and line.HasField("amount"):
        spacing = {0: "pct", 1: "min", 2: "exact", 3: "max", 4: "between"}.get(line.mode, "pct")
        amount = line.amount if (line.amount or spacing != "pct") else 1.0
    return {"align": ALIGNMENTS.get(get("alignment"), "l"), "before": get("space_before") or 0.0,
            "after": get("space_after") or 0.0, "spacing": (spacing, amount), "left": get("left_indent") or 0.0,
            "right": get("right_indent") or 0.0, "first": get("first_line_indent") or 0.0}


def run_style(model, char_id, para_id, tools):
    """{font, size, b, i, u, st, color, alpha, script, caps} of text with a character style (0 for
    none) in a paragraph style; whatever the character style sets wins."""
    def get(name):
        for style_id in (char_id, para_id):
            if style_id:
                value = text_value(model, style_id, "char_properties", name)
                if value is not None and value != "" and not (name == "font_size" and value <= 0):
                    return value
        return None

    name = get("font_name")
    tools.note(name)
    colour = get("font_color")
    colour, alpha = colour_of(colour) if colour is not None else ("#000000", 1.0)
    return {"font": tools.font(name) or "Arial", "size": round(get("font_size") or 12.0, 1),
            "b": bool(get("bold")) or tools.bold(name), "i": bool(get("italic")),
            "u": {1: "sng", 2: "dbl", 3: "wavy"}.get(get("underline")),
            "st": {1: "sngStrike", 2: "dblStrike", 3: "dblStrike"}.get(get("strikethru")),
            "color": colour, "alpha": round(alpha, 3), "script": {1: "super", 2: "sub"}.get(get("superscript")),
            "caps": {1: "all", 2: "small"}.get(get("capitalization"))}


def bullet(model, list_id, level):
    """How a paragraph in a list is marked: {"char": "•"} or {"auto": "arabicPeriod"}, with "indent"
    (where the text starts) and "hang" (how far the mark sits before it), or None."""
    if not list_id or list_id not in model.objects:
        return None
    style = model.objects[list_id]
    pick = lambda values, default: values[level] if level < len(values) else (values[-1] if len(values) else default)  # noqa: E731
    kind = pick(style.label_types, 0)
    if kind == 0:
        return None
    hang = pick(style.text_indents, 0.0) or 18.0
    mark = {"indent": pick(style.indents, 0.0) + hang, "hang": hang}
    if kind == 3:
        mark["auto"] = NUMBERING.get(pick(style.number_types, 0), "arabicPeriod")
    else:
        mark["char"] = (pick(style.strings, "•") or "•") if kind == 2 else "•"
    return mark


def text_width(text, style):
    """Roughly how wide text is, in points (for working out where it wraps)."""
    width = 0.0
    for ch in text:
        width += (0.28 if ch == " " else 0.3 if ch in NARROW else 0.85 if ch in WIDE else 0.56 if ch.isdigit()
                  else 0.66 if ch.isupper() else 0.52)
    return width * style["size"] * (1.06 if style["b"] else 1.0) * (1.08 if style["font"] == "Segoe UI" else 1.0)


def line_height(p, style):
    natural = style["size"] * LINE_HEIGHTS.get(style["font"], 1.2)
    kind, amount = p["spacing"]
    return {"pct": natural * amount, "exact": amount, "min": max(natural, amount), "max": min(natural, amount),
            "between": natural + amount}.get(kind, natural)


def text_height(paragraphs, width, padding):
    """Roughly how tall text is in a box `width` points wide, laid out in the Windows fonts Excel uses."""
    left, top, right, bottom = padding
    height = top + bottom
    for p in paragraphs:
        runs = p["runs"] or [("", p["end"])]
        tallest = max((style for _, style in runs), key=lambda s: s["size"])
        indent = p["left"] + (p["bullet"]["indent"] if p["bullet"] else 0)
        room = max(10.0, width - left - right - indent - p["right"])
        height += p["before"] + count_lines(runs, room) * line_height(p, tallest) + p["after"]
    return height


def count_lines(runs, room):
    lines, x = 1, 0.0
    for text, style in runs:
        for i, part in enumerate(text.split("\u2028")):
            if i:
                lines, x = lines + 1, 0.0
            for word in re.findall(r"\S+\s*|\s+", part):
                if x > 0 and x + text_width(word.rstrip(), style) > room:
                    lines, x = lines + 1, 0.0
                x += text_width(word, style)
                while x > room * 1.02:  # a word longer than the line breaks across lines
                    lines, x = lines + 1, x - room
    return lines


def text_width_needed(d):
    """The width a text box that grows sideways needs for its longest line."""
    longest = 0.0
    for p in d.paragraphs:
        for line in "".join(text for text, _ in p["runs"]).split("\u2028"):
            style = p["runs"][0][1] if p["runs"] else p["end"]
            longest = max(longest, text_width(line, style))
    return longest + d.padding[0] + d.padding[2] + 2


# ---------------------------------------------------------------- pictures


def package_file(model, data_id):
    """(bytes, file name) of a file in the package's Data folder, or None (some, like a theme's
    pictures, are kept by Numbers itself rather than in the file)."""
    info = next((d for d in model.objects[PACKAGE_ID].datas if d.identifier == data_id), None)
    if info is None or not info.file_name:
        return None
    store = model.objects.file_store
    for key in (f"Data/{info.file_name}", f"Data\\{info.file_name}"):
        if key in store and isinstance(store[key], (bytes, bytearray)):
            return bytes(store[key]), info.preferred_file_name or info.file_name
    return None


def picture_size(data):
    """(width, height, horizontal dpi, vertical dpi) of a picture Excel can show, else None (Excel
    can't show HEIC, TIFF, PDF or SVG pictures)."""
    try:
        image = Image(io.BytesIO(data))
        return image.width, image.height, image.x_dpi, image.y_dpi
    except Exception:
        return None


def read_image(model, oid, obj, place):
    """A picture, or a movie's poster frame: its file (the adjusted picture where Numbers keeps one,
    or the preview it keeps of a picture Excel can't show), where it's shown, and how it's cropped."""
    movie = obj.DESCRIPTOR.full_name == "TSD.MovieArchive"
    names = ("posterImageData",) if movie else ("adjustedImageData", "enhancedImageData", "data",
                                                  "thumbnailAdjustedImageData", "thumbnailData")
    picture = None
    for name in names:
        if obj.HasField(name):
            found = package_file(model, getattr(obj, name).identifier)
            size = picture_size(found[0]) if found else None
            if size:
                picture = found + (size,)
                break
    if picture is None:
        return None
    d = Drawing(oid, "image")
    d.image, d.image_name, d.pixels = picture
    geometry = geometry_of(obj)
    x, y, w, h, angle, flags = geometry
    if w <= 0 or h <= 0:
        return None
    if not movie and obj.HasField("mask") and obj.mask.identifier in model.objects:
        mask = model.objects[obj.mask.identifier]
        mx, my, mw, mh = geometry_of(mask)[:4]
        if not (-1 <= mx <= w and -1 <= my <= h) and x - 1 <= mx <= x + w and y - 1 <= my <= y + h:
            mx, my = mx - x, my - y  # measured from the sheet's corner rather than the picture's
        left, top = max(0.0, mx), max(0.0, my)
        right, bottom = min(w, mx + mw), min(h, my + mh)
        if right - left > 0.5 and bottom - top > 0.5:
            d.crop = (left / w, top / h, 1 - right / w, 1 - bottom / h)
            outline = Drawing(0, "shape")
            read_outline(outline, mask.pathsource)
            if outline.outline == "path":
                d.outline, d.path, d.path_size = "path", outline.path, outline.path_size
            geometry = (x + left, y + top, right - left, bottom - top, angle, flags)
    (fx, fy, fw, fh), d.rotation, d.flip = place(geometry)
    d.frame = (fx, fy, fw, fh)
    styles = styles_of(model, obj.style.identifier) if obj.HasField("style") else []
    d.line = read_stroke(first(styles, shape_value("stroke")))
    opacity = first(styles, shape_value("opacity"))
    d.opacity = max(0.0, min(1.0, opacity)) if opacity is not None else 1.0
    d.name = one_line(drawable_field(obj, "accessibility_description")) or d.image_name
    d.link = web_link(obj)
    d.size = (max(1, round(fw * PT_TO_PX)), max(1, round(fh * PT_TO_PX)))
    return d


def drawable_field(obj, name):
    """A DrawableArchive setting of a text box, shape or picture (accessibility_description: its
    description for screen readers; hyperlink_url: the link it opens), or ""."""
    for _ in range(3):
        if obj.DESCRIPTOR.full_name == "TSD.DrawableArchive":
            return getattr(obj, name) if obj.HasField(name) else ""
        obj = obj.super
    return ""


def web_link(obj):
    """The web or mail link a text box, shape or picture opens when clicked, else None."""
    link = drawable_field(obj, "hyperlink_url").strip()
    return link if re.match(r"(https?|mailto):", link, re.I) and len(link) <= 2000 else None


# ---------------------------------------------------------------- writing


def write_drawings(layout, worksheets):
    """Adds every text box, shape and picture to the worksheet the layout puts it on; returns how
    many were written (any that fail are counted as not converted)."""
    written = 0
    for plan in layout.sheets:
        ws = worksheets[plan.name]
        for d, (x, y) in plan.drawings:
            options = {"x_offset": x, "y_offset": y, "object_position": 1, **({"url": d.link} if d.link else {})}
            try:
                if d.kind == "image":
                    width, height, x_dpi, y_dpi = d.pixels
                    picture = Image(io.BytesIO(d.image))
                    picture.image_name = d.image_name
                    ws.insert_image(0, 0, picture, {**options, "description": d.name,
                                                    "x_scale": d.size[0] / (width * 96.0 / x_dpi),
                                                    "y_scale": d.size[1] / (height * 96.0 / y_dpi)})
                    ws.images[-1].numbers = d
                else:
                    ws.insert_textbox(0, 0, d.text, {**options, "width": d.size[0], "height": d.size[1], "numbers": d})
                written += 1
            except Exception:
                layout.book.unsupported[d.counts_as] += 1
    return written


def emu(points):
    return int(round(points * EMU))


def write_colour(xml, colour, alpha):
    """<a:srgbClr>, with how opaque it is."""
    if alpha < 0.995:
        xml._xml_start_tag("a:srgbClr", [("val", colour[1:])])
        xml._xml_empty_tag("a:alpha", [("val", int(round(max(0.0, alpha) * 100000)))])
        xml._xml_end_tag("a:srgbClr")
    else:
        xml._xml_empty_tag("a:srgbClr", [("val", colour[1:])])


def write_fill(xml, fill, opacity):
    if fill is None:
        xml._xml_empty_tag("a:noFill")
    elif fill[0] == "solid":
        xml._xml_start_tag("a:solidFill")
        write_colour(xml, fill[1], fill[2] * opacity)
        xml._xml_end_tag("a:solidFill")
    else:
        _, stops, angle, radial = fill
        xml._xml_start_tag("a:gradFill", [("rotWithShape", "1")])
        xml._xml_start_tag("a:gsLst")
        for position, colour, alpha in stops:
            xml._xml_start_tag("a:gs", [("pos", int(round(position * 100000)))])
            write_colour(xml, colour, alpha * opacity)
            xml._xml_end_tag("a:gs")
        xml._xml_end_tag("a:gsLst")
        if radial:
            xml._xml_start_tag("a:path", [("path", "circle")])
            xml._xml_empty_tag("a:fillToRect", [("l", 50000), ("t", 50000), ("r", 50000), ("b", 50000)])
            xml._xml_end_tag("a:path")
        else:
            xml._xml_empty_tag("a:lin", [("ang", int(round((360 - angle) % 360 * 60000))), ("scaled", "0")])
        xml._xml_end_tag("a:gradFill")


def write_line(xml, line, opacity):
    if line is None:
        xml._xml_start_tag("a:ln")
        xml._xml_empty_tag("a:noFill")
        xml._xml_end_tag("a:ln")
        return
    xml._xml_start_tag("a:ln", [("w", emu(line["width"]))])
    xml._xml_start_tag("a:solidFill")
    write_colour(xml, line["color"], line["alpha"] * opacity)
    xml._xml_end_tag("a:solidFill")
    xml._xml_empty_tag("a:prstDash", [("val", line["dash"] or "solid")])
    for end in ("head", "tail"):
        if line.get(end):
            xml._xml_empty_tag(f"a:{end}End", [("type", line[end])])
    xml._xml_end_tag("a:ln")


def write_geometry(xml, d):
    """<a:prstGeom> for rectangles and lines, <a:custGeom> with the outline for anything else."""
    if d.outline != "path" or not d.path:
        xml._xml_start_tag("a:prstGeom", [("prst", "line" if d.outline == "line" else "rect")])
        xml._xml_empty_tag("a:avLst")
        xml._xml_end_tag("a:prstGeom")
        return
    w, h = d.path_size
    scale = lambda v: int(round(v * 100))  # noqa: E731  (hundredths of a point keep curves smooth)
    xml._xml_start_tag("a:custGeom")
    for tag in ("a:avLst", "a:gdLst", "a:ahLst", "a:cxnLst"):
        xml._xml_empty_tag(tag)
    xml._xml_empty_tag("a:rect", [("l", "l"), ("t", "t"), ("r", "r"), ("b", "b")])
    xml._xml_start_tag("a:pathLst")
    closed = any(cmd == "Z" for cmd, _ in d.path)
    xml._xml_start_tag("a:path", [("w", max(1, scale(w))), ("h", max(1, scale(h)))] + ([] if closed else [("fill", "none")]))
    tags = {"M": "a:moveTo", "L": "a:lnTo", "Q": "a:quadBezTo", "C": "a:cubicBezTo"}
    for cmd, points in d.path:
        if cmd == "Z":
            xml._xml_empty_tag("a:close")
            continue
        xml._xml_start_tag(tags[cmd])
        for px, py in points:
            xml._xml_empty_tag("a:pt", [("x", scale(px)), ("y", scale(py))])
        xml._xml_end_tag(tags[cmd])
    xml._xml_end_tag("a:path")
    xml._xml_end_tag("a:pathLst")
    xml._xml_end_tag("a:custGeom")


def write_transform(xml, d, col_absolute, row_absolute, width, height):
    attributes = []
    rotation = round(d.rotation, 3) % 360
    if rotation and abs(rotation - 360) > 0.001:
        attributes.append(("rot", int(round(rotation * 60000))))
    if d.flip[0]:
        attributes.append(("flipH", "1"))
    if d.flip[1]:
        attributes.append(("flipV", "1"))
    xml._xml_start_tag("a:xfrm", attributes)
    xml._xml_empty_tag("a:off", [("x", col_absolute), ("y", row_absolute)])
    xml._xml_empty_tag("a:ext", [("cx", width), ("cy", height)])
    xml._xml_end_tag("a:xfrm")


def write_run_properties(xml, tag, style, opacity):
    attributes = [("lang", "en-US"), ("sz", max(100, min(400000, int(round(style["size"] * 100))))),  # 1 to 4000 pt
                  ("b", "1" if style["b"] else "0"), ("i", "1" if style["i"] else "0")]
    if style["u"]:
        attributes.append(("u", style["u"]))
    if style["st"]:
        attributes.append(("strike", style["st"]))
    if style["caps"]:
        attributes.append(("cap", style["caps"]))
    if style["script"]:
        attributes.append(("baseline", "30000" if style["script"] == "super" else "-25000"))
    xml._xml_start_tag(tag, attributes)
    xml._xml_start_tag("a:solidFill")
    write_colour(xml, style["color"], style["alpha"] * opacity)
    xml._xml_end_tag("a:solidFill")
    xml._xml_empty_tag("a:latin", [("typeface", style["font"])])
    xml._xml_empty_tag("a:cs", [("typeface", style["font"])])
    xml._xml_end_tag(tag)


def write_paragraph(xml, p, opacity):
    attributes = [("algn", p["align"])]
    mark = p["bullet"]
    left = p["left"] + (mark["indent"] if mark else 0)
    if left:
        attributes.append(("marL", emu(left)))
    if p["right"]:
        attributes.append(("marR", emu(p["right"])))
    first_line = -mark["hang"] if mark else p["first"]
    if first_line:
        attributes.append(("indent", emu(first_line)))
    xml._xml_start_tag("a:p")
    xml._xml_start_tag("a:pPr", attributes)
    kind, amount = p["spacing"]
    xml._xml_start_tag("a:lnSpc")
    if kind == "pct":
        xml._xml_empty_tag("a:spcPct", [("val", int(round(amount * 100000)))])
    else:
        xml._xml_empty_tag("a:spcPts", [("val", int(round(line_height(p, p["end"]) * 100)))])
    xml._xml_end_tag("a:lnSpc")
    for tag, value in (("a:spcBef", p["before"]), ("a:spcAft", p["after"])):
        xml._xml_start_tag(tag)
        xml._xml_empty_tag("a:spcPts", [("val", int(round(value * 100)))])
        xml._xml_end_tag(tag)
    if mark and "auto" in mark:
        xml._xml_empty_tag("a:buAutoNum", [("type", mark["auto"])])
    elif mark:
        xml._xml_empty_tag("a:buFont", [("typeface", "Arial")])
        xml._xml_empty_tag("a:buChar", [("char", mark["char"][:1])])
    else:
        xml._xml_empty_tag("a:buNone")
    xml._xml_end_tag("a:pPr")
    for text, style in p["runs"]:
        for i, piece in enumerate(text.split("\u2028")):
            if i:
                xml._xml_start_tag("a:br")
                write_run_properties(xml, "a:rPr", style, opacity)
                xml._xml_end_tag("a:br")
            if piece:
                xml._xml_start_tag("a:r")
                write_run_properties(xml, "a:rPr", style, opacity)
                xml._xml_data_element("a:t", piece)
                xml._xml_end_tag("a:r")
    write_run_properties(xml, "a:endParaRPr", p["end"], opacity)
    xml._xml_end_tag("a:p")


def _write_shape(xml, index, col_absolute, row_absolute, drawing, d):
    """<xdr:sp> for a text box, shape or line: its outline, fill and line, and its text with each
    paragraph's alignment and spacing and each run's font."""
    xml._xml_start_tag("xdr:sp", [("macro", ""), ("textlink", "")])
    xml._xml_start_tag("xdr:nvSpPr")
    xml._write_c_nv_pr(index + 1, drawing, d.name or f"Shape {index}")
    xml._xml_empty_tag("xdr:cNvSpPr", [("txBox", "1")] if d.kind == "text" else [])
    xml._xml_end_tag("xdr:nvSpPr")
    xml._xml_start_tag("xdr:spPr")
    write_transform(xml, d, col_absolute, row_absolute, drawing._width, drawing._height)
    write_geometry(xml, d)
    write_fill(xml, d.fill, d.opacity)
    write_line(xml, d.line, d.opacity)
    xml._xml_end_tag("xdr:spPr")
    if d.kind != "line":
        left, top, right, bottom = d.padding
        xml._xml_start_tag("xdr:txBody")
        xml._xml_start_tag("a:bodyPr", [("wrap", "square"), ("lIns", emu(left)), ("tIns", emu(top)), ("rIns", emu(right)),
                                        ("bIns", emu(bottom)), ("rtlCol", "0"), ("anchor", d.anchor)])
        xml._xml_empty_tag("a:spAutoFit" if d.autofit else "a:noAutofit")
        xml._xml_end_tag("a:bodyPr")
        xml._xml_empty_tag("a:lstStyle")
        for p in d.paragraphs:
            write_paragraph(xml, p, d.opacity)
        if not d.paragraphs:
            xml._xml_empty_tag("a:p")
        xml._xml_end_tag("xdr:txBody")
    xml._xml_end_tag("xdr:sp")


def _write_picture(xml, index, col_absolute, row_absolute, drawing, d):
    """<xdr:pic> for a picture: cropped, see-through, turned and outlined as in Numbers."""
    xml._xml_start_tag("xdr:pic")
    xml._xml_start_tag("xdr:nvPicPr")
    xml._write_c_nv_pr(index + 1, drawing, f"Picture {index}")
    xml._write_c_nv_pic_pr()
    xml._xml_end_tag("xdr:nvPicPr")
    xml._xml_start_tag("xdr:blipFill")
    blip = [("xmlns:r", "http://schemas.openxmlformats.org/officeDocument/2006/relationships"),
            ("r:embed", f"rId{drawing._rel_index}")]
    if d.opacity < 0.995:
        xml._xml_start_tag("a:blip", blip)
        xml._xml_empty_tag("a:alphaModFix", [("amt", int(round(d.opacity * 100000)))])
        xml._xml_end_tag("a:blip")
    else:
        xml._xml_empty_tag("a:blip", blip)
    if d.crop:
        xml._xml_empty_tag("a:srcRect", [(side, int(round(value * 100000))) for side, value in zip("ltrb", d.crop) if value > 0.0005])
    xml._write_a_stretch()
    xml._xml_end_tag("xdr:blipFill")
    xml._xml_start_tag("xdr:spPr")
    write_transform(xml, d, col_absolute, row_absolute, drawing._width, drawing._height)
    write_geometry(xml, d)
    if d.line is not None:
        write_line(xml, d.line, d.opacity)
    xml._xml_end_tag("xdr:spPr")
    xml._xml_end_tag("xdr:pic")


# XlsxWriter writes the shapes and pictures this module adds with the functions above; everything
# else (charts, comments, its own text boxes) as it always has.
_set_options, _write_sp, _write_pic, _prepare_image = (Shape._set_options, SheetDrawing._write_sp,
                                                       SheetDrawing._write_pic, Worksheet._prepare_image)


def _shape_options(self, options):
    _set_options(self, options)
    self.numbers = options.get("numbers")


def _sp(self, index, col_absolute, row_absolute, drawing):
    d = getattr(drawing._shape, "numbers", None) if drawing._shape is not None else None
    if d is None:
        return _write_sp(self, index, col_absolute, row_absolute, drawing)
    return _write_shape(self, index, col_absolute, row_absolute, drawing, d)


def _pic(self, index, col_absolute, row_absolute, drawing):
    d = getattr(drawing, "numbers", None)
    if d is None:
        return _write_pic(self, index, col_absolute, row_absolute, drawing)
    return _write_picture(self, index, col_absolute, row_absolute, drawing, d)


def _prepare(self, image, image_id, drawing_id):
    _prepare_image(self, image, image_id, drawing_id)
    if getattr(image, "numbers", None) is not None:
        self.drawing.drawings[-1].numbers = image.numbers


Shape._set_options, SheetDrawing._write_sp, SheetDrawing._write_pic, Worksheet._prepare_image = (
    _shape_options, _sp, _pic, _prepare)


# ---------------------------------------------------------------- preview

PREVIEW_PICTURE_LIMIT = 6_000_000  # bytes of pictures sent with one sheet's preview


def describe(d, position, budget):
    """What the app's preview needs to draw a text box, shape or picture as it will look in Excel:
    its place (`position`) and size in pixels from the sheet's top-left cell, outline ("path": [[M, x,
    y], [C, x1, y1, x2, y2, x, y], [Z]...] in a box "box", for anything but rectangles and lines),
    fill ({color} or {stops: [[0-1, color]], angle (clockwise, as Excel's), radial}), line and text,
    or its picture as a data: URL (while `budget`, a one-item list of bytes left, lasts)."""
    def css(colour, alpha):
        r, g, b = (int(colour[i:i + 2], 16) for i in (1, 3, 5))
        return f"rgba({r},{g},{b},{round(alpha * d.opacity, 3)})" if alpha * d.opacity < 0.995 else colour

    out = {"kind": d.kind, "x": position[0], "y": position[1], "w": d.size[0], "h": d.size[1],
           "rotation": round(d.rotation % 360, 2), "flipH": d.flip[0], "flipV": d.flip[1], "outline": d.outline}
    if d.outline == "path" and d.path:
        out["path"] = [[cmd] + [round(v, 2) for point in points for v in point] for cmd, points in d.path]
        out["box"] = [max(0.01, d.path_size[0]), max(0.01, d.path_size[1])]
    if d.fill is not None and d.fill[0] == "solid":
        out["fill"] = {"color": css(d.fill[1], d.fill[2])}
    elif d.fill is not None:
        _, stops, angle, radial = d.fill
        out["fill"] = {"stops": [[position, css(colour, alpha)] for position, colour, alpha in stops],
                       "angle": round((360 - angle) % 360, 2), "radial": radial}
    if d.line is not None:
        out["line"] = {"color": css(d.line["color"], d.line["alpha"]), "width": round(d.line["width"] * PT_TO_PX, 2),
                       "dash": d.line["dash"]}
    if d.kind == "image":
        if d.image is not None and len(d.image) <= budget[0]:
            budget[0] -= len(d.image)
            kind = {b"\x89PNG": "png", b"GIF8": "gif"}.get(d.image[:4], "jpeg" if d.image[:2] == b"\xff\xd8" else None)
            if kind:
                out["src"] = f"data:image/{kind};base64," + base64.b64encode(d.image).decode("ascii")
        out["crop"] = list(d.crop) if d.crop else None
        out["name"] = d.name
        return out
    out["padding"] = [round(v * PT_TO_PX, 2) for v in d.padding]
    out["anchor"] = d.anchor
    out["paragraphs"] = [{"align": p["align"], "before": p["before"], "after": p["after"],
                          "lineHeight": round(line_height(p, p["end"]), 2) if p["spacing"][0] != "pct" else None,
                          "lineScale": p["spacing"][1] if p["spacing"][0] == "pct" else None,
                          "left": p["left"] + (p["bullet"]["indent"] if p["bullet"] else 0), "first": p["first"],
                          "bullet": (p["bullet"].get("char") or "1.") if p["bullet"] else None,
                          "size": p["end"]["size"],
                          "runs": [[text, {k: v for k, v in style.items() if v not in (None, False)}] for text, style in p["runs"]]}
                         for p in d.paragraphs]
    return out
