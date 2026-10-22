"""
Makes test/fixtures/images.numbers: test-comments.numbers with pictures and shapes added to its sheet,
as Numbers stores them, for the tests (no sample file from Numbers has pictures).

    .venv/Scripts/python tools/make_images_fixture.py

Added, besides the table at the top left:
  * "Sunrise.png" (120 x 80 pt) right of the table, with a 2 pt border, a description and a link;
  * "Cropped.png" (200 x 100 pt) masked by a rounded rectangle to its middle 100 x 50 pt;
  * "Scan.tiff", which Excel can't show, with the PNG preview Numbers keeps of it;
  * "Photo.heic", which Excel can't show and has no preview;
  * a group below the table holding a picture and a filled oval with the text "Grouped";
  * a rounded rectangle turned 30 degrees, with a fill and the text "Turned".

Only numbers-parser has read the result; it hasn't been opened in Numbers.
"""

import hashlib
import os
import struct
import sys
import warnings
import zlib

from numbers_parser import Document
from numbers_parser.constants import PACKAGE_ID
from numbers_parser.generated import TSDArchives_pb2 as TSD
from numbers_parser.generated import TSPArchiveMessages_pb2 as TSP
from numbers_parser.generated import TSWPArchives_pb2 as TSWP

HERE = os.path.dirname(os.path.abspath(__file__))
FIXTURES = os.path.join(HERE, "..", "test", "fixtures")


def png(width, height, colour):
    """A PNG of `colour(x, y) -> (r, g, b)`."""
    raw = b"".join(b"\x00" + bytes(v for x in range(width) for v in colour(x, y)) for y in range(height))

    def chunk(tag, data):
        return struct.pack(">I", len(data)) + tag + data + struct.pack(">I", zlib.crc32(tag + data) & 0xFFFFFFFF)
    return (b"\x89PNG\r\n\x1a\n" + chunk(b"IHDR", struct.pack(">IIBBBBB", width, height, 8, 2, 0, 0, 0))
            + chunk(b"IDAT", zlib.compress(raw, 9)) + chunk(b"IEND", b""))


def main():
    warnings.simplefilter("ignore")
    doc = Document(os.path.join(FIXTURES, "test-comments.numbers"))
    model = doc._model
    objects = model.objects
    sheet_id = doc.sheets[0]._sheet_id
    sheet = objects[sheet_id]
    styles = {}
    for oid, obj in objects._objects.items():
        base = obj
        while "super" in base.DESCRIPTOR.fields_by_name and "style_identifier" not in base.DESCRIPTOR.fields_by_name:
            base = base.super
        if "style_identifier" in base.DESCRIPTOR.fields_by_name and base.style_identifier:
            styles[base.style_identifier] = oid
    stylesheet = next(oid for oid, obj in objects._objects.items() if type(obj).__name__ == "StylesheetArchive")
    body = next(oid for oid, obj in objects._objects.items()
                if type(obj).__name__ == "ParagraphStyleArchive" and obj.super.name == "Body")
    no_list = next(oid for oid, obj in objects._objects.items()
                   if type(obj).__name__ == "ListStyleArchive" and obj.super.name == "None")
    package = objects[PACKAGE_ID]

    def add(cls, fields):
        return objects.create_object_from_dict("Index/Document", fields, cls)[0]

    def data(name, blob):
        identifier = max(d.identifier for d in package.datas) + 1
        stored = f"{os.path.splitext(name)[0].lower()}-{identifier}{os.path.splitext(name)[1]}"
        package.datas.append(TSP.DataInfo(identifier=identifier, digest=hashlib.sha1(blob).digest(),  # noqa: S324
                                          preferred_file_name=name, file_name=stored))
        objects.file_store[f"Data/{stored}"] = blob
        return identifier

    def geometry(x, y, w, h, angle=0.0):
        return {"position": {"x": x, "y": y}, "size": {"width": w, "height": h}, "flags": 3, "angle": angle}

    def image(parent, frame, blob_id, style="image-0-imageStyle", **more):
        return add(TSD.ImageArchive, {"super": {"geometry": geometry(*frame), "parent": {"identifier": parent}},
                                      "data": {"identifier": blob_id}, "style": {"identifier": styles[style]},
                                      "naturalSize": {"width": frame[2], "height": frame[3]}, **more})

    def shape(parent, frame, path, text, style="shape-1-shapestyle"):
        storage = add(TSWP.StorageArchive, {
            "kind": 3, "style_sheet": {"identifier": stylesheet}, "text": [text], "in_document": True,
            "table_para_style": {"entries": [{"character_index": 0, "object": {"identifier": body}}]},
            "table_list_style": {"entries": [{"character_index": 0, "object": {"identifier": no_list}}]},
            "table_para_data": {"entries": [{"character_index": 0, "first": 0, "second": 0}]},
            "table_char_style": {"entries": []}})
        return add(TSWP.ShapeInfoArchive, {
            "super": {"super": {"geometry": geometry(*frame), "parent": {"identifier": parent}},
                      "style": {"identifier": styles[style]}, "pathsource": path},
            "owned_storage": {"identifier": storage}, "deprecated_storage": {"identifier": storage}})

    sunrise = data("Sunrise.png", png(48, 32, lambda x, y: (255, 140 + y * 3, 40 + x * 4)))
    cropped = data("Cropped.png", png(40, 20, lambda x, y: (40 + x * 5, 90, 200 - y * 5)))
    scan_preview = data("Scan-preview.png", png(16, 16, lambda x, y: (120, 120, 120)))
    scan = data("Scan.tiff", b"II*\x00" + b"\x08\x00\x00\x00" + bytes(64))
    photo = data("Photo.heic", b"\x00\x00\x00\x18ftypheic\x00\x00\x00\x00mif1heic" + bytes(64))
    grouped = data("Grouped.png", png(20, 20, lambda x, y: (20, 160, 90)))

    added = [image(sheet_id, (320, 0, 120, 80), sunrise, "image-1-imageStyle")]
    objects[added[0]].super.accessibility_description = "A sunrise"
    objects[added[0]].super.hyperlink_url = "https://example.com/sunrise"
    masked = image(sheet_id, (320, 100, 200, 100), cropped)
    mask = add(TSD.MaskArchive, {
        "super": {"geometry": geometry(50, 25, 100, 50), "parent": {"identifier": masked}},
        "pathsource": {"scalar_path_source": {"type": 0, "scalar": 12, "naturalSize": {"width": 100, "height": 50}}}})
    objects[masked].mask.identifier = mask
    added.append(masked)
    added.append(image(sheet_id, (460, 0, 60, 60), scan, thumbnailData={"identifier": scan_preview}))
    added.append(image(sheet_id, (540, 0, 60, 60), photo))

    group = add(TSD.GroupArchive, {"super": {"geometry": geometry(0, 150, 220, 80), "parent": {"identifier": sheet_id}}})
    oval = {"bezier_path_source": {"naturalSize": {"width": 140, "height": 80}, "path": {"elements": [
        {"type": 1, "points": [{"x": 70, "y": 0}]},
        {"type": 4, "points": [{"x": 108.7, "y": 0}, {"x": 140, "y": 17.9}, {"x": 140, "y": 40}]},
        {"type": 4, "points": [{"x": 140, "y": 62.1}, {"x": 108.7, "y": 80}, {"x": 70, "y": 80}]},
        {"type": 4, "points": [{"x": 31.3, "y": 80}, {"x": 0, "y": 62.1}, {"x": 0, "y": 40}]},
        {"type": 4, "points": [{"x": 0, "y": 17.9}, {"x": 31.3, "y": 0}, {"x": 70, "y": 0}]},
        {"type": 5}]}}}
    for child in (image(group, (0, 0, 80, 80), grouped), shape(group, (80, 0, 140, 80), oval, "Grouped")):
        objects[group].children.add(identifier=child)
    added.append(group)

    rounded = {"scalar_path_source": {"type": 0, "scalar": 10, "naturalSize": {"width": 120, "height": 50}}}
    added.append(shape(sheet_id, (320, 230, 120, 50, 30.0), rounded, "Turned"))

    for oid in added:
        sheet.drawable_infos.add(identifier=oid)
    out = os.path.join(FIXTURES, "images.numbers")
    doc.save(out)
    print(out)


if __name__ == "__main__":
    sys.exit(main())
