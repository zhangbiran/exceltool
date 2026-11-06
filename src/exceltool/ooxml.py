import os
import re
import shutil
import tempfile
import zipfile
from pathlib import Path


_SHEET_VIEW_TAG = re.compile(
    rb"<(?:[A-Za-z_][A-Za-z0-9_.-]*:)?sheetView\b[^>]*>"
)
_DEFAULT_GRID_COLOR = re.compile(
    rb"\s+defaultGridColor\s*=\s*([\"'])(?:true|1)\1"
)
_DEFAULT_COLOR_ID = re.compile(
    rb"\s+colorId\s*=\s*([\"'])64\1"
)


def _normalize_sheet_view_xml(data):
    normalized_count = 0

    def normalize_tag(match):
        nonlocal normalized_count
        tag = match.group(0)
        default_grid_color = _DEFAULT_GRID_COLOR.search(tag)
        color_id = _DEFAULT_COLOR_ID.search(tag)
        if default_grid_color is None and color_id is None:
            return tag

        ranges = [
            attribute.span()
            for attribute in (default_grid_color, color_id)
            if attribute is not None
        ]
        ranges.sort(reverse=True)
        for begin, end in ranges:
            tag = tag[:begin] + tag[end:]
        normalized_count += 1
        return tag

    return _SHEET_VIEW_TAG.sub(normalize_tag, data), normalized_count


def _count_redundant_default_grid_attributes(data):
    count = 0
    for match in _SHEET_VIEW_TAG.finditer(data):
        tag = match.group(0)
        if _DEFAULT_GRID_COLOR.search(tag) or _DEFAULT_COLOR_ID.search(tag):
            count += 1
    return count


def normalize_xlsx_default_grid_color(path):
    """Remove redundant default grid-color attributes written by LibreOffice."""
    path = Path(path)
    if path.suffix.lower() != ".xlsx":
        return 0

    replacement_names = set()
    normalized_count = 0
    with zipfile.ZipFile(str(path), "r") as source:
        for entry in source.infolist():
            if not (
                entry.filename.startswith("xl/worksheets/")
                and entry.filename.endswith(".xml")
            ):
                continue
            data = source.read(entry)
            count = _count_redundant_default_grid_attributes(data)
            if count:
                replacement_names.add(entry.filename)
                normalized_count += count

    if not replacement_names:
        return 0

    handle, temporary = tempfile.mkstemp(
        prefix=".%s." % path.name,
        suffix=".xlsx",
        dir=str(path.parent),
    )
    os.close(handle)
    try:
        with zipfile.ZipFile(str(path), "r") as source:
            with zipfile.ZipFile(temporary, "w") as target:
                target.comment = source.comment
                for entry in source.infolist():
                    data = source.read(entry)
                    if entry.filename in replacement_names:
                        data, _ = _normalize_sheet_view_xml(data)
                    target.writestr(entry, data)
        shutil.copystat(str(path), temporary)
        os.replace(temporary, str(path))
    finally:
        if os.path.exists(temporary):
            os.unlink(temporary)

    return normalized_count
