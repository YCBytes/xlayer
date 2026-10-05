"""Test-only raw OOXML inputs, never production-derived expected answers."""

from __future__ import annotations

import zipfile
from collections.abc import Mapping
from pathlib import Path
from xml.sax.saxutils import quoteattr


def write_workbook_case(
    path: Path,
    sheet_bodies: Mapping[str, str],
    *,
    names_xml: str = "",
    zip_order: tuple[str, ...] | None = None,
) -> Path:
    main = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
    rel = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/"
    relns = "http://schemas.openxmlformats.org/package/2006/relationships"
    names = f"<definedNames>{names_xml}</definedNames>" if names_xml else ""
    tabs = "".join(
        f'<sheet name={quoteattr(name)} sheetId="{i}" r:id="r{i}"/>'
        for i, name in enumerate(sheet_bodies, 1)
    )
    relationships = "".join(
        f'<Relationship Id="r{i}" Type="{rel}worksheet" Target="worksheets/{i}.xml"/>'
        for i in range(1, len(sheet_bodies) + 1)
    )
    parts = {
        "[Content_Types].xml": '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
        '<Default Extension="xml" ContentType="application/xml"/>'
        '<Default Extension="rels" '
        'ContentType="application/vnd.openxmlformats-package.relationships+xml"/>'
        '<Override PartName="/xl/workbook.xml" '
        'ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>'
        "</Types>",
        "_rels/.rels": f'<Relationships xmlns="{relns}">'
        f'<Relationship Id="r" Type="{rel}officeDocument" Target="xl/workbook.xml"/>'
        "</Relationships>",
        "xl/workbook.xml": f'<workbook xmlns="{main}" xmlns:r="{rel[:-1]}">'
        f"<sheets>{tabs}</sheets>{names}</workbook>",
        "xl/_rels/workbook.xml.rels": f'<Relationships xmlns="{relns}">'
        f"{relationships}</Relationships>",
    }
    for i, body in enumerate(sheet_bodies.values(), 1):
        parts[f"xl/worksheets/{i}.xml"] = f'<worksheet xmlns="{main}">{body}</worksheet>'
    with zipfile.ZipFile(path, "w") as archive:
        for name in zip_order or tuple(parts):
            archive.writestr(name, parts[name].encode("utf-8"))
    return path
