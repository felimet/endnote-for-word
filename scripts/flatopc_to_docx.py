#!/usr/bin/env python3
"""Convert a Word flat-OPC XML file (wdFormatFlatXML / Range.WordOpenXML) into a .docx package.

Why this exists: some Word installations never return from SaveAs2 when the document contains
EndNote CWYW fields (the save-time "format bibliography" step spins forever), but
`Range.WordOpenXML` serializes the same document in ~0.2 s without touching the save path.
This converter turns that flat XML into a normal OPC (zip) .docx, so the render pipeline can
still deliver a field-bearing document on such machines.

Usage:
    python scripts/flatopc_to_docx.py <input.xml> <output.docx> [--base <base.docx>]

Part XML is extracted *verbatim* from the flat package — no re-serialization — because Word
rejects packages whose parts lost their default-namespace form. The optional --base merges the
flat parts into an existing DOCX package: everything the flat XML does not carry (docProps,
customXml, ...) is taken from the base, and the base's content-type declarations are preserved.
"""

from __future__ import annotations

import base64
import re
import sys
import zipfile
from pathlib import Path

CONTENT_TYPES_NS = "http://schemas.openxmlformats.org/package/2006/content-types"
XML_DECLARATION = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n'

PART_PATTERN = re.compile(r"<pkg:part\b([^>]*)>(.*?)</pkg:part>", re.DOTALL)
ATTR_PATTERN = re.compile(r'pkg:(name|contentType)="([^"]*)"')
XML_DATA_PATTERN = re.compile(r"<pkg:xmlData\b[^>]*>(.*?)</pkg:xmlData>", re.DOTALL)
BINARY_DATA_PATTERN = re.compile(r"<pkg:binaryData\b[^>]*>(.*?)</pkg:binaryData>", re.DOTALL)
CONTENT_TYPE_DECL_PATTERN = re.compile(
    r'<(?:[A-Za-z0-9_.-]+:)?(Default|Override)\b([^>]*)/?>'
)
DECL_ATTR_PATTERN = re.compile(r'(Extension|PartName|ContentType)="([^"]*)"')


def xml_escape(value: str) -> str:
    return value.replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;").replace('"', "&quot;")


def parse_content_types(raw: bytes) -> tuple[list[tuple[str, str]], list[tuple[str, str]]]:
    text = raw.decode("utf-8-sig", "replace")
    defaults: list[tuple[str, str]] = []
    overrides: list[tuple[str, str]] = []
    for kind, attributes in CONTENT_TYPE_DECL_PATTERN.findall(text):
        fields = dict(DECL_ATTR_PATTERN.findall(attributes))
        if kind == "Default" and fields.get("Extension") and fields.get("ContentType"):
            defaults.append((fields["Extension"], fields["ContentType"]))
        elif kind == "Override" and fields.get("PartName") and fields.get("ContentType"):
            overrides.append((fields["PartName"], fields["ContentType"]))
    return defaults, overrides


def build_content_types(defaults: list[tuple[str, str]], overrides: list[tuple[str, str]]) -> str:
    seen_defaults: set[str] = set()
    parts = [XML_DECLARATION, f'<Types xmlns="{CONTENT_TYPES_NS}">']
    for extension, content_type in defaults:
        key = extension.lower()
        if key in seen_defaults:
            continue
        seen_defaults.add(key)
        parts.append(f'<Default Extension="{xml_escape(extension)}" ContentType="{xml_escape(content_type)}"/>')
    seen_overrides: set[str] = set()
    for part_name, content_type in overrides:
        if part_name in seen_overrides:
            continue
        seen_overrides.add(part_name)
        parts.append(f'<Override PartName="{xml_escape(part_name)}" ContentType="{xml_escape(content_type)}"/>')
    parts.append("</Types>")
    return "".join(parts)


def extract_parts(flat_xml: str) -> list[tuple[str, str, bytes]]:
    parts: list[tuple[str, str, bytes]] = []
    for attributes, body in PART_PATTERN.findall(flat_xml):
        fields = dict(ATTR_PATTERN.findall(attributes))
        name = fields.get("name", "")
        content_type = fields.get("contentType", "")
        if not name or not content_type:
            raise SystemExit("a pkg:part is missing pkg:name or pkg:contentType")
        xml_match = XML_DATA_PATTERN.search(body)
        if xml_match:
            inner = xml_match.group(1).strip()
            payload = (XML_DECLARATION + inner).encode("utf-8")
        else:
            binary_match = BINARY_DATA_PATTERN.search(body)
            if not binary_match or not binary_match.group(1).strip():
                raise SystemExit(f"part {name} has neither xmlData nor binaryData")
            payload = base64.b64decode(binary_match.group(1).strip())
        parts.append((name.lstrip("/"), content_type, payload))
    return parts


def main() -> int:
    arguments = sys.argv[1:]
    base_path: Path | None = None
    if "--base" in arguments:
        index = arguments.index("--base")
        if index + 1 >= len(arguments):
            print(__doc__)
            return 2
        base_path = Path(arguments[index + 1])
        del arguments[index:index + 2]
    if len(arguments) != 2:
        print(__doc__)
        return 2
    source = Path(arguments[0])
    destination = Path(arguments[1])
    if not source.exists():
        raise SystemExit(f"input not found: {source}")
    if base_path is not None and not base_path.exists():
        raise SystemExit(f"base package not found: {base_path}")

    flat_parts = extract_parts(source.read_text(encoding="utf-8"))
    if not flat_parts:
        raise SystemExit("no pkg:part elements found; is this really a flat OPC document?")

    merged: dict[str, bytes] = {}
    flat_content_types: dict[str, str] = {}
    defaults: list[tuple[str, str]] = []
    overrides: list[tuple[str, str]] = []
    if base_path is not None:
        with zipfile.ZipFile(base_path) as base_archive:
            for entry in base_archive.namelist():
                if entry == "[Content_Types].xml":
                    defaults, overrides = parse_content_types(base_archive.read(entry))
                    continue
                merged[entry] = base_archive.read(entry)
        print(f"base package parts: {len(merged)}")

    if not defaults:
        defaults = [
            ("rels", "application/vnd.openxmlformats-package.relationships+xml"),
            ("xml", "application/xml"),
        ]

    for name, content_type, payload in flat_parts:
        merged[name] = payload
        flat_content_types[name] = content_type

    declared = {part_name for part_name, _ in overrides}
    for name, content_type in sorted(flat_content_types.items()):
        if name.endswith(".rels"):
            continue
        if "/" + name in declared:
            continue
        overrides.append(("/" + name, content_type))

    content_types = build_content_types(defaults, overrides)

    if destination.exists():
        destination.unlink()
    with zipfile.ZipFile(destination, "w", zipfile.ZIP_DEFLATED) as archive:
        archive.writestr("[Content_Types].xml", content_types.encode("utf-8"))
        for name, payload in sorted(merged.items()):
            archive.writestr(name, payload)

    # verification: reopen the package and sanity-check the parts we care about
    with zipfile.ZipFile(destination) as archive:
        names = archive.namelist()
        missing = [want for want in ("word/document.xml", "word/settings.xml") if want not in names]
        if missing:
            raise SystemExit(f"converted package is missing: {missing}")
        document = archive.read("word/document.xml").decode("utf-8")
        settings = archive.read("word/settings.xml").decode("utf-8")
        citation_fields = document.count("ADDIN EN.CITE")
        has_layout = "EN.Layout" in settings
        has_libraries = "EN.Libraries" in settings
        print(f"wrote {destination} ({destination.stat().st_size} bytes)")
        print(f"parts: {len(names)}; citation field markers: {citation_fields}; "
              f"EN.Layout: {has_layout}; EN.Libraries: {has_libraries}")
        if citation_fields == 0:
            raise SystemExit("no ADDIN EN.CITE content in the converted document")
    return 0


if __name__ == "__main__":
    sys.exit(main())
