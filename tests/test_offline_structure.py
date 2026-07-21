"""Offline OOXML readers for EndNote CWYW output, plus their unit tests.

Nothing in this module needs Word, EndNote or COM. Everything is implemented with
``zipfile`` and ``xml.etree`` so the same helpers can be reused by the live
integration test to cross-check what the COM bridge reports.

The counting rule implemented by :func:`count_citation_fields` is the one measured
on the reference machine: a merged citation is a single ``ADDIN EN.CITE`` field, and
EndNote moves large traveling-library payloads into a nested ``ADDIN EN.CITE.DATA``
field, so the nested fields must be subtracted.
"""

from __future__ import annotations

import logging
import os
import unittest
import zipfile
from pathlib import Path
from xml.etree import ElementTree as ET
from xml.sax.saxutils import escape, quoteattr

LOGGER = logging.getLogger(__name__)

ROOT = Path(__file__).resolve().parents[1]
FIXTURES = Path(__file__).resolve().parent / "fixtures"

W_NAMESPACE = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
PACKAGE_RELATIONSHIP_NAMESPACE = "http://schemas.openxmlformats.org/package/2006/relationships"
OFFICE_RELATIONSHIP_NAMESPACE = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
CONTENT_TYPES_NAMESPACE = "http://schemas.openxmlformats.org/package/2006/content-types"

DOCUMENT_PART = "word/document.xml"
SETTINGS_PART = "word/settings.xml"

CITATION_FIELD_PREFIX = "ADDIN EN.CITE"
CITATION_DATA_FIELD_PREFIX = "ADDIN EN.CITE.DATA"
BIBLIOGRAPHY_FIELD_PREFIX = "ADDIN EN.REFLIST"
HYPERLINK_FIELD_PREFIX = "HYPERLINK"

LAYOUT_VARIABLE_NAME = "EN.Layout"
LIBRARIES_VARIABLE_NAME = "EN.Libraries"

FIXTURE_DOCX_ENVIRONMENT_VARIABLE = "ENDNOTE_FIXTURE_DOCX"


def _tag(local_name: str) -> str:
    return "{" + W_NAMESPACE + "}" + local_name


def _read_part(docx_path: Path | str, part_name: str, required: bool = True) -> str | None:
    """Return the decoded text of one package part, or ``None`` when optional and absent."""
    path = Path(docx_path)
    try:
        with zipfile.ZipFile(path) as package:
            if part_name not in package.namelist():
                if required:
                    raise ValueError(f"{path} is not a Word package: missing part '{part_name}'.")
                LOGGER.debug("Optional part '%s' is absent from %s.", part_name, path)
                return None
            return package.read(part_name).decode("utf-8")
    except zipfile.BadZipFile as error:
        raise ValueError(f"{path} is not a readable OOXML (zip) package.") from error
    except OSError as error:
        raise ValueError(f"{path} could not be opened: {error}") from error


def read_field_codes(docx_path: Path | str) -> list[str]:
    """Return every Word field code in ``word/document.xml``, in document order.

    Complex fields are reconstructed from their ``w:instrText`` runs, which Word is
    free to split across several runs. Nesting is tracked with an explicit stack so
    a nested ``ADDIN EN.CITE.DATA`` field is reported as its own code rather than
    being folded into the parent ``ADDIN EN.CITE`` code.
    """
    xml_text = _read_part(docx_path, DOCUMENT_PART, required=True)
    try:
        root = ET.fromstring(xml_text)
    except ET.ParseError as error:
        raise ValueError(f"{docx_path}: {DOCUMENT_PART} is not well-formed XML.") from error

    codes: list[str] = []
    open_fields: list[list[str]] = []
    field_char_tag = _tag("fldChar")
    instr_text_tag = _tag("instrText")
    field_simple_tag = _tag("fldSimple")
    field_char_type_attribute = _tag("fldCharType")
    instruction_attribute = _tag("instr")

    for element in root.iter():
        if element.tag == field_char_tag:
            field_char_type = element.get(field_char_type_attribute)
            if field_char_type == "begin":
                open_fields.append([])
            elif field_char_type == "end":
                if open_fields:
                    codes.append("".join(open_fields.pop()).strip())
                else:
                    LOGGER.warning("Field end without a matching begin in %s.", docx_path)
        elif element.tag == instr_text_tag:
            if open_fields:
                open_fields[-1].append(element.text or "")
            else:
                LOGGER.warning("Field instruction outside a field in %s.", docx_path)
        elif element.tag == field_simple_tag:
            instruction = element.get(instruction_attribute)
            if instruction is not None:
                codes.append(instruction.strip())

    for unterminated in open_fields:
        LOGGER.warning("Unterminated field in %s; counting it anyway.", docx_path)
        codes.append("".join(unterminated).strip())
    return codes


def count_citation_fields(docx_path: Path | str) -> int:
    """Return the number of real EndNote citations in a DOCX.

    ``citation_count = count("ADDIN EN.CITE") - count("ADDIN EN.CITE.DATA")``
    """
    codes = read_field_codes(docx_path)
    citations = sum(1 for code in codes if code.startswith(CITATION_FIELD_PREFIX))
    payloads = sum(1 for code in codes if code.startswith(CITATION_DATA_FIELD_PREFIX))
    return citations - payloads


def count_citation_data_fields(docx_path: Path | str) -> int:
    """Return the number of nested traveling-library payload fields."""
    return sum(
        1
        for code in read_field_codes(docx_path)
        if code.startswith(CITATION_DATA_FIELD_PREFIX)
    )


def count_bibliography_fields(docx_path: Path | str) -> int:
    """Return the number of ``ADDIN EN.REFLIST`` bibliography fields."""
    return sum(
        1
        for code in read_field_codes(docx_path)
        if code.startswith(BIBLIOGRAPHY_FIELD_PREFIX)
    )


def read_document_variables(docx_path: Path | str) -> dict[str, str]:
    """Return the ``w:docVar`` name/value pairs from ``word/settings.xml``."""
    xml_text = _read_part(docx_path, SETTINGS_PART, required=False)
    if xml_text is None:
        return {}
    try:
        root = ET.fromstring(xml_text)
    except ET.ParseError as error:
        raise ValueError(f"{docx_path}: {SETTINGS_PART} is not well-formed XML.") from error

    variables: dict[str, str] = {}
    name_attribute = _tag("name")
    value_attribute = _tag("val")
    for element in root.iter(_tag("docVar")):
        name = element.get(name_attribute)
        if name is None:
            LOGGER.warning("Ignoring an unnamed docVar in %s.", docx_path)
            continue
        variables[name] = element.get(value_attribute, "")
    return variables


def read_layout_style(docx_path: Path | str) -> str | None:
    """Return the EndNote output style recorded in the ``EN.Layout`` document variable.

    Returns ``None`` when the document has never been formatted by EndNote.
    """
    value = read_document_variables(docx_path).get(LAYOUT_VARIABLE_NAME)
    if value is None:
        return None
    try:
        layout = ET.fromstring(value)
    except ET.ParseError as error:
        raise ValueError(
            f"{docx_path}: the {LAYOUT_VARIABLE_NAME} document variable is not well-formed XML."
        ) from error
    style = layout.findtext("Style")
    if style is None:
        raise ValueError(
            f"{docx_path}: the {LAYOUT_VARIABLE_NAME} document variable has no <Style> element."
        )
    return style.strip()


def read_libraries(docx_path: Path | str) -> list[dict[str, object]]:
    """Return the libraries recorded in the ``EN.Libraries`` document variable."""
    value = read_document_variables(docx_path).get(LIBRARIES_VARIABLE_NAME)
    if value is None:
        return []
    try:
        libraries = ET.fromstring(value)
    except ET.ParseError as error:
        raise ValueError(
            f"{docx_path}: the {LIBRARIES_VARIABLE_NAME} document variable is not well-formed XML."
        ) from error

    entries: list[dict[str, object]] = []
    for item in libraries.findall("item"):
        record_ids: list[int] = []
        for record in item.findall("record-ids/item"):
            raw = (record.text or "").strip()
            if not raw:
                continue
            try:
                record_ids.append(int(raw))
            except ValueError as error:
                raise ValueError(
                    f"{docx_path}: non-numeric record id '{raw}' in {LIBRARIES_VARIABLE_NAME}."
                ) from error
        entries.append(
            {
                "db_id": item.get("db-id", ""),
                "name": (item.text or "").strip(),
                "record_ids": record_ids,
            }
        )
    return entries


def read_library_record_ids(docx_path: Path | str) -> list[int]:
    """Return every EndNote record id recorded in ``EN.Libraries``, sorted and unique."""
    record_ids: set[int] = set()
    for library in read_libraries(docx_path):
        record_ids.update(library["record_ids"])  # type: ignore[arg-type]
    return sorted(record_ids)


# --------------------------------------------------------------------------------------
# Minimal OOXML writer used to build synthetic fixtures for these tests and the live test.
# --------------------------------------------------------------------------------------

_CONTENT_TYPES_XML = (
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
    f'<Types xmlns="{CONTENT_TYPES_NAMESPACE}">'
    '<Default Extension="rels" '
    'ContentType="application/vnd.openxmlformats-package.relationships+xml"/>'
    '<Default Extension="xml" ContentType="application/xml"/>'
    '<Override PartName="/word/document.xml" '
    'ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>'
    '<Override PartName="/word/settings.xml" '
    'ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.settings+xml"/>'
    "</Types>"
)

_PACKAGE_RELATIONSHIPS_XML = (
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
    f'<Relationships xmlns="{PACKAGE_RELATIONSHIP_NAMESPACE}">'
    f'<Relationship Id="rId1" Type="{OFFICE_RELATIONSHIP_NAMESPACE}/officeDocument" '
    'Target="word/document.xml"/>'
    "</Relationships>"
)

_DOCUMENT_RELATIONSHIPS_XML = (
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
    f'<Relationships xmlns="{PACKAGE_RELATIONSHIP_NAMESPACE}">'
    f'<Relationship Id="rId1" Type="{OFFICE_RELATIONSHIP_NAMESPACE}/settings" '
    'Target="settings.xml"/>'
    "</Relationships>"
)


def paragraph_xml(*runs: str) -> str:
    """Wrap raw run XML in a ``w:p`` element."""
    return "<w:p>" + "".join(runs) + "</w:p>"


def text_run_xml(text: str) -> str:
    """Return a ``w:r`` run carrying literal text."""
    return f'<w:r><w:t xml:space="preserve">{escape(text)}</w:t></w:r>'


def field_xml(code: str, result: str = "", nested: str = "", split_code: bool = False) -> str:
    """Return the run sequence for one complex Word field.

    ``nested`` is raw run XML inserted between the field code and the separator, which
    is how EndNote stores an ``ADDIN EN.CITE.DATA`` payload inside an ``ADDIN EN.CITE``
    field. ``split_code`` reproduces Word splitting one code across several
    ``w:instrText`` runs.
    """
    if split_code and len(code) >= 3:
        third = max(1, len(code) // 3)
        chunks = [code[:third], code[third : third * 2], code[third * 2 :]]
    else:
        chunks = [code]
    instructions = "".join(
        f'<w:r><w:instrText xml:space="preserve">{escape(chunk)}</w:instrText></w:r>'
        for chunk in chunks
    )
    return (
        '<w:r><w:fldChar w:fldCharType="begin"/></w:r>'
        + instructions
        + nested
        + '<w:r><w:fldChar w:fldCharType="separate"/></w:r>'
        + text_run_xml(result)
        + '<w:r><w:fldChar w:fldCharType="end"/></w:r>'
    )


def write_minimal_docx(
    path: Path | str,
    body_xml: str,
    document_variables: dict[str, str] | None = None,
) -> Path:
    """Write a small but structurally valid DOCX package and return its path."""
    target = Path(path)
    document_xml = (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        f'<w:document xmlns:w="{W_NAMESPACE}"><w:body>'
        + body_xml
        + '<w:sectPr><w:pgSz w:w="11906" w:h="16838"/></w:sectPr>'
        "</w:body></w:document>"
    )
    variables = document_variables or {}
    if variables:
        doc_vars = (
            "<w:docVars>"
            + "".join(
                f"<w:docVar w:name={quoteattr(name)} w:val={quoteattr(value)}/>"
                for name, value in variables.items()
            )
            + "</w:docVars>"
        )
    else:
        doc_vars = ""
    settings_xml = (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        f'<w:settings xmlns:w="{W_NAMESPACE}">' + doc_vars + "</w:settings>"
    )

    try:
        target.parent.mkdir(parents=True, exist_ok=True)
        with zipfile.ZipFile(target, "w", zipfile.ZIP_DEFLATED) as package:
            package.writestr("[Content_Types].xml", _CONTENT_TYPES_XML)
            package.writestr("_rels/.rels", _PACKAGE_RELATIONSHIPS_XML)
            package.writestr("word/_rels/document.xml.rels", _DOCUMENT_RELATIONSHIPS_XML)
            package.writestr(DOCUMENT_PART, document_xml)
            package.writestr(SETTINGS_PART, settings_xml)
    except OSError as error:
        raise ValueError(f"Could not write the synthetic DOCX at {target}: {error}") from error
    return target


_SAMPLE_LAYOUT = (
    "<ENLayout><Style>IEEE</Style><LeftDelim>{</LeftDelim><RightDelim>}</RightDelim>"
    "<FontName>Times New Roman</FontName><FontSize>12</FontSize><ReflistTitle></ReflistTitle>"
    "<StartingRefnum>1</StartingRefnum><FirstLineIndent>0</FirstLineIndent>"
    "<HangingIndent>720</HangingIndent><LineSpacing>0</LineSpacing><SpaceAfter>0</SpaceAfter>"
    "<HyperlinksEnabled>0</HyperlinksEnabled><HyperlinksVisible>0</HyperlinksVisible>"
    "<EnableBibliographyCategories>0</EnableBibliographyCategories></ENLayout>"
)

_SAMPLE_LIBRARIES = (
    '<Libraries><item db-id="fixturedbid">FixtureLibrary<record-ids>'
    "<item>1</item><item>2</item><item>3</item>"
    "</record-ids></item></Libraries>"
)


def build_sample_document(path: Path | str) -> Path:
    """Build the synthetic document the offline readers are unit-tested against.

    It contains one plain citation, one citation whose payload lives in a nested
    ``ADDIN EN.CITE.DATA`` field, one bibliography field, one hyperlink inside the
    bibliography, and a field code deliberately split across three runs.
    """
    plain_citation = field_xml(
        " ADDIN EN.CITE <EndNote><Cite><Author>Thomsen</Author><Year>2023</Year>"
        "<RecNum>1</RecNum><DisplayText>[1]</DisplayText><record>"
        "<rec-number>1</rec-number></record></Cite></EndNote> ",
        result="[1]",
    )
    nested_payload = field_xml(
        " ADDIN EN.CITE.DATA <EndNote><Cite><record><rec-number>2</rec-number></record>"
        "<record><rec-number>3</rec-number></record></Cite></EndNote> ",
        result="",
    )
    merged_citation = field_xml(
        " ADDIN EN.CITE <EndNote><Cite><Author>Magrin</Author><Year>2023</Year>"
        "<RecNum>2</RecNum></Cite><Cite><Author>Bernhard</Author><Year>2020</Year>"
        "<RecNum>3</RecNum></Cite></EndNote> ",
        result="[2, 3]",
        nested=nested_payload,
        split_code=True,
    )
    bibliography = field_xml(" ADDIN EN.REFLIST ", result="Thomsen, P. T. (2023).")
    hyperlink = field_xml(' HYPERLINK "https://doi.org/10.0000/fixture.0001" ', result="doi")

    body = (
        paragraph_xml(text_run_xml("First sentence "), plain_citation, text_run_xml(" ends."))
        + paragraph_xml(text_run_xml("Second sentence "), merged_citation, text_run_xml(" ends."))
        + paragraph_xml(bibliography, hyperlink)
    )
    return write_minimal_docx(
        path,
        body,
        {LAYOUT_VARIABLE_NAME: _SAMPLE_LAYOUT, LIBRARIES_VARIABLE_NAME: _SAMPLE_LIBRARIES},
    )


def _locate_rendered_fixture() -> Path | None:
    """Return a rendered DOCX to cross-check, or ``None`` when none is available."""
    configured = os.environ.get(FIXTURE_DOCX_ENVIRONMENT_VARIABLE)
    if configured:
        candidate = Path(configured)
        if candidate.is_file():
            return candidate
        LOGGER.warning(
            "%s points at '%s', which is not a file; ignoring it.",
            FIXTURE_DOCX_ENVIRONMENT_VARIABLE,
            configured,
        )
    if FIXTURES.is_dir():
        for candidate in sorted(FIXTURES.glob("*.docx")):
            return candidate
    return None


class OfflineStructureTests(unittest.TestCase):
    """Unit tests for the offline readers against a document this test builds itself."""

    @classmethod
    def setUpClass(cls) -> None:
        import tempfile

        cls._directory = tempfile.TemporaryDirectory(prefix="endnote-offline-")
        cls.document = build_sample_document(Path(cls._directory.name) / "sample.docx")

    @classmethod
    def tearDownClass(cls) -> None:
        cls._directory.cleanup()

    def test_sample_document_is_a_real_package(self) -> None:
        self.assertTrue(zipfile.is_zipfile(self.document))
        with zipfile.ZipFile(self.document) as package:
            names = set(package.namelist())
        self.assertIn(DOCUMENT_PART, names)
        self.assertIn(SETTINGS_PART, names)
        self.assertIn("[Content_Types].xml", names)

    def test_field_codes_are_reconstructed_in_document_order(self) -> None:
        codes = read_field_codes(self.document)
        self.assertEqual(len(codes), 5)
        self.assertTrue(codes[0].startswith(CITATION_FIELD_PREFIX))
        self.assertTrue(codes[1].startswith(CITATION_DATA_FIELD_PREFIX))
        self.assertTrue(codes[2].startswith(CITATION_FIELD_PREFIX))
        self.assertTrue(codes[3].startswith(BIBLIOGRAPHY_FIELD_PREFIX))
        self.assertTrue(codes[4].startswith(HYPERLINK_FIELD_PREFIX))

    def test_split_field_code_is_joined_before_matching(self) -> None:
        codes = read_field_codes(self.document)
        merged = codes[2]
        self.assertIn("<Author>Magrin</Author>", merged)
        self.assertIn("<Author>Bernhard</Author>", merged)
        self.assertNotIn("ADDIN EN.CITE.DATA", merged)

    def test_count_citation_fields_subtracts_nested_data_fields(self) -> None:
        self.assertEqual(count_citation_fields(self.document), 2)
        self.assertEqual(count_citation_data_fields(self.document), 1)

    def test_count_bibliography_fields_ignores_hyperlinks(self) -> None:
        self.assertEqual(count_bibliography_fields(self.document), 1)

    def test_field_simple_elements_are_counted(self) -> None:
        directory = Path(self.__class__._directory.name)
        simple = directory / "simple-field.docx"
        body = paragraph_xml(
            '<w:fldSimple w:instr=" ADDIN EN.CITE simple ">'
            '<w:r><w:t>(A, 2020)</w:t></w:r></w:fldSimple>'
        )
        write_minimal_docx(simple, body)
        self.assertEqual(count_citation_fields(simple), 1)

    def test_read_layout_style(self) -> None:
        self.assertEqual(read_layout_style(self.document), "IEEE")

    def test_read_layout_style_returns_none_without_the_variable(self) -> None:
        directory = Path(self.__class__._directory.name)
        bare = write_minimal_docx(directory / "bare.docx", paragraph_xml(text_run_xml("No EndNote.")))
        self.assertIsNone(read_layout_style(bare))
        self.assertEqual(read_library_record_ids(bare), [])

    def test_read_layout_style_rejects_malformed_variable(self) -> None:
        directory = Path(self.__class__._directory.name)
        broken = write_minimal_docx(
            directory / "broken-layout.docx",
            paragraph_xml(text_run_xml("Broken.")),
            {LAYOUT_VARIABLE_NAME: "<ENLayout><Style>IEEE</ENLayout>"},
        )
        with self.assertRaises(ValueError):
            read_layout_style(broken)

    def test_read_library_record_ids(self) -> None:
        self.assertEqual(read_library_record_ids(self.document), [1, 2, 3])
        libraries = read_libraries(self.document)
        self.assertEqual(len(libraries), 1)
        self.assertEqual(libraries[0]["db_id"], "fixturedbid")
        self.assertEqual(libraries[0]["name"], "FixtureLibrary")

    def test_non_package_input_raises_a_clear_error(self) -> None:
        directory = Path(self.__class__._directory.name)
        not_a_package = directory / "plain.txt"
        not_a_package.write_text("not a zip archive", encoding="utf-8")
        with self.assertRaises(ValueError):
            read_field_codes(not_a_package)


class RenderedFixtureTests(unittest.TestCase):
    """Cross-checks against a real rendered DOCX when one is available."""

    def setUp(self) -> None:
        self.fixture = _locate_rendered_fixture()
        if self.fixture is None:
            self.skipTest(
                "No rendered DOCX fixture found; set "
                f"{FIXTURE_DOCX_ENVIRONMENT_VARIABLE} or place one in tests/fixtures."
            )

    def test_rendered_fixture_carries_endnote_structure(self) -> None:
        assert self.fixture is not None
        self.assertGreaterEqual(count_citation_fields(self.fixture), 1)
        self.assertEqual(count_bibliography_fields(self.fixture), 1)
        style = read_layout_style(self.fixture)
        self.assertIsNotNone(style)
        self.assertTrue(style)
        self.assertGreaterEqual(len(read_library_record_ids(self.fixture)), 1)


if __name__ == "__main__":
    unittest.main()
