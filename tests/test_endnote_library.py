"""Offline tests for scripts/endnote_library.py.

The tests build their own synthetic EndNote library in a temporary directory: an
``.enl`` stub whose first four bytes are ``Eds1`` and a ``<stem>.Data/sdb/sdb.eni``
SQLite database whose ``refs`` table carries the column list recorded in the build
spec, including the custom ``ENCI_Base`` and ``ENCIN_zh_TW`` collations. No real
library is ever touched and EndNote does not need to be installed.
"""

from __future__ import annotations

import json
import sqlite3
import subprocess
import sys
import tempfile
import unittest
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
LIBRARY_SCRIPT = ROOT / "scripts" / "endnote_library.py"
FIXTURES = Path(__file__).resolve().parent / "fixtures"
FIXTURE_RIS = FIXTURES / "fixture-references.ris"

CLI_TIMEOUT_SECONDS = 120
ABSTRACT_HEAD_LIMIT = 300
ENL_MAGIC = b"Eds1"

# Column list taken from the build spec. Text columns that EndNote declares with its
# private collations keep them here so a reader that forgets to register the no-op
# collations fails the same way it would fail against a real library.
REFS_COLUMNS = """
    id INTEGER PRIMARY KEY,
    trash_state INTEGER NOT NULL DEFAULT 0,
    reference_type INTEGER,
    author TEXT COLLATE ENCI_Base,
    year TEXT,
    title TEXT COLLATE ENCI_Base,
    pages TEXT,
    secondary_title TEXT COLLATE ENCI_Base,
    volume TEXT,
    number TEXT,
    publisher TEXT,
    place_published TEXT,
    abstract TEXT,
    keywords TEXT,
    url TEXT,
    isbn TEXT,
    electronic_resource_number TEXT,
    short_title TEXT,
    label TEXT,
    notes TEXT,
    edition TEXT,
    date TEXT,
    language TEXT COLLATE ENCIN_zh_TW,
    accession_number TEXT,
    custom_1 TEXT,
    custom_2 TEXT,
    custom_3 TEXT,
    custom_4 TEXT,
    custom_5 TEXT,
    custom_6 TEXT,
    custom_7 TEXT,
    added_to_library TEXT,
    record_last_updated TEXT
"""

LONG_ABSTRACT = (
    "Synthetic abstract for the first fixture record. "
    + ("This sentence pads the abstract so the reader has to truncate it. " * 12)
)

FIXTURE_RECORDS = [
    {
        "id": 1,
        "trash_state": 0,
        "reference_type": 17,
        "author": "Thomsen, Peter T.\rShearer, Jane K.",
        "year": "2023",
        "title": "Synthetic fixture record one for automated citation tests",
        "pages": "45-46",
        "secondary_title": "Journal of Test Fixtures",
        "volume": "12",
        "number": "3",
        "abstract": LONG_ABSTRACT,
        "keywords": "fixture\rautomation",
        "url": "https://doi.org/10.0000/fixture.0001",
        "electronic_resource_number": "10.0000/fixture.0001",
        "publisher": "Fixture Press",
    },
    {
        "id": 2,
        "trash_state": 0,
        "reference_type": 17,
        "author": "Luca Magrin",
        "year": "2023",
        "title": "Synthetic fixture record two for merged citation tests",
        "pages": "1-14",
        "secondary_title": "Journal of Test Fixtures",
        "volume": "11",
        "number": "1",
        "abstract": "Synthetic abstract for the merged citation fixture record.",
        "keywords": "fixture\rmerged citation",
        "url": "https://doi.org/10.0000/fixture.0002",
        "electronic_resource_number": "10.0000/fixture.0002",
        "publisher": "Fixture Press",
    },
    {
        "id": 3,
        "trash_state": 0,
        "reference_type": 17,
        "author": "Bernhard, Anna\rMüller, Karin",
        "year": "2020",
        "title": "Synthetic fixture record three for style switching tests",
        "pages": "201-218",
        "secondary_title": "Journal of Test Fixtures",
        "volume": "8",
        "number": "4",
        "abstract": "Synthetic abstract for the style switching fixture record.",
        "keywords": "fixture\routput style",
        "url": "https://doi.org/10.0000/fixture.0003",
        "electronic_resource_number": "10.0000/fixture.0003",
        "publisher": "Fixture Press",
    },
    {
        "id": 4,
        "trash_state": 0,
        "reference_type": 17,
        "author": "Müller, Karin",
        "year": "2019",
        "title": "Synthetic fixture record four with a diacritic in the first author",
        "pages": "77-90",
        "secondary_title": "Zeitschrift fur Testdaten",
        "volume": "3",
        "number": "2",
        "abstract": "Synthetic abstract used to test accent tolerant author matching.",
        "keywords": "fixture\rdiacritics",
        "url": "https://doi.org/10.0000/fixture.0004",
        "electronic_resource_number": "10.0000/fixture.0004",
        "publisher": "Fixture Press",
    },
    {
        "id": 5,
        "trash_state": 1,
        "reference_type": 17,
        "author": "Deleted, Record",
        "year": "2018",
        "title": "Synthetic fixture record five that lives in the trash",
        "pages": "5-6",
        "secondary_title": "Journal of Test Fixtures",
        "volume": "1",
        "number": "1",
        "abstract": "Synthetic abstract for a trashed record that must never be returned.",
        "keywords": "fixture\rtrash",
        "url": "https://doi.org/10.0000/fixture.0005",
        "electronic_resource_number": "10.0000/fixture.0005",
        "publisher": "Fixture Press",
    },
]

XML_EXPORT = """<?xml version="1.0" encoding="UTF-8"?>
<xml><records>
  <record>
    <rec-number>1</rec-number>
    <ref-type name="Journal Article">17</ref-type>
    <contributors><authors>
      <author>Thomsen, Peter T.</author>
      <author>Shearer, Jane K.</author>
    </authors></contributors>
    <titles>
      <title>Synthetic fixture record one for automated citation tests</title>
      <secondary-title>Journal of Test Fixtures</secondary-title>
    </titles>
    <dates><year>2023</year></dates>
    <electronic-resource-num>10.0000/fixture.0001</electronic-resource-num>
  </record>
</records></xml>
"""


def _no_op_collation(left: str, right: str) -> int:
    """Ordering stand-in for EndNote's private collations."""
    return (left > right) - (left < right)


def build_synthetic_library(directory: Path, name: str = "FixtureLibrary") -> Path:
    """Create ``<name>.enl`` plus ``<name>.Data/sdb/sdb.eni`` and return the .enl path."""
    enl_path = directory / f"{name}.enl"
    data_directory = directory / f"{name}.Data" / "sdb"
    data_directory.mkdir(parents=True, exist_ok=True)
    enl_path.write_bytes(ENL_MAGIC + b"\x00" * 28)

    database_path = data_directory / "sdb.eni"
    connection = sqlite3.connect(database_path)
    try:
        connection.create_collation("ENCI_Base", _no_op_collation)
        connection.create_collation("ENCIN_zh_TW", _no_op_collation)
        connection.execute(f"CREATE TABLE refs ({REFS_COLUMNS})")
        columns = [row[1] for row in connection.execute("PRAGMA table_info(refs)")]
        for record in FIXTURE_RECORDS:
            values = [record.get(column) for column in columns]
            placeholders = ", ".join("?" for _ in columns)
            connection.execute(
                f"INSERT INTO refs ({', '.join(columns)}) VALUES ({placeholders})", values
            )
        connection.commit()
    except sqlite3.Error as error:
        raise RuntimeError(
            f"Could not build the synthetic library at {database_path}: {error}"
        ) from error
    finally:
        connection.close()
    return enl_path


class EndNoteLibraryCliTests(unittest.TestCase):
    """Drive the documented endnote_library.py command line, offline."""

    @classmethod
    def setUpClass(cls) -> None:
        if not LIBRARY_SCRIPT.is_file():
            raise unittest.SkipTest(f"{LIBRARY_SCRIPT} is missing.")
        cls._directory = tempfile.TemporaryDirectory(prefix="endnote-library-")
        cls.workspace = Path(cls._directory.name)
        cls.library = build_synthetic_library(cls.workspace)

    @classmethod
    def tearDownClass(cls) -> None:
        cls._directory.cleanup()

    def run_cli(self, *arguments: str) -> dict:
        """Run the library CLI and return its single JSON object."""
        completed = subprocess.run(
            [sys.executable, str(LIBRARY_SCRIPT), *arguments],
            capture_output=True,
            text=True,
            encoding="utf-8",
            timeout=CLI_TIMEOUT_SECONDS,
        )
        stdout = (completed.stdout or "").strip()
        self.assertTrue(
            stdout,
            f"No stdout from {' '.join(arguments)}; stderr was: {completed.stderr}",
        )
        try:
            return json.loads(stdout)
        except json.JSONDecodeError:
            lines = [line for line in stdout.splitlines() if line.strip()]
            return json.loads(lines[-1])

    def write_manifest(self, name: str, citations: list[dict]) -> Path:
        path = self.workspace / name
        manifest = {
            "schema_version": 1,
            "library": {"path": str(self.library)},
            "style": "APA 7th",
            "citations": citations,
            "bibliography": {"expected": True},
        }
        path.write_text(json.dumps(manifest, indent=2), encoding="utf-8")
        return path

    # -- search ------------------------------------------------------------------

    def test_search_by_author_returns_the_documented_record_shape(self) -> None:
        result = self.run_cli("search", "--library", str(self.library), "--author", "Thomsen")
        self.assertTrue(result["ok"], result)
        records = result["references"]
        self.assertEqual(len(records), 1)
        record = records[0]
        self.assertEqual(record["rec_number"], 1)
        self.assertEqual(record["reference_type"], 17)
        self.assertEqual(record["authors"], ["Thomsen, Peter T.", "Shearer, Jane K."])
        self.assertEqual(record["first_author_surname"], "Thomsen")
        self.assertEqual(record["year"], "2023")
        self.assertEqual(record["journal"], "Journal of Test Fixtures")
        self.assertEqual(record["volume"], "12")
        self.assertEqual(record["issue"], "3")
        self.assertEqual(record["pages"], "45-46")
        self.assertEqual(record["doi"], "10.0000/fixture.0001")
        self.assertLessEqual(len(record["abstract_head"]), ABSTRACT_HEAD_LIMIT)
        self.assertLess(len(record["abstract_head"]), len(LONG_ABSTRACT))

    def test_search_query_is_case_insensitive_across_title_and_abstract(self) -> None:
        by_title = self.run_cli("search", "--library", str(self.library), "--query", "MERGED CITATION")
        self.assertTrue(by_title["ok"], by_title)
        self.assertEqual([record["rec_number"] for record in by_title["references"]], [2])
        self.assertEqual(by_title["references"][0]["first_author_surname"], "Magrin")

        by_abstract = self.run_cli(
            "search", "--library", str(self.library), "--query", "accent tolerant"
        )
        self.assertTrue(by_abstract["ok"], by_abstract)
        self.assertEqual([record["rec_number"] for record in by_abstract["references"]], [4])

    def test_search_year_is_exact(self) -> None:
        result = self.run_cli("search", "--library", str(self.library), "--year", "2020")
        self.assertTrue(result["ok"], result)
        self.assertEqual([record["rec_number"] for record in result["references"]], [3])

    def test_trashed_records_are_never_returned(self) -> None:
        searched = self.run_cli("search", "--library", str(self.library), "--query", "trash")
        self.assertTrue(searched["ok"], searched)
        self.assertEqual(searched["references"], [])

        fetched = self.run_cli("get", "--library", str(self.library), "--rec-number", "5")
        self.assertFalse(fetched["ok"], fetched)

    def test_search_limit_is_applied(self) -> None:
        result = self.run_cli(
            "search", "--library", str(self.library), "--query", "Synthetic", "--limit", "2"
        )
        self.assertTrue(result["ok"], result)
        self.assertEqual(len(result["references"]), 2)

    def test_search_accepts_an_endnote_xml_export(self) -> None:
        export = self.workspace / "export.xml"
        export.write_text(XML_EXPORT, encoding="utf-8")
        result = self.run_cli("search", "--library", str(export), "--author", "Thomsen")
        self.assertTrue(result["ok"], result)
        self.assertEqual(len(result["references"]), 1)
        record = result["references"][0]
        self.assertEqual(record["rec_number"], 1)
        self.assertEqual(record["first_author_surname"], "Thomsen")
        self.assertEqual(record["year"], "2023")
        self.assertEqual(record["doi"], "10.0000/fixture.0001")

    # -- get ---------------------------------------------------------------------

    def test_get_returns_one_record(self) -> None:
        result = self.run_cli("get", "--library", str(self.library), "--rec-number", "3")
        self.assertTrue(result["ok"], result)
        record = result["reference"]
        self.assertEqual(record["rec_number"], 3)
        self.assertEqual(record["first_author_surname"], "Bernhard")
        self.assertEqual(record["authors"], ["Bernhard, Anna", "Müller, Karin"])

    def test_get_missing_record_fails_cleanly(self) -> None:
        result = self.run_cli("get", "--library", str(self.library), "--rec-number", "9999")
        self.assertFalse(result["ok"], result)
        self.assertIn("error", result)

    # -- precheck ----------------------------------------------------------------

    def test_precheck_accepts_a_matching_manifest(self) -> None:
        manifest = self.write_manifest(
            "good-manifest.json",
            [
                {
                    "id": "single",
                    "marker": "{{CITE:single}}",
                    "targets": [{"rec_number": 1, "author": "Thomsen", "year": 2023}],
                },
                {
                    "id": "merged",
                    "marker": "{{CITE:merged}}",
                    "targets": [
                        {"rec_number": 2, "author": "Magrin", "year": 2023},
                        {"rec_number": 3, "author": "Bernhard", "year": 2020},
                    ],
                },
            ],
        )
        result = self.run_cli(
            "precheck", "--library", str(self.library), "--manifest", str(manifest)
        )
        self.assertTrue(result["ok"], result)
        self.assertEqual(result.get("problems", []), [])

    def test_precheck_reports_every_problem_not_only_the_first(self) -> None:
        manifest = self.write_manifest(
            "bad-manifest.json",
            [
                {
                    "id": "missingRecord",
                    "marker": "{{CITE:missingRecord}}",
                    "targets": [{"rec_number": 9999, "author": "Nobody", "year": 1900}],
                },
                {
                    "id": "wrongAuthor",
                    "marker": "{{CITE:wrongAuthor}}",
                    "targets": [{"rec_number": 1, "author": "Bernhard", "year": 2023}],
                },
                {
                    "id": "wrongYear",
                    "marker": "{{CITE:wrongYear}}",
                    "targets": [{"rec_number": 2, "author": "Magrin", "year": 1999}],
                },
                {
                    "id": "trashedRecord",
                    "marker": "{{CITE:trashedRecord}}",
                    "targets": [{"rec_number": 5, "author": "Deleted", "year": 2018}],
                },
            ],
        )
        result = self.run_cli(
            "precheck", "--library", str(self.library), "--manifest", str(manifest)
        )
        self.assertFalse(result["ok"], result)
        problems = result["problems"]
        self.assertGreaterEqual(len(problems), 4)
        serialised = json.dumps(problems, ensure_ascii=False)
        for token in ("9999", "wrongAuthor", "wrongYear", "trashedRecord"):
            self.assertIn(token, serialised)

    def test_precheck_author_matching_is_accent_tolerant(self) -> None:
        manifest = self.write_manifest(
            "accent-manifest.json",
            [
                {
                    "id": "diacritic",
                    "marker": "{{CITE:diacritic}}",
                    "targets": [{"rec_number": 4, "author": "Muller", "year": 2019}],
                }
            ],
        )
        result = self.run_cli(
            "precheck", "--library", str(self.library), "--manifest", str(manifest)
        )
        self.assertTrue(result["ok"], result)

    # -- failure surfaces --------------------------------------------------------

    def test_missing_data_directory_fails_with_a_clear_message(self) -> None:
        orphan = self.workspace / "Orphan.enl"
        orphan.write_bytes(ENL_MAGIC + b"\x00" * 28)
        result = self.run_cli("search", "--library", str(orphan), "--query", "anything")
        self.assertFalse(result["ok"], result)
        self.assertIn("Orphan.Data", json.dumps(result))

    def test_missing_library_fails_with_a_clear_message(self) -> None:
        absent = self.workspace / "DoesNotExist.enl"
        result = self.run_cli("search", "--library", str(absent), "--query", "anything")
        self.assertFalse(result["ok"], result)
        self.assertIn("error", result)

    # -- export-ris --------------------------------------------------------------

    def test_export_ris_writes_crlf_records_bounded_by_ty_and_er(self) -> None:
        references = [
            {
                "type": "JOUR",
                # The diacritic is deliberate: without it nothing in this test would fail if
                # the writer emitted cp1252 instead of the UTF-8 the spec requires.
                "authors": ["Thomsen, Peter T.", "Shearer, Jane K.", "Müller, Karin"],
                "year": "2023",
                "title": "Synthetic fixture record one for automated citation tests",
                "journal": "Journal of Test Fixtures",
                "volume": "12",
                "issue": "3",
                "pages": "45-46",
                "doi": "10.0000/fixture.0001",
                "url": "https://doi.org/10.0000/fixture.0001",
                "abstract": "Synthetic abstract that describes no real study.",
                "keywords": ["fixture", "automation"],
                "publisher": "Fixture Press",
            }
        ]
        references_path = self.workspace / "references.json"
        references_path.write_text(json.dumps(references), encoding="utf-8")
        output_path = self.workspace / "exported.ris"

        result = self.run_cli(
            "export-ris", "--output", str(output_path), "--references", str(references_path)
        )
        self.assertTrue(result["ok"], result)
        self.assertTrue(output_path.is_file())

        raw = output_path.read_bytes()
        self.assertIn(b"\r\n", raw)
        self.assertEqual(raw.count(b"\n"), raw.count(b"\r\n"), "every LF must be part of a CRLF")
        text = raw.decode("utf-8")
        lines = [line for line in text.split("\r\n") if line.strip()]
        self.assertTrue(lines[0].startswith("TY  - "))
        self.assertTrue(lines[-1].startswith("ER  -"))
        self.assertIn("AU  - Thomsen, Peter T.", lines)
        self.assertIn("AU  - Shearer, Jane K.", lines)
        self.assertIn("AU  - Müller, Karin", lines)
        self.assertIn("DO  - 10.0000/fixture.0001", lines)

        exported_tags = {line[:2] for line in lines}
        fixture_tags = {
            line[:2]
            for line in FIXTURE_RIS.read_bytes().decode("utf-8").split("\r\n")
            if line.strip()
        }
        self.assertTrue(
            exported_tags.issubset(fixture_tags),
            f"Unexpected RIS tags: {sorted(exported_tags - fixture_tags)}",
        )

    def test_export_ris_rejects_an_empty_reference_list(self) -> None:
        references_path = self.workspace / "empty.json"
        references_path.write_text("[]", encoding="utf-8")
        output_path = self.workspace / "empty.ris"
        result = self.run_cli(
            "export-ris", "--output", str(output_path), "--references", str(references_path)
        )
        self.assertFalse(result["ok"], result)
        self.assertFalse(output_path.exists())


class FixtureRisTests(unittest.TestCase):
    """The shipped RIS fixture is the reference for the RIS conventions used here."""

    def test_fixture_is_utf8_crlf_and_record_bounded(self) -> None:
        raw = FIXTURE_RIS.read_bytes()
        text = raw.decode("utf-8")  # read_bytes avoids universal-newline translation

        # A pure ASCII fixture would decode identically under cp1252, so the encoding claim
        # would assert nothing. One diacritic makes the decode a real check.
        self.assertTrue(any(byte > 127 for byte in raw), "the fixture exercises no UTF-8 byte")
        self.assertIn("Müller, Karin", text)
        self.assertIn("\r\n", text)
        self.assertEqual(text.count("\n"), text.count("\r\n"))
        records = [block for block in text.split("\r\n\r\n") if block.strip()]
        self.assertEqual(len(records), 3)
        for record in records:
            lines = [line for line in record.split("\r\n") if line.strip()]
            self.assertTrue(lines[0].startswith("TY  - "))
            self.assertTrue(lines[-1].startswith("ER  -"))
            self.assertTrue(any(line.startswith("AU  - ") for line in lines))
            self.assertTrue(any(line.startswith("TI  - ") for line in lines))


if __name__ == "__main__":
    unittest.main()
