#!/usr/bin/env python3
"""Read-only EndNote library access plus RIS export, standard library only.

This module is the library read side of the endnote-for-word plugin. It never writes to an
EndNote library. The SQLite store is copied to a temporary file and the copy is opened
read-only, so a running EndNote instance is never disturbed and the original bytes are never
touched.

Two library inputs are understood, chosen by file extension:

``<Library>.enl``
    The real library. The ``.enl`` file itself is a small stub; the records live in the
    sibling ``<Library>.Data/sdb/sdb.eni`` file, which is a SQLite 3 database whose schema
    declares custom collations (``ENCI_Base``, ``ENCIN_zh_TW``) that plain sqlite3 cannot
    resolve. Neutral collations are registered before any query runs.

``<Library>.xml``
    An EndNote XML export, used as a fallback when the SQLite store is unavailable, for
    example on a machine that only received the export.

``<Library>.eni`` is also accepted so that a data file can be pointed at directly.

Every public function returns a plain ``dict`` with an ``ok`` key, so the MCP layer can
import this module and forward the result without translation. Diagnostics go to stderr
through ``logging``; stdout carries the single JSON result and nothing else.
"""

from __future__ import annotations

import argparse
import json
import logging
import os
import re
import shutil
import sqlite3
import sys
import tempfile
import unicodedata
import xml.etree.ElementTree as ElementTree
from contextlib import contextmanager
from pathlib import Path
from typing import Any, Iterable, Iterator, Mapping, Sequence


LOGGER = logging.getLogger("endnote_library")

# --- library layout -------------------------------------------------------------------
DATA_DIRECTORY_SUFFIX = ".Data"
SQLITE_RELATIVE_PARTS = ("sdb", "sdb.eni")
LIBRARY_SUFFIX_ENL = ".enl"
LIBRARY_SUFFIX_XML = ".xml"
LIBRARY_SUFFIX_SQLITE = ".eni"
SUPPORTED_LIBRARY_SUFFIXES = (LIBRARY_SUFFIX_ENL, LIBRARY_SUFFIX_XML, LIBRARY_SUFFIX_SQLITE)
LIBRARY_FORMAT_SQLITE = "sqlite"
LIBRARY_FORMAT_XML = "xml"

# --- SQLite access --------------------------------------------------------------------
TEMP_DIRECTORY_PREFIX = "endnote-library-"
REFERENCE_TABLE = "refs"
LIVE_TRASH_STATE = 0
ENDNOTE_COLLATIONS = ("ENCI_Base", "ENCIN_zh_TW")
LIKE_ESCAPE_CHARACTER = "\\"
SELECT_COLUMNS = (
    "id",
    "reference_type",
    "author",
    "year",
    "title",
    "secondary_title",
    "volume",
    "number",
    "pages",
    "electronic_resource_number",
    "url",
    "abstract",
    "keywords",
    "trash_state",
)
FREE_TEXT_SEARCH_COLUMNS = (
    "author",
    "title",
    "secondary_title",
    "electronic_resource_number",
    "keywords",
    "abstract",
)

# --- result shaping -------------------------------------------------------------------
DEFAULT_SEARCH_LIMIT = 25
MAX_SEARCH_LIMIT = 200
ABSTRACT_HEAD_CHARACTERS = 300
AUTHOR_SPLIT_PATTERN = re.compile(r"[\r\n]+")
WHITESPACE_PATTERN = re.compile(r"\s+")
SURNAME_TRIM_CHARACTERS = " .,;"
YEAR_TOKEN_PATTERN = re.compile(r"\b\d{4}\b")

# --- manifest -------------------------------------------------------------------------
MANIFEST_SCHEMA_VERSION = 1

# --- XML export hardening -------------------------------------------------------------
DOCUMENT_TYPE_MARKER = b"<!DOCTYPE"
DOCUMENT_TYPE_PATTERN = re.compile(re.escape(DOCUMENT_TYPE_MARKER), re.IGNORECASE)
XML_SCAN_CHUNK_BYTES = 65536

# --- RIS export -----------------------------------------------------------------------
RIS_TAG_SEPARATOR = "  - "
RIS_LINE_ENDING = "\r\n"
RIS_DEFAULT_TYPE = "JOUR"
# RIS type codes run from two letters (EQ is not used, but ART, PAT, GEN are) up to seven
# (MANSCPT, ANCIENT). A four letter cap would silently rewrite CPAPER, EDBOOK, EJOUR,
# GOVDOC, MANSCPT, ENCYC, FIGURE, SLIDE, SOUND, STAND, UNBILL and VIDEO as JOUR, including
# codes this module's own alias table emits.
RIS_TYPE_CODE_PATTERN = re.compile(r"^[A-Z]{2,8}$")
RIS_PAGE_RANGE_PATTERN = re.compile(r"^\s*([^\s-]+)\s*-{1,2}\s*(\S.*?)\s*$")
RIS_TYPE_ALIASES = {
    "journal article": "JOUR",
    "journal": "JOUR",
    "article": "JOUR",
    "electronic article": "EJOUR",
    "book": "BOOK",
    "edited book": "EDBOOK",
    "book section": "CHAP",
    "chapter": "CHAP",
    "conference paper": "CPAPER",
    "conference proceedings": "CONF",
    "conference": "CONF",
    "thesis": "THES",
    "dissertation": "THES",
    "report": "RPRT",
    "web page": "ELEC",
    "webpage": "ELEC",
    "online database": "DBASE",
    "dataset": "DATA",
    "patent": "PAT",
    "standard": "STAND",
    "manuscript": "MANSCPT",
    "generic": "GEN",
}

# --- logging --------------------------------------------------------------------------
LOG_LEVEL_ENVIRONMENT_VARIABLE = "ENDNOTE_LIBRARY_LOG_LEVEL"
DEFAULT_LOG_LEVEL = "INFO"
LOG_FORMAT = "[endnote_library] %(levelname)s %(message)s"

EXIT_OK = 0
EXIT_FAILURE = 1


# ======================================================================================
# Small text helpers
# ======================================================================================


def _text(value: Any) -> str:
    """Coerce a database or XML value into a stripped string, mapping None to empty."""
    if value is None:
        return ""
    if isinstance(value, str):
        return value.strip()
    return str(value).strip()


def _collapse_whitespace(value: str) -> str:
    return WHITESPACE_PATTERN.sub(" ", value).strip()


def _clean_argument(value: Any) -> str | None:
    """Normalise an optional CLI or tool argument into a non-empty string or None."""
    if value is None:
        return None
    text = str(value).strip()
    return text or None


def split_authors(author_block: Any) -> list[str]:
    """Split the EndNote author column, which stores one ``Last, First`` name per CR line."""
    block = _text(author_block)
    if not block:
        return []
    return [_collapse_whitespace(part) for part in AUTHOR_SPLIT_PATTERN.split(block) if part.strip()]


def author_surname(author: str) -> str:
    """Return the surname of a single ``Last, First`` author entry.

    EndNote stores corporate authors with a trailing comma ("World Health Organization,"),
    which this handles naturally because everything before the first comma is the surname.
    """
    return _collapse_whitespace(_text(author).split(",")[0])


def normalize_name(value: Any) -> str:
    """Fold a surname for tolerant comparison.

    Case is folded and combining marks are removed after NFKD decomposition, so "Munoz"
    matches "Munoz" with an accent and "MUNOZ". Surrounding punctuation is trimmed.
    """
    decomposed = unicodedata.normalize("NFKD", _text(value))
    without_marks = "".join(char for char in decomposed if not unicodedata.combining(char))
    return _collapse_whitespace(without_marks).casefold().strip(SURNAME_TRIM_CHARACTERS)


def _failure(action: str, message: str, **extra: Any) -> dict[str, Any]:
    LOGGER.error("%s failed: %s", action, message)
    result: dict[str, Any] = {"ok": False, "action": action, "error": message}
    result.update(extra)
    return result


# ======================================================================================
# Library path resolution
# ======================================================================================


def resolve_data_directory(enl_path: str | os.PathLike[str]) -> Path:
    """Map ``<Library>.enl`` to its sibling ``<Library>.Data`` directory.

    Raises:
        FileNotFoundError: when the data directory is absent, which is the common symptom
            of an .enl stub that was copied without its data folder.
    """
    path = Path(enl_path).expanduser()
    data_directory = path.with_name(path.stem + DATA_DIRECTORY_SUFFIX)
    if not data_directory.is_dir():
        raise FileNotFoundError(
            f"EndNote data directory not found: {data_directory}. "
            "An .enl file is only a stub; the records live in the sibling .Data folder, "
            "which must be kept next to it."
        )
    return data_directory


def resolve_sqlite_path(enl_path: str | os.PathLike[str]) -> Path:
    """Return the SQLite store inside ``<Library>.Data`` for an ``.enl`` library."""
    data_directory = resolve_data_directory(enl_path)
    database = data_directory.joinpath(*SQLITE_RELATIVE_PARTS)
    if not database.is_file():
        raise FileNotFoundError(
            f"EndNote SQLite store not found: {database}. "
            "Open the library once in EndNote to rebuild it, or supply an EndNote XML "
            "export instead."
        )
    return database


def _resolve_library(library: Any) -> tuple[Path, Path, str]:
    """Resolve a caller-supplied library path.

    Returns:
        A tuple of (library path as given, readable source file, library format).
    """
    text = _clean_argument(library)
    if text is None:
        raise ValueError(
            "A library path is required. Pass the .enl library file or an EndNote XML export."
        )
    path = Path(text).expanduser()
    suffix = path.suffix.lower()
    if suffix not in SUPPORTED_LIBRARY_SUFFIXES:
        raise ValueError(
            f"Unsupported library file '{path.name}'. Supported extensions: "
            + ", ".join(SUPPORTED_LIBRARY_SUFFIXES)
        )
    if not path.is_file():
        raise FileNotFoundError(f"Library file not found: {path}")
    if suffix == LIBRARY_SUFFIX_XML:
        return path, path, LIBRARY_FORMAT_XML
    if suffix == LIBRARY_SUFFIX_SQLITE:
        return path, path, LIBRARY_FORMAT_SQLITE
    return path, resolve_sqlite_path(path), LIBRARY_FORMAT_SQLITE


# ======================================================================================
# Record shaping
# ======================================================================================


def _raw_record(
    rec_number: int,
    reference_type: Any,
    author_block: Any,
    year: Any,
    title: Any,
    secondary_title: Any,
    volume: Any,
    issue: Any,
    pages: Any,
    doi: Any,
    url: Any,
    abstract: Any,
    keywords: Any,
    trash_state: int,
) -> dict[str, Any]:
    """Build the internal record used for matching, before it is trimmed for output."""
    return {
        "rec_number": rec_number,
        "reference_type": reference_type,
        "author_block": _text(author_block),
        "year": _text(year),
        "title": _collapse_whitespace(_text(title)),
        "secondary_title": _collapse_whitespace(_text(secondary_title)),
        "volume": _text(volume),
        "issue": _text(issue),
        "pages": _text(pages),
        "doi": _text(doi),
        "url": _text(url),
        "abstract": _collapse_whitespace(_text(abstract)),
        "keywords": _collapse_whitespace(_text(keywords)),
        "trash_state": trash_state,
    }


def _reference_type_number(value: Any) -> int:
    """Return the numeric EndNote reference type, or -1 when it cannot be determined."""
    text = _text(value)
    if not text:
        return -1
    try:
        return int(text)
    except ValueError:
        LOGGER.debug("non numeric reference type %r, reporting -1", text)
        return -1


def _public_record(raw: Mapping[str, Any]) -> dict[str, Any]:
    """Trim an internal record into the documented output shape.

    Only the first 300 characters of an abstract are exposed; whole abstracts would flood
    an agent's context for no benefit.
    """
    authors = split_authors(raw["author_block"])
    return {
        "rec_number": raw["rec_number"],
        "reference_type": _reference_type_number(raw["reference_type"]),
        "authors": authors,
        "first_author_surname": author_surname(authors[0]) if authors else "",
        "year": raw["year"],
        "title": raw["title"],
        "journal": raw["secondary_title"],
        "volume": raw["volume"],
        "issue": raw["issue"],
        "pages": raw["pages"],
        "doi": raw["doi"],
        "url": raw["url"],
        "abstract_head": raw["abstract"][:ABSTRACT_HEAD_CHARACTERS],
    }


# ======================================================================================
# SQLite reader
# ======================================================================================


def _neutral_collation(left: str, right: str) -> int:
    """Stand in for EndNote's custom collations, which sqlite3 cannot load."""
    return (left > right) - (left < right)


@contextmanager
def _sqlite_reader(database_path: Path) -> Iterator[sqlite3.Connection]:
    """Copy the EndNote SQLite store to a temporary file and open the copy read-only.

    The original file is never opened, so EndNote can keep the library open while this runs.
    """
    temp_directory = Path(tempfile.mkdtemp(prefix=TEMP_DIRECTORY_PREFIX))
    connection: sqlite3.Connection | None = None
    try:
        working_copy = temp_directory / database_path.name
        shutil.copy2(database_path, working_copy)
        # Journal, WAL and shared memory siblings must travel with the database, otherwise
        # recently committed records are invisible in the copy.
        for sibling in sorted(database_path.parent.glob(database_path.name + "-*")):
            if sibling.is_file():
                shutil.copy2(sibling, temp_directory / sibling.name)
        connection = sqlite3.connect(f"{working_copy.as_uri()}?mode=ro", uri=True)
        for collation in ENDNOTE_COLLATIONS:
            connection.create_collation(collation, _neutral_collation)
        yield connection
    finally:
        if connection is not None:
            try:
                connection.close()
            except sqlite3.Error as exc:
                LOGGER.warning("could not close the temporary library copy: %s", exc)
        try:
            shutil.rmtree(temp_directory)
        except OSError as exc:
            LOGGER.warning("could not remove temporary directory %s: %s", temp_directory, exc)


def _row_to_raw(row: Sequence[Any]) -> dict[str, Any]:
    values = dict(zip(SELECT_COLUMNS, row))
    return _raw_record(
        rec_number=int(values["id"]),
        reference_type=values["reference_type"],
        author_block=values["author"],
        year=values["year"],
        title=values["title"],
        secondary_title=values["secondary_title"],
        volume=values["volume"],
        issue=values["number"],
        pages=values["pages"],
        doi=values["electronic_resource_number"],
        url=values["url"],
        abstract=values["abstract"],
        keywords=values["keywords"],
        trash_state=int(values["trash_state"] or LIVE_TRASH_STATE),
    )


def _like_pattern(term: str) -> str:
    """Build a contains pattern with LIKE metacharacters escaped."""
    escaped = term.replace(LIKE_ESCAPE_CHARACTER, LIKE_ESCAPE_CHARACTER * 2)
    for wildcard in ("%", "_"):
        escaped = escaped.replace(wildcard, LIKE_ESCAPE_CHARACTER + wildcard)
    return f"%{escaped}%"


def _like_clause(column: str) -> str:
    return f"{column} LIKE ? ESCAPE '{LIKE_ESCAPE_CHARACTER}'"


def _search_sqlite(
    connection: sqlite3.Connection, filters: Mapping[str, str | None], limit: int
) -> list[dict[str, Any]]:
    clauses = ["trash_state = ?"]
    parameters: list[Any] = [LIVE_TRASH_STATE]

    query = filters.get("query")
    if query:
        clauses.append("(" + " OR ".join(_like_clause(column) for column in FREE_TEXT_SEARCH_COLUMNS) + ")")
        parameters.extend([_like_pattern(query)] * len(FREE_TEXT_SEARCH_COLUMNS))
    for key, column in (("author", "author"), ("title", "title"), ("doi", "electronic_resource_number")):
        value = filters.get(key)
        if value:
            clauses.append(_like_clause(column))
            parameters.append(_like_pattern(value))
    year = filters.get("year")
    if year:
        # CAST keeps the comparison working whether the column holds text or an integer.
        clauses.append("TRIM(CAST(year AS TEXT)) = ?")
        parameters.append(year)

    statement = (
        f"SELECT {', '.join(SELECT_COLUMNS)} FROM {REFERENCE_TABLE} "
        f"WHERE {' AND '.join(clauses)} ORDER BY id LIMIT ?"
    )
    parameters.append(limit)
    LOGGER.debug("sqlite search: %s", statement)
    cursor = connection.execute(statement, parameters)
    try:
        return [_row_to_raw(row) for row in cursor.fetchall()]
    finally:
        cursor.close()


def _fetch_sqlite(
    connection: sqlite3.Connection, rec_number: int, include_trashed: bool
) -> dict[str, Any] | None:
    clauses = ["id = ?"]
    parameters: list[Any] = [rec_number]
    if not include_trashed:
        clauses.append("trash_state = ?")
        parameters.append(LIVE_TRASH_STATE)
    statement = (
        f"SELECT {', '.join(SELECT_COLUMNS)} FROM {REFERENCE_TABLE} "
        f"WHERE {' AND '.join(clauses)} LIMIT 1"
    )
    cursor = connection.execute(statement, parameters)
    try:
        row = cursor.fetchone()
    finally:
        cursor.close()
    return _row_to_raw(row) if row is not None else None


# ======================================================================================
# EndNote XML export reader
# ======================================================================================


def _xml_joined_text(element: ElementTree.Element | None) -> str:
    """Flatten an EndNote XML element, which usually wraps its text in <style> children."""
    if element is None:
        return ""
    return _collapse_whitespace("".join(element.itertext()))


def _xml_child_text(record: ElementTree.Element, path: str) -> str:
    return _xml_joined_text(record.find(path))


def _xml_record_to_raw(record: ElementTree.Element, position: int) -> dict[str, Any] | None:
    rec_text = _xml_child_text(record, "rec-number")
    try:
        rec_number = int(rec_text)
    except ValueError:
        LOGGER.warning("skipping XML record %d with unusable rec-number %r", position, rec_text)
        return None
    authors = [
        _xml_joined_text(author) for author in record.findall("contributors/authors/author")
    ]
    keywords = [_xml_joined_text(keyword) for keyword in record.findall("keywords/keyword")]
    urls = [_xml_joined_text(url) for url in record.findall("urls/related-urls/url")]
    return _raw_record(
        rec_number=rec_number,
        reference_type=_xml_child_text(record, "ref-type"),
        author_block="\r".join(author for author in authors if author),
        year=_xml_child_text(record, "dates/year"),
        title=_xml_child_text(record, "titles/title"),
        secondary_title=_xml_child_text(record, "titles/secondary-title"),
        volume=_xml_child_text(record, "volume"),
        issue=_xml_child_text(record, "number"),
        pages=_xml_child_text(record, "pages"),
        doi=_xml_child_text(record, "electronic-resource-num"),
        url=next((url for url in urls if url), ""),
        abstract=_xml_child_text(record, "abstract"),
        keywords="; ".join(keyword for keyword in keywords if keyword),
        trash_state=LIVE_TRASH_STATE,
    )


def _reject_document_type_declaration(xml_path: Path) -> None:
    """Refuse an XML export that carries a DOCTYPE declaration.

    xml.etree.ElementTree never resolves external entities, but expat does expand internal
    entities declared in an inline DTD, which is the billion laughs and quadratic blowup
    class of input. Genuine EndNote exports have no DTD, so rejecting the declaration
    removes that risk without leaving the standard library for defusedxml. The parser
    itself cannot be hardened instead: CPython's C accelerated XMLParser stopped exposing
    the underlying expat parser, so the defusedxml style StartDoctypeDeclHandler hook is
    unavailable here.

    The whole file is scanned rather than a fixed prolog window. A comment or processing
    instruction may legally precede the declaration and can be padded to any length, so a
    windowed scan is trivially bypassed. Outside the prolog the literal bytes can only
    occur inside a comment or a CDATA section, and refusing those costs nothing.

    ponytail: chunked scan keeps memory flat; the file is read twice in total, which is
    irrelevant next to building the whole ElementTree in memory afterwards.
    """
    overlap = len(DOCUMENT_TYPE_MARKER) - 1
    tail = b""
    with xml_path.open("rb") as handle:
        while True:
            chunk = handle.read(XML_SCAN_CHUNK_BYTES)
            if not chunk:
                return
            window = tail + chunk
            if DOCUMENT_TYPE_PATTERN.search(window):
                raise ValueError(
                    f"{xml_path} contains a DOCTYPE declaration. EndNote XML exports have "
                    "none, and inline entity declarations are refused."
                )
            tail = window[-overlap:]


def _read_xml(xml_path: Path) -> list[dict[str, Any]]:
    """Parse an EndNote XML export into internal records.

    ponytail: the whole export is parsed into memory, which is fine for personal libraries;
    switch to ElementTree.iterparse if multi hundred megabyte exports ever show up.
    """
    _reject_document_type_declaration(xml_path)
    tree = ElementTree.parse(xml_path)
    records: list[dict[str, Any]] = []
    for position, element in enumerate(tree.getroot().iter("record")):
        raw = _xml_record_to_raw(element, position)
        if raw is not None:
            records.append(raw)
    LOGGER.debug("parsed %d records from %s", len(records), xml_path)
    return records


def _matches(raw: Mapping[str, Any], filters: Mapping[str, str | None]) -> bool:
    """Apply the search filters in Python, mirroring the SQL used for the SQLite reader."""
    query = filters.get("query")
    if query:
        haystack = " \n".join(
            (
                raw["author_block"],
                raw["title"],
                raw["secondary_title"],
                raw["doi"],
                raw["keywords"],
                raw["abstract"],
            )
        ).casefold()
        if query.casefold() not in haystack:
            return False
    for key, field in (("author", "author_block"), ("title", "title"), ("doi", "doi")):
        value = filters.get(key)
        if value and value.casefold() not in str(raw[field]).casefold():
            return False
    year = filters.get("year")
    if year and raw["year"].strip() != year:
        return False
    return True


def _search_xml(
    xml_path: Path, filters: Mapping[str, str | None], limit: int
) -> list[dict[str, Any]]:
    matched: list[dict[str, Any]] = []
    for raw in sorted(_read_xml(xml_path), key=lambda record: record["rec_number"]):
        if _matches(raw, filters):
            matched.append(raw)
            if len(matched) >= limit:
                break
    return matched


# ======================================================================================
# Shared lookup
# ======================================================================================


def _fetch_records(
    source: Path, library_format: str, rec_numbers: Iterable[int], include_trashed: bool
) -> dict[int, dict[str, Any]]:
    """Fetch several records in one pass, opening the library exactly once."""
    wanted = sorted({int(rec_number) for rec_number in rec_numbers})
    if not wanted:
        return {}
    if library_format == LIBRARY_FORMAT_XML:
        wanted_set = set(wanted)
        return {
            raw["rec_number"]: raw
            for raw in _read_xml(source)
            if raw["rec_number"] in wanted_set
        }
    found: dict[int, dict[str, Any]] = {}
    with _sqlite_reader(source) as connection:
        for rec_number in wanted:
            raw = _fetch_sqlite(connection, rec_number, include_trashed=include_trashed)
            if raw is not None:
                found[rec_number] = raw
    return found


def _coerce_limit(limit: Any) -> int:
    if limit is None:
        return DEFAULT_SEARCH_LIMIT
    try:
        value = int(limit)
    except (TypeError, ValueError) as exc:
        raise ValueError(f"limit must be an integer, received {limit!r}") from exc
    if value < 1:
        raise ValueError(f"limit must be 1 or greater, received {value}")
    if value > MAX_SEARCH_LIMIT:
        LOGGER.warning("limit %d exceeds the cap, using %d", value, MAX_SEARCH_LIMIT)
        return MAX_SEARCH_LIMIT
    return value


def _coerce_rec_number(value: Any, context: str) -> int:
    try:
        rec_number = int(value)
    except (TypeError, ValueError) as exc:
        raise ValueError(f"{context} must be an integer, received {value!r}") from exc
    if rec_number < 1:
        raise ValueError(f"{context} must be greater than 0, received {rec_number}")
    return rec_number


def _sqlite_error_message(source: Path | None, exc: sqlite3.Error) -> str:
    """Describe a SQLite failure, naming the store only when it was actually resolved."""
    location = f" {source}" if source is not None else ""
    return (
        f"Could not read the EndNote SQLite store{location}: {exc}. "
        "Close EndNote and retry, or supply an EndNote XML export as the library instead."
    )


# ======================================================================================
# Public operations
# ======================================================================================


def search_references(
    library: str | os.PathLike[str],
    query: str | None = None,
    author: str | None = None,
    year: str | int | None = None,
    title: str | None = None,
    doi: str | None = None,
    limit: int | None = DEFAULT_SEARCH_LIMIT,
) -> dict[str, Any]:
    """Search a library and return matching references in the documented record shape."""
    action = "search"
    source: Path | None = None
    try:
        effective_limit = _coerce_limit(limit)
        filters = {
            "query": _clean_argument(query),
            "author": _clean_argument(author),
            "year": _clean_argument(year),
            "title": _clean_argument(title),
            "doi": _clean_argument(doi),
        }
        library_path, source, library_format = _resolve_library(library)
        if library_format == LIBRARY_FORMAT_XML:
            raws = _search_xml(source, filters, effective_limit)
        else:
            with _sqlite_reader(source) as connection:
                raws = _search_sqlite(connection, filters, effective_limit)
    except (FileNotFoundError, ValueError) as exc:
        return _failure(action, str(exc))
    except sqlite3.Error as exc:
        return _failure(action, _sqlite_error_message(source, exc))
    except ElementTree.ParseError as exc:
        return _failure(action, f"Could not parse the EndNote XML export: {exc}")
    except OSError as exc:
        return _failure(action, f"Could not read the library: {exc}")

    references = [_public_record(raw) for raw in raws]
    LOGGER.info("search returned %d reference(s) from %s", len(references), library_path)
    return {
        "ok": True,
        "action": action,
        "library": str(library_path),
        "library_format": library_format,
        "filters": {key: value for key, value in filters.items() if value},
        "limit": effective_limit,
        "count": len(references),
        "truncated": len(references) >= effective_limit,
        "references": references,
    }


def get_reference(library: str | os.PathLike[str], rec_number: int) -> dict[str, Any]:
    """Fetch one reference by its EndNote record number."""
    action = "get"
    source: Path | None = None
    try:
        wanted = _coerce_rec_number(rec_number, "rec_number")
        library_path, source, library_format = _resolve_library(library)
        found = _fetch_records(source, library_format, [wanted], include_trashed=True)
    except (FileNotFoundError, ValueError) as exc:
        return _failure(action, str(exc))
    except sqlite3.Error as exc:
        return _failure(action, _sqlite_error_message(source, exc))
    except ElementTree.ParseError as exc:
        return _failure(action, f"Could not parse the EndNote XML export: {exc}")
    except OSError as exc:
        return _failure(action, f"Could not read the library: {exc}")

    raw = found.get(wanted)
    if raw is None:
        return _failure(
            action,
            f"Record {wanted} does not exist in {library_path}. "
            "Citing a missing record makes EndNote open a blocking dialog, so this must be "
            "fixed before rendering.",
            library=str(library_path),
            rec_number=wanted,
            found=False,
        )
    if raw["trash_state"] != LIVE_TRASH_STATE:
        return _failure(
            action,
            f"Record {wanted} is in the trash of {library_path} and cannot be cited.",
            library=str(library_path),
            rec_number=wanted,
            found=True,
            in_trash=True,
        )
    return {
        "ok": True,
        "action": action,
        "library": str(library_path),
        "library_format": library_format,
        "found": True,
        "reference": _public_record(raw),
    }


def _load_json_document(value: Any, description: str) -> Any:
    """Accept either an already parsed JSON document or a path to a JSON file."""
    if isinstance(value, (dict, list)):
        return value
    text = _clean_argument(value)
    if text is None:
        raise ValueError(f"A {description} is required.")
    path = Path(text).expanduser()
    if not path.is_file():
        raise FileNotFoundError(f"{description.capitalize()} file not found: {path}")
    try:
        with path.open("r", encoding="utf-8-sig") as handle:
            return json.load(handle)
    except json.JSONDecodeError as exc:
        raise ValueError(f"{description.capitalize()} {path} is not valid JSON: {exc}") from exc


def _problem(
    kind: str,
    message: str,
    citation_id: Any = None,
    target_index: int | None = None,
    rec_number: Any = None,
    **extra: Any,
) -> dict[str, Any]:
    """Build one precheck problem entry in the shape the agent contract expects."""
    record: dict[str, Any] = {
        "citation_id": citation_id,
        "target_index": target_index,
        "rec_number": rec_number,
        "problem": kind,
        "message": message,
    }
    record.update(extra)
    return record


def _manifest_targets(manifest: Mapping[str, Any]) -> tuple[list[dict[str, Any]], list[dict[str, Any]]]:
    """Flatten manifest citations into targets, collecting structural problems as it goes."""
    targets: list[dict[str, Any]] = []
    problems: list[dict[str, Any]] = []

    schema_version = manifest.get("schema_version")
    if schema_version != MANIFEST_SCHEMA_VERSION:
        problems.append(
            _problem(
                "invalid_manifest",
                f"schema_version must be {MANIFEST_SCHEMA_VERSION}, found {schema_version!r}.",
            )
        )

    citations = manifest.get("citations")
    if not isinstance(citations, list) or not citations:
        problems.append(_problem("invalid_manifest", "citations must be a non-empty array."))
        return targets, problems

    for citation_index, citation in enumerate(citations):
        if not isinstance(citation, dict):
            problems.append(
                _problem("invalid_manifest", f"citations[{citation_index}] must be an object.")
            )
            continue
        citation_id = citation.get("id")
        citation_targets = citation.get("targets")
        if not isinstance(citation_targets, list) or not citation_targets:
            problems.append(
                _problem(
                    "invalid_manifest",
                    f"citations[{citation_index}].targets must be a non-empty array.",
                    citation_id=citation_id,
                )
            )
            continue
        for target_index, target in enumerate(citation_targets):
            if not isinstance(target, dict):
                problems.append(
                    _problem(
                        "invalid_manifest",
                        "each target must be an object with rec_number, author and year.",
                        citation_id=citation_id,
                        target_index=target_index,
                    )
                )
                continue
            targets.append(
                {"citation_id": citation_id, "target_index": target_index, "target": target}
            )
    return targets, problems


def _year_matches(expected: Any, actual: str) -> bool:
    """Compare manifest and record years.

    A record year is usually just "2023" but EndNote also accepts values such as
    "2023 Jan", so a standalone four digit token counts as a match.
    """
    expected_text = _text(expected)
    actual_text = _text(actual)
    if not expected_text:
        return False
    if expected_text == actual_text:
        return True
    return expected_text in YEAR_TOKEN_PATTERN.findall(actual_text)


def precheck_manifest(
    library: str | os.PathLike[str] | None, manifest: Any
) -> dict[str, Any]:
    """Verify that every manifest target resolves to a live, matching library record.

    Every mismatch is reported, never only the first, because a single unresolved record
    number makes EndNote open a modal dialog that blocks the whole render.
    """
    action = "precheck"
    source: Path | None = None
    try:
        document = _load_json_document(manifest, "manifest")
        if not isinstance(document, dict):
            raise ValueError("The manifest must be a JSON object.")
        effective_library = _clean_argument(library)
        if effective_library is None:
            declared = document.get("library")
            if isinstance(declared, dict):
                effective_library = _clean_argument(declared.get("path"))
        if effective_library is None:
            raise ValueError(
                "No library path supplied and the manifest does not declare library.path."
            )
        library_path, source, library_format = _resolve_library(effective_library)
        targets, problems = _manifest_targets(document)

        rec_numbers: list[int] = []
        for entry in targets:
            try:
                rec_numbers.append(_coerce_rec_number(entry["target"].get("rec_number"), "rec_number"))
            except ValueError:
                # Reported per target below; the lookup simply skips it.
                continue
        records = _fetch_records(source, library_format, rec_numbers, include_trashed=True)
    except (FileNotFoundError, ValueError) as exc:
        return _failure(action, str(exc))
    except sqlite3.Error as exc:
        return _failure(action, _sqlite_error_message(source, exc))
    except ElementTree.ParseError as exc:
        return _failure(action, f"Could not parse the EndNote XML export: {exc}")
    except OSError as exc:
        return _failure(action, f"Could not read the library: {exc}")

    resolved: list[dict[str, Any]] = []
    for entry in targets:
        citation_id = entry["citation_id"]
        target_index = entry["target_index"]
        target = entry["target"]
        where = {
            "citation_id": citation_id,
            "target_index": target_index,
            "rec_number": target.get("rec_number"),
        }

        try:
            rec_number = _coerce_rec_number(target.get("rec_number"), "rec_number")
        except ValueError as exc:
            problems.append(_problem("invalid_rec_number", str(exc), **where))
            continue

        raw = records.get(rec_number)
        if raw is None:
            problems.append(
                _problem(
                    "record_not_found",
                    f"Record {rec_number} does not exist in {library_path}.",
                    **where,
                )
            )
            continue
        if raw["trash_state"] != LIVE_TRASH_STATE:
            problems.append(
                _problem(
                    "record_in_trash", f"Record {rec_number} is in the library trash.", **where
                )
            )
            continue

        authors = split_authors(raw["author_block"])
        actual_surname = author_surname(authors[0]) if authors else ""
        expected_author = _text(target.get("author"))
        if not expected_author:
            problems.append(
                _problem("missing_author", "The target has no author surname.", **where)
            )
        elif normalize_name(expected_author) != normalize_name(actual_surname):
            problems.append(
                _problem(
                    "author_mismatch",
                    f"Record {rec_number} has first author surname '{actual_surname}', "
                    f"the manifest says '{expected_author}'.",
                    expected=expected_author,
                    actual=actual_surname,
                    **where,
                )
            )

        expected_year = target.get("year")
        if _text(expected_year) == "":
            problems.append(_problem("missing_year", "The target has no year.", **where))
        elif not _year_matches(expected_year, raw["year"]):
            problems.append(
                _problem(
                    "year_mismatch",
                    f"Record {rec_number} has year '{raw['year']}', "
                    f"the manifest says '{_text(expected_year)}'.",
                    expected=_text(expected_year),
                    actual=raw["year"],
                    **where,
                )
            )

        resolved.append(
            {
                "citation_id": citation_id,
                "target_index": target_index,
                "rec_number": rec_number,
                "first_author_surname": actual_surname,
                "year": raw["year"],
                "title": raw["title"],
            }
        )

    ok = not problems
    if not ok:
        LOGGER.error("precheck found %d problem(s) in the manifest", len(problems))
    return {
        "ok": ok,
        "action": action,
        "library": str(library_path),
        "library_format": library_format,
        "checked": len(targets),
        "resolved": resolved,
        "problems": problems,
    }


# ======================================================================================
# RIS export
# ======================================================================================


def _ris_value(value: Any) -> str:
    """Flatten a value onto one line; embedded newlines break RIS parsers."""
    return _collapse_whitespace(_text(value))


def _ris_type_code(value: Any) -> str:
    text = _ris_value(value)
    if not text:
        return RIS_DEFAULT_TYPE
    alias = RIS_TYPE_ALIASES.get(text.casefold())
    if alias:
        return alias
    upper = text.upper()
    if RIS_TYPE_CODE_PATTERN.match(upper):
        return upper
    LOGGER.warning("unknown reference type %r, exporting as %s", text, RIS_DEFAULT_TYPE)
    return RIS_DEFAULT_TYPE


def _ris_string_list(value: Any, field: str, index: int) -> list[str]:
    if value is None:
        return []
    if not isinstance(value, list):
        raise ValueError(f"references[{index}].{field} must be an array of strings.")
    items: list[str] = []
    for item in value:
        if not isinstance(item, str):
            raise ValueError(f"references[{index}].{field} must contain strings only.")
        text = _ris_value(item)
        if text:
            items.append(text)
    return items


def _ris_record_lines(reference: Mapping[str, Any], index: int) -> list[str]:
    authors = _ris_string_list(reference.get("authors"), "authors", index)
    keywords = _ris_string_list(reference.get("keywords"), "keywords", index)
    title = _ris_value(reference.get("title"))
    if not title and not authors:
        raise ValueError(
            f"references[{index}] has neither a title nor authors, which cannot be imported."
        )

    lines = [f"TY{RIS_TAG_SEPARATOR}{_ris_type_code(reference.get('type'))}"]
    lines.extend(f"AU{RIS_TAG_SEPARATOR}{author}" for author in authors)
    for tag, key in (
        ("PY", "year"),
        ("TI", "title"),
        ("T2", "journal"),
        ("VL", "volume"),
        ("IS", "issue"),
        ("PB", "publisher"),
        ("DO", "doi"),
        ("UR", "url"),
        ("AB", "abstract"),
    ):
        value = _ris_value(reference.get(key))
        if value:
            lines.append(f"{tag}{RIS_TAG_SEPARATOR}{value}")
    pages = _ris_value(reference.get("pages"))
    if pages:
        match = RIS_PAGE_RANGE_PATTERN.match(pages)
        if match:
            lines.append(f"SP{RIS_TAG_SEPARATOR}{match.group(1)}")
            lines.append(f"EP{RIS_TAG_SEPARATOR}{match.group(2)}")
        else:
            lines.append(f"SP{RIS_TAG_SEPARATOR}{pages}")
    lines.extend(f"KW{RIS_TAG_SEPARATOR}{keyword}" for keyword in keywords)
    lines.append(f"ER{RIS_TAG_SEPARATOR}")
    return lines


def export_ris(output: str | os.PathLike[str], references: Any) -> dict[str, Any]:
    """Write a UTF-8 RIS file that EndNote can import.

    Args:
        output: destination ``.ris`` path; parent directories are created.
        references: a list of reference objects, or a path to a JSON file holding one.
    """
    action = "export-ris"
    try:
        document = _load_json_document(references, "references document")
        if not isinstance(document, list):
            raise ValueError("The references document must be a JSON array of objects.")
        if not document:
            raise ValueError("The references document is empty; nothing to export.")
        for index, reference in enumerate(document):
            if not isinstance(reference, dict):
                raise ValueError(f"references[{index}] must be an object.")

        output_text = _clean_argument(output)
        if output_text is None:
            raise ValueError("An output path is required.")
        output_path = Path(output_text).expanduser()
        if output_path.is_dir():
            raise ValueError(f"The output path is a directory: {output_path}")

        blocks: list[str] = []
        for index, reference in enumerate(document):
            lines = _ris_record_lines(reference, index)
            blocks.append(RIS_LINE_ENDING.join(lines) + RIS_LINE_ENDING)
        payload = RIS_LINE_ENDING.join(blocks).encode("utf-8")

        parent = output_path.parent
        if str(parent):
            parent.mkdir(parents=True, exist_ok=True)
        # Binary mode: the CRLF endings are already explicit and must not be translated
        # again, and no BOM may be added.
        with output_path.open("wb") as handle:
            handle.write(payload)
    except (ValueError, FileNotFoundError) as exc:
        return _failure(action, str(exc))
    except OSError as exc:
        return _failure(action, f"Could not write the RIS file: {exc}")

    LOGGER.info("wrote %d reference(s) to %s", len(document), output_path)
    return {
        "ok": True,
        "action": action,
        "output_path": str(output_path.resolve()),
        "reference_count": len(document),
        "encoding": "utf-8",
        "line_ending": "CRLF",
        "bytes_written": len(payload),
    }


# ======================================================================================
# Command line interface
# ======================================================================================


def _configure_logging() -> None:
    level_name = os.environ.get(LOG_LEVEL_ENVIRONMENT_VARIABLE, DEFAULT_LOG_LEVEL).upper()
    level = getattr(logging, level_name, None)
    if not isinstance(level, int):
        level = logging.INFO
    handler = logging.StreamHandler(stream=sys.stderr)
    handler.setFormatter(logging.Formatter(LOG_FORMAT))
    logging.basicConfig(level=level, handlers=[handler], force=True)
    if level_name != DEFAULT_LOG_LEVEL and not isinstance(getattr(logging, level_name, None), int):
        LOGGER.warning("unknown log level %r, falling back to INFO", level_name)


def _force_utf8_streams() -> None:
    """Console output defaults to a legacy code page on Windows; force UTF-8."""
    for stream in (sys.stdout, sys.stderr):
        reconfigure = getattr(stream, "reconfigure", None)
        if reconfigure is None:
            continue
        try:
            reconfigure(encoding="utf-8")
        except (OSError, ValueError) as exc:
            LOGGER.debug("could not force UTF-8 on %r: %s", stream, exc)


def _build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(
        prog="endnote_library.py",
        description=(
            "Read an EndNote library without opening it for writing, and write RIS import "
            "files. Emits exactly one JSON object on stdout."
        ),
    )
    subparsers = parser.add_subparsers(dest="command", required=True)

    search = subparsers.add_parser("search", help="Search the library for references.")
    search.add_argument("--library", required=True, help="Path to the .enl library or an EndNote XML export.")
    search.add_argument("--query", help="Free text matched across author, title, journal, DOI, keywords and abstract.")
    search.add_argument("--author", help="Substring matched against the author column.")
    search.add_argument("--year", help="Exact publication year.")
    search.add_argument("--title", help="Substring matched against the title column.")
    search.add_argument("--doi", help="Substring matched against the DOI column.")
    search.add_argument(
        "--limit",
        type=int,
        default=DEFAULT_SEARCH_LIMIT,
        help=f"Maximum results, capped at {MAX_SEARCH_LIMIT} (default {DEFAULT_SEARCH_LIMIT}).",
    )

    get = subparsers.add_parser("get", help="Fetch one reference by record number.")
    get.add_argument("--library", required=True, help="Path to the .enl library or an EndNote XML export.")
    get.add_argument("--rec-number", type=int, required=True, help="EndNote record number.")

    precheck = subparsers.add_parser(
        "precheck", help="Verify every manifest target against the library before rendering."
    )
    precheck.add_argument(
        "--library",
        help="Path to the .enl library. Defaults to library.path from the manifest.",
    )
    precheck.add_argument("--manifest", required=True, help="Path to a schema v1 manifest.")

    export = subparsers.add_parser("export-ris", help="Write a RIS file for import into EndNote.")
    export.add_argument("--output", required=True, help="Destination .ris path.")
    export.add_argument("--references", required=True, help="Path to a JSON array of reference objects.")

    return parser


def _dispatch(args: argparse.Namespace) -> dict[str, Any]:
    if args.command == "search":
        return search_references(
            library=args.library,
            query=args.query,
            author=args.author,
            year=args.year,
            title=args.title,
            doi=args.doi,
            limit=args.limit,
        )
    if args.command == "get":
        return get_reference(library=args.library, rec_number=args.rec_number)
    if args.command == "precheck":
        return precheck_manifest(library=args.library, manifest=args.manifest)
    if args.command == "export-ris":
        return export_ris(output=args.output, references=args.references)
    raise ValueError(f"Unknown command: {args.command}")


def main(argv: Sequence[str] | None = None) -> int:
    _configure_logging()
    _force_utf8_streams()
    args = _build_parser().parse_args(argv)
    try:
        result = _dispatch(args)
    except Exception as exc:  # CLI boundary: report as JSON instead of a traceback on stdout.
        LOGGER.exception("unhandled error while running %s", args.command)
        result = {"ok": False, "action": args.command, "error": f"Unhandled error: {exc}"}
    sys.stdout.write(json.dumps(result, ensure_ascii=False) + "\n")
    sys.stdout.flush()
    return EXIT_OK if result.get("ok") else EXIT_FAILURE


if __name__ == "__main__":
    raise SystemExit(main())
