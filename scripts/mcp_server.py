#!/usr/bin/env python3
"""Dependency-free stdio MCP server for EndNote 21 Cite While You Write automation in Word.

Two backends sit behind the tool table:

* The PowerShell bridge ``endnote-word.ps1`` drives Word and the CWYW add-in over COM.
  It is spawned as a child process and answers with exactly one compact JSON line.
* ``endnote_library`` is imported in-process for the read-only library tools and the RIS
  export, because those never need Word and must stay fast.

Everything here is standard library only so the plugin runs on a bare CPython install.
"""

from __future__ import annotations

import json
import os
import re
import shutil
import subprocess
import sys
import traceback
from pathlib import Path
from typing import Any, BinaryIO


SERVER_NAME = "endnote-for-word"
SERVER_VERSION = "1.2.0"
SCRIPT_PATH = Path(__file__).with_name("endnote-word.ps1")
SCRIPT_DIRECTORY = SCRIPT_PATH.parent
DEFAULT_PROTOCOL_VERSION = "2025-06-18"

# PowerShell selection. PowerShell 7 is preferred, Windows PowerShell 5.1 is accepted;
# the bridge is written to parse and run under both.
POWERSHELL_ENV_VARIABLE = "ENDNOTE_WORD_POWERSHELL"
POWERSHELL_CANDIDATES = ("pwsh.exe", "powershell.exe")
POWERSHELL_INSTALL_URL = (
    "https://learn.microsoft.com/powershell/scripting/install/installing-powershell-on-windows"
)

# Bridge process budget. EndNote can take a long time to open a large library before the
# first ENFormatBibliography call returns, hence the generous default.
TIMEOUT_ENV_VARIABLE = "ENDNOTE_WORD_TIMEOUT_SECONDS"
PYTHON_ENV_VARIABLE = "ENDNOTE_WORD_PYTHON"
DEFAULT_TIMEOUT_SECONDS = 240
DEFAULT_RENDER_BASE_TIMEOUT_SECONDS = 60
DEFAULT_RENDER_SECONDS_PER_CITATION = 8
MAX_RENDER_TIMEOUT_SECONDS = 1800

# Processes the timeout watchdog is allowed to clean up. Only PIDs that appeared during the
# call are ever terminated, so a user's own EndNote session is never killed.
WATCHED_PROCESS_NAMES = ("WINWORD.EXE", "EndNote.exe")

# Win32 constants used by the ctypes process enumeration.
PROCESS_QUERY_LIMITED_INFORMATION = 0x1000
MAX_ENUMERATED_PROCESSES = 4096
MAX_IMAGE_PATH_CHARACTERS = 32768

# Library search bounds. The bridge and the library reader both cap results so a broad
# query can never dump a whole library into the model context.
DEFAULT_SEARCH_LIMIT = 25
MAX_SEARCH_LIMIT = 200
SEARCH_FILTER_KEYS = ("query", "author", "year", "title", "doi", "limit")

STDERR_EXCERPT_CHARACTERS = 4000

# EndNote reference fields accepted by the RIS exporter (see references/library-import.md).
RIS_REFERENCE_PROPERTIES: dict[str, Any] = {
    "type": {"type": "string", "description": "RIS reference type, for example JOUR, BOOK, CHAP, CONF."},
    "authors": {"type": "array", "items": {"type": "string"}, "description": "Authors in 'Last, First' form."},
    "year": {"type": "string"},
    "title": {"type": "string"},
    "journal": {"type": "string"},
    "volume": {"type": "string"},
    "issue": {"type": "string"},
    "pages": {"type": "string"},
    "doi": {"type": "string"},
    "url": {"type": "string"},
    "abstract": {"type": "string"},
    "keywords": {"type": "array", "items": {"type": "string"}},
    "publisher": {"type": "string"},
}


# Tools answered by spawning the PowerShell bridge, mapped to its -Action value.
BRIDGE_TOOL_ACTIONS = {
    "probe_endnote_word": "probe",
    "list_endnote_styles": "list-styles",
    "configure_endnote_defaults": "configure-defaults",
    "render_endnote_word_document": "render",
    "validate_endnote_word_document": "validate",
    "update_endnote_word_fields": "update",
    "set_endnote_citation_style": "set-style",
    "unformat_endnote_word_document": "unformat",
    "import_reference_file": "import-ris",
}
# There is deliberately no remove-field-codes action. ENRemoveFieldCodes was measured to raise
# a modal confirmation dialog that never returns under automation, so converting a document to
# plain text stays a manual EndNote ribbon operation.

# Tools answered in-process by endnote_library, mapped to their reported action name.
LIBRARY_TOOL_ACTIONS = {
    "search_endnote_library": "search",
    "get_endnote_reference": "get",
    "precheck_endnote_manifest": "precheck",
    "export_reference_import_file": "export-ris",
}


TOOLS: list[dict[str, Any]] = [
    {
        "name": "probe_endnote_word",
        "description": (
            "Check Windows, Microsoft Word COM, the EndNote 21 executable, the Cite While You Write "
            "add-in registration, the CWYW global template and whether its macros are callable. "
            "Pass library_path to also confirm the .enl library and its .Data directory. Read-only; "
            "run this before any render."
        ),
        "inputSchema": {
            "type": "object",
            "properties": {
                "library_path": {
                    "type": "string",
                    "description": "Optional absolute path to an EndNote .enl library to check.",
                }
            },
            "required": [],
            "additionalProperties": False,
        },
        "annotations": {
            "title": "Probe EndNote and Word",
            "readOnlyHint": True,
            "destructiveHint": False,
            "idempotentHint": True,
            "openWorldHint": False,
        },
    },
    {
        "name": "list_endnote_styles",
        "description": (
            "List installed EndNote output styles (.ens) from the program styles folder and the "
            "personal styles folder. Pass style to filter by substring. Always confirm a style name "
            "here before rendering: an unknown style makes EndNote open a modal dialog that hangs "
            "the automation run."
        ),
        "inputSchema": {
            "type": "object",
            "properties": {
                "style": {
                    "type": "string",
                    "description": "Optional case-insensitive substring filter, for example 'IEEE' or 'APA'.",
                }
            },
            "required": [],
            "additionalProperties": False,
        },
        "annotations": {
            "title": "List EndNote Output Styles",
            "readOnlyHint": True,
            "destructiveHint": False,
            "idempotentHint": True,
            "openWorldHint": False,
        },
    },
    {
        "name": "search_endnote_library",
        "description": (
            "Search an EndNote library read-only and return matching references with their "
            "rec_number, authors, year, title, journal, pages, bare DOI and canonical DOI URL. Accepts a .enl library "
            "(read through a temporary copy of its SQLite store) or an EndNote XML export. Every "
            "rec_number used in a manifest must come from this tool."
        ),
        "inputSchema": {
            "type": "object",
            "properties": {
                "library_path": {
                    "type": "string",
                    "description": "Absolute path to the .enl library or an EndNote .xml export.",
                },
                "query": {"type": "string", "description": "Free text matched across author, title, journal, DOI, keywords and abstract."},
                "author": {"type": "string", "description": "Author substring, matched case-insensitively."},
                "year": {"type": "integer", "description": "Exact publication year."},
                "title": {"type": "string", "description": "Title substring, matched case-insensitively."},
                "doi": {"type": "string", "description": "DOI substring, matched case-insensitively."},
                "limit": {
                    "type": "integer",
                    "minimum": 1,
                    "maximum": MAX_SEARCH_LIMIT,
                    "default": DEFAULT_SEARCH_LIMIT,
                    "description": "Maximum number of references to return.",
                },
            },
            "required": ["library_path"],
            "additionalProperties": False,
        },
        "annotations": {
            "title": "Search EndNote Library",
            "readOnlyHint": True,
            "destructiveHint": False,
            "idempotentHint": True,
            "openWorldHint": False,
        },
    },
    {
        "name": "get_endnote_reference",
        "description": (
            "Fetch one EndNote reference by its rec_number (the library record id used inside "
            "temporary citations such as {Thomsen, 2023 #1}). Read-only."
        ),
        "inputSchema": {
            "type": "object",
            "properties": {
                "library_path": {
                    "type": "string",
                    "description": "Absolute path to the .enl library or an EndNote .xml export.",
                },
                "rec_number": {
                    "type": "integer",
                    "minimum": 1,
                    "description": "EndNote record id, the number after '#' in a temporary citation.",
                },
            },
            "required": ["library_path", "rec_number"],
            "additionalProperties": False,
        },
        "annotations": {
            "title": "Get EndNote Reference",
            "readOnlyHint": True,
            "destructiveHint": False,
            "idempotentHint": True,
            "openWorldHint": False,
        },
    },
    {
        "name": "precheck_endnote_manifest",
        "description": (
            "Verify every citation target in a manifest against the library before Word is opened: "
            "the record must exist, must not be in the trash, and its first-author surname and year "
            "must match. Every target must have a matching bare DOI, and every DOI must resolve through "
            "the official doi.org Proxy REST API. A missing record makes EndNote open a blocking "
            "'select matching reference' dialog, so this check is mandatory before render. Read-only; "
            "reports every problem, not just the first."
        ),
        "inputSchema": {
            "type": "object",
            "properties": {
                "library_path": {
                    "type": "string",
                    "description": "Absolute path to the .enl library or an EndNote .xml export.",
                },
                "manifest_path": {"type": "string", "description": "Absolute path to a schema v1 JSON manifest."},
            },
            "required": ["library_path", "manifest_path"],
            "additionalProperties": False,
        },
        "annotations": {
            "title": "Precheck EndNote Manifest",
            "readOnlyHint": True,
            "destructiveHint": False,
            "idempotentHint": True,
            "openWorldHint": True,
        },
    },
    {
        "name": "configure_endnote_defaults",
        "description": (
            "Persist this plugin's defaults (default output style, timeouts and bibliography "
            "expectations) to the per-user configuration directory and report the effective values. "
            "EndNote's own preferences and UI state are left untouched."
        ),
        "inputSchema": {"type": "object", "properties": {}, "required": [], "additionalProperties": False},
        "annotations": {
            "title": "Configure EndNote for Word Defaults",
            "readOnlyHint": False,
            "destructiveHint": False,
            "idempotentHint": True,
            "openWorldHint": False,
        },
    },
    {
        "name": "render_endnote_word_document",
        "description": (
            "Replace {{CITE:id}} manifest markers in a DOCX with EndNote temporary citations, then "
            "run the CWYW macro ENFormatBibliography so EndNote produces genuine ADDIN EN.CITE "
            "fields with a traveling library and an ADDIN EN.REFLIST bibliography. Per-citation "
            "display supports 'default' for [x] and 'author-year' for Author et al. [x] under IEEE. "
            "The bridge repeats the DOI precheck, and Word plus any EndNote instance it starts remain "
            "hidden. Writes a new output file; pass overwrite=true only to replace an existing output. "
            "Run precheck_endnote_manifest and list_endnote_styles first."
        ),
        "inputSchema": {
            "type": "object",
            "properties": {
                "input_path": {"type": "string", "description": "Absolute path to the source DOCX."},
                "output_path": {"type": "string", "description": "Absolute path for the rendered DOCX."},
                "manifest_path": {"type": "string", "description": "Absolute path to a schema v1 JSON manifest."},
                "overwrite": {"type": "boolean", "default": False},
            },
            "required": ["input_path", "output_path", "manifest_path"],
            "additionalProperties": False,
        },
        "annotations": {
            "title": "Render EndNote Word Document",
            "readOnlyHint": False,
            "destructiveHint": True,
            "idempotentHint": False,
            "openWorldHint": True,
        },
    },
    {
        "name": "validate_endnote_word_document",
        "description": (
            "Open a DOCX read-only and verify the citation gates: no leftover {{CITE:}} markers or "
            "temporary citation text, the expected number of ADDIN EN.CITE fields, a non-empty "
            "ADDIN EN.REFLIST bibliography with every visible canonical DOI link, no doubled DOI "
            "resolver, every requested Author (Year) display, every manifest rec_number present "
            "in the traveling library XML, and the EN.Layout and EN.Libraries document variables. Supply "
            "manifest_path to enable the manifest-dependent gates."
        ),
        "inputSchema": {
            "type": "object",
            "properties": {
                "document_path": {"type": "string", "description": "Absolute path to the DOCX."},
                "manifest_path": {
                    "type": "string",
                    "description": "Optional absolute path to the manifest that produced the document.",
                },
            },
            "required": ["document_path"],
            "additionalProperties": False,
        },
        "annotations": {
            "title": "Validate EndNote Word Document",
            "readOnlyHint": True,
            "destructiveHint": False,
            "idempotentHint": True,
            "openWorldHint": False,
        },
    },
    {
        "name": "update_endnote_word_fields",
        "description": (
            "Re-run ENFormatBibliography over an existing document so all EndNote citations and the "
            "bibliography are reformatted from their traveling library. Write to a new output path, "
            "or pass overwrite=true to update the input in place."
        ),
        "inputSchema": {
            "type": "object",
            "properties": {
                "input_path": {"type": "string", "description": "Absolute path to the DOCX to update."},
                "output_path": {"type": "string", "description": "Absolute path for the updated DOCX."},
                "overwrite": {"type": "boolean", "default": False},
            },
            "required": ["input_path"],
            "additionalProperties": False,
        },
        "annotations": {
            "title": "Update EndNote Citations and Bibliography",
            "readOnlyHint": False,
            "destructiveHint": True,
            "idempotentHint": True,
            "openWorldHint": False,
        },
    },
    {
        "name": "set_endnote_citation_style",
        "description": (
            "Switch a document to another EndNote output style, for example from APA 7th to IEEE. "
            "Patches the EN.Layout document variable and re-runs ENFormatBibliography, which "
            "rewrites both the in-text citations and the whole bibliography. The style name must "
            "exist in list_endnote_styles."
        ),
        "inputSchema": {
            "type": "object",
            "properties": {
                "input_path": {"type": "string", "description": "Absolute path to the DOCX."},
                "output_path": {"type": "string", "description": "Absolute path for the restyled DOCX."},
                "style": {"type": "string", "description": "Installed EndNote output style name, for example 'IEEE'."},
                "overwrite": {"type": "boolean", "default": False},
            },
            "required": ["input_path", "style"],
            "additionalProperties": False,
        },
        "annotations": {
            "title": "Set EndNote Citation Style",
            "readOnlyHint": False,
            "destructiveHint": True,
            "idempotentHint": True,
            "openWorldHint": False,
        },
    },
    {
        "name": "unformat_endnote_word_document",
        "description": (
            "Run ENUnformatCitations to revert formatted citations back to editable temporary "
            "citations such as {Thomsen, 2023 #1} and remove the bibliography. The traveling "
            "library is preserved, so the document can be reformatted later. Always writes a new "
            "file, never in place."
        ),
        "inputSchema": {
            "type": "object",
            "properties": {
                "input_path": {"type": "string", "description": "Absolute path to the formatted DOCX."},
                "output_path": {"type": "string", "description": "Absolute path for the unformatted DOCX."},
                "overwrite": {"type": "boolean", "default": False},
            },
            "required": ["input_path", "output_path"],
            "additionalProperties": False,
        },
        "annotations": {
            "title": "Unformat EndNote Citations",
            "readOnlyHint": False,
            "destructiveHint": True,
            "idempotentHint": False,
            "openWorldHint": False,
        },
    },
    {
        "name": "export_reference_import_file",
        "description": (
            "Write a UTF-8 RIS file from reference objects. DOI is required and is normalised to "
            "one bare DO tag even when the input contains one or more resolver prefixes. Use "
            "import_reference_file next when the records should be imported into a library."
        ),
        "inputSchema": {
            "type": "object",
            "properties": {
                "output_path": {"type": "string", "description": "Absolute path for the .ris file to write."},
                "references": {
                    "type": "array",
                    "minItems": 1,
                    "description": "Reference objects to serialise, one RIS record each.",
                    "items": {
                        "type": "object",
                        "properties": RIS_REFERENCE_PROPERTIES,
                        "required": ["title", "doi"],
                        "additionalProperties": False,
                    },
                },
            },
            "required": ["output_path", "references"],
            "additionalProperties": False,
        },
        "annotations": {
            "title": "Export Reference Import File",
            "readOnlyHint": False,
            "destructiveHint": False,
            "idempotentHint": True,
            "openWorldHint": False,
        },
    },
    {
        "name": "import_reference_file",
        "description": (
            "Import a RIS file into the exact EndNote library supplied by library_path through "
            "EndNote's native Direct Export route, then report the before/after live-reference "
            "counts. If library_path does not exist, empty_library_template_path may name a "
            "closed EndNote library with zero stored references, including trash; EndNote must "
            "be fully closed while its .enl and .Data pair is copied "
            "to create the target before import. Never guesses rec_numbers and never writes "
            "sdb.eni directly; search the library again after this call."
        ),
        "inputSchema": {
            "type": "object",
            "properties": {
                "library_path": {
                    "type": "string",
                    "description": "Absolute target .enl path. It may be absent only when an empty template is supplied.",
                },
                "ris_path": {"type": "string", "description": "Absolute path to a non-empty UTF-8 .ris file."},
                "empty_library_template_path": {
                    "type": "string",
                    "description": (
                        "Optional absolute path to a closed, empty .enl library to clone when "
                        "library_path does not yet exist. Ignored when the target exists."
                    ),
                },
            },
            "required": ["library_path", "ris_path"],
            "additionalProperties": False,
        },
        "annotations": {
            "title": "Import References into EndNote",
            "readOnlyHint": False,
            "destructiveHint": True,
            "idempotentHint": False,
            "openWorldHint": False,
        },
    },
]


def _log(message: str) -> None:
    print(f"[{SERVER_NAME}] {message}", file=sys.stderr, flush=True)


def _force_utf8_streams() -> None:
    """Force UTF-8 on stdout and stderr.

    Windows gives a piped stdout the legacy ANSI code page (cp950 on the development machine),
    and every JSON-RPC frame is serialised with ensure_ascii=False. Without this, one accented
    surname or CJK title in a search result raises UnicodeEncodeError inside _write_message and
    the response is lost. Reading is unaffected: _read_message decodes sys.stdin.buffer itself.
    """
    for stream in (sys.stdout, sys.stderr):
        reconfigure = getattr(stream, "reconfigure", None)
        if reconfigure is None:
            continue
        try:
            reconfigure(encoding="utf-8", errors="replace")
        except (OSError, ValueError) as exc:
            print(f"[{SERVER_NAME}] could not force UTF-8 on a stream: {exc}", file=sys.stderr)


def _powershell_executable() -> str:
    """Return the PowerShell host to run the bridge with.

    Preference order is the explicit override, then PowerShell 7, then Windows PowerShell 5.1.
    The bridge itself is 5.1 compatible, so accepting the older host is safe.
    """
    override = os.environ.get(POWERSHELL_ENV_VARIABLE, "").strip()
    if override:
        return override
    for candidate in POWERSHELL_CANDIDATES:
        resolved = shutil.which(candidate)
        if resolved:
            return resolved
    raise RuntimeError(
        "No PowerShell host was found. This plugin needs PowerShell 7 (pwsh.exe) or "
        "Windows PowerShell 5.1 (powershell.exe) on PATH. Install PowerShell 7 from "
        f"{POWERSHELL_INSTALL_URL}, or set {POWERSHELL_ENV_VARIABLE} to the full path of a "
        "PowerShell executable."
    )


def _default_timeout_seconds(action: str, arguments: dict[str, Any]) -> int:
    """Scale the default render budget for manifests with many citations."""
    if action != "render":
        return DEFAULT_TIMEOUT_SECONDS
    manifest_path = arguments.get("manifest_path")
    if not isinstance(manifest_path, str) or not manifest_path.strip():
        return DEFAULT_TIMEOUT_SECONDS
    try:
        manifest = json.loads(Path(manifest_path).read_text(encoding="utf-8"))
        citations = manifest.get("citations") if isinstance(manifest, dict) else None
        citation_count = len(citations) if isinstance(citations, list) else 0
    except (OSError, UnicodeError, json.JSONDecodeError):
        return DEFAULT_TIMEOUT_SECONDS
    estimate = DEFAULT_RENDER_BASE_TIMEOUT_SECONDS + (
        DEFAULT_RENDER_SECONDS_PER_CITATION * citation_count
    )
    return min(MAX_RENDER_TIMEOUT_SECONDS, max(DEFAULT_TIMEOUT_SECONDS, estimate))


def _timeout_seconds(action: str, arguments: dict[str, Any]) -> int:
    """Read an explicit budget or calculate a citation-aware default."""
    raw = os.environ.get(TIMEOUT_ENV_VARIABLE, "").strip()
    if not raw:
        return _default_timeout_seconds(action, arguments)
    try:
        value = int(raw)
    except ValueError:
        fallback = _default_timeout_seconds(action, arguments)
        _log(f"{TIMEOUT_ENV_VARIABLE}={raw!r} is not an integer; using {fallback} seconds.")
        return fallback
    if value <= 0:
        fallback = _default_timeout_seconds(action, arguments)
        _log(f"{TIMEOUT_ENV_VARIABLE}={value} is not positive; using {fallback} seconds.")
        return fallback
    return value


def _extract_word_pid(stderr: str) -> int | None:
    """Pick up the WORD_PID marker the bridge writes to stderr when it starts Word."""
    matches = re.findall(r"\bWORD_PID=(\d+)\b", stderr)
    return int(matches[-1]) if matches else None


def _process_pids(executable_name: str) -> set[int]:
    """Return the PIDs of running processes whose image name matches executable_name."""
    if os.name != "nt":
        return set()
    try:
        import ctypes
        from ctypes import wintypes

        process_ids = (wintypes.DWORD * MAX_ENUMERATED_PROCESSES)()
        bytes_returned = wintypes.DWORD()
        if not ctypes.windll.psapi.EnumProcesses(
            ctypes.byref(process_ids), ctypes.sizeof(process_ids), ctypes.byref(bytes_returned)
        ):
            return set()
        count = bytes_returned.value // ctypes.sizeof(wintypes.DWORD)
        matches: set[int] = set()
        for process_id in process_ids[:count]:
            handle = ctypes.windll.kernel32.OpenProcess(
                PROCESS_QUERY_LIMITED_INFORMATION, False, process_id
            )
            if not handle:
                continue
            try:
                capacity = wintypes.DWORD(MAX_IMAGE_PATH_CHARACTERS)
                buffer = ctypes.create_unicode_buffer(capacity.value)
                if ctypes.windll.kernel32.QueryFullProcessImageNameW(
                    handle, 0, buffer, ctypes.byref(capacity)
                ) and Path(buffer.value).name.upper() == executable_name.upper():
                    matches.add(int(process_id))
            finally:
                ctypes.windll.kernel32.CloseHandle(handle)
        return matches
    except (AttributeError, OSError, ValueError) as exc:
        _log(f"could not enumerate {executable_name} processes: {exc}")
        return set()


def _snapshot_watched_pids() -> dict[str, set[int]]:
    """Record the PIDs of Word and EndNote that already existed before a bridge call."""
    return {name: _process_pids(name) for name in WATCHED_PROCESS_NAMES}


def _terminate_process_pid(process_id: int) -> bool:
    try:
        completed = subprocess.run(
            ["taskkill.exe", "/PID", str(process_id), "/T", "/F"],
            stdin=subprocess.DEVNULL,
            stdout=subprocess.PIPE,
            stderr=subprocess.PIPE,
            text=True,
            encoding="utf-8",
            errors="replace",
            check=False,
            creationflags=getattr(subprocess, "CREATE_NO_WINDOW", 0),
        )
    except OSError as exc:
        _log(f"taskkill failed for PID {process_id}: {exc}")
        return False
    return completed.returncode == 0


def _reap_orphans(before: dict[str, set[int]]) -> dict[str, list[int]]:
    """Kill only the Word and EndNote processes that appeared during the timed-out call.

    A pre-existing EndNote session belongs to the user and is never touched, which is why the
    snapshot difference, and not a blanket taskkill by image name, drives the cleanup.
    """
    observed: dict[str, list[int]] = {}
    terminated: list[int] = []
    for name in WATCHED_PROCESS_NAMES:
        new_pids = sorted(_process_pids(name) - before.get(name, set()))
        observed[name] = new_pids
        for process_id in new_pids:
            if _terminate_process_pid(process_id):
                terminated.append(process_id)
    observed["terminated"] = sorted(terminated)
    return observed


def _build_command(action: str, arguments: dict[str, Any], executable: str) -> list[str]:
    command = [
        executable,
        "-NoLogo",
        "-NoProfile",
        "-NonInteractive",
        "-ExecutionPolicy",
        "Bypass",
        "-File",
        str(SCRIPT_PATH),
        "-Action",
        action,
    ]
    mapping = {
        "input_path": "-InputPath",
        "document_path": "-InputPath",
        "output_path": "-OutputPath",
        "manifest_path": "-ManifestPath",
        "library_path": "-LibraryPath",
        "style": "-Style",
        "ris_path": "-RisPath",
        "empty_library_template_path": "-EmptyLibraryTemplatePath",
    }
    for key, switch in mapping.items():
        value = arguments.get(key)
        if value is not None and value != "":
            command.extend([switch, str(value)])
    if arguments.get("overwrite"):
        command.append("-Overwrite")
    return command


def _invoke_bridge(action: str, arguments: dict[str, Any], timeout: int | None = None) -> dict[str, Any]:
    """Run one bridge action and parse its single JSON result line."""
    if not SCRIPT_PATH.is_file():
        return {
            "ok": False,
            "action": action,
            "error": f"Bridge script not found: {SCRIPT_PATH}",
        }
    try:
        executable = _powershell_executable()
    except RuntimeError as exc:
        return {"ok": False, "action": action, "error": str(exc)}

    command = _build_command(action, arguments, executable)
    effective_timeout = timeout or _timeout_seconds(action, arguments)
    before = _snapshot_watched_pids()
    try:
        completed = subprocess.run(
            command,
            stdin=subprocess.DEVNULL,
            stdout=subprocess.PIPE,
            stderr=subprocess.PIPE,
            text=True,
            encoding="utf-8",
            errors="replace",
            timeout=effective_timeout,
            check=False,
            env={**os.environ, PYTHON_ENV_VARIABLE: sys.executable},
            creationflags=getattr(subprocess, "CREATE_NO_WINDOW", 0),
        )
    except subprocess.TimeoutExpired as exc:
        stderr = exc.stderr or ""
        if isinstance(stderr, bytes):
            stderr = stderr.decode("utf-8", errors="replace")
        word_process_id = _extract_word_pid(stderr)
        observed = _reap_orphans(before)
        new_word_pids = observed.get("WINWORD.EXE", [])
        if word_process_id is None and len(new_word_pids) == 1:
            word_process_id = new_word_pids[0]
        terminated = observed.get("terminated", [])
        return {
            "ok": False,
            "action": action,
            "error": (
                f"EndNote/Word bridge timed out after {effective_timeout} seconds. A modal EndNote "
                "or Word dialog is the usual cause; see references/troubleshooting.md. Raise "
                f"{TIMEOUT_ENV_VARIABLE} for very large libraries."
            ),
            "isolated_word_pid": word_process_id,
            "isolated_word_process_terminated": bool(word_process_id and word_process_id in terminated),
            "new_word_pids_observed": new_word_pids,
            "new_endnote_pids_observed": observed.get("EndNote.exe", []),
            "isolated_processes_terminated": terminated,
            "stderr_tail": stderr[-STDERR_EXCERPT_CHARACTERS:],
        }
    except OSError as exc:
        return {
            "ok": False,
            "action": action,
            "error": f"Could not start the PowerShell bridge with {executable}: {exc}",
        }

    if completed.stderr.strip():
        _log(completed.stderr.strip())
    lines = [line.strip() for line in completed.stdout.splitlines() if line.strip()]
    if not lines:
        return {
            "ok": False,
            "action": action,
            "error": f"Bridge returned no JSON (exit {completed.returncode}).",
            "stderr_tail": completed.stderr[-STDERR_EXCERPT_CHARACTERS:],
        }
    try:
        result = json.loads(lines[-1])
    except json.JSONDecodeError as exc:
        return {
            "ok": False,
            "action": action,
            "error": f"Bridge returned invalid JSON: {exc}",
            "stdout": completed.stdout[-STDERR_EXCERPT_CHARACTERS:],
        }
    if not isinstance(result, dict):
        return {
            "ok": False,
            "action": action,
            "error": "Bridge returned a JSON value that is not an object.",
            "stdout": completed.stdout[-STDERR_EXCERPT_CHARACTERS:],
        }
    if completed.returncode and result.get("ok", True):
        result["ok"] = False
        result["error"] = f"Bridge exited with code {completed.returncode}."
    return result


def _import_library_module() -> Any:
    """Import endnote_library from the scripts directory next to this server."""
    directory = str(SCRIPT_DIRECTORY)
    if directory not in sys.path:
        sys.path.insert(0, directory)
    import endnote_library

    return endnote_library


def _require(arguments: dict[str, Any], key: str, tool: str) -> Any:
    value = arguments.get(key)
    if value is None or (isinstance(value, str) and not value.strip()):
        raise ValueError(f"Tool {tool} requires a non-empty '{key}' argument.")
    return value


def _invoke_library(name: str, arguments: dict[str, Any]) -> dict[str, Any]:
    """Answer a library tool in-process; no Word and no subprocess are involved."""
    action = LIBRARY_TOOL_ACTIONS[name]
    try:
        library = _import_library_module()
        if name == "search_endnote_library":
            filters = {
                key: arguments[key]
                for key in SEARCH_FILTER_KEYS
                if arguments.get(key) is not None and arguments.get(key) != ""
            }
            result = library.search_references(_require(arguments, "library_path", name), **filters)
        elif name == "get_endnote_reference":
            result = library.get_reference(
                _require(arguments, "library_path", name), _require(arguments, "rec_number", name)
            )
        elif name == "precheck_endnote_manifest":
            result = library.precheck_manifest(
                _require(arguments, "library_path", name), _require(arguments, "manifest_path", name)
            )
        elif name == "export_reference_import_file":
            result = library.export_ris(
                _require(arguments, "output_path", name), _require(arguments, "references", name)
            )
        else:
            raise ValueError(f"Unrouted library tool: {name}")
    except Exception as exc:  # Tool boundary: report the failure instead of dropping the call.
        _log(traceback.format_exc())
        return {"ok": False, "action": action, "error": f"{type(exc).__name__}: {exc}"}
    if not isinstance(result, dict):
        return {
            "ok": False,
            "action": action,
            "error": f"endnote_library.{action} returned {type(result).__name__}, expected an object.",
        }
    result.setdefault("action", action)
    return result


def _call_tool(name: str, arguments: dict[str, Any]) -> dict[str, Any]:
    if name in BRIDGE_TOOL_ACTIONS:
        result = _invoke_bridge(BRIDGE_TOOL_ACTIONS[name], arguments)
    elif name in LIBRARY_TOOL_ACTIONS:
        result = _invoke_library(name, arguments)
    else:
        raise ValueError(f"Unknown tool: {name}")
    rendered = json.dumps(result, ensure_ascii=False, indent=2)
    return {
        "content": [{"type": "text", "text": rendered}],
        "structuredContent": result,
        "isError": not bool(result.get("ok")),
    }


def _read_message(stream: BinaryIO) -> dict[str, Any] | None:
    line = stream.readline()
    if not line:
        return None
    while line in (b"\r\n", b"\n"):
        line = stream.readline()
        if not line:
            return None
    if line.lower().startswith(b"content-length:"):
        length = int(line.split(b":", 1)[1].strip())
        while True:
            header = stream.readline()
            if header in (b"\r\n", b"\n", b""):
                break
        payload = stream.read(length)
        return json.loads(payload.decode("utf-8"))
    return json.loads(line.decode("utf-8"))


def _write_message(message: dict[str, Any]) -> None:
    data = json.dumps(message, ensure_ascii=False, separators=(",", ":"))
    sys.stdout.write(data + "\n")
    sys.stdout.flush()


def _success(request_id: Any, result: Any) -> dict[str, Any]:
    return {"jsonrpc": "2.0", "id": request_id, "result": result}


def _error(request_id: Any, code: int, message: str, data: Any = None) -> dict[str, Any]:
    error: dict[str, Any] = {"code": code, "message": message}
    if data is not None:
        error["data"] = data
    return {"jsonrpc": "2.0", "id": request_id, "error": error}


def _handle(message: dict[str, Any]) -> dict[str, Any] | None:
    method = message.get("method")
    request_id = message.get("id")
    params = message.get("params") or {}
    if method in ("notifications/initialized", "notifications/cancelled"):
        return None
    if method == "initialize":
        requested = params.get("protocolVersion") or DEFAULT_PROTOCOL_VERSION
        return _success(
            request_id,
            {
                "protocolVersion": requested,
                "capabilities": {"tools": {"listChanged": False}},
                "serverInfo": {"name": SERVER_NAME, "version": SERVER_VERSION},
                "instructions": (
                    "Probe first, then resolve every rec_number with search_endnote_library and "
                    "confirm the output style with list_endnote_styles. Run "
                    "precheck_endnote_manifest before rendering. Every cited record must carry a bare "
                    "DOI that resolves at doi.org. Use IEEE by default. Set display='author-year' only "
                    "when the author is a grammatical part of the sentence; use display='default' for "
                    "claim support and merged citations. Render into a new file, then validate. Citations "
                    "must be CWYW-produced ADDIN EN.CITE fields with a traveling library and an ADDIN "
                    "EN.REFLIST bibliography; never type citation text by hand. Word and EndNote editing "
                    "runs silently. Library reads are read-only; explicit RIS import is performed by "
                    "EndNote itself against the exact requested library."
                ),
            },
        )
    if method == "ping":
        return _success(request_id, {})
    if method == "tools/list":
        return _success(request_id, {"tools": TOOLS})
    if method == "tools/call":
        return _success(request_id, _call_tool(str(params.get("name", "")), params.get("arguments") or {}))
    if request_id is None:
        return None
    return _error(request_id, -32601, f"Method not found: {method}")


def main() -> int:
    _force_utf8_streams()
    _log(f"starting {SERVER_VERSION}")
    while True:
        message: dict[str, Any] | None = None
        try:
            message = _read_message(sys.stdin.buffer)
            if message is None:
                return 0
            response = _handle(message)
            if response is not None:
                _write_message(response)
        except json.JSONDecodeError as exc:
            _write_message(_error(None, -32700, "Parse error", str(exc)))
        except Exception as exc:  # MCP boundary: convert failures to JSON-RPC errors.
            _log(traceback.format_exc())
            request_id = message.get("id") if isinstance(message, dict) else None
            _write_message(_error(request_id, -32603, str(exc)))


if __name__ == "__main__":
    raise SystemExit(main())
