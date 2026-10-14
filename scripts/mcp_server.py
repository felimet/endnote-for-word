#!/usr/bin/env python3
"""Dependency-free stdio MCP server for EndNote 21 Cite While You Write automation in Word.

Two backends sit behind the tool table:

* The PowerShell bridge ``endnote-word.ps1`` drives Word and the CWYW add-in over COM.
  It is spawned as a child process and answers with exactly one compact JSON line.
* ``endnote_library`` is imported in-process for the read-only library tools and the RIS
  export, because those never need Word and must stay fast.

Every bridge action that reaches EndNote runs as a job on one worker thread, because the user's
single EndNote process serves every CWYW call. Render, update and set-style return a job id at
once by default; the worker opens the document's library, watches for dialogs while the bridge
runs, closes what the job opened, and closes an EndNote it started once the queue has been idle.

Everything here is standard library only so the plugin runs on a bare CPython install.
"""

from __future__ import annotations

import json
import os
import queue
import re
import shutil
import subprocess
import sys
import tempfile
import threading
import time
import traceback
import uuid
from dataclasses import dataclass, field
from pathlib import Path
from typing import Any, BinaryIO


SERVER_NAME = "endnote-for-word"
SERVER_VERSION = "1.5.0"
# Absolute: main() may leave the folder the server was started in.
SCRIPT_PATH = Path(__file__).resolve().with_name("endnote-word.ps1")
SCRIPT_DIRECTORY = SCRIPT_PATH.parent
PLUGIN_ROOT = SCRIPT_DIRECTORY.parent
DEFAULT_PROTOCOL_VERSION = "2025-06-18"

# PowerShell selection. PowerShell 7 is preferred, Windows PowerShell 5.1 is accepted;
# the bridge is written to parse and run under both.
POWERSHELL_ENV_VARIABLE = "ENDNOTE_WORD_POWERSHELL"
POWERSHELL_7_EXECUTABLE = "pwsh.exe"
POWERSHELL_CANDIDATES = (POWERSHELL_7_EXECUTABLE, "powershell.exe")
# Where a PowerShell 7 can appear while the server runs, even if PATH was captured before:
# the per-user MSIX app execution alias that winget installs, then the MSI location.
POWERSHELL_7_KNOWN_LOCATIONS = (
    ("LOCALAPPDATA", ("Microsoft", "WindowsApps", POWERSHELL_7_EXECUTABLE)),
    ("ProgramFiles", ("PowerShell", "7", POWERSHELL_7_EXECUTABLE)),
)
POWERSHELL_INSTALL_URL = (
    "https://learn.microsoft.com/powershell/scripting/install/installing-powershell-on-windows"
)

# Bridge process budget. EndNote can take a long time to open a large library before the
# first ENFormatBibliography call returns, hence the generous default.
# The bridge also prints its PowerShell 7 notice to stderr behind this marker, so the notice
# survives a call that times out before the JSON result line is written.
POWERSHELL_NOTICE_MARKER = "POWERSHELL_NOTICE="
TIMEOUT_ENV_VARIABLE = "ENDNOTE_WORD_TIMEOUT_SECONDS"
PYTHON_ENV_VARIABLE = "ENDNOTE_WORD_PYTHON"
DEFAULT_TIMEOUT_SECONDS = 240
DEFAULT_RENDER_BASE_TIMEOUT_SECONDS = 60
DEFAULT_RENDER_SECONDS_PER_CITATION = 8
MAX_RENDER_TIMEOUT_SECONDS = 1800

# Processes the timeout watchdog is allowed to clean up. Only PIDs that appeared during the
# call are ever terminated, so a user's own EndNote session is never killed.
WATCHED_PROCESS_NAMES = ("WINWORD.EXE", "EndNote.exe")
WORD_PROCESS_NAME = "WINWORD.EXE"
ENDNOTE_PROCESS_NAME = "EndNote.exe"

# Jobs. Every action that reaches EndNote is queued, so two calls never drive CWYW at once; the
# named mutex extends that to every server of this Windows session, such as Claude Code and Codex.
QUEUED_ACTIONS = frozenset({"probe", "render", "update", "set-style", "unformat", "import-ris"})
BACKGROUND_TOOLS = frozenset({
    "render_endnote_word_document", "update_endnote_word_fields", "set_endnote_citation_style",
})
JOB_STATUS_TOOL = "get_endnote_job_status"
JOB_CANCEL_TOOL = "cancel_endnote_job"
BRIDGE_MUTEX_NAME = "Local\\EndNoteForWordAgentBridge"
TERMINAL_JOB_STATES = frozenset({"succeeded", "failed", "cancelled"})
MAX_STATUS_WAIT_SECONDS = 120
JOB_POLL_SECONDS = 1.0
LOCK_POLL_MILLISECONDS = 500
# A dialog is reported once it has stayed this long, so brief progress windows are not noise.
DIALOG_REPORT_AFTER_SECONDS = 3
# CWYW choosers that never close by themselves. Measured: with two open libraries whose record
# numbers overlap, CWYW cannot tell which record a temporary citation means and asks with Select
# Matching Reference. The job presses Cancel, which picks no record, and lets the bridge finish:
# ending Word at the chooser instead leaves EndNote unable to exit.
CANCELLED_DIALOG_TITLES = ("Select Matching Reference",)
CANCELLED_DIALOG_ADVICE = (
    "CWYW found more than one open record for a temporary citation, which happens when EndNote has "
    "another library open whose record numbers overlap the document's library. The job pressed "
    "Cancel, so no record was chosen. Close the other libraries in EndNote, or run the job while "
    "only the document's library is open, then retry."
)
# An EndNote this server started is closed when no job has run for this long.
IDLE_CLOSE_ENV_VARIABLE = "ENDNOTE_IDLE_CLOSE_SECONDS"
DEFAULT_IDLE_CLOSE_SECONDS = 120
READY_TIMEOUT_ENV_VARIABLE = "ENDNOTE_READY_TIMEOUT_MS"
DEFAULT_READY_TIMEOUT_MILLISECONDS = 90000
LIBRARY_CLOSE_WAIT_SECONDS = 10
ENDNOTE_EXIT_WAIT_SECONDS = 30

# Win32 job object: every bridge and every EndNote this server starts ends with the server.
JOB_OBJECT_EXTENDED_LIMIT_INFORMATION = 9
JOB_OBJECT_LIMIT_KILL_ON_JOB_CLOSE = 0x2000
WAIT_OBJECT_0 = 0
WAIT_ABANDONED = 0x80

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
            "Pass empty_library_template_path once to register the zero-record library that "
            "import_reference_file clones when its target library does not exist; a registered "
            "template is kept when this tool runs again without it. "
            "EndNote's own preferences and UI state are left untouched."
        ),
        "inputSchema": {
            "type": "object",
            "properties": {
                "empty_library_template_path": {
                    "type": "string",
                    "description": (
                        "Optional absolute path to an EndNote library with zero stored references, "
                        "including trash, created once in EndNote with File > New."
                    ),
                },
            },
            "required": [],
            "additionalProperties": False,
        },
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
                "background": {
                    "type": "boolean",
                    "default": True,
                    "description": (
                        "Queue the job and return its job_id at once; follow it with "
                        "get_endnote_job_status. false waits for the result in this call."
                    ),
                },
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
                "background": {
                    "type": "boolean",
                    "default": True,
                    "description": (
                        "Queue the job and return its job_id at once; follow it with "
                        "get_endnote_job_status. false waits for the result in this call."
                    ),
                },
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
                "background": {
                    "type": "boolean",
                    "default": True,
                    "description": (
                        "Queue the job and return its job_id at once; follow it with "
                        "get_endnote_job_status. false waits for the result in this call."
                    ),
                },
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
            "counts. When the user names no library, omit library_path and pass document_path: "
            "the target is <document stem>.enl beside the DOCX, and it is refused when it does "
            "not exist yet but the DOCX already carries EndNote fields. A missing target is "
            "created by copying empty_library_template_path, or the template registered with "
            "configure_endnote_defaults, which must hold zero stored references, including "
            "trash; EndNote must be fully closed while its .enl and .Data pair is copied. "
            "Never guesses rec_numbers and never writes sdb.eni directly; search the returned "
            "library_path again after this call."
        ),
        "inputSchema": {
            "type": "object",
            "properties": {
                "library_path": {
                    "type": "string",
                    "description": (
                        "Absolute target .enl path. It may be absent when an empty template is "
                        "supplied or registered. Omit it to use document_path's library."
                    ),
                },
                "document_path": {
                    "type": "string",
                    "description": (
                        "Absolute DOCX path whose library is <document stem>.enl beside it. "
                        "Used only when library_path is omitted."
                    ),
                },
                "ris_path": {"type": "string", "description": "Absolute path to a non-empty UTF-8 .ris file."},
                "empty_library_template_path": {
                    "type": "string",
                    "description": (
                        "Optional absolute path to a closed, empty .enl library to clone when "
                        "the target does not yet exist. Overrides the registered template; "
                        "ignored when the target exists."
                    ),
                },
            },
            "required": ["ris_path"],
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
    {
        "name": JOB_STATUS_TOOL,
        "description": (
            "Report a queued EndNote job: state (queued, waiting_for_lock, running, succeeded, "
            "failed, cancelled), the tool result once it finished, notices it dismissed, and any "
            "dialog that has stayed open, which a user may have to answer. Pass wait_seconds to "
            "block until the job finishes or that many seconds pass. Without job_id it lists this "
            "session's jobs."
        ),
        "inputSchema": {
            "type": "object",
            "properties": {
                "job_id": {"type": "string", "description": "The job_id a background call returned."},
                "wait_seconds": {
                    "type": "integer",
                    "minimum": 0,
                    "maximum": MAX_STATUS_WAIT_SECONDS,
                    "default": 0,
                    "description": "Wait up to this many seconds for the job to finish.",
                },
            },
            "required": [],
            "additionalProperties": False,
        },
        "annotations": {
            "title": "Get EndNote Job Status",
            "readOnlyHint": True,
            "destructiveHint": False,
            "idempotentHint": True,
            "openWorldHint": False,
        },
    },
    {
        "name": JOB_CANCEL_TOOL,
        "description": (
            "Cancel a queued or running EndNote job. A running job's bridge and hidden Word are "
            "ended and the library the job opened is closed; no output file is written, because "
            "the bridge saves atomically at the end."
        ),
        "inputSchema": {
            "type": "object",
            "properties": {"job_id": {"type": "string", "description": "The job to cancel."}},
            "required": ["job_id"],
            "additionalProperties": False,
        },
        "annotations": {
            "title": "Cancel EndNote Job",
            "readOnlyHint": False,
            "destructiveHint": False,
            "idempotentHint": True,
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


def _powershell_notice_from_stderr(stderr: str) -> dict[str, Any] | None:
    """Recover the bridge's PowerShell 7 notice from stderr when no result line arrived."""
    for line in reversed(stderr.splitlines()):
        if line.startswith(POWERSHELL_NOTICE_MARKER):
            try:
                notice = json.loads(line[len(POWERSHELL_NOTICE_MARKER):])
            except json.JSONDecodeError:
                _log(f"Ignoring a malformed PowerShell 7 notice on stderr: {line[:200]}")
                return None
            return notice if isinstance(notice, dict) else None
    return None


def _with_powershell_notice(result: dict[str, Any], stderr: str) -> dict[str, Any]:
    notice = _powershell_notice_from_stderr(stderr)
    if notice is not None:
        result["powershell_notice"] = notice
    return result


def _known_powershell_7_paths() -> list[Path]:
    """Return the fixed PowerShell 7 locations whose root variable is set."""
    paths = []
    for variable, parts in POWERSHELL_7_KNOWN_LOCATIONS:
        root = os.environ.get(variable, "").strip()
        if root:
            paths.append(Path(root, *parts))
    return paths


def _powershell_executable() -> str:
    """Return the PowerShell host to run the bridge with.

    Preference order is the explicit override, then PowerShell 7, then Windows PowerShell 5.1.
    The bridge itself is 5.1 compatible, so accepting the older host is safe. Resolution runs
    on every call, so a PowerShell 7 installed by the bridge's background install takes over
    from the next call without a server restart.
    """
    override = os.environ.get(POWERSHELL_ENV_VARIABLE, "").strip()
    if override:
        return override
    for candidate in POWERSHELL_CANDIDATES:
        resolved = shutil.which(candidate)
        if not resolved and candidate == POWERSHELL_7_EXECUTABLE:
            resolved = next(
                (str(path) for path in _known_powershell_7_paths() if path.is_file()), None
            )
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
        manifest = json.loads(Path(manifest_path).read_text(encoding="utf-8-sig"))
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
        return _with_powershell_notice({
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
        }, stderr)
    except OSError as exc:
        return {
            "ok": False,
            "action": action,
            "error": f"Could not start the PowerShell bridge with {executable}: {exc}",
        }

    return _parse_bridge_output(action, completed.returncode, completed.stdout, completed.stderr)


def _parse_bridge_output(action: str, returncode: int, stdout: str, stderr: str) -> dict[str, Any]:
    """Turn a finished bridge run into its result object: the last stdout line is the JSON."""
    if stderr.strip():
        _log(stderr.strip())
    lines = [line.strip() for line in stdout.splitlines() if line.strip()]
    if not lines:
        return _with_powershell_notice({
            "ok": False,
            "action": action,
            "error": f"Bridge returned no JSON (exit {returncode}).",
            "stderr_tail": stderr[-STDERR_EXCERPT_CHARACTERS:],
        }, stderr)
    try:
        result = json.loads(lines[-1])
    except json.JSONDecodeError as exc:
        return {
            "ok": False,
            "action": action,
            "error": f"Bridge returned invalid JSON: {exc}",
            "stdout": stdout[-STDERR_EXCERPT_CHARACTERS:],
        }
    if not isinstance(result, dict):
        return {
            "ok": False,
            "action": action,
            "error": "Bridge returned a JSON value that is not an object.",
            "stdout": stdout[-STDERR_EXCERPT_CHARACTERS:],
        }
    if returncode and result.get("ok", True):
        result["ok"] = False
        result["error"] = f"Bridge exited with code {returncode}."
    return result


# ---------------------------------------------------------------------------------------------
# Jobs: one worker thread, one EndNote session per job, cleanup when the queue goes idle.
# ---------------------------------------------------------------------------------------------


class _JobError(Exception):
    """A job failed before or around the bridge run; details go into the result."""

    def __init__(self, message: str, details: dict[str, Any] | None = None) -> None:
        super().__init__(message)
        self.details = details or {}


def _environment_integer(name: str, default: int) -> int:
    raw = os.environ.get(name, "").strip()
    if not raw:
        return default
    try:
        value = int(raw)
    except ValueError:
        _log(f"{name}={raw!r} is not an integer; using {default}.")
        return default
    if value < 0:
        _log(f"{name}={value} is negative; using {default}.")
        return default
    return value


def _session_module() -> Any:
    """endnote_session needs user32 and winreg, so it exists only on Windows."""
    if os.name != "nt":
        return None
    directory = str(SCRIPT_DIRECTORY)
    if directory not in sys.path:
        sys.path.insert(0, directory)
    import endnote_session

    return endnote_session


_LIFETIME_JOB: int | None = None
_LIFETIME_JOB_LOCK = threading.Lock()


def _job_object_api() -> tuple[Any, Any]:
    """kernel32 with the job-object signatures, and a function that sets a job's limit flags."""
    import ctypes
    from ctypes import wintypes

    class BasicLimits(ctypes.Structure):
        _fields_ = [
            ("PerProcessUserTimeLimit", ctypes.c_int64), ("PerJobUserTimeLimit", ctypes.c_int64),
            ("LimitFlags", wintypes.DWORD), ("MinimumWorkingSetSize", ctypes.c_size_t),
            ("MaximumWorkingSetSize", ctypes.c_size_t), ("ActiveProcessLimit", wintypes.DWORD),
            ("Affinity", ctypes.c_size_t), ("PriorityClass", wintypes.DWORD), ("SchedulingClass", wintypes.DWORD),
        ]

    class ExtendedLimits(ctypes.Structure):
        _fields_ = [
            ("BasicLimitInformation", BasicLimits), ("IoInfo", ctypes.c_ulonglong * 6),
            ("ProcessMemoryLimit", ctypes.c_size_t), ("JobMemoryLimit", ctypes.c_size_t),
            ("PeakProcessMemoryUsed", ctypes.c_size_t), ("PeakJobMemoryUsed", ctypes.c_size_t),
        ]

    kernel32 = ctypes.WinDLL("kernel32", use_last_error=True)
    kernel32.CreateJobObjectW.restype = wintypes.HANDLE
    kernel32.CreateJobObjectW.argtypes = (ctypes.c_void_p, wintypes.LPCWSTR)
    kernel32.SetInformationJobObject.argtypes = (wintypes.HANDLE, ctypes.c_int, ctypes.c_void_p, wintypes.DWORD)
    kernel32.AssignProcessToJobObject.argtypes = (wintypes.HANDLE, wintypes.HANDLE)
    kernel32.CloseHandle.argtypes = (wintypes.HANDLE,)

    def set_limit_flags(job: int, flags: int) -> bool:
        limits = ExtendedLimits()
        limits.BasicLimitInformation.LimitFlags = flags
        if kernel32.SetInformationJobObject(job, JOB_OBJECT_EXTENDED_LIMIT_INFORMATION, ctypes.byref(limits), ctypes.sizeof(limits)):
            return True
        _log(f"could not set the job object limits: {ctypes.WinError(ctypes.get_last_error())}")
        return False

    return kernel32, set_limit_flags


def _bind_to_server_lifetime(process: subprocess.Popen) -> None:
    """Windows keeps a child alive when its parent dies. Every bridge and every EndNote this server
    starts joins one kill-on-close job object, so even a hard kill of the server leaves none behind."""
    global _LIFETIME_JOB
    if os.name != "nt":
        return
    import ctypes

    kernel32, set_limit_flags = _job_object_api()
    with _LIFETIME_JOB_LOCK:
        if _LIFETIME_JOB is None:
            job = kernel32.CreateJobObjectW(None, None)
            if not job:
                _log(f"kill-on-close job object unavailable: {ctypes.WinError(ctypes.get_last_error())}")
                return
            if not set_limit_flags(job, JOB_OBJECT_LIMIT_KILL_ON_JOB_CLOSE):
                kernel32.CloseHandle(job)
                return
            _LIFETIME_JOB = job  # held for the server's lifetime; the OS closes it on exit
        if not kernel32.AssignProcessToJobObject(_LIFETIME_JOB, int(process._handle)):
            _log(f"could not bind PID {process.pid} to the server's lifetime: {ctypes.WinError(ctypes.get_last_error())}")


def _release_server_lifetime() -> None:
    """Let every process bound so far outlive the server; later ones join a new job object.
    Called only from the worker between jobs, when the one process still bound is an EndNote
    that the user has started working in."""
    global _LIFETIME_JOB
    if os.name != "nt":
        return
    kernel32, set_limit_flags = _job_object_api()
    with _LIFETIME_JOB_LOCK:
        if _LIFETIME_JOB is None or not set_limit_flags(_LIFETIME_JOB, 0):
            return
        kernel32.CloseHandle(_LIFETIME_JOB)
        _LIFETIME_JOB = None


class _BridgeLock:
    """Named mutex shared by every server in this Windows session. Only the worker thread owns it."""

    def __init__(self) -> None:
        self._handle = None
        if os.name != "nt":
            return
        import ctypes
        from ctypes import wintypes

        self._kernel32 = ctypes.WinDLL("kernel32", use_last_error=True)
        self._kernel32.CreateMutexW.restype = wintypes.HANDLE
        self._kernel32.CreateMutexW.argtypes = (ctypes.c_void_p, wintypes.BOOL, wintypes.LPCWSTR)
        self._kernel32.WaitForSingleObject.argtypes = (wintypes.HANDLE, wintypes.DWORD)
        self._kernel32.ReleaseMutex.argtypes = (wintypes.HANDLE,)
        self._handle = self._kernel32.CreateMutexW(None, False, BRIDGE_MUTEX_NAME)
        if not self._handle:
            _log(f"could not open the bridge lock: {ctypes.WinError(ctypes.get_last_error())}")

    def acquire(self, should_stop) -> bool:
        if not self._handle:
            return True
        while True:
            if self._kernel32.WaitForSingleObject(self._handle, LOCK_POLL_MILLISECONDS) in (WAIT_OBJECT_0, WAIT_ABANDONED):
                return True
            if should_stop():
                return False

    def release(self) -> None:
        if self._handle:
            self._kernel32.ReleaseMutex(self._handle)


@dataclass
class _Job:
    id: str
    tool: str
    action: str
    arguments: dict[str, Any]
    created: float = field(default_factory=time.time)
    state: str = "queued"
    started: float | None = None
    finished: float | None = None
    result: dict[str, Any] | None = None
    notices_dismissed: list[str] = field(default_factory=list)
    attention: list[dict[str, Any]] = field(default_factory=list)
    cancel_requested: bool = False
    process: subprocess.Popen | None = None
    word_pid: int | None = None


@dataclass
class _Session:
    library: Path
    process_id: int
    window_preexisting: bool
    recent_before: dict[str, str]


def _timestamp(value: float | None) -> str | None:
    return None if value is None else time.strftime("%Y-%m-%dT%H:%M:%SZ", time.gmtime(value))


def _job_library(job: _Job) -> Path | None:
    """The library a job formats against: render's manifest library, or an explicit library_path."""
    explicit = job.arguments.get("library_path")
    candidate = str(explicit) if explicit else None
    if candidate is None and job.action == "render":
        try:
            manifest = json.loads(Path(str(job.arguments.get("manifest_path", ""))).read_text(encoding="utf-8-sig"))
            library = manifest.get("library") if isinstance(manifest, dict) else None
            candidate = library.get("path") if isinstance(library, dict) else None
        except (OSError, UnicodeError, json.JSONDecodeError):
            return None  # the bridge reports the unreadable manifest itself
    if not candidate:
        return None
    path = Path(candidate)
    return path if path.suffix.lower() == ".enl" and path.is_file() else None


class _JobRunner:
    def __init__(self) -> None:
        self._condition = threading.Condition(threading.RLock())
        self._jobs: dict[str, _Job] = {}
        self._pending: queue.Queue[_Job | None] = queue.Queue()
        self._worker: threading.Thread | None = None
        self._stopping = False
        # (PID, recent-library list before it started) of an EndNote this server started, and the
        # library windows this server opened in it. EndNote is single-instance, so a library the
        # user opens meanwhile lands in that same process and makes it the user's.
        self._owned_endnote: tuple[int, dict[str, str]] | None = None
        self._owned_windows: set[int] = set()

    # -- public, called from the MCP thread ------------------------------------------------

    def submit(self, tool: str, action: str, arguments: dict[str, Any]) -> _Job:
        job = _Job(uuid.uuid4().hex[:12], tool, action, dict(arguments))
        with self._condition:
            self._jobs[job.id] = job
            if self._worker is None:
                self._worker = threading.Thread(target=self._work, name="endnote-jobs", daemon=True)
                self._worker.start()
        self._pending.put(job)
        return job

    def get(self, job_id: str) -> _Job | None:
        with self._condition:
            return self._jobs.get(job_id)

    def wait(self, job: _Job, seconds: float | None) -> None:
        with self._condition:
            self._condition.wait_for(lambda: job.state in TERMINAL_JOB_STATES, timeout=seconds)

    def cancel(self, job: _Job) -> None:
        with self._condition:
            if job.state in TERMINAL_JOB_STATES:
                return
            job.cancel_requested = True
            if job.state == "queued":
                self._finish(job, "cancelled", {"ok": False, "action": job.action, "error": "Cancelled before it started."})

    def describe(self, job: _Job, include_result: bool = True) -> dict[str, Any]:
        with self._condition:
            data: dict[str, Any] = {
                "job_id": job.id,
                "tool": job.tool,
                "action": job.action,
                "state": job.state,
                "created_at": _timestamp(job.created),
                "started_at": _timestamp(job.started),
                "finished_at": _timestamp(job.finished),
                "notices_dismissed": list(job.notices_dismissed),
                "attention": list(job.attention),
            }
            if job.state == "queued":
                data["queue_position"] = 1 + sum(
                    1 for other in self._jobs.values() if other.state == "queued" and other.created < job.created
                )
            if include_result and job.result is not None:
                data["result"] = job.result
            return data

    def all_jobs(self) -> list[_Job]:
        with self._condition:
            return sorted(self._jobs.values(), key=lambda job: job.created)

    def shutdown(self) -> None:
        """stdin closed: cancel everything, then let the worker close what this server opened."""
        with self._condition:
            self._stopping = True
            worker = self._worker
            for job in self._jobs.values():
                if job.state not in TERMINAL_JOB_STATES:
                    self.cancel(job)
        if worker is None:
            return
        self._pending.put(None)
        worker.join(LIBRARY_CLOSE_WAIT_SECONDS + ENDNOTE_EXIT_WAIT_SECONDS)

    # -- worker thread ----------------------------------------------------------------------

    def _finish(self, job: _Job, state: str, result: dict[str, Any]) -> None:
        with self._condition:
            job.state = state
            job.finished = time.time()
            job.result = result
            self._condition.notify_all()

    def _set_state(self, job: _Job, state: str) -> None:
        with self._condition:
            job.state = state
            if state == "running":
                job.started = time.time()
            self._condition.notify_all()

    def _work(self) -> None:
        lock = _BridgeLock()
        while True:
            idle = _environment_integer(IDLE_CLOSE_ENV_VARIABLE, DEFAULT_IDLE_CLOSE_SECONDS)
            try:
                job = self._pending.get(timeout=idle if self._owned_endnote else None)
            except queue.Empty:
                self._close_owned_endnote(lock)
                continue
            if job is None:
                self._close_owned_endnote(lock)
                return
            if job.state in TERMINAL_JOB_STATES:
                continue
            try:
                self._run(job, lock)
            except Exception as exc:  # Worker boundary: one broken job must not stop the queue.
                _log(traceback.format_exc())
                self._finish(job, "failed", {"ok": False, "action": job.action, "error": f"{type(exc).__name__}: {exc}"})

    def _run(self, job: _Job, lock: _BridgeLock) -> None:
        self._set_state(job, "waiting_for_lock")
        if not lock.acquire(lambda: job.cancel_requested or self._stopping):
            self._finish(job, "cancelled", {
                "ok": False, "action": job.action,
                "error": "Cancelled while another EndNote job in this Windows session held the bridge lock.",
            })
            return
        try:
            self._set_state(job, "running")
            if job.action == "import-ris":
                # Creating a library needs EndNote fully closed, so an idle EndNote kept by this
                # server must not refuse it. The named mutex is re-entrant for this thread.
                self._close_owned_endnote(lock)
            sessions = _session_module()
            recent_before = sessions.recent_libraries() if sessions else {}
            endnote_before = _process_pids(ENDNOTE_PROCESS_NAME)
            session = None
            try:
                session = self._open_session(job)
                result = self._execute(job)
            except _JobError as exc:
                result = {"ok": False, "action": job.action, "error": str(exc), **exc.details}
            finally:
                if session is not None:
                    self._close_session(session, job)
            self._adopt_new_endnote(endnote_before, recent_before)
            if job.notices_dismissed:
                result["notices_dismissed"] = list(job.notices_dismissed)
            if job.attention and not result.get("ok"):
                result["blocking_dialogs"] = list(job.attention)
            if job.cancel_requested:
                state = "cancelled"
            else:
                state = "succeeded" if result.get("ok") else "failed"
            self._finish(job, state, result)
        finally:
            lock.release()

    def _open_session(self, job: _Job) -> _Session | None:
        """Open the job's library in EndNote, beside any library the user has open."""
        sessions = _session_module()
        library = _job_library(job)
        if sessions is None or library is None:
            return None
        executable = sessions.endnote_executable()
        if executable is None:
            return None  # the bridge reports the missing installation itself
        recent = sessions.recent_libraries()
        running = _process_pids(ENDNOTE_PROCESS_NAME)
        job.notices_dismissed.extend(sessions.dismiss_notices(running))
        preexisting = bool(running) and sessions.library_window(running, library) is not None

        def forward() -> None:
            # Also for an open library: forwarding it makes it the one CWYW formats against.
            _bind_to_server_lifetime(sessions.launch_endnote(executable, library, hidden=True))

        forward()
        timeout = _environment_integer(READY_TIMEOUT_ENV_VARIABLE, DEFAULT_READY_TIMEOUT_MILLISECONDS) / 1000
        window, dismissed = sessions.wait_for_library(lambda: _process_pids(ENDNOTE_PROCESS_NAME), library, timeout, forward)
        job.notices_dismissed.extend(dismissed)
        if window is None:
            dialogs = sessions.blocking_dialogs(_process_pids(ENDNOTE_PROCESS_NAME))
            raise _JobError(
                f"EndNote did not open {library} within {timeout:.0f} s. Raise {READY_TIMEOUT_ENV_VARIABLE} "
                "for a very large library, or answer the dialog EndNote shows.",
                {"blocking_dialogs": [{"title": item["title"], "buttons": item["buttons"]} for item in dialogs]},
            )
        with self._condition:
            if not running:
                self._owned_endnote = (window.process_id, recent)
            if self._owned_endnote is not None and self._owned_endnote[0] == window.process_id:
                self._owned_windows.add(window.handle)
        return _Session(library, window.process_id, preexisting, recent)

    def _close_session(self, session: _Session, job: _Job) -> None:
        """Close the library this job opened; an EndNote this server started stays for the idle timer."""
        sessions = _session_module()
        owned = self._owned_endnote is not None and self._owned_endnote[0] == session.process_id
        if not session.window_preexisting:
            window = sessions.library_window({session.process_id}, session.library)
            only_library = len(sessions.library_windows({session.process_id})) <= 1
            if window is not None and not (owned and only_library):
                sessions.close_window(window.handle)
                deadline = time.monotonic() + LIBRARY_CLOSE_WAIT_SECONDS
                while sessions.window_exists(window.handle) and time.monotonic() < deadline:
                    job.notices_dismissed.extend(sessions.dismiss_notices({session.process_id}))
                    time.sleep(sessions.POLL_SECONDS)
                if sessions.window_exists(window.handle):
                    _log(f"EndNote kept the library window of {session.library} open after WM_CLOSE.")
        if not owned:
            try:
                sessions.restore_recent_libraries(session.recent_before)
            except OSError as exc:
                _log(f"could not restore EndNote's recent-library list: {exc}")

    def _adopt_new_endnote(self, endnote_before: set[int], recent_before: dict[str, str]) -> None:
        """An EndNote that CWYW or import-ris started during the job is closed by the idle timer too."""
        if self._owned_endnote is not None:
            return
        new = _process_pids(ENDNOTE_PROCESS_NAME) - endnote_before
        sessions = _session_module()
        if len(new) == 1 and sessions is not None:
            process_id = new.pop()
            with self._condition:
                self._owned_endnote = (process_id, recent_before)
                self._owned_windows = {window.handle for window in sessions.library_windows({process_id})}

    def _close_owned_endnote(self, lock: _BridgeLock) -> None:
        owned = self._owned_endnote
        sessions = _session_module()
        if owned is None or sessions is None:
            return
        process_id, recent = owned

        def alive() -> bool:
            return process_id in _process_pids(ENDNOTE_PROCESS_NAME)

        def wait_for_exit(seconds: float) -> None:
            deadline = time.monotonic() + seconds
            while alive() and time.monotonic() < deadline:
                sessions.dismiss_notices({process_id})
                time.sleep(sessions.POLL_SECONDS)

        lock.acquire(lambda: False)
        try:
            # A visible frame or a library this server did not open means the user works in it now.
            foreign = [window.title for window in sessions.library_windows({process_id}) if window.handle not in self._owned_windows]
            if foreign or any(window.visible for window in sessions.frame_windows({process_id})):
                _log(f"EndNote PID {process_id}, started by this server, now holds the user's work {foreign}; leaving it open.")
                with self._condition:
                    self._owned_endnote = None
                    self._owned_windows.clear()
                _release_server_lifetime()
                return
            # Closing the last library usually ends EndNote; the frame is closed only when it does not.
            for window in sessions.library_windows({process_id}):
                sessions.close_window(window.handle)
            wait_for_exit(LIBRARY_CLOSE_WAIT_SECONDS)
            if alive():
                for window in sessions.frame_windows({process_id}):
                    sessions.close_window(window.handle)
                wait_for_exit(ENDNOTE_EXIT_WAIT_SECONDS)
            if alive():
                # Measured after a Word was ended mid-CWYW (cancel or timeout): EndNote answers
                # messages but ignores every close request for minutes. Only an EndNote this server
                # started is ended this way; the user's own EndNote is never terminated.
                if not _terminate_process_pid(process_id) and alive():
                    _log(f"Could not end EndNote PID {process_id}, started by this server; retrying after the next idle period.")
                    return
                _log(f"Ended EndNote PID {process_id}, started by this server, which ignored every close request.")
            else:
                _log(f"Closed EndNote PID {process_id}, which this server had started.")
            with self._condition:
                self._owned_endnote = None
                self._owned_windows.clear()
            try:
                sessions.restore_recent_libraries(recent)
            except OSError as exc:
                _log(f"could not restore EndNote's recent-library list: {exc}")
        finally:
            lock.release()

    def _execute(self, job: _Job) -> dict[str, Any]:
        """Run the bridge, watching for dialogs, cancellation and the timeout while it works."""
        if not SCRIPT_PATH.is_file():
            raise _JobError(f"Bridge script not found: {SCRIPT_PATH}")
        try:
            executable = _powershell_executable()
        except RuntimeError as exc:
            raise _JobError(str(exc)) from exc
        command = _build_command(job.action, job.arguments, executable)
        timeout = _timeout_seconds(job.action, job.arguments)
        words_before = _process_pids(WORD_PROCESS_NAME)
        try:
            process = subprocess.Popen(
                command, stdin=subprocess.DEVNULL, stdout=subprocess.PIPE, stderr=subprocess.PIPE,
                text=True, encoding="utf-8", errors="replace",
                env={**os.environ, PYTHON_ENV_VARIABLE: sys.executable},
                creationflags=getattr(subprocess, "CREATE_NO_WINDOW", 0),
            )
        except OSError as exc:
            raise _JobError(f"Could not start the PowerShell bridge with {executable}: {exc}") from exc
        _bind_to_server_lifetime(process)
        job.process = process
        stdout: list[str] = []
        stderr: list[str] = []

        def drain(stream, lines: list[str], watch_word: bool) -> None:
            for line in stream:
                lines.append(line)
                if watch_word:
                    word_pid = _extract_word_pid(line)
                    if word_pid is not None:
                        job.word_pid = word_pid

        readers = [
            threading.Thread(target=drain, args=(process.stdout, stdout, False), daemon=True),
            threading.Thread(target=drain, args=(process.stderr, stderr, True), daemon=True),
        ]
        for reader in readers:
            reader.start()
        deadline = time.monotonic() + timeout
        first_seen: dict[int, float] = {}
        cancelled_dialogs: list[str] = []
        pressed: set[int] = set()
        outcome = None
        while True:
            try:
                process.wait(timeout=JOB_POLL_SECONDS)
                break
            except subprocess.TimeoutExpired:
                pass
            if job.cancel_requested:
                outcome = "cancelled"
                break
            if time.monotonic() >= deadline:
                outcome = "timeout"
                break
            chooser = self._watch_dialogs(job, first_seen)
            if chooser is not None and chooser["handle"] not in pressed and _session_module().press_cancel(chooser["handle"]):
                _log(f"Pressed Cancel on the CWYW dialog '{chooser['title']}'.")
                pressed.add(chooser["handle"])
                cancelled_dialogs.append(chooser["title"])
        if outcome is not None:
            self._stop_bridge(job, words_before)
        process.wait()
        for reader in readers:
            reader.join()
        job.process = None
        stderr_text = "".join(stderr)
        if outcome == "cancelled":
            return {"ok": False, "action": job.action, "error": "Cancelled while running; no output file was written.",
                    "stderr_tail": stderr_text[-STDERR_EXCERPT_CHARACTERS:]}
        if outcome == "timeout":
            return _with_powershell_notice({
                "ok": False,
                "action": job.action,
                "error": (
                    f"EndNote/Word bridge timed out after {timeout} seconds. blocking_dialogs lists any "
                    f"dialog that stayed open; see references/troubleshooting.md. Raise {TIMEOUT_ENV_VARIABLE} "
                    "for very large libraries."
                ),
                "isolated_word_pid": job.word_pid,
                "stderr_tail": stderr_text[-STDERR_EXCERPT_CHARACTERS:],
            }, stderr_text)
        result = _parse_bridge_output(job.action, process.returncode, "".join(stdout), stderr_text)
        if cancelled_dialogs:
            result["dialogs_cancelled"] = cancelled_dialogs
            if not result.get("ok"):
                result["bridge_error"] = result.get("error")
                result["error"] = f"CWYW asked '{cancelled_dialogs[0]}'. {CANCELLED_DIALOG_ADVICE}"
        return result

    def _stop_bridge(self, job: _Job, words_before: set[int]) -> None:
        """End the bridge tree and its hidden Word. Word comes from COM, not from the bridge, so its
        PID is the bridge's WORD_PID marker; a PID merely seen meanwhile could be the user's Word."""
        if job.process is not None:
            _terminate_process_pid(job.process.pid)
        word_pid = job.word_pid
        if word_pid is None:
            new_words = _process_pids(WORD_PROCESS_NAME) - words_before
            word_pid = new_words.pop() if len(new_words) == 1 else None
        if word_pid is not None:
            _terminate_process_pid(word_pid)

    def _watch_dialogs(self, job: _Job, first_seen: dict[int, float]) -> dict[str, Any] | None:
        """Dismiss news notices, report lasting dialogs, and return a CWYW chooser to cancel."""
        sessions = _session_module()
        if sessions is None:
            return None
        endnote = _process_pids(ENDNOTE_PROCESS_NAME)
        dismissed = sessions.dismiss_notices(endnote)
        targets = set(endnote) | ({job.word_pid} if job.word_pid else set())
        now = time.monotonic()
        dialogs = sessions.blocking_dialogs(targets)
        open_handles = {item["handle"] for item in dialogs}
        for handle in list(first_seen):
            if handle not in open_handles:
                del first_seen[handle]
        attention = []
        for item in dialogs:
            seen = first_seen.setdefault(item["handle"], now)
            if now - seen >= DIALOG_REPORT_AFTER_SECONDS:
                attention.append({
                    "application": "Word" if item["process_id"] == job.word_pid else "EndNote",
                    "title": item["title"],
                    "buttons": item["buttons"],
                    "open_seconds": round(now - seen),
                })
        chooser = next(
            (item for item in dialogs if any(title in item["title"] for title in CANCELLED_DIALOG_TITLES)), None
        )
        with self._condition:
            job.notices_dismissed.extend(dismissed)
            job.attention = attention
        return chooser


_JOB_RUNNER = _JobRunner()


def _job_status(arguments: dict[str, Any]) -> dict[str, Any]:
    job_id = arguments.get("job_id")
    if not job_id:
        return {"ok": True, "action": "job-status",
                "jobs": [_JOB_RUNNER.describe(job, include_result=False) for job in _JOB_RUNNER.all_jobs()]}
    job = _JOB_RUNNER.get(str(job_id))
    if job is None:
        return {"ok": False, "action": "job-status",
                "error": f"Unknown job_id '{job_id}'. Jobs live only as long as this MCP server."}
    try:
        wait_seconds = int(arguments.get("wait_seconds") or 0)
    except (TypeError, ValueError):
        wait_seconds = 0
    wait_seconds = max(0, min(MAX_STATUS_WAIT_SECONDS, wait_seconds))
    if wait_seconds:
        _JOB_RUNNER.wait(job, wait_seconds)
    return {"ok": True, "action": "job-status", "job": _JOB_RUNNER.describe(job)}


def _job_cancel(arguments: dict[str, Any]) -> dict[str, Any]:
    job = _JOB_RUNNER.get(str(_require(arguments, "job_id", JOB_CANCEL_TOOL)))
    if job is None:
        return {"ok": False, "action": "job-cancel", "error": f"Unknown job_id '{arguments.get('job_id')}'."}
    _JOB_RUNNER.cancel(job)
    return {"ok": True, "action": "job-cancel", "job": _JOB_RUNNER.describe(job, include_result=False)}


def _call_bridge_tool(name: str, arguments: dict[str, Any]) -> dict[str, Any]:
    action = BRIDGE_TOOL_ACTIONS[name]
    if action not in QUEUED_ACTIONS:
        return _invoke_bridge(action, arguments)
    job = _JOB_RUNNER.submit(name, action, arguments)
    if name in BACKGROUND_TOOLS and arguments.get("background", True):
        return {
            "ok": True,
            "action": action,
            "job_id": job.id,
            "state": job.state,
            "next": f"Call {JOB_STATUS_TOOL} with this job_id and wait_seconds until the state is terminal.",
        }
    _JOB_RUNNER.wait(job, None)
    return job.result or {"ok": False, "action": action, "error": "The job finished without a result."}


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
    if name == JOB_STATUS_TOOL:
        result = _job_status(arguments)
    elif name == JOB_CANCEL_TOOL:
        result = _job_cancel(arguments)
    elif name in BRIDGE_TOOL_ACTIONS:
        result = _call_bridge_tool(name, arguments)
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
                    "claim support and merged citations. Render into a new file, then validate. Render, "
                    "update and set-style return a job_id at once; follow it with get_endnote_job_status "
                    "and wait_seconds until the state is succeeded, failed or cancelled, and tell the user "
                    "about any dialog listed under attention. Jobs run one at a time against the "
                    "document's own library and close what they opened. Citations "
                    "must be CWYW-produced ADDIN EN.CITE fields with a traveling library and an ADDIN "
                    "EN.REFLIST bibliography; never type citation text by hand. Word and EndNote editing "
                    "runs silently. Library reads are read-only; explicit RIS import is performed by "
                    "EndNote itself against the exact requested library. When a result carries "
                    "powershell_notice, tell the user its message: under Windows PowerShell 5.1 "
                    "without pwsh.exe, read-only calls announce a PowerShell 7 install and the first "
                    "call other than probe, list-styles or validate starts it in the background."
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
    # Codex starts the server in the plugin's install folder, and Windows cannot rename or delete
    # a process's working folder, so staying there would block plugin updates. Any other folder,
    # such as a Claude Code project, stays the base for relative paths.
    if Path.cwd().resolve() == PLUGIN_ROOT:
        os.chdir(tempfile.gettempdir())
    try:
        return _serve()
    finally:
        # The client closed the conversation: cancel jobs and close what this server opened.
        _JOB_RUNNER.shutdown()


def _serve() -> int:
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
