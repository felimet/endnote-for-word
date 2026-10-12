"""Static contract tests: no Word, no EndNote, no PowerShell execution.

These tests read the committed files only. They guard the release metadata, the
agent-facing contracts that other tracks depend on, and the Windows PowerShell 5.1
compatibility of every .ps1 in the repository.
"""

from __future__ import annotations

import json
import re
import unittest
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]

PACKAGE_NAME = "endnote-for-word"
RELEASE_VERSION = "1.3.0"
RELEASE_AUTHOR = "Jia-Ming Zhou (Felimet)"
SKILL_DIRECTORY = ROOT / "skills" / PACKAGE_NAME
SKILL_PATH = SKILL_DIRECTORY / "SKILL.md"
BRIDGE_PATH = ROOT / "scripts" / "endnote-word.ps1"
SERVER_PATH = ROOT / "scripts" / "mcp_server.py"

SKILL_MAXIMUM_LINES = 500
MINIMUM_FILE_BACKED_EVALS = 6

PROSE_SUFFIXES = {".md", ".json", ".yaml", ".yml", ".py", ".ps1", ".txt", ".ris"}
IGNORED_DIRECTORIES = {".git", ".claude", "__pycache__", ".venv", "node_modules"}

EM_DASH = "—"
EN_DASH = "–"

# Windows PowerShell 5.1 has no ternary, null-coalescing, null-conditional, clean block,
# ForEach-Object -Parallel or Join-String. Section 3 of the build spec makes 5.1 support
# mandatory, so any of these in a committed .ps1 is a hard failure.
POWERSHELL_SEVEN_ONLY_PATTERNS = (
    (re.compile(r"(?im)^\s*clean\s*\{"), "clean {} block (PowerShell 7.3+)"),
    (re.compile(r"(?i)-Parallel\b"), "ForEach-Object -Parallel (PowerShell 7+)"),
    (re.compile(r"(?i)\bJoin-String\b"), "Join-String (PowerShell 7+)"),
)

# Helper names the bridge contract fixes, because the MCP server, the skill and these
# tests all reason about them.
REQUIRED_BRIDGE_FUNCTIONS = (
    "Write-Log",
    "Write-JsonResult",
    "Resolve-ExistingFile",
    "Resolve-OutputFile",
    "Release-ComObject",
    "Start-Word",
    "Stop-Word",
    "Get-AgentDefaults",
    "Get-DefaultOutputStyle",
    "Read-Manifest",
    "Assert-Manifest",
    "Invoke-ManifestPrecheck",
    "Get-TemporaryCitationText",
    "Get-CitationMarkerPlans",
    "Assert-NoUnresolvedCitationMarkers",
    "Set-CitationDisplayFormats",
    "Find-UniqueMarkerRange",
    "Get-InstalledStyles",
    "Resolve-StyleName",
    "Start-EndNoteWithLibrary",
    "Get-EndNoteLayout",
    "Set-EndNoteLayout",
    "Invoke-CwywFormat",
    "Get-CitationFields",
    "Get-ValidationReport",
    "Get-EndNoteRecordCount",
    "Resolve-EndNoteImportLibrary",
    "Open-EndNoteLibraryForImport",
    "Save-DocumentAtomically",
)

REQUIRED_BRIDGE_ACTIONS = (
    "probe",
    "list-styles",
    "configure-defaults",
    "render",
    "validate",
    "update",
    "set-style",
    "unformat",
    "import-ris",
)

# Build spec section 0.6b: ENRemoveFieldCodes raises a modal confirmation dialog and never
# returns under automation, so it is not part of the v1 surface. Section 3 requires the action
# to be absent from the bridge, so re-adding it is a regression these tests must catch.
REJECTED_BRIDGE_ACTIONS = ("remove-field-codes",)

JSON_FILES = (
    Path(".claude-plugin/plugin.json"),
    Path(".codex-plugin/plugin.json"),
    Path(".mcp.json"),
    Path(".codex-mcp.json"),
    Path("config/defaults.json"),
    Path("examples/example-manifest.json"),
    Path("evals/evals.json"),
    Path("tests/fixtures/single-manifest.json"),
    Path("tests/fixtures/live-manifest.json"),
)


def iter_repository_files(suffixes: set[str]) -> list[Path]:
    """Return every committed file with one of ``suffixes``, ignoring build directories."""
    found: list[Path] = []
    for path in ROOT.rglob("*"):
        if not path.is_file():
            continue
        # Relative parts: the checkout itself may live under .claude/worktrees.
        if IGNORED_DIRECTORIES.intersection(path.relative_to(ROOT).parts):
            continue
        if path.suffix.lower() in suffixes:
            found.append(path)
    return sorted(found)


def _blank_span(consumed: str) -> str:
    """Return the replacement for one stripped literal.

    The newlines the literal spanned are kept so that the line numbers reported for the
    code after a block comment or a here-string still match the real file.
    """
    return " " + "\n" * consumed.count("\n")


def strip_powershell_literals(text: str) -> str:
    """Blank out comments, here-strings and quoted strings in PowerShell source.

    The remaining text can be scanned for operators without the false positives that a
    naive regex hits on regular expressions, URLs and documentation comments.
    """
    output: list[str] = []
    index = 0
    length = len(text)
    while index < length:
        pair = text[index : index + 2]
        character = text[index]
        start = index

        if pair == "<#":
            depth = 1
            index += 2
            while index < length and depth > 0:
                if text[index : index + 2] == "<#":
                    depth += 1
                    index += 2
                elif text[index : index + 2] == "#>":
                    depth -= 1
                    index += 2
                else:
                    index += 1
            output.append(_blank_span(text[start:index]))
            continue

        if pair in ("@'", '@"') and (index == 0 or text[index - 1] in "\r\n \t=,(&|"):
            terminator = "\n'@" if pair == "@'" else '\n"@'
            end = text.find(terminator, index + 2)
            index = length if end == -1 else end + len(terminator)
            output.append(_blank_span(text[start:index]))
            continue

        if character == "'":
            index += 1
            while index < length:
                if text[index] == "'":
                    if text[index + 1 : index + 2] == "'":
                        index += 2
                        continue
                    index += 1
                    break
                index += 1
            output.append(_blank_span(text[start:index]))
            continue

        if character == '"':
            index += 1
            while index < length:
                if text[index] == "`":
                    index += 2
                    continue
                if text[index] == '"':
                    if text[index + 1 : index + 2] == '"':
                        index += 2
                        continue
                    index += 1
                    break
                index += 1
            output.append(_blank_span(text[start:index]))
            continue

        if character == "#":
            while index < length and text[index] != "\n":
                index += 1
            output.append(_blank_span(text[start:index]))
            continue

        output.append(character)
        index += 1
    return "".join(output)


def find_powershell_seven_only_syntax(text: str) -> list[str]:
    """Return human-readable descriptions of PowerShell 7 only syntax found in ``text``."""
    findings: list[str] = []
    code = strip_powershell_literals(text)
    for line_number, line in enumerate(code.splitlines(), start=1):
        if "?" in line:
            findings.append(
                f"line {line_number}: '?' is ternary, null-coalescing, null-conditional "
                "or the Where-Object alias, none of which are safe on 5.1"
            )
    for pattern, description in POWERSHELL_SEVEN_ONLY_PATTERNS:
        for match in pattern.finditer(code):
            line_number = code.count("\n", 0, match.start()) + 1
            findings.append(f"line {line_number}: {description}")
    return findings


class ReleaseMetadataTests(unittest.TestCase):
    def test_json_files_parse(self) -> None:
        for relative in JSON_FILES:
            with self.subTest(path=relative):
                path = ROOT / relative
                self.assertTrue(path.is_file(), f"{relative} is missing")
                json.loads(path.read_text(encoding="utf-8"))

    def test_release_versions_and_author_are_aligned(self) -> None:
        for relative in (Path(".claude-plugin/plugin.json"), Path(".codex-plugin/plugin.json")):
            with self.subTest(path=relative):
                manifest = json.loads((ROOT / relative).read_text(encoding="utf-8"))
                self.assertEqual(manifest["name"], PACKAGE_NAME)
                self.assertEqual(manifest["version"], RELEASE_VERSION)
                self.assertEqual(manifest["author"]["name"], RELEASE_AUTHOR)

        skill = SKILL_PATH.read_text(encoding="utf-8")
        self.assertIn(f"version: {RELEASE_VERSION}", skill)
        self.assertIn(f"author: {RELEASE_AUTHOR}", skill)

        server = SERVER_PATH.read_text(encoding="utf-8")
        self.assertIn(f'SERVER_NAME = "{PACKAGE_NAME}"', server)
        self.assertIn(f'SERVER_VERSION = "{RELEASE_VERSION}"', server)

        for relative in (Path("LICENSE"), Path(f"skills/{PACKAGE_NAME}/LICENSE.txt")):
            with self.subTest(path=relative):
                self.assertIn(RELEASE_AUTHOR, (ROOT / relative).read_text(encoding="utf-8"))

    def test_skill_frontmatter_and_line_cap(self) -> None:
        text = SKILL_PATH.read_text(encoding="utf-8")
        self.assertTrue(text.startswith("---\n"))
        self.assertRegex(text, rf"(?m)^name: {PACKAGE_NAME}$")
        self.assertRegex(text, r"(?m)^description: .*EndNote.*$")
        self.assertRegex(text, r"(?m)^license: MIT$")
        self.assertLessEqual(len(text.splitlines()), SKILL_MAXIMUM_LINES)

    def test_shared_output_coordination_contract(self) -> None:
        text = SKILL_PATH.read_text(encoding="utf-8")
        for phrase in (
            "## Shared output coordination",
            "user's current working directory",
            "Operate as a complete standalone skill",
            "Do not assume, require, install, or prompt for `mathtype-for-word`",
            "endnote-for-word-manifest.json",
            "<source-stem>-mathtype-endnote.docx",
            "Verify that `mathtype-for-word` is available",
            "complete only the EndNote scope",
            "one shared final DOCX",
            "validate the final shared document with both skills",
        ):
            self.assertIn(phrase, text)

    def test_skill_states_the_non_negotiable_invariants(self) -> None:
        text = SKILL_PATH.read_text(encoding="utf-8")
        for phrase in (
            "ADDIN EN.CITE",
            "ADDIN EN.REFLIST",
            "traveling library",
            "search_endnote_library",
            "precheck_endnote_manifest",
            "list_endnote_styles",
            "modal dialog",
            "powershell.exe",
            "pwsh.exe",
        ):
            with self.subTest(phrase=phrase):
                self.assertIn(phrase, text)

    def test_platform_specific_plugin_root_variables(self) -> None:
        claude = (ROOT / ".mcp.json").read_text(encoding="utf-8")
        codex = (ROOT / ".codex-mcp.json").read_text(encoding="utf-8")
        self.assertIn("CLAUDE_PLUGIN_ROOT", claude)
        self.assertIn("PLUGIN_ROOT", claude)
        self.assertIn("${PLUGIN_ROOT}", codex)
        for text in (claude, codex):
            servers = json.loads(text)["mcpServers"]
            self.assertIn(PACKAGE_NAME, servers.keys())
            # powershell.exe ships with every supported Windows, so the server starts before
            # PowerShell 7 exists; the bridge itself still runs under pwsh.exe when present.
            self.assertEqual(servers[PACKAGE_NAME]["command"], "powershell.exe")


class DefaultsTests(unittest.TestCase):
    def test_defaults_match_the_measured_temporary_citation_syntax(self) -> None:
        data = json.loads((ROOT / "config" / "defaults.json").read_text(encoding="utf-8"))
        self.assertEqual(data["schema_version"], 1)

        style = data["style"]["default_output_style"]
        self.assertEqual(style, "IEEE")

        temporary = data["temporary_citation"]
        self.assertEqual(temporary["open"], "{")
        self.assertEqual(temporary["close"], "}")
        self.assertEqual(temporary["record_prefix"], "#")
        self.assertEqual(temporary["pages_prefix"], "@")

        validation = data["validation"]
        for gate in (
            "reject_unresolved_markers",
            "reject_residual_temporary_citations",
            "require_native_en_cite_fields",
            "require_native_reflist_field",
            "require_visible_verified_doi_links",
            "reject_duplicate_doi_resolver_prefixes",
        ):
            with self.subTest(gate=gate):
                self.assertTrue(validation[gate])

        self.assertTrue(data["formatting"]["silent_word_and_endnote"])
        self.assertTrue(data["doi"]["required"])
        self.assertTrue(data["doi"]["verify_before_citing"])
        self.assertEqual(data["doi"]["bibliography_prefix"], "doi: https://doi.org/")

    def test_bridge_publishes_defaults_to_the_documented_location(self) -> None:
        text = BRIDGE_PATH.read_text(encoding="utf-8")
        self.assertIn("EndNoteForWordAgent", text)
        self.assertIn("defaults.json", text)
        self.assertIn("Get-DefaultOutputStyle", text)


class BridgeContractTests(unittest.TestCase):
    def test_bridge_exists(self) -> None:
        self.assertTrue(BRIDGE_PATH.is_file(), f"{BRIDGE_PATH} is missing")

    def test_bridge_defines_every_contractual_helper(self) -> None:
        text = BRIDGE_PATH.read_text(encoding="utf-8")
        for name in REQUIRED_BRIDGE_FUNCTIONS:
            with self.subTest(function=name):
                self.assertRegex(text, rf"(?im)^\s*function\s+{re.escape(name)}\b")

    def test_bridge_exposes_every_contractual_action(self) -> None:
        text = BRIDGE_PATH.read_text(encoding="utf-8")
        for action in REQUIRED_BRIDGE_ACTIONS:
            with self.subTest(action=action):
                self.assertIn(f"'{action}'", text)

    def test_bridge_rejects_the_actions_dropped_after_measurement(self) -> None:
        text = BRIDGE_PATH.read_text(encoding="utf-8")
        validate_set = re.search(r"\[ValidateSet\((.*?)\)\]", text, re.DOTALL)
        self.assertIsNotNone(validate_set, "the -Action parameter has no ValidateSet")
        assert validate_set is not None
        for action in REJECTED_BRIDGE_ACTIONS:
            with self.subTest(action=action):
                self.assertNotIn(f"'{action}'", validate_set.group(1))

    def test_bridge_uses_strict_mode_and_stops_on_error(self) -> None:
        text = BRIDGE_PATH.read_text(encoding="utf-8")
        self.assertIn("Set-StrictMode -Version Latest", text)
        self.assertIn("$ErrorActionPreference = 'Stop'", text)

    def test_bridge_reads_text_files_through_the_strict_utf8_reader(self) -> None:
        # Get-Content without -Encoding uses the ANSI code page on Windows PowerShell 5.1, and
        # File.ReadAllText turns lenient as soon as it sees a byte order mark.
        code = strip_powershell_literals(BRIDGE_PATH.read_text(encoding="utf-8"))
        self.assertIn("New-Object System.Text.UTF8Encoding($true, $true)", code)
        for reader in (r"\bGet-Content\b", r"\bReadAllText\b", r"\bReadAllLines\b"):
            with self.subTest(reader=reader):
                self.assertIsNone(re.search(reader, code))

    def test_bridge_hands_out_plain_file_system_paths(self) -> None:
        # (Resolve-Path).Path is provider-qualified for UNC files, which Python and Word reject.
        code = strip_powershell_literals(BRIDGE_PATH.read_text(encoding="utf-8"))
        self.assertIsNone(re.search(r"\bResolve-Path\b", code))

    def test_bridge_never_calls_a_modal_endnote_macro(self) -> None:
        text = BRIDGE_PATH.read_text(encoding="utf-8")
        for macro in (
            "ENFormatBibliographyDlg",
            "ENConfigureInstantFormattingDlg",
            "ENPreferences",
            "ENRemoveFieldCodes",
        ):
            with self.subTest(macro=macro):
                self.assertNotRegex(text, rf"Run\(\s*'{macro}'")

    def test_bridge_keeps_office_automation_hidden(self) -> None:
        text = BRIDGE_PATH.read_text(encoding="utf-8")
        self.assertIn("$script:Word.Visible = $false", text)
        self.assertRegex(text, r"Start-Process[^\r\n]+-WindowStyle Hidden")


class PowerShellCompatibilityTests(unittest.TestCase):
    """Windows PowerShell 5.1 compatibility lint over every .ps1 in the repository."""

    def test_literal_stripper_understands_powershell_quoting(self) -> None:
        sample = "\n".join(
            [
                "# a comment with a ? in it",
                "<# block ? comment #>",
                "$pattern = '\\d+\\?'",
                '$text = "a ? inside a string"',
                "@'",
                "here ? string",
                "'@",
                "$plain = 1",
            ]
        )
        self.assertEqual(find_powershell_seven_only_syntax(sample), [])
        self.assertNotEqual(find_powershell_seven_only_syntax("$a = $b ?? 'fallback'"), [])
        self.assertNotEqual(find_powershell_seven_only_syntax("$a = $test ? 1 : 2"), [])
        self.assertNotEqual(find_powershell_seven_only_syntax("$a = $object?.Member"), [])
        self.assertNotEqual(find_powershell_seven_only_syntax("clean {\n  $x = 1\n}"), [])
        self.assertNotEqual(
            find_powershell_seven_only_syntax("$items | ForEach-Object -Parallel { $_ }"), []
        )

    def test_reported_line_numbers_survive_a_multi_line_literal(self) -> None:
        source = "@'\nhere\nstring\n'@\n$a = $b ?? 'fallback'\n"
        findings = find_powershell_seven_only_syntax(source)
        self.assertEqual(len(findings), 1, findings)
        self.assertTrue(findings[0].startswith("line 5:"), findings[0])

    def test_every_script_is_windows_powershell_51_compatible(self) -> None:
        scripts = iter_repository_files({".ps1"})
        self.assertTrue(scripts, "no .ps1 files were found to lint")
        for path in scripts:
            with self.subTest(path=path.relative_to(ROOT)):
                findings = find_powershell_seven_only_syntax(
                    path.read_text(encoding="utf-8-sig")
                )
                self.assertEqual(findings, [], f"{path.relative_to(ROOT)}: {findings}")

    def test_word_save_calls_pass_a_plain_string_path(self) -> None:
        # Under Windows PowerShell 5.1, Word never returned from SaveAs2 when the path was uncast
        # Join-Path output; a [string] cast made it return. The cause inside Word is unknown, so
        # every save call keeps the cast.
        calls = []
        for path in iter_repository_files({".ps1"}):
            code = strip_powershell_literals(path.read_text(encoding="utf-8-sig"))
            calls.extend(
                (path, match)
                for match in re.finditer(r"\.SaveAs2?\(\s*([^,)]*)", code, re.IGNORECASE)
            )
        self.assertTrue(calls, "no Word save call was found to check")
        for path, match in calls:
            with self.subTest(path=path.relative_to(ROOT), argument=match.group(1)):
                self.assertRegex(match.group(1), r"(?i)^\(*\s*\[string\]", match.group(0))

    def test_every_script_forces_utf8_console_output(self) -> None:
        for path in iter_repository_files({".ps1"}):
            text = path.read_text(encoding="utf-8-sig")
            if "ConvertTo-Json" not in text:
                continue
            with self.subTest(path=path.relative_to(ROOT)):
                self.assertIn("UTF8Encoding", text)


class RepositoryHygieneTests(unittest.TestCase):
    def test_no_todo_or_placeholder_markers(self) -> None:
        findings: list[str] = []
        pattern = re.compile(r"\[(?:TODO|PLACEHOLDER):", re.IGNORECASE)
        for path in iter_repository_files(PROSE_SUFFIXES):
            if path == Path(__file__).resolve():
                continue
            if pattern.search(path.read_text(encoding="utf-8", errors="replace")):
                findings.append(str(path.relative_to(ROOT)))
        self.assertEqual(findings, [])

    def test_prose_uses_hyphens_not_dashes(self) -> None:
        findings: list[str] = []
        for path in iter_repository_files(PROSE_SUFFIXES):
            if path == Path(__file__).resolve():
                continue
            text = path.read_text(encoding="utf-8", errors="replace")
            if EM_DASH in text or EN_DASH in text:
                findings.append(str(path.relative_to(ROOT)))
        self.assertEqual(findings, [], "em dash or en dash found; the house style is hyphen only")

    def test_no_machine_specific_home_directory_is_committed(self) -> None:
        home = str(Path.home())
        if not home or len(home) < 4:
            self.skipTest("home directory could not be resolved")
        findings: list[str] = []
        for path in iter_repository_files(PROSE_SUFFIXES):
            if path == Path(__file__).resolve():
                continue
            text = path.read_text(encoding="utf-8", errors="replace")
            if home.lower() in text.lower() or home.replace("\\", "\\\\").lower() in text.lower():
                findings.append(str(path.relative_to(ROOT)))
        self.assertEqual(findings, [], "a personal home directory path is committed")


class ManifestFixtureTests(unittest.TestCase):
    def test_fixture_manifests_follow_schema_v1(self) -> None:
        for name in ("single-manifest.json", "live-manifest.json"):
            with self.subTest(fixture=name):
                path = ROOT / "tests" / "fixtures" / name
                manifest = json.loads(path.read_text(encoding="utf-8"))
                self.assertEqual(manifest["schema_version"], 1)
                self.assertTrue(manifest["library"]["path"])
                self.assertTrue(manifest["style"])
                self.assertEqual(manifest["style"], "IEEE")
                self.assertIn("expected", manifest["bibliography"])

                identifiers = set()
                markers = set()
                for citation in manifest["citations"]:
                    self.assertRegex(citation["id"], r"^[A-Za-z][A-Za-z0-9_-]*$")
                    self.assertEqual(citation["marker"], "{{CITE:" + citation["id"] + "}}")
                    self.assertNotIn(citation["id"], identifiers)
                    self.assertNotIn(citation["marker"], markers)
                    identifiers.add(citation["id"])
                    markers.add(citation["marker"])
                    self.assertTrue(citation["targets"])
                    self.assertIn(citation.get("display", "default"), {"default", "author-year"})
                    for target in citation["targets"]:
                        self.assertIsInstance(target["rec_number"], int)
                        self.assertGreater(target["rec_number"], 0)
                        self.assertTrue(target["author"])
                        self.assertIsInstance(target["year"], int)
                        self.assertRegex(target["doi"], r"(?i)^10\.\d{4,9}/[!-~]+$")
                        self.assertNotIn("https://doi.org/", target["doi"].lower())
                    if citation.get("display") == "author-year":
                        self.assertEqual(len(citation["targets"]), 1)
                    self.assertNotIn("prefix", citation, "v1 rejects citation prefixes")
                    if "pages" in citation:
                        self.assertRegex(citation["pages"], r"^[0-9]+(\s*-\s*[0-9]+)?$")
                    for key in ("pages", "suffix"):
                        # Both apply to the last target only, so EndNote semantics allow them
                        # on a single-target citation and nowhere else.
                        if key in citation:
                            self.assertEqual(
                                len(citation["targets"]),
                                1,
                                f"'{key}' requires exactly one target",
                            )

    def test_fixture_manifests_cover_the_supported_syntax(self) -> None:
        manifest = json.loads(
            (ROOT / "tests" / "fixtures" / "live-manifest.json").read_text(encoding="utf-8")
        )
        citations = manifest["citations"]
        self.assertTrue(any(len(entry["targets"]) > 1 for entry in citations), "no merged citation")
        self.assertTrue(any("pages" in entry for entry in citations), "no cited-pages citation")
        self.assertTrue(
            any(entry.get("display") == "author-year" for entry in citations),
            "no Author (Year) citation",
        )
        self.assertTrue(
            any(entry.get("display") == "default" for entry in citations),
            "no default-format citation",
        )
        self.assertTrue(
            any(len(entry["targets"]) == 1 and "pages" not in entry for entry in citations),
            "no plain single citation",
        )


class EvalTests(unittest.TestCase):
    def test_eval_files_exist_and_scenarios_are_file_backed(self) -> None:
        data = json.loads((ROOT / "evals" / "evals.json").read_text(encoding="utf-8"))
        self.assertEqual(data["skill_name"], PACKAGE_NAME)
        file_backed = 0
        identifiers = set()
        for evaluation in data["evals"]:
            with self.subTest(eval_id=evaluation["id"]):
                self.assertNotIn(evaluation["id"], identifiers)
                identifiers.add(evaluation["id"])
                self.assertTrue(evaluation["prompt"].strip())
                self.assertTrue(evaluation["expected_output"].strip())
                self.assertTrue(evaluation["expectations"])
                if evaluation["files"]:
                    file_backed += 1
                for relative in evaluation["files"]:
                    self.assertTrue((ROOT / relative).is_file(), relative)
        self.assertGreaterEqual(file_backed, MINIMUM_FILE_BACKED_EVALS)

    def test_evals_cover_the_required_scenarios(self) -> None:
        text = (ROOT / "evals" / "evals.json").read_text(encoding="utf-8")
        data = json.loads(text)
        tags = {tag for evaluation in data["evals"] for tag in evaluation.get("tags", [])}
        for tag in (
            "chinese-draft",
            "english-draft",
            "merged-citations",
            "cited-pages",
            "style-switch",
            "probe-failure",
        ):
            with self.subTest(tag=tag):
                self.assertIn(tag, tags)


if __name__ == "__main__":
    unittest.main()
