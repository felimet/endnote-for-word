---
name: endnote-for-word
description: "Find DOI-verified research papers with the Agent's available scholarly-search or web tools; export EndNote-compatible RIS for users without a library; and insert, format, update, restyle, inspect, or validate genuine EndNote Cite While You Write (CWYW) citations and bibliographies in Microsoft Word DOCX. Use for IEEE citations, visible https://doi.org links, Default Format [x], Display as Author (Year) forms such as Author et al. [x], silent Word and EndNote automation, ADDIN EN.CITE fields with traveling library payloads, ADDIN EN.REFLIST reference lists, rec-number lookups, output-style switching, merged citations, cited pages, and repairs that must avoid typed citation text, fabricated field codes, or fake plain-text bibliographies."
license: MIT
metadata:
  author: Jia-Ming Zhou (Felimet)
  version: 1.1.0
---

# EndNote for Word

Create real EndNote citations and a real bibliography in DOCX by driving EndNote 21 Cite While You Write through Word COM. Treat the source DOCX and the EndNote library as user data: write a new output file unless the user explicitly authorizes in-place replacement, and never write to the library at all.

## Non-negotiable invariants

- Citations must be CWYW-produced `ADDIN EN.CITE` fields carrying a traveling library payload, and the bibliography must be an `ADDIN EN.REFLIST` field. Nothing else counts as a citation.
- Never type citation text, never fabricate field codes in OOXML, and never pass off a plain-text bibliography as a real one. A hand-written `(Author, 2023)` or a typed reference list is a failed deliverable, not a fallback.
- The library is read-only. Every `rec_number` must come from `search_endnote_library` and must pass `precheck_endnote_manifest` before rendering. Never guess a rec-number from a PDF, a DOI, or memory.
- IEEE is the default output style. Every cited target must have a bare DOI such as `10.1109/5.771073`, and `precheck_endnote_manifest` must successfully resolve it through `https://doi.org/` before citation.
- The visible bibliography form is exactly `doi: https://doi.org/<DOI>`. Reject missing DOI links and duplicated resolver prefixes such as `https://doi.org/https://doi.org/...`.
- The output style must be verified with `list_endnote_styles` first. An unknown style name makes EndNote open a modal dialog and hangs the run, and no `DisplayAlerts` setting suppresses it.
- v1 does not support citation prefixes or excluded authors. The backslash forms `{see \Author, 2023 #1}` and `{\Author, 2023 #1}` open modal dialogs, so `Assert-Manifest` rejects a `prefix` key outright.
- Source DOCX is user data. Default to writing a new output file, and only reuse the input path when the user asks for it and `-Overwrite` is supplied.
- A rec-number that does not exist in the target library triggers the blocking "select matching reference" dialog. The pre-check exists to make that impossible, so never skip it to save a step.
- A merged citation such as `{A, 2023 #1; B, 2020 #2}` is one field, not two. Count citations as `count(codes starting with "ADDIN EN.CITE") - count(codes starting with "ADDIN EN.CITE.DATA")`, because a large traveling library payload is moved into a nested `ADDIN EN.CITE.DATA` field.
- Only the measured non-interactive CWYW macros are callable. `ENFormatBibliography`, `ENAuthorYearFormat`, and `ENUnformatCitations` are allowed. Every `*Dlg` macro opens a modal dialog and must never be invoked. `ENRemoveFieldCodes` also raises a modal confirmation dialog and never returns under automation, so converting to plain text stays a manual ribbon operation. `ENToggleInstantFormatting` toggles blindly with no read-back and is not used.
- Word is always created with `Visible = false`; an EndNote process started by the bridge uses a hidden window. Never surface either application during Agent editing.
- Never call `Application.AddIns.Add` on the CWYW template. `EndNote Cwyw.dotm` auto-loads as a global template as soon as a document is open, and adding it again double-loads it.
- `COMAddIns("EndNote.WordAddin.Connect").Object` returns null on this platform. `Application.Run('<macro>')` is the only working entry point.

## Terminal selection

- Detect the active terminal and the available executables before invoking the bridge. The bridge is written to run unchanged on both PowerShell 7 and Windows PowerShell 5.1.
- When the active host is `pwsh.exe` 7 or later, run the bridge directly.
- Windows PowerShell 5.1 (`powershell.exe`) is supported and allowed. Prefer `pwsh.exe` when it is present, but do not refuse to run in 5.1.
- From Bash on Windows, including Git Bash, or from CMD, try `pwsh.exe` first and fall back to `powershell.exe`.
- From WSL Bash, call the Windows-side executable. A Linux `pwsh` cannot automate Windows Office COM.
- If neither `pwsh.exe` nor `powershell.exe` is available, stop and ask the user to install PowerShell using <https://learn.microsoft.com/powershell/scripting/install/installing-powershell-on-windows>. Never silently downgrade the workflow to typed citations.
- Read [installation-matrix.md](references/installation-matrix.md) for cross-shell commands and agent-specific installation paths.

## Mandatory workflow

1. Preserve the input DOCX and choose an explicit output path. Reject anything that is not `.docx`.
2. Call `probe_endnote_word`. Pass `library_path` when the library is already known so the three `library_*` checks are filled in. Stop on a failed prerequisite and report the exact failed check. Exit code 2 means the environment is not ready, not that the document is wrong.
3. Call `list_endnote_styles` and resolve the requested output style to an installed `.ens` name before anything else touches Word. If the user named a style that does not resolve, present the nearest matches and ask; do not substitute a style silently.
4. Scan the complete document twice. First pass inventories every place that needs a citation and every existing citation already present. Second pass classifies each place as a single citation or a merged citation, chooses `display`, decides whether cited pages or a suffix are needed, and assigns one stable id and one unique marker per citation. Never start writing a manifest during the first pass.
5. For every citation, call `search_endnote_library` to find candidate records and `get_endnote_reference` to confirm the exact record. Copy its `rec_number`, `first_author_surname`, year, and bare DOI into the target. A missing DOI makes the record ineligible.
6. Build a schema v1 JSON manifest. Read [workflow-and-schema.md](references/workflow-and-schema.md) for the exact contract and worked examples, and [temporary-citations.md](references/temporary-citations.md) for the supported temporary citation grammar.
7. Call `precheck_endnote_manifest`. It verifies every target against the library and verifies each DOI online through doi.org before Word is opened. Fix every reported problem and re-run it; the pre-check reports all mismatches at once.
8. Call `render_endnote_word_document`. Do not edit the same DOCX in Word while the render runs, and do not run two render jobs concurrently in the same Windows profile.
9. Call `validate_endnote_word_document` with the same manifest. A render is incomplete until validation returns `ok: true`. An output file that exists but fails validation is a diagnostic artifact, not a deliverable.
10. Report the output path, the citation and target counts, the resolved style, whether the bibliography is present, whether EndNote was already running before the run (`endnote_preexisting`), the validation result, and any limitation that applies.

## MCP tools

Use the bundled local MCP server when available:

- `probe_endnote_word`: read-only prerequisite check for Word COM, the EndNote executable, the CWYW add-in registration, the CWYW template load state, macro callability, the styles directory, and the optional library.
- `list_endnote_styles`: read-only enumeration of installed and personal `.ens` output styles, with an optional case-insensitive filter.
- `search_endnote_library`: read-only search over the library across author, title, journal, DOI, keywords, and abstract.
- `get_endnote_reference`: read-only fetch of one record by rec-number.
- `precheck_endnote_manifest`: read-only library verification plus online doi.org verification for every bare DOI.
- `configure_endnote_defaults`: persist the packaged default profile. Invoke only when the user asks to configure defaults or during an explicitly requested installation. It does not touch EndNote's own preferences.
- `render_endnote_word_document`: marker-driven insertion of temporary citations followed by CWYW formatting.
- `validate_endnote_word_document`: read-only structural validation against the gates below.
- `update_endnote_word_fields`: re-run CWYW formatting after the document text changed.
- `set_endnote_citation_style`: switch the output style of an already formatted document.
- `unformat_endnote_word_document`: revert formatted citations to temporary citation text, into a new file only.
- `export_reference_import_file`: write a UTF-8 RIS file so the user can import new references into EndNote themselves.

If MCP is unavailable, invoke the same bridge directly from the plugin root. Resolve every variable from the user's attached or repository files and never reuse a sample path literally.

```powershell
# PowerShell 7 preferred. Replace pwsh.exe with powershell.exe on Windows PowerShell 5.1.
pwsh.exe -NoLogo -NoProfile -NonInteractive -ExecutionPolicy Bypass -File scripts\endnote-word.ps1 -Action probe -LibraryPath $library
pwsh.exe -NoLogo -NoProfile -NonInteractive -ExecutionPolicy Bypass -File scripts\endnote-word.ps1 -Action list-styles -Style $styleFilter
pwsh.exe -NoLogo -NoProfile -NonInteractive -ExecutionPolicy Bypass -File scripts\endnote-word.ps1 -Action render -InputPath $inputDocx -OutputPath $outputDocx -ManifestPath $manifestJson
pwsh.exe -NoLogo -NoProfile -NonInteractive -ExecutionPolicy Bypass -File scripts\endnote-word.ps1 -Action validate -InputPath $outputDocx -ManifestPath $manifestJson
```

```powershell
# The offline library reader is stdlib-only Python and needs no Word and no EndNote.
python scripts\endnote_library.py search --library $library --author "Thomsen" --year 2023 --limit 20
python scripts\endnote_library.py precheck --library $library --manifest $manifestJson
```

## Manifest decisions

- Use stable ASCII ids such as `thomsen2023` or `merged_welfare`. An id matches `^[A-Za-z][A-Za-z0-9_-]*$` and is globally unique.
- Use collision-resistant visible markers such as `{{CITE:thomsen2023}}`. Each marker must occur exactly once in the document. Never reuse a marker.
- One `citations[]` entry with several `targets[]` is a merged citation and renders as one field. Do not emit two adjacent citations when the user wants one merged parenthesis.
- Each target needs an integer `rec_number` greater than zero, a non-empty `author` surname, an integer `year`, and a bare `doi`. The DOI field must not contain `doi:` or a resolver URL.
- Set `display` to `author-year` only when the citation itself is a grammatical subject or named actor, for example `{{CITE:smith}} demonstrated ...`, which IEEE renders as `Smith et al. [1] demonstrated ...`. The marker replaces the author phrase; do not write `Smith {{CITE:smith}}` or the name will be duplicated.
- Set `display` to `default` when the reference supports a claim, normally at the clause or sentence boundary, for example `The method is stable {{CITE:smith}}`, which IEEE renders as `The method is stable [1]`.
- Merged citations always use `display: default`. `author-year` requires exactly one target. When the prose is ambiguous, rewrite the sentence so the grammatical role is explicit before choosing.
- `pages` supplies EndNote cited pages and must match `^[0-9]+(\s*-\s*[0-9]+)?$`. It renders as `@45-46` inside the temporary citation.
- `suffix` supplies free text after the citation, such as `, p. 99`.
- `pages` and `suffix` bind to the last target only in EndNote semantics, so the manifest may only carry them when `targets` has exactly one entry. `Assert-Manifest` enforces that.
- `prefix` is rejected in v1. Read [temporary-citations.md](references/temporary-citations.md) for the reason and the supported alternatives.
- Unknown top-level and per-citation keys are rejected. A typo is an error, not a silently ignored field.
- `library.path` must point at an existing `.enl` file that has a sibling `<name>.Data` directory. The stub `.enl` alone is not a library.
- Set `bibliography.expected` to true whenever a reference list is wanted, which is almost always. It turns on validation gate 4.

## Style policy

- Use `IEEE` unless the user explicitly requests another installed style.
- The custom IEEE bibliography templates must print `doi: https://doi.org/DOI`; the DOI field itself stays bare. Never add a resolver prefix to an already prefixed DOI value.
- The document's output style lives in a Word document variable named `EN.Layout` holding an `<ENLayout>` XML fragment. The bridge reads it, replaces the `<Style>` element, and re-runs `ENFormatBibliography`. No dialog is involved.
- Always resolve the style name against the installed `.ens` files first. `list_endnote_styles` is the authority, not the user's spelling and not a remembered style list.
- Switching style rewrites both the in-text citations and the whole bibliography. Under IEEE, `display: default` yields `[1]`; `display: author-year` invokes EndNote's `ENAuthorYearFormat` and yields a form such as `Thomsen et al. [1]`.
- Use `set_endnote_citation_style` on an already formatted document rather than re-rendering from the source. Re-rendering repeats the marker work for no gain.
- A second document variable named `EN.Libraries` records which library and which record ids were used. Do not edit it by hand; it is the provenance evidence used by validation gate 8.
- Both variables persist in `word/settings.xml` as `w:docVar` entries, so the style and the record ids are also readable offline without Word.
- Read [styles-and-formatting.md](references/styles-and-formatting.md) before diagnosing a style that will not take effect.

## Editing existing documents

- Use markers for deterministic edits. If the document has no markers, create a working copy and insert unique `{{CITE:id}}` markers at the intended ranges before rendering.
- Never use broad find and replace over citation text or field results. A formatted citation's visible text is generated output and will be regenerated.
- After adding, deleting, or reordering text around existing citations, call `update_endnote_word_fields`, then validate again.
- Do not unlink, flatten, or hand-edit `ADDIN EN.CITE`, `ADDIN EN.CITE.DATA`, or `ADDIN EN.REFLIST` fields. The traveling library inside them is what lets EndNote update the document on another machine.
- `unformat_endnote_word_document` is the correct way to get back to editable `{Author, Year #RecNum}` text. It writes a new file and never edits in place.
- Converting citations to plain text is not automated. The macro behind it hangs on a modal dialog, so ask the user to run **EndNote > Convert Citations and Bibliography > Convert to Plain Text** from the Word ribbon, and warn them that the converted copy can never be updated by EndNote again.
- Do not overwrite a locked file. Ask the user to close it or choose another output path.
- Read [field-codes-and-traveling-library.md](references/field-codes-and-traveling-library.md) before inspecting or repairing a field.

## Validation gates

Treat any of these as a failed deliverable. `validate` returns `ok: false` and exit code 3 when any gate fails.

1. Any remaining `{{CITE:` marker text anywhere in the document.
2. Any residual temporary citation text, matched by `\{[^{}]*#\d+[^{}]*\}` over the document text. Leftover braces mean CWYW did not format that citation.
3. A native citation field count, computed with the counting rule above, that does not equal the manifest citation count. Without a manifest the count must still be at least 1.
4. When `bibliography.expected` is true, anything other than exactly one `ADDIN EN.REFLIST` field whose result text is non-empty after trimming.
5. Any cited DOI missing from the bibliography's visible `doi: https://doi.org/<DOI>` text, or any duplicated DOI resolver prefix.
6. A manifest `rec_number` that does not appear as `<rec-number>N</rec-number>` in some EN.CITE traveling library XML together with that target's author surname. When the payload moved into a nested `ADDIN EN.CITE.DATA` field, match against that nested field's own code text.
7. An `author-year` citation whose native field result does not visibly contain its author surname.
8. `Error!`, `Reference source not found`, `{Citation}`, or `{Bibliography}` placeholder strings in the document text.
9. A missing `EN.Layout` document variable, or a style in it that does not match the manifest style case-insensitively.
10. A missing `EN.Libraries` document variable, or one whose `<record-ids>` does not contain every manifest `rec_number`.

`HYPERLINK` fields inside the bibliography are DOI links, not citations. Never count them as citations and never treat their presence or absence as a failure.

## Failure handling

- If probe fails, read [troubleshooting.md](references/troubleshooting.md) and report the exact failed check. Do not silently fall back to typed citations.
- If the run hangs, a modal EndNote or Word dialog is blocking `Application.Run`. `DisplayAlerts = 0` does not suppress add-in dialogs. The MCP watchdog enumerates `WINWORD.EXE` and `EndNote.exe` PIDs before and after the call and kills only the ones the call created. A pre-existing EndNote process is never killed.
- The two usual causes of a hang are a rec-number that does not exist in the library and a style name that does not resolve to an installed `.ens` file. Both are prevented by steps 3 and 7 of the workflow, so a hang means a step was skipped.
- If the bridge reports a timeout, record the isolated PIDs and the cleanup result, re-run `precheck_endnote_manifest`, then retry only the smallest failing citation once before escalating.
- If EndNote was already running when the render started, the bridge reports `endnote_preexisting: true` and does not force it to switch libraries. If the citations formatted against the wrong library, ask the user to open the intended library and re-run rather than closing their EndNote for them.
- If a needed reference is genuinely absent from the library, do not invent a rec-number. Use `export_reference_import_file` to produce a RIS file and ask the user to import it, then search again. Read [library-import.md](references/library-import.md).
- Word COM is single-user desktop automation. It needs an unlocked interactive desktop session and one job at a time in a given Windows profile.
- Console output on Windows may default to a legacy code page. Every entry point forces UTF-8; if output is mojibake, the caller bypassed an entry point.
- Read [cwyw-entrypoints.md](references/cwyw-entrypoints.md) before attempting any CWYW call the bridge does not already wrap.

## Reference routing

- Manifest contract, schema v1 rules, and end-to-end examples: [workflow-and-schema.md](references/workflow-and-schema.md)
- Supported and unsupported temporary citation syntax, cited pages, suffixes, merged citations: [temporary-citations.md](references/temporary-citations.md)
- Field anatomy for `ADDIN EN.CITE`, nested `ADDIN EN.CITE.DATA`, `ADDIN EN.REFLIST`, and the counting rule: [field-codes-and-traveling-library.md](references/field-codes-and-traveling-library.md)
- CWYW macro inventory, the working `Application.Run` entry point, and the dead `COMAddIns` route: [cwyw-entrypoints.md](references/cwyw-entrypoints.md)
- Output style resolution, the `EN.Layout` and `EN.Libraries` document variables, and style switching: [styles-and-formatting.md](references/styles-and-formatting.md)
- Library storage, the read-only SQLite reader, search fields, and RIS export for import: [library-import.md](references/library-import.md)
- Terminal selection plus Claude Code, Claude Desktop, Codex, and ChatGPT Desktop setup: [installation-matrix.md](references/installation-matrix.md)
- Errors, hangs, watchdog recovery, and the fallback boundary: [troubleshooting.md](references/troubleshooting.md)
