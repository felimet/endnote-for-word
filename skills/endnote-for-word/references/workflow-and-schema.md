# Workflow and manifest schema

## Contents

1. End to end flow
2. Manifest schema v1
3. Assert-Manifest rules
4. Temporary citation text generated from a manifest entry
5. Complete worked example
6. Bridge command lines for every action
7. Exit codes and environment variables
8. Output and validation contract

## End to end flow

The bridge performs a deterministic transaction. Steps 1 to 4 run before Word is ever
opened, so that a bad manifest can never reach the point where EndNote raises a modal
dialog and blocks the run.

1. Resolve and validate the paths. A non `.docx` input or output is rejected.
2. `Read-Manifest` and `Assert-Manifest`.
3. Verify the requested style resolves to an installed `.ens` file. When it does not, fail
   hard and report the nearest matches. An unknown style name makes EndNote open a modal
   dialog, which hangs the automation permanently.
4. Verify the library file exists together with its sibling `.Data` directory.
5. `Start-EndNoteWithLibrary`. When no `EndNote` process owns a main window, start the
   EndNote executable with the library path and poll every 250 ms until a process reports
   `MainWindowHandle -ne 0`. When EndNote is already running, do not start a second
   instance and do not force it to switch libraries; the result records
   `endnote_preexisting: true`. This step is not optional even though
   `ENFormatBibliography` was measured to work with EndNote closed, because CWYW then
   auto-launches EndNote with the most recently used library, and rec-numbers are
   library-local, so the wrong library silently resolves the wrong records. When the bridge
   observes an EndNote process that did not exist before the call, it records
   `endnote_autostarted: true`.
6. `Start-Word` on a copy of the input written to the output temporary path. The input is
   never opened for writing.
7. Set the `EN.Layout` document variable to the requested style. The variable is created
   when missing, otherwise only its `<Style>` element is replaced.
8. For each citation, `Find-UniqueMarkerRange` locates the marker and the range text is set
   to the generated temporary citation string.
9. `Invoke-CwywFormat` calls `ENFormatBibliography` through `Application.Run`.
10. Poll until the EN.CITE field count reaches the expected value or the automation timeout
    expires.
11. Save atomically: `SaveAs2` to a `.<name>.<guid>.tmp.docx` sibling, stop Word, then move
    the temporary file onto the output path.

Source documents and libraries are user data. The bridge never writes to `InputPath`
unless `-Overwrite` is supplied, and never writes to the library at all.

## Manifest schema v1

```jsonc
{
  "schema_version": 1,
  "library": { "path": "C:\\Users\\<you>\\Documents\\ProjectLibrary.enl" },
  "style": "APA 7th",
  "citations": [
    {
      "id": "thomsen2023",
      "marker": "{{CITE:thomsen2023}}",
      "targets": [ { "rec_number": 1, "author": "Thomsen", "year": 2023 } ],
      "pages": "45-46",
      "suffix": ", p. 99"
    },
    {
      "id": "merged",
      "marker": "{{CITE:merged}}",
      "targets": [
        { "rec_number": 2, "author": "Magrin",   "year": 2023 },
        { "rec_number": 3, "author": "Bernhard", "year": 2020 }
      ]
    }
  ],
  "bibliography": { "expected": true }
}
```

The build spec does not state whether `bibliography` is required or optional, so treat it
as present in every manifest you author and do not rely on a default. Validation gate 4
only fires when `bibliography.expected` is true.

The first entry above carries both `pages` and `suffix`, mirroring the build spec's own
example. That combination is permitted by `Assert-Manifest` and by the string building
algorithm, but it is not one of the measured rows in the syntax table, so its rendered
output is unverified. Prefer one or the other. See `temporary-citations.md`.

## Assert-Manifest rules

Every rule below is a hard failure on the bridge side. There are no warnings in
`Assert-Manifest`.

| Rule | Requirement |
|---|---|
| `schema_version` | Must be exactly `1`. |
| `citations` | Must be a non empty array. |
| `id` | Must match `^[A-Za-z][A-Za-z0-9_-]*$` and be globally unique. |
| `marker` | Must be globally unique and must occur exactly once in the document. |
| `targets` | Must be non empty. |
| `targets[].rec_number` | Integer greater than 0. |
| `targets[].author` | Non empty string. |
| `targets[].year` | Integer. |
| `pages` | Optional string, must match `^[0-9]+(\s*-\s*[0-9]+)?$`. |
| `suffix` | Optional string. |
| `pages` and `suffix` arity | Either key may only be supplied when `targets` has exactly one entry, because EndNote applies both to the last target only. |
| `prefix` | Rejected in v1. The error message points at `temporary-citations.md`, because the backslash prefix syntax opens a modal dialog. |
| `library.path` | Must exist and must have a sibling `<name>.Data` directory. |
| `style` | Must be a non empty string. |
| Unknown keys | Any unknown top level key or unknown per citation key is rejected. |

Two further pre conditions are enforced outside `Assert-Manifest`, both before Word opens:

1. Every `rec_number` must exist in the target library. Use `precheck_endnote_manifest`,
   which reads the library SQLite store directly. A citation pointing at a record that does
   not exist makes EndNote raise a "select matching reference" dialog and hang.
2. The requested style name must resolve to an installed `.ens` file. Use
   `list_endnote_styles` first.

## Temporary citation text generated from a manifest entry

```
{ + join("; ", for each target: "<author>, <year> #<rec_number>")
  + [ "@" + pages ]
  + [ suffix ]
  + }
```

Worked results:

| Manifest entry | Generated text |
|---|---|
| One target, rec 1, Thomsen, 2023 | `{Thomsen, 2023 #1}` |
| One target plus `pages: "45-46"` | `{Thomsen, 2023 #1@45-46}` |
| One target plus `suffix: ", p. 99"` | `{Thomsen, 2023 #1, p. 99}` |
| Two targets, rec 2 and rec 3 | `{Magrin, 2023 #2; Bernhard, 2020 #3}` |

`temporary-citations.md` holds the measured syntax table and the full string building
algorithm, including the forms that are rejected in v1.

## Complete worked example

Source document text, three markers, one per sentence:

```text
Automated lameness scoring has been validated on commercial herds {{CITE:thomsen2023}}.

Two independent replications reported comparable agreement {{CITE:merged}}.

The scoring rubric is reproduced verbatim from the original protocol {{CITE:rubric_pages}}.
```

Manifest, covering a single citation, a merged citation and a cited pages citation:

```json
{
  "schema_version": 1,
  "library": { "path": "C:\\Users\\<you>\\Documents\\ProjectLibrary.enl" },
  "style": "APA 7th",
  "citations": [
    {
      "id": "thomsen2023",
      "marker": "{{CITE:thomsen2023}}",
      "targets": [ { "rec_number": 1, "author": "Thomsen", "year": 2023 } ]
    },
    {
      "id": "merged",
      "marker": "{{CITE:merged}}",
      "targets": [
        { "rec_number": 2, "author": "Magrin", "year": 2023 },
        { "rec_number": 3, "author": "Bernhard", "year": 2020 }
      ]
    },
    {
      "id": "rubric_pages",
      "marker": "{{CITE:rubric_pages}}",
      "targets": [ { "rec_number": 1, "author": "Thomsen", "year": 2023 } ],
      "pages": "45-46"
    }
  ],
  "bibliography": { "expected": true }
}
```

Two manifest entries may cite the same `rec_number`. Uniqueness is required of `id` and
`marker`, not of the records they point at.

Intermediate document text after step 8, before `ENFormatBibliography` runs:

```text
Automated lameness scoring has been validated on commercial herds {Thomsen, 2023 #1}.

Two independent replications reported comparable agreement {Magrin, 2023 #2; Bernhard, 2020 #3}.

The scoring rubric is reproduced verbatim from the original protocol {Thomsen, 2023 #1@45-46}.
```

Expected output structure after `ENFormatBibliography`:

- Three `ADDIN EN.CITE` fields. The merged citation is one field, not two.
- Zero or more nested `ADDIN EN.CITE.DATA` fields. EndNote creates one whenever a
  traveling library payload is too large to sit inside the parent field code. Their number
  is not predictable from the manifest, so never assert on it.
- Exactly one `ADDIN EN.REFLIST` field carrying the bibliography, at the end of document.
- Possibly several `HYPERLINK` fields inside the bibliography for DOI links. They are not
  citations and must not be counted as such.
- `EN.Layout` reads back `<Style>APA 7th</Style>`.
- `EN.Libraries` lists record ids 1, 2 and 3.

The displayed citation text depends on the active output style. The build spec records
measured display strings such as `(Thomsen et al., 2023)` and
`(Thomsen et al., 2023, pp. 45-46)` in section 0.5, but does not record which style
produced them, so do not treat any particular rendering as guaranteed for `APA 7th`.
The spec does record one verified style round trip: switching from `Annotated` to `IEEE`
changed the body text from `(Thomsen, Shearer et al. 2023)` to `[1]` and rewrote the whole
bibliography.

## Bridge command lines for every action

The bridge is `scripts\endnote-word.ps1` at the repository root. It runs on both
PowerShell 7 (`pwsh.exe`) and Windows PowerShell 5.1 (`powershell.exe`); substitute the
executable name as the terminal policy in `SKILL.md` dictates and leave every other token
unchanged. Prefer `pwsh.exe` when both are present.

The skill directory also ships `scripts/invoke-bridge.ps1`. The build spec lists that file
in the repository layout but does not define its flag surface, so the command lines below
address the repository bridge directly.

```powershell
# Readiness probe, no document required. Exit code 2 means not ready.
pwsh.exe -NoLogo -NoProfile -NonInteractive -ExecutionPolicy Bypass -File scripts\endnote-word.ps1 -Action probe

# Probe including the three library_* checks.
pwsh.exe -NoLogo -NoProfile -NonInteractive -ExecutionPolicy Bypass -File scripts\endnote-word.ps1 -Action probe -LibraryPath $libraryEnl

# Enumerate installed output styles.
pwsh.exe -NoLogo -NoProfile -NonInteractive -ExecutionPolicy Bypass -File scripts\endnote-word.ps1 -Action list-styles

# Enumerate styles matching a case insensitive substring.
pwsh.exe -NoLogo -NoProfile -NonInteractive -ExecutionPolicy Bypass -File scripts\endnote-word.ps1 -Action list-styles -Style "IEEE"

# Install the packaged defaults into %APPDATA%\EndNoteForWordAgent\defaults.json.
pwsh.exe -NoLogo -NoProfile -NonInteractive -ExecutionPolicy Bypass -File scripts\endnote-word.ps1 -Action configure-defaults

# Render. Library and style come from the manifest.
pwsh.exe -NoLogo -NoProfile -NonInteractive -ExecutionPolicy Bypass -File scripts\endnote-word.ps1 -Action render -InputPath $inputDocx -OutputPath $outputDocx -ManifestPath $manifestJson

# Render over an existing output file.
pwsh.exe -NoLogo -NoProfile -NonInteractive -ExecutionPolicy Bypass -File scripts\endnote-word.ps1 -Action render -InputPath $inputDocx -OutputPath $outputDocx -ManifestPath $manifestJson -Overwrite

# Validate against the manifest. Exit code 3 means at least one gate failed.
pwsh.exe -NoLogo -NoProfile -NonInteractive -ExecutionPolicy Bypass -File scripts\endnote-word.ps1 -Action validate -InputPath $outputDocx -ManifestPath $manifestJson

# Validate without a manifest. Only the manifest independent gates run and a warning is recorded.
pwsh.exe -NoLogo -NoProfile -NonInteractive -ExecutionPolicy Bypass -File scripts\endnote-word.ps1 -Action validate -InputPath $outputDocx

# Re-run ENFormatBibliography into a new file.
pwsh.exe -NoLogo -NoProfile -NonInteractive -ExecutionPolicy Bypass -File scripts\endnote-word.ps1 -Action update -InputPath $inputDocx -OutputPath $outputDocx

# Re-run ENFormatBibliography in place. -Overwrite is mandatory for in place writes.
pwsh.exe -NoLogo -NoProfile -NonInteractive -ExecutionPolicy Bypass -File scripts\endnote-word.ps1 -Action update -InputPath $documentDocx -Overwrite

# Switch the output style, then reformat.
pwsh.exe -NoLogo -NoProfile -NonInteractive -ExecutionPolicy Bypass -File scripts\endnote-word.ps1 -Action set-style -InputPath $inputDocx -OutputPath $outputDocx -Style "IEEE"

# Revert formatted citations to temporary citation text. New file only.
pwsh.exe -NoLogo -NoProfile -NonInteractive -ExecutionPolicy Bypass -File scripts\endnote-word.ps1 -Action unformat -InputPath $inputDocx -OutputPath $outputDocx
```

There are eight actions and no more:
`probe`, `list-styles`, `configure-defaults`, `render`, `validate`, `update`, `set-style`,
`unformat`.

There is deliberately **no `remove-field-codes` action**. Stripping field codes to plain
text would need the `ENRemoveFieldCodes` macro, and that macro was measured to raise a
modal confirmation dialog and never return under automation, which is unrecoverable. It
was removed from the v1 surface. When a user needs a plain text copy with no field codes,
they do it themselves in Word through the EndNote ribbon: **EndNote tab, Convert Citations
and Bibliography, Convert to Plain Text**, which produces a new document and leaves the
original intact. Explain that path to the user; do not try to automate it and do not offer
a bridge action or an MCP tool for it.

Action notes that the command lines alone do not convey:

- `unformat` refuses to write in place even when `-Overwrite` is supplied. It always needs
  a distinct `OutputPath`.
- `set-style` reports both `previous_style` and `style`.
- `configure-defaults` never touches EndNote's own registry preferences. EndNote's UI state
  belongs to the user.
- Exactly one line of compact JSON appears on stdout and it is the last stdout line. All
  logging goes to stderr.

## Exit codes and environment variables

| Exit code | Meaning |
|---:|---|
| 0 | Success. |
| 1 | Unhandled error. |
| 2 | `probe` ran but the environment is not ready. |
| 3 | `validate` ran and at least one gate failed. |

| Variable | Default | Effect |
|---|---|---|
| `ENDNOTE_READY_TIMEOUT_MS` | 90000 | How long `Start-EndNoteWithLibrary` polls for an EndNote main window. |
| `ENDNOTE_AUTOMATION_TIMEOUT_MS` | 60000 | How long the render step polls for the expected EN.CITE field count. |
| `ENDNOTE_WORD_POWERSHELL` | unset | Overrides which PowerShell executable the MCP server spawns. |
| `ENDNOTE_WORD_TIMEOUT_SECONDS` | 240 | MCP server subprocess timeout before the watchdog kills orphaned processes. |
| `ENDNOTE_LIVE_TEST` | unset | Set to `1` to enable the live integration test. |
| `ENDNOTE_TEST_LIBRARY` | unset | Library path used by the live integration test. |

## Output and validation contract

A successful render response is not a substitute for validation. The build spec requires
the render result to carry these fields:

| Field | Meaning |
|---|---|
| `citations` | Number of manifest citations processed. |
| `targets` | Number of individual targets across those citations. |
| `style` | Style written into `EN.Layout`. |
| `bibliography_present` | Whether an `ADDIN EN.REFLIST` field exists. |
| `endnote_preexisting` | True when EndNote was already running and the bridge did not start it. |
| `endnote_autostarted` | True when a new EndNote process appeared during the call, whether the bridge started it or CWYW auto-launched it. |

Always run `validate` afterwards. It applies eight gates, and any failure sets
`ok: false` and exit code 3:

1. No `{{CITE:` marker text remains anywhere in the document.
2. No residual temporary citation text remains. The check is the regex
   `\{[^{}]*#\d+[^{}]*\}` over `Document.Content.Text`.
3. `native_citation_fields` equals the manifest citation count when a manifest is supplied,
   otherwise it must be at least 1. The counting rule is in
   `field-codes-and-traveling-library.md`.
4. When `bibliography.expected` is true, exactly one `ADDIN EN.REFLIST` field exists and
   its `Result.Text` is non empty after trimming.
5. Every manifest `rec_number` appears in some EN.CITE traveling library XML as
   `<rec-number>N</rec-number>`, and its `author` surname appears in the same field's XML.
   A record whose payload was moved into a nested `ADDIN EN.CITE.DATA` field is matched
   against that nested field's own code text.
6. Document text contains none of `Error!`, `Reference source not found`, `{Citation}` or
   `{Bibliography}`.
7. The `EN.Layout` document variable exists and its style equals the manifest `style`,
   compared case insensitively, when a manifest is supplied.
8. The `EN.Libraries` document variable exists and contains every manifest `rec_number`
   inside `<record-ids>`.

Report shape:

```jsonc
{ "ok": true, "action": "validate", "document_path": "...",
  "counts": { "native_citation_fields": 2, "citation_data_fields": 1,
              "bibliography_fields": 1, "hyperlink_fields": 3 },
  "style": "IEEE", "libraries": [{"db_id": "...", "name": "...", "record_ids": [1,2,3]}],
  "errors": [], "warnings": [] }
```
