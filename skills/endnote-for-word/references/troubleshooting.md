# Troubleshooting

## Contents

1. Quick table
2. The modal dialog hang and watchdog recovery
3. Unknown style name
4. A rec-number that does not exist
5. EndNote is not running
6. The wrong library is open
7. Locked output file
8. cp950 console mangling
9. Orphaned WINWORD processes
10. Library reader failures
11. Validation failures
12. Recovery rules
13. Fallback boundary

## Quick table

| Symptom | Cause | Fix |
|---|---|---|
| Bridge call never returns, Word visible in Task Manager | A modal EndNote dialog is blocking `Application.Run` | Watchdog kills the isolated Word PID. Find which pre-check was skipped. See section 2. |
| Bridge fails before Word opens, message lists similar style names | The manifest style does not resolve to an installed `.ens` | Use a name from `list_endnote_styles`. See section 3. |
| Bridge fails before Word opens, message names a `rec_number` | The record is absent or in the trash | Re-query the library, fix the manifest, re-run `precheck`. See section 4. |
| `probe` returns `endnote_running: false` | EndNote desktop has no process with a main window | `render` starts it. For other actions start it manually. See section 5. |
| Citations format against records that are not the ones you chose | EndNote has a different library open | Close the other library, open the manifest library, re-render. See section 6. |
| Save fails with an access or sharing error | The output DOCX is open in Word or held by another process | Close it or pick a new output path. See section 7. |
| `powershell_notice` reports `failed` or `unavailable` | The automatic PowerShell 7 install could not run or finish | The bridge keeps working under Windows PowerShell 5.1. Install PowerShell 7 by hand, or delete the state file the notice names to retry once. See installation-matrix.md. |
| Chinese text in the console renders as garbage | Console code page is cp950 | Force UTF-8 in the terminal. See section 8. |
| `WINWORD.EXE` remains after a failed job | A COM call blocked and the process was orphaned | Kill only the PID the failed job created. See section 9. |
| Timeout whose `stderr_tail` ends at `Saving temporary DOCX` | Word's `SaveAs2` never returned; on Windows PowerShell 5.1 it followed passing uncast `Join-Path` output as the file name | Update the bridge, prefer PowerShell 7, and clear Word's recovery state after the kill. See section 9. |
| `no such collation sequence: ENCI_Base` | The SQLite reader queried without registering the custom collations | Register the two no-op collations before querying. See section 10. |
| Library read fails with a lock or malformed database error | The original `sdb.eni` was opened, or EndNote was mid-write | Copy to a temp file first, open the copy read-only, retry. See section 10. |
| `validate` exits 3 with `{{CITE:` still present | A marker was never replaced | The marker did not occur exactly once in the document. See section 11. |
| `validate` exits 3 on the citation count | A merged citation was counted twice | Apply the counting rule, subtracting `ADDIN EN.CITE.DATA` fields. See section 11. |
| `validate` exits 3 on style | `EN.Layout` style does not equal the manifest style | Run `set-style`, then validate again. |
| Manifest rejected for a `prefix` key | Prefix and excluded author syntax are not supported in v1 | Both open modal dialogs. See `temporary-citations.md`. |
| Precheck reports `library_doi_not_canonical` | The EndNote DOI field contains `doi:` or a resolver URL | Keep only the bare DOI in EndNote, then rerun precheck. |
| Precheck reports `doi_not_resolvable` | doi.org does not return a successful handle | Correct the DOI from an authoritative source; do not render while the gate fails. |
| Bibliography validation reports a missing DOI link | The active style's Bibliography Template does not print `doi: https://doi.org/DOI` for that reference type | Correct the style with Save As, select the corrected style, rerender, and validate. |
| Bibliography validation reports a duplicated resolver | A prefixed EndNote DOI was combined with the style's resolver literal | Replace the record DOI with its bare value and rerender. |
| MCP client reports a protocol parse error | Something wrote to stdout that was not JSON-RPC | Protocol output on stdout must be JSON only; all diagnostics belong on stderr. |
| `render` reports fewer citations than the manifest declares | `ENFormatBibliography` did not finish inside the automation timeout | Raise `ENDNOTE_AUTOMATION_TIMEOUT_MS`, re-run, and check for a dialog. |
| Bibliography is missing after a style switch | `ENFormatBibliography` was interrupted | Run `update`, then validate again. |
| Citations became plain text and cannot be updated | The EndNote ribbon command Convert to Plain Text was run on that copy | Irreversible in that file. Go back to the version with field codes. See section 13. |

## The modal dialog hang and watchdog recovery

This is the dominant failure mode of the whole system, so treat it as the default hypothesis
whenever a call does not return. The exception: when `stderr_tail` ends at
`Saving temporary DOCX` with no `Word SaveAs2 returned` line, no dialog is involved; see
section 9.

**What happens.** An EndNote add-in dialog opens in the Word instance the bridge is driving.
The dialog owns the message loop. `Application.Run('ENFormatBibliography')` never returns.
The bridge is blocked inside a COM call, so it cannot time out on its own and cannot close
the dialog.

**What does not help.** `Application.DisplayAlerts = 0` suppresses Word's own alerts. It has
no effect on dialogs raised by the CWYW add-in. There is no property, no registry value and
no macro in the CWYW inventory that disables them.

**What does help.** The watchdog in the MCP layer:

1. Before the bridge call, enumerate all `WINWORD.EXE` process ids.
2. Run the bridge with the explicit `ENDNOTE_WORD_TIMEOUT_SECONDS` when set. Otherwise use
   240 seconds, with render automatically scaled to `max(240, 60 + 8 * citation_count)` and
   capped at 1800 seconds.
3. On timeout, enumerate `WINWORD.EXE` again and compute the difference. Terminate only the
   process ids that appeared during this call.
4. Do the same for `EndNote.exe`, and never terminate an EndNote process that already
   existed before the call. That process is the user's session and may hold unsaved state.

The result reports `isolated_word_pid` and whether the isolated process was terminated. Read
those fields before retrying anything.

**Known triggers, all measured.** Every one of them has a pre-check that runs before Word is
opened, so a hang inside `Application.Run` means a pre-check was skipped or bypassed:

| Trigger | Guard |
|---|---|
| A `rec_number` that does not exist in the library | `precheck_endnote_manifest`, plus the bridge's own library check |
| A style name that does not resolve to an installed `.ens` | `Resolve-StyleName` at render step 3 |
| Backslash prefix syntax, for example `{see \Thomsen, 2023 #1}` | `Assert-Manifest` rejects the `prefix` key in v1 |
| Excluded author syntax, for example `{\Thomsen, 2023 #1}` | Same rejection; the bridge never emits a backslash form |
| Any `*Dlg` macro, for example `ENFormatBibliographyDlg` | The bridge calls no `*Dlg` macro at all |

**After a hang.** Do not immediately re-run the same job. Clear the Word recovery state the
watchdog kill left behind (section 9). Establish which guard was not satisfied, fix the
manifest or the style, then re-run the smallest possible fixture before the real document.

## Unknown style name

**Symptom.** With the guard in place: the bridge exits before opening Word, with a message
naming the style that failed and the nearest matching installed names. Without the guard, or
if a style is deleted between the check and the run: a modal dialog and a hang.

**Cause.** The style string in the manifest, or in `-Style`, does not match the base name of
any `.ens` file in the installed styles directory or the personal styles directory.

**Fix.**

```console
pwsh.exe -NoLogo -NoProfile -ExecutionPolicy Bypass -File "<REPO_ROOT>\scripts\endnote-word.ps1" -Action list-styles -Style "APA"
```

Take the name exactly as returned. `APA 7th` and `APA7` are different styles and only one of
them may exist. Comparison is case-insensitive, but spacing and ordinal suffixes are part of
the name. Never guess a style name because the format is well known; the machine only has
the styles that are installed on it. Details in `styles-and-formatting.md`.

## A rec-number that does not exist

**Symptom.** With the guard in place: `precheck` or the bridge fails with an explicit
message naming every bad target. Without it: EndNote opens a "select matching reference"
dialog and the run hangs.

**Cause.** One of these:

- The record was never in the library.
- The record is in the trash. The reader filters `trash_state = 0`, so a trashed record
  behaves exactly like a missing one.
- The `rec_number` was guessed or computed rather than read from a search result.
- The manifest points at a different library from the one being searched.

**Fix.**

```console
python scripts\endnote_library.py search   --library "<LIBRARY>.enl" --author Thomsen --year 2023
python scripts\endnote_library.py precheck --library "<LIBRARY>.enl" --manifest "<PATH>\manifest.json"
```

`precheck` reports every problem it finds, not just the first, so one round trip fixes a
manifest with several bad targets. It also checks that the record's first author surname and
year match the manifest, accent-tolerantly and case-insensitively. A mismatch there is a
real warning: it usually means the manifest was written against a different library or a
stale search result.

The `rec_number` is the SQLite `refs.id` primary key. It is assigned by EndNote and is never
predictable. See `library-import.md`.

## EndNote is not running

**Symptom.** `probe` reports `endnote_running: false`. Or `render` waits and then fails with
a readiness timeout.

**Cause.** No `EndNote` process has a main window handle. On its own that does not stop a
format: `ENFormatBibliography` was measured to complete with EndNote closed, because CWYW
silently auto-launches EndNote with the most recently used library. The risk is which
library that auto-launch picks. Rec-numbers are library-local, so resolving citations
against the most recently used library instead of the manifest library produces a wrong
document rather than an error, which is why the bridge still opens the manifest library
explicitly. When a new EndNote process appears during a call, the result reports
`endnote_autostarted: true`.

**Fix.** `render` handles this itself: when no EndNote process has a main window, it starts
the EndNote executable with the library path and polls every 250 milliseconds up to
`ENDNOTE_READY_TIMEOUT_MS`, default 90000, for a process whose `MainWindowHandle` is not
zero. Raise that value on a slow machine or a very large library.

Two behaviours to know about:

- When EndNote is already running, the bridge does **not** start a second instance and does
  **not** force it to switch libraries. The result records `endnote_preexisting: true`.
  Forcing a library switch would disrupt the user's session and could discard unsaved edits.
- The readiness poll checks for a main window, not merely for a process. An EndNote process
  that is still initialising cannot service CWYW requests.

For actions other than `render`, start EndNote and open the library first.

## The wrong library is open

**Symptom.** The render succeeds and validation passes, but the formatted citations name
authors and years that are not the ones intended, or the traveling library payload carries a
`db-id` that does not belong to the manifest library.

**Cause.** `endnote_preexisting` was true and the already running EndNote had a different
library open. CWYW resolves a temporary citation against whatever library is open, and
`rec_number` values are per library. Record 1 exists in almost every library, so the wrong
library rarely produces an error; it produces a wrong document that looks right.

**Fix.**

1. Ask the user to close the other library and open the one the manifest names.
2. Re-run `render`.
3. Confirm with `validate`. Gate 8 requires the `EN.Libraries` document variable to contain
   every manifest `rec_number` inside `<record-ids>`, and the variable also records the
   library `db-id` and name. Compare that name against the manifest library.

This is the failure mode with the highest chance of reaching a submitted manuscript
unnoticed, which is why `EN.Libraries` is checked as a gate rather than reported as
information.

## Locked output file

**Symptom.** The save step fails with a sharing violation or an access denied error, or the
atomic move at the end of the job fails.

**Cause.** The output DOCX is open in Word, in a preview pane, or held by a sync client or a
virus scanner.

**Fix.** Close the document and retry, or choose a different output path. Note what the
bridge does and does not do here:

- It never writes to `InputPath` unless `-Overwrite` is passed explicitly.
- It saves through a temporary sibling file named `.<name>.<guid>.tmp.docx`, closes Word,
  and then moves the temporary file into place. A failed job therefore leaves the previous
  output intact rather than half written.
- `unformat` refuses to write in place at all, even with `-Overwrite`, because it is a
  destructive transformation of user content.

On the next run for the same destination, the bridge safely sweeps abandoned
`.<name>.<guid>.work.docx` and `.<name>.<guid>.tmp.docx` siblings older than 24 hours. It checks
the exact destination-specific name, refuses directories and reparse points, deletes only from
the destination directory, and verifies that each removed file is gone. A newer artifact is
left alone because another run may still own it.

## cp950 console mangling

**Symptom.** Chinese or accented characters appear as question marks or as unrelated
glyphs. JSON output looks corrupted even though the underlying data is fine.

**Cause.** The console code page on this class of machine defaults to cp950, and the
terminal decodes UTF-8 bytes as cp950.

**Fix.** Every entry point in this project already forces UTF-8 on its own output, inside a
try/catch because Windows PowerShell 5.1 can throw when stdout is redirected. What remains is
the terminal itself:

| Terminal | Command |
|---|---|
| PowerShell 7 and 5.1 | `[Console]::OutputEncoding = New-Object Text.UTF8Encoding $false` |
| CMD | `chcp 65001` |
| Git Bash | Usually already UTF-8. If not, set `LANG=en_US.UTF-8`. |

Do not fix mangled output by transliterating data. If a file was written with the wrong
encoding, regenerate it. The RIS exporter writes UTF-8 with CRLF line endings, and the
automated importer rejects invalid UTF-8 before EndNote opens. See `library-import.md`.

## Orphaned WINWORD processes

**Symptom.** `WINWORD.EXE` entries remain in Task Manager after a failed or timed out job.
Subsequent runs behave oddly, or Word prompts about document recovery.

**Cause.** A COM call blocked, the bridge could not reach its `finally` block cleanly, or the
watchdog terminated the process without Word getting a chance to shut down.

**Fix.** Kill only the process the failed job created. The MCP layer enumerates
`WINWORD.EXE` process ids before and after each call precisely so that this difference is
known, and it reports the isolated process id in the result.

Never kill every Word process. The user very likely has their manuscript open in another
Word window, and terminating it discards their unsaved work. The same applies with more
force to `EndNote.exe`: only a process that this call started may be terminated, never a
pre-existing one.

After cleaning up:

1. Confirm no `.tmp.docx` sibling remains next to the output path.
2. Confirm the input document is untouched.
3. Re-run the smallest fixture before the real job.

**A save that never returns.** When a timed out result shows `stderr_tail` ending at
`Saving temporary DOCX` with no `Word SaveAs2 returned` line, the job stalled inside Word's
own `SaveAs2`. What was measured on Word 16.0.20430 with EndNote 21.5: under Windows
PowerShell 5.1, every run that passed the output of `Join-Path` straight to `SaveAs2` stalled
(part of one core busy, no file I/O), with or without EndNote fields in the document. The
same path cast to `[string]`, or built by string concatenation, saved at once, and under
PowerShell 7 both forms saved. Why Word stalls is not established: the value is wrapped the
same way on both hosts, and a stand-in COM server that presents Word's own type information
receives the same by-value string in every form. The bridge casts the path, and a static
test keeps every save call that way. If you still see this stall:

1. Confirm the bridge in use contains the cast in `Save-DocumentAtomically`.
2. Prefer PowerShell 7. The bridge host is `pwsh.exe` unless `ENDNOTE_WORD_POWERSHELL` names
   another one; the probe result's `checks.powershell` field shows the version in use.
3. Report the PowerShell edition, the Word and EndNote builds, and whether the same render
   completes under the other PowerShell host.

**Word misbehaves after a forced kill.** Terminating a Word process that had a document open
leaves an entry under `HKCU\Software\Microsoft\Office\16.0\Word\Resiliency\DocumentRecovery`.
Observed once, with no Word process left running: while five such entries existed, every
automated run failed early with "the 'EndNote Cwyw.dotm' global template is not loaded";
after they were cleared, runs passed that check again. To clear them, open Word, close the
documents listed in the Document Recovery pane without saving, and quit Word. Alternatively,
with Word closed, back up the key with `reg export` and delete only the `DocumentRecovery`
subkeys that point at the job's `.work.docx` or `.tmp.docx` files.

`isolated_word_process_terminated: true` only means `taskkill /F` accepted the request, not
that the process exited. Check `tasklist /FI "PID eq <isolated_word_pid>"`; a PID still
listed after a few minutes is blocked in kernel mode, and usually only a reboot clears it.

## Library reader failures

| Symptom | Cause | Fix |
|---|---|---|
| `no such collation sequence: ENCI_Base` or `ENCIN_zh_TW` | The `refs` schema declares custom collations that plain SQLite does not implement | Register both as no-op collations on the connection before any query. |
| `no such function: EN_MAKE_SORT_KEY` | Same class of problem, a custom SQL function | Do not issue queries that call it. The reader's own queries do not need it. |
| Database is locked | The original `sdb.eni` was opened while EndNote had it open | Copy the file to a temporary path first and open the copy with `mode=ro`. Delete the copy in a `finally` block. |
| Database disk image is malformed | The copy was taken while EndNote was mid-write | Let EndNote finish, then copy and read again. The failure is safe: it aborts a read, it cannot damage the original. |
| `.Data` directory not found | The `.enl` file was moved or copied without its sibling directory | Both `<name>.enl` and `<name>.Data` must be present. The `.enl` on its own is a stub. |
| Search returns nothing for a reference the user can see in EndNote | The record is in the trash, or a different library is being searched | The reader filters `trash_state = 0`. Confirm the library path. |

The reader never opens the original database file, under any error condition. If you are
debugging a reader problem and are tempted to open the original just once to look at it,
open a copy instead.

## Validation failures

`validate` exits 3 and sets `ok: false` when any gate fails. Each gate maps to a distinct
cause:

| Gate | Failure meaning | First thing to check |
|---|---|---|
| 1 | A `{{CITE:` marker is still in the document | The marker did not occur exactly once, so it was never replaced. |
| 2 | Temporary citation text like `{Thomsen, 2023 #1}` remains | `ENFormatBibliography` did not run, or it stopped partway. |
| 3 | Citation field count does not match the manifest | Apply the counting rule below before concluding the render failed. |
| 4 | Bibliography missing or empty | Formatting was interrupted. Run `update` and validate again. |
| 5 | A manifest `rec_number` or author does not appear in any traveling library payload | The wrong library was open, or the wrong record was cited. |
| 6 | `Error!`, `Reference source not found`, or `{Citation}` placeholder text | A field was broken by hand editing or by a copy between documents. |
| 7 | `EN.Layout` style does not equal the manifest style | Run `set-style` with the manifest style, then validate again. |
| 8 | `EN.Libraries` does not contain every manifest `rec_number` | Strong signal that the wrong library was open. See section 6. |

**The counting rule.** A merged citation such as `{A, 2023 #1; B, 2020 #2}` is one field,
not two. When the traveling library payload is large, EndNote moves part of it into a nested
`ADDIN EN.CITE.DATA` field, which naively counts as an extra citation. The correct count is:

```text
citation_count = count(codes starting with "ADDIN EN.CITE")
               - count(codes starting with "ADDIN EN.CITE.DATA")
```

`HYPERLINK` fields inside the bibliography are DOI links emitted by the style. They are
counted separately as `hyperlink_fields` and are never citations.

For gate 5, a record whose payload was moved into a nested `ADDIN EN.CITE.DATA` field must be
matched against that nested field's own code text, not against the parent field.

## Recovery rules

- Leave the input document untouched. Every destructive action writes elsewhere by default.
- Record the citation id, the marker, the `rec_number`, the style, and the exact bridge
  error text before retrying anything.
- When a call times out, assume a modal dialog until proven otherwise, and find the
  pre-check that was skipped.
- Clean up only the Word process the failed job created. Never a pre-existing Word or
  EndNote process.
- Re-run the smallest fixture before re-running a large document.
- Re-run `validate` after any manual edit to a rendered document. Hand edits to field codes
  are the usual cause of gate 6 failures.

## Fallback boundary

Typed citation text, a hand written reference list, a Word bibliography feature, or a
citation copied as plain text from another document are not acceptable substitutes. They
look correct and are structurally dead: EndNote cannot update them, a style switch will not
touch them, and the traveling library is absent.

Before offering any fallback:

1. Show the failed prerequisite, the exact error, or the failed validation gate.
2. State plainly that the output will not be EndNote native and cannot be updated by CWYW.
3. Get explicit user acceptance.
4. Label the deliverable as a fallback. Never describe it as a formatted citation or a real
   bibliography.

One irreversible case deserves its own warning, and it is deliberately not automated.
Converting a document to plain text is a manual EndNote ribbon operation: **EndNote >
Convert Citations and Bibliography > Convert to Plain Text**, which produces a new copy and
leaves the original document's field codes intact. There is no `remove-field-codes` bridge
action and no MCP tool for it, because the macro behind that ribbon command,
`ENRemoveFieldCodes`, raises a modal confirmation dialog and never returns under automation.
It is the same hang as section 2, so it was removed from the v1 surface rather than shipped
with a watchdog around it.

Tell the user what the conversion costs before they run it: inside the converted copy,
EndNote can no longer update the citations, a style switch will not touch them, and the
traveling library is gone. Keep the version with field codes; it is the only editable copy.
