# CWYW entry points

This file is the M0 spike deliverable. It records how the bridge reaches EndNote's Cite
While You Write engine, which candidate routes were tried, which one won, and what the
losing routes did instead. Every measurement below was taken on the machine described in
section 1.

## Contents

1. Environment tested
2. Candidate routes
3. Chosen route and why
4. Why `COMAddIns(...).Object` is dead
5. Macro inventory
6. Macros this project calls, and the two it deliberately does not
7. Primary and fallback order the bridge uses
8. Probe checks that prove the route is live
9. Observed failure modes
10. Version discovery, never hardcode 21

## Environment tested

| Fact | Value |
|---|---|
| OS | Windows 11 Pro 26200 |
| Word | 16.0, build 16.0.20131, COM `Word.Application` works |
| EndNote | EndNote 21, `C:\Program Files (x86)\EndNote 21\EndNote.exe` |
| EndNote registry root | `HKCU:\Software\ISI ResearchSoft\EndNote` |
| CWYW COM add-in ProgID | `EndNote.WordAddin.Connect` |
| CWYW COM add-in CLSID | `{C7E89D1A-C877-426B-AE98-27E86872622E}` |
| CWYW DLL, 64 bit | `C:\Program Files\Common Files\ResearchSoft\Cwyw\21\EndNote Cwyw.dll` |
| CWYW macro template | `C:\Program Files\Common Files\ResearchSoft\Cwyw\21\EndNote Cwyw.dotm` |
| Installed styles folder | `C:\Program Files (x86)\EndNote 21\Styles`, 7448 `.ens` files |
| Personal styles folder | `%USERPROFILE%\...\Documents\EndNote\Styles`, may not exist |
| PowerShell | `pwsh.exe` 7.6.3 and `powershell.exe` 5.1, both present |
| Python | 3.14 and 3.13, both present |

The macro inventory in section 5 was extracted from the template's `vbaProject.bin`.

## Candidate routes

Four routes to CWYW were considered. The build spec names Route A and Route B explicitly.
Routes C and D are labelled here for symmetry; both correspond to prohibitions the spec
records rather than to named routes, so the labels are this document's and the verdicts
are the spec's.

| Route | Mechanism | Verdict |
|---|---|---|
| A | `Application.COMAddIns("EndNote.WordAddin.Connect").Object`, then call methods on the returned automation object | Dead. `.Object` returns `$null`. |
| B | `Application.Run('<macro>')` against the `EndNote Cwyw.dotm` global template that Word auto loads | Primary, verified working end to end. |
| C | Force load the template with `Application.AddIns.Add('...EndNote Cwyw.dotm')` first, then `Application.Run` | Rejected. It double loads the template. |
| D | Write `ADDIN EN.CITE` field codes directly into the OOXML package, bypassing CWYW entirely | Rejected for writing. Permitted for reading only. |

Route D deserves the longer explanation, because it looks attractive: no Word, no EndNote,
no COM, no dialogs. It is rejected because a hand written field code cannot carry a real
traveling library. The `<record>` payload inside `ADDIN EN.CITE` is a verbatim copy of the
library record, and the `EN.Libraries` document variable records provenance for it.
Synthesising both is fabrication, and a document built that way is not one EndNote can
update. The skill invariant is explicit: never fabricate field codes in OOXML, never pass
off a plain text bibliography as a real one.

Reading the OOXML package is a different matter and is fully sanctioned. The citation
counting rule and the `EN.Layout` and `EN.Libraries` parsers are specified to work
identically over COM and over the unzipped DOCX, which is what lets the offline structure
tests run on a machine with neither Word nor EndNote installed. See
`field-codes-and-traveling-library.md`.

## Chosen route and why

**Route B, `Application.Run` against the auto loaded global template.**

The facts that decided it:

- `EndNote Cwyw.dotm` is auto loaded as a global template as soon as any Word document is
  open. `$word.Templates` then contains `EndNote Cwyw.dotm`. No loading step is needed.
- `$word.Run('ENFormatBibliography')` formats every temporary citation in the document and
  rebuilds the bibliography. This was verified end to end: two temporary citations, one of
  them merged, produced two `ADDIN EN.CITE` fields with traveling library payloads, one
  nested `ADDIN EN.CITE.DATA` field and one `ADDIN EN.REFLIST` bibliography.
- The macro surface is large and named, so the other macro the bridge needs
  (`ENUnformatCitations`, which backs the `unformat` action) is reachable by the same
  mechanism with no new plumbing. Reachable is not the same as usable: `ENRemoveFieldCodes`
  is reachable too and still had to be dropped because it hangs, see section 6.
- It is non interactive. Nothing on this path opens a dialog by itself; the dialogs that
  do appear come from bad input data, and those are prevented by pre checks.

```powershell
# The whole entry point, once a document is open in $word.
$script:Word.Run('ENFormatBibliography')
```

Route C was rejected even though it also ends in `Application.Run`, because
`$word.AddIns.Add(...)` double loads a template that Word has already loaded. Do not call
it. Route B needs no load step at all.

## Why `COMAddIns(...).Object` is dead

`Application.COMAddIns` does list the add-in under ProgID `EndNote.WordAddin.Connect`, so
the add-in is registered and present. But its `.Object` property returns `$null` on this
installation, so there is no automation object to call methods on. The route ends there.

Practical guidance:

- Use `COMAddIns` and the registry only as evidence that CWYW is installed, which is what
  the `cwyw_addin_registered` and `cwyw_addin_progid` probe checks report.
- Never dereference `.Object`. A null dereference here is a confusing failure far away from
  its cause.
- Do not read the null as "CWYW is broken". It is not; the spike proved Route B works on
  the same installation, at the same moment.

## Macro inventory

Twenty nine macros, extracted from `vbaProject.bin` in `EndNote Cwyw.dotm`, sorted:

```
ENAuthorYearFormat
ENConfigureInstantFormattingDlg
ENConfigureReferenceCategoriesDlg
ENConvertWordCitations
ENDefaultFormat
ENEditCitations
ENEditLibraryReferences
ENExcludeAuthor
ENExcludeAuthorYear
ENExcludeYear
ENExportTravelingLibrary
ENExportWordCitations
ENExportWordMasterList
ENFindCitations
ENFormatBibliography
ENFormatBibliographyDlg
ENGoToEndNote
ENHelp
ENInsertFigure
ENInsertNote
ENInsertSelectedCitations
ENManuscriptMatcher
ENPreferences
ENRemoveFieldCodes
ENShowOnlyInBibliography
ENToggleCategorizeReferences
ENToggleInstantFormatting
ENUnformatCitations
ENUpdateEditCitations
```

The build spec records the names only. It does not record parameters, return values, or
non interactive behaviour for any macro outside the four listed in the next section, so
treat the rest as unverified and do not call them from automation on the assumption that
they behave.

## Macros this project calls, and the two it deliberately does not

| Macro | Use | Notes |
|---|---|---|
| `ENFormatBibliography` | Format every temporary citation and rebuild the bibliography | Core of `render`, `update` and `set-style`. Non interactive. |
| `ENUnformatCitations` | Revert formatted citations back to `{Author, Year #RecNum}` text | Backs the `unformat` action. New output file only. Measured non interactive: it dropped the spike document from four fields to zero and restored the temporary citation text. |
| `ENRemoveFieldCodes` | Strip field codes to plain text | **Never called.** Measured to raise a modal confirmation dialog and never return under automation. |
| `ENToggleInstantFormatting` | Toggle instant formatting | Present, but not used in v1 automation. |

Only the first two rows are ever invoked. Three prohibitions:

- `ENRemoveFieldCodes` hangs, so there is no `remove-field-codes` bridge action and no
  `remove_endnote_field_codes` MCP tool. Converting a document to plain text stays a manual
  operation the user performs in Word: EndNote tab, Convert Citations and Bibliography,
  Convert to Plain Text. Describe that path when a user asks for it; never automate it.
- `ENToggleInstantFormatting` toggles and offers no read back, so calling it blindly can
  leave the setting in either state. v1 never calls it.
- Every macro whose name ends in `Dlg` opens a modal dialog. `ENConfigureInstantFormattingDlg`,
  `ENConfigureReferenceCategoriesDlg` and `ENFormatBibliographyDlg` must never be called
  from the bridge. A modal dialog is unrecoverable, see section 9.

## Primary and fallback order the bridge uses

The build spec establishes exactly one working route for creating and updating citations.
There is therefore no second code path to fall back to, and inventing one would mean
inventing behaviour the spike did not measure. The order below is an ordered set of
preconditions, not a set of alternatives.

1. Start Word through COM on a copy of the document. Opening a document is what makes Word
   auto load `EndNote Cwyw.dotm`.
2. Confirm `$word.Templates` contains `EndNote Cwyw.dotm`. Do not call
   `$word.AddIns.Add(...)` when it is missing; report the condition instead.
3. Call `$script:Word.Run('<macro>')` for the action's macro.
4. Poll until the expected `ADDIN EN.CITE` field count is reached or
   `ENDNOTE_AUTOMATION_TIMEOUT_MS` expires, default 60000 ms.
5. On failure at any step, stop and report. Do not retry through Route A, C or D. Run
   `probe` to produce the diagnostic detail.

For reading a document, the fallback order is different and does have a genuine second
path: prefer COM when Word is available, otherwise unzip the DOCX and apply the same
counting rule and document variable parsers to `word/document.xml` and
`word/settings.xml`. That fallback is read only by construction.

## Probe checks that prove the route is live

`probe` is read only and exits with code 2 when the environment is not ready. The checks
that speak to the entry point specifically:

| Check | Meaning |
|---|---|
| `word_com`, `word_version` | `Word.Application` can be created over COM. |
| `endnote_executable`, `endnote_executable_path`, `endnote_product_version` | The EndNote executable was discovered. |
| `cwyw_addin_registered`, `cwyw_addin_progid` | The COM add-in is registered. This is evidence of installation only, see section 4. |
| `cwyw_template_present`, `cwyw_template_path` | `EndNote Cwyw.dotm` exists on disk. |
| `cwyw_template_loaded` | Word actually auto loaded the template, which is the precondition for Route B. |
| `cwyw_macro_callable` | A macro call through `Application.Run` succeeded. |
| `endnote_running` | Whether an EndNote process already exists, which changes how `render` starts EndNote. |
| `styles_directory`, `installed_style_count` | Where styles live and how many were found. |

`cwyw_macro_callable` is probed on a throwaway empty document, never on user data, so a
probe can never corrupt a manuscript.

## Observed failure modes

**A modal dialog blocks `Application.Run` forever.** This is the dominant failure mode and
the reason for every pre check in this project.

- `$word.DisplayAlerts = 0` does not suppress add-in dialogs. There is no COM setting that
  does.
- The call does not time out on its own. It waits for a human who is not there.
- The only recovery is a watchdog. The Python MCP layer enumerates `WINWORD.EXE` PIDs
  before and after the call and kills the isolated one on timeout. It also kills
  `EndNote.exe` PIDs created during the call, and never kills a pre-existing EndNote PID,
  because that would be the user's own session.
- Known triggers, all preventable before Word opens: a `rec_number` that does not exist in
  the library, an output style name that does not resolve to an installed `.ens` file, the
  backslash prefix syntax, the excluded author syntax, and any `Dlg` suffixed macro.
  See `temporary-citations.md` for the measured syntax cases.

**Console encoding defaults to cp950 on this machine.** Every entry point must force
UTF-8 explicitly, or non ASCII author names and titles are mangled on the way out. In
PowerShell that is `[Console]::OutputEncoding = New-Object Text.UTF8Encoding $false`,
wrapped in try/catch because 5.1 can throw when stdout is redirected.

**`AddIns.Add` double loads the template.** Calling `$word.AddIns.Add(...)` on
`EndNote Cwyw.dotm` loads a template Word has already loaded. Never call it.

**EndNote may already be running with a different library.** When an EndNote process owns
a main window, the bridge does not start a second instance and does not force a library
switch. It records `endnote_preexisting: true` and proceeds. The user's open library is
their business, and forcing a switch would be a destructive act on user state.

**CWYW silently auto-launches EndNote when it is closed.** `ENFormatBibliography` was
measured to succeed with no EndNote process running: CWYW starts EndNote itself, with the
most recently used library. This is a failure mode, not a convenience, because rec-numbers
are library-local, so a run against the wrong library resolves the wrong records without
raising anything. It is why `Start-EndNoteWithLibrary` still opens the manifest's library
explicitly before Word is touched. The bridge compares the EndNote PID set before and after
and reports `endnote_autostarted` when a new process appeared.

## Version discovery, never hardcode 21

Every path in section 1 contains the literal `21`, and none of it may be hardcoded. The
bridge discovers the EndNote executable and the styles directory by scanning:

- the subkeys of `HKLM:\SOFTWARE\WOW6432Node\ISI ResearchSoft\EndNote`, and
- the `C:\Program Files (x86)\EndNote *` directories,

newest version first. The same applies to the CWYW folder under
`C:\Program Files\Common Files\ResearchSoft\Cwyw\<version>\`.

The build spec records measurements from EndNote 21 only. It says nothing about EndNote
20 or earlier, or about any release after 21, so the discovery logic is written to find
whatever is installed rather than to assert that other versions behave the same way. If
the bridge runs against a version other than 21, the macro inventory and the entry point
should be re-verified before the results are trusted.
