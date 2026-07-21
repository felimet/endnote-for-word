# Temporary citations

A temporary citation is plain document text that EndNote recognises and converts into a
real `ADDIN EN.CITE` field when `ENFormatBibliography` runs. It is the only supported way
this skill creates citations. Nothing here is guessed; every row of the table below was
measured on the target machine during the M0 spike.

## Contents

1. Measured syntax table
2. Syntaxes rejected in v1 and why
3. The nonexistent record trap
4. String building algorithm
5. What `unformat` writes back
6. Delimiters
7. What the skill never does

## Measured syntax table

| Written text | Result | EN.CITE XML | Status |
|---|---|---|---|
| `{Thomsen, 2023 #1}` | `(Thomsen et al., 2023)` | `<Author>...</Author><Year>...</Year><RecNum>1</RecNum>` | SUPPORTED |
| `{Thomsen, 2023 #1@45-46}` | `(Thomsen et al., 2023, pp. 45-46)` | adds `<Pages>45-46</Pages>` | SUPPORTED, cited pages |
| `{Thomsen, 2023 #1, p. 99}` | `(Thomsen et al., 2023, p. 99)` | adds a `<Suffix>` element | SUPPORTED, suffix |
| `{A, 2023 #1; B, 2023 #2; C, 2020 #3}` | one merged citation, style ordered | one field, nested `EN.CITE.DATA` | SUPPORTED, merged |
| `{see \Thomsen, 2023 #1}` | HANGS on a modal dialog | not reached | NOT SUPPORTED in v1 |
| `{\Thomsen, 2023 #1}` | HANGS on a modal dialog | not reached | NOT SUPPORTED in v1 |
| `{Nobody, 1900 #9999}` | HANGS on the "select matching reference" dialog | not reached | MUST be prevented by pre-check |

Two facts follow directly from this table and drive the whole design:

- A merged citation is a single field. `{A, 2023 #1; B, 2023 #2}` yields one
  `ADDIN EN.CITE` field, not two. Field counting must never assume one field per target.
- The displayed result strings above were measured, but the build spec does not record
  which output style was active when they were produced. Treat them as evidence that the
  syntax works, not as a promise about how any particular style renders.

## Syntaxes rejected in v1 and why

Both rejected forms use the backslash prefix character.

| Form | Intent | Observed behaviour |
|---|---|---|
| `{see \Thomsen, 2023 #1}` | Citation prefix, for example "see" before the author | Hangs. EndNote opens a modal dialog. |
| `{\Thomsen, 2023 #1}` | Exclude the author from the formatted citation | Hangs. EndNote opens a modal dialog. |

A modal EndNote or Word dialog blocks `Application.Run` forever. Setting
`$word.DisplayAlerts = 0` does not suppress add-in dialogs, so there is no way for the
bridge to answer or dismiss the dialog from inside the automation. The only recovery is
the watchdog: enumerate `WINWORD.EXE` PIDs before and after the call and `taskkill` the
isolated one when the timeout expires. That leaves the document unsaved and the run
wasted.

Because there is no non interactive escape, v1 refuses these forms at the earliest
possible point:

- `Assert-Manifest` rejects any `prefix` key in a citation entry, with an error message
  pointing at this file.
- There is no manifest key that produces an excluded author or an excluded year. The
  `ENExcludeAuthor`, `ENExcludeAuthorYear` and `ENExcludeYear` macros exist in the CWYW
  template but are not called by the bridge; the spike did not establish a non interactive
  call path for them.

For narrative citations, use the manifest's `display: author-year` mode instead of the
backslash syntax. The bridge first creates a native field, selects that field, and invokes
EndNote's non-interactive `ENAuthorYearFormat` macro. Under IEEE the result is a form such
as `Thomsen et al. [1]`. The marker must replace the entire author phrase.

## The nonexistent record trap

`{Nobody, 1900 #9999}` is syntactically valid. It hangs because EndNote cannot resolve
record 9999 in the open library and opens a "select matching reference" dialog, which is
the same unrecoverable state as above.

This is a data problem, not a syntax problem, so the guard is a data check. Before Word is
opened at all:

1. Every `rec_number` in the manifest is looked up in the library SQLite store
   (`<Library>.Data\sdb\sdb.eni`) through `precheck_endnote_manifest`.
2. The record must exist, must have `trash_state = 0`, must have a first author surname
   matching the manifest `author` case insensitively and accent tolerantly, and must have
   a matching `year` and bare DOI.
3. The DOI must resolve successfully through the official doi.org Proxy API.
4. The precheck reports every mismatch, not just the first, so a manifest can be fixed in
   one pass.

The same class of guard applies to the output style. An unknown style name also opens a
modal dialog, so the style is resolved against the installed `.ens` files before the
render begins.

## String building algorithm

One `citations[]` entry produces exactly one temporary citation string.

```
build(entry):
    parts = []
    for target in entry.targets:                       # order is preserved as authored
        parts.append(target.author + ", " + target.year + " #" + target.rec_number)
    body = join("; ", parts)                           # "; " separator, merged citations
    if entry has pages:
        body = body + "@" + entry.pages
    if entry has suffix:
        body = body + entry.suffix                     # appended verbatim, no separator
    return "{" + body + "}"
```

Rules the algorithm depends on:

- The separator between targets is a semicolon followed by a single space.
- Inside one target the separator between author and year is a comma followed by a single
  space, and the record number is prefixed by a space and then `#`.
- `pages` is introduced by `@` with no surrounding whitespace, and must already have passed
  the `^[0-9]+(\s*-\s*[0-9]+)?$` check in `Assert-Manifest`.
- `suffix` is appended verbatim. The author of the manifest owns its leading punctuation
  and spacing, which is why the schema example writes `", p. 99"` with the comma and space
  included.
- EndNote applies `pages` and `suffix` to the last target only. `Assert-Manifest`
  therefore permits either key only when `targets` has exactly one entry, so the effect is
  never ambiguous.

Worked results:

| Entry | Generated string |
|---|---|
| one target, rec 1, Thomsen, 2023 | `{Thomsen, 2023 #1}` |
| the same plus `"pages": "45-46"` | `{Thomsen, 2023 #1@45-46}` |
| the same plus `"suffix": ", p. 99"` | `{Thomsen, 2023 #1, p. 99}` |
| two targets, rec 2 Magrin 2023 and rec 3 Bernhard 2020 | `{Magrin, 2023 #2; Bernhard, 2020 #3}` |
| three targets, rec 1, rec 2, rec 3 | `{A, 2023 #1; B, 2023 #2; C, 2020 #3}` |

Supplying both `pages` and `suffix` on the same single target entry follows from the
algorithm and yields `{Thomsen, 2023 #1@45-46, p. 99}`. The build spec's schema example
does show both keys on one entry, but the combined form is not one of the measured rows in
the syntax table, so its rendered output is unverified. Prefer one or the other until it
has been measured.

The generated string replaces the marker text in place. The bridge locates the marker with
`Find-UniqueMarkerRange`, which requires the marker to occur exactly once in the document,
and sets the range text to the generated string.

## What `unformat` writes back

The `unformat` action calls `ENUnformatCitations`, which is verified working and non
interactive. On the spike document it dropped the field count from four to zero and
restored plain temporary citation text. The restored text is not byte identical to what
the bridge wrote:

| Written by the bridge | Written back by `ENUnformatCitations` |
|---|---|
| `{Thomsen, 2023 #1}` | `{Thomsen, 2023 #1}` |
| `{Magrin, 2023 #2; Bernhard, 2020 #3}` | `{Magrin, 2023 #2;Bernhard, 2020 #3}` |

A merged citation round trips **without the space after the semicolon**. Anything that
parses temporary citation text back out of a document must accept both `"; "` and `";"` as
the target separator. The generator still emits `"; "`, as the algorithm above specifies;
only the parser is required to be tolerant.

## Delimiters

The `{` and `}` delimiters are not hardcoded EndNote behaviour, they are recorded per
document in the `EN.Layout` document variable as `<LeftDelim>{</LeftDelim>` and
`<RightDelim>}</RightDelim>`. The verified spike value uses the defaults.

The build spec does not record behaviour for a document whose `EN.Layout` declares
different delimiters, and the bridge does not rewrite them. Assume the defaults and, if a
document arrives with custom delimiters, stop and report it rather than guessing.

A practical consequence for authoring: braces are ordinary characters in prose. Validation
gate 2 scans the finished document with `\{[^{}]*#\d+[^{}]*\}`, so a literal brace pair
containing a hash and digits will be reported as a residual temporary citation even if the
author meant it literally.

## What the skill never does

- Never type formatted citation text such as `(Thomsen et al., 2023)` into the document.
  Only CWYW may produce that text, and only as the display result of an `ADDIN EN.CITE`
  field.
- Never fabricate field codes in OOXML.
- Never write a bibliography by hand and present it as a real one.
- Never use the backslash prefix or excluded author forms.
- Never emit a temporary citation for a record that has not passed the precheck.
