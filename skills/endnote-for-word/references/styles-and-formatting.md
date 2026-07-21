# Output styles and bibliography formatting

## Contents

1. Where the style is stored
2. The `EN.Layout` element table
3. Switching a style without a dialog
4. Verified round trip: Annotated to IEEE
5. Why an unknown style name hangs the run
6. Discovering installed styles
7. Reading `EN.Layout` offline
8. Validation coupling
9. IEEE DOI template contract
10. Default Format and Author (Year)

## Where the style is stored

EndNote does not keep the document's output style in the citation fields. It keeps it in a
Word **document variable** named `EN.Layout` whose value is a single XML string. The
document variable travels with the DOCX, so a document carries its own style selection and
bibliography layout independently of whatever style the EndNote desktop application happens
to be showing.

This is the whole reason style control in this project is fully programmatic. There is no
need to drive the Format Bibliography dialog, and the dialog macro
`ENFormatBibliographyDlg` must never be called from the bridge because every `*Dlg` macro
opens a modal window that blocks automation.

A second document variable, `EN.Libraries`, records which library and which record ids were
used. It is provenance, not configuration, and it is covered in
`field-codes-and-traveling-library.md`.

## The `EN.Layout` element table

Value read back verbatim from the spike document:

```xml
<ENLayout><Style>Annotated</Style><LeftDelim>{</LeftDelim><RightDelim>}</RightDelim><FontName>Times New Roman</FontName><FontSize>12</FontSize><ReflistTitle></ReflistTitle><StartingRefnum>1</StartingRefnum><FirstLineIndent>0</FirstLineIndent><HangingIndent>720</HangingIndent><LineSpacing>0</LineSpacing><SpaceAfter>0</SpaceAfter><HyperlinksEnabled>0</HyperlinksEnabled><HyperlinksVisible>0</HyperlinksVisible><EnableBibliographyCategories>0</EnableBibliographyCategories></ENLayout>
```

| Element | Observed value | What it controls | Written by this tool |
|---|---|---|---|
| `Style` | `Annotated` | Output style name. Must match the base name of an installed `.ens` file. | Yes. The only element this tool ever rewrites. |
| `LeftDelim` | `{` | Opening character EndNote scans for when it looks for temporary citations. | No |
| `RightDelim` | `}` | Closing character of a temporary citation. | No |
| `FontName` | `Times New Roman` | Font applied to the generated bibliography. | No |
| `FontSize` | `12` | Point size applied to the generated bibliography. | No |
| `ReflistTitle` | empty | Heading text emitted above the `ADDIN EN.REFLIST` result. Empty means no heading. | No |
| `StartingRefnum` | `1` | First number used by a numbered style such as IEEE or Vancouver. | No |
| `FirstLineIndent` | `0` | First line indent of each bibliography entry. | No |
| `HangingIndent` | `720` | Hanging indent of each bibliography entry. | No |
| `LineSpacing` | `0` | Line spacing selector for the bibliography. | No |
| `SpaceAfter` | `0` | Space after each bibliography entry. | No |
| `HyperlinksEnabled` | `0` | Whether formatted in-text citations link into the bibliography entry. | No |
| `HyperlinksVisible` | `0` | Whether those links are rendered with visible link formatting. | No |
| `EnableBibliographyCategories` | `0` | Categorised bibliography sections, the state paired with `ENToggleCategorizeReferences` and `ENConfigureReferenceCategoriesDlg`. | No |

Read the table honestly. The "Observed value" column is measured: it is the exact string
read back from the spike document. The "What it controls" column follows the element names
and the observed formatted output. Only `Style` was varied and re-measured during the spike,
so treat any other element as documented but untested. If you need a different bibliography
font or indent, change it in the EndNote desktop application and re-read `EN.Layout` to see
what value it produced rather than guessing a number.

Two notes that prevent common misreadings:

- `HangingIndent` of `720` is consistent with Word's twips unit, where 720 twips is half an
  inch. No other value was written or measured, so do not treat that conversion as verified
  behaviour of EndNote's parser.
- `HyperlinksEnabled` was `0` in the spike document and the resulting bibliography still
  contained `HYPERLINK` fields for DOIs. Those links come from the style's DOI or URL
  output, not from this element. Validation counts them separately as `hyperlink_fields`
  and never treats them as citations.

Because only `<Style>` is rewritten, a document that a user already configured by hand keeps
its font, indent and hyperlink choices when this tool switches styles. Never rebuild the
whole `EN.Layout` value from a template when the variable already exists; that would
silently discard the user's layout settings.

## Switching a style without a dialog

Three steps, verified end to end and with no modal window at any point:

1. Read the document variable `EN.Layout`. Create it with a complete default value only when
   it is absent.
2. Replace the contents of the `<Style>` element with the target style name. When the
   variable already exists, do a targeted replacement of that one element and leave every
   other element byte for byte unchanged.
3. Run the CWYW macro `ENFormatBibliography` through `Application.Run`. It reformats every
   citation in the document and rebuilds the bibliography in the new style.

Bridge and MCP surfaces for this:

| Surface | Call |
|---|---|
| Bridge | `-Action set-style -InputPath <docx> -OutputPath <docx> -Style "IEEE"` |
| Bridge helpers | `Get-EndNoteLayout`, `Set-EndNoteLayout`, `Resolve-StyleName`, `Invoke-CwywFormat` |
| MCP tool | `set_endnote_citation_style` |

`set-style` reports both `previous_style` and `style`, so the caller can prove a change
actually happened rather than assuming it. The action is destructive in the sense that it
rewrites the document body, and it obeys the same rule as every other destructive action:
it will not write over `InputPath` unless `-Overwrite` is passed.

During a `render`, the same mechanism runs as step 7 of the render sequence, before any
marker is replaced. The style is therefore already correct the first time
`ENFormatBibliography` runs, and no second reformat pass is needed.

## Verified round trip: Annotated to IEEE

Starting state, style `Annotated`:

```text
body:          (Thomsen, Shearer et al. 2023)
EN.Layout:     <Style>Annotated</Style>
```

After patching `<Style>` to `IEEE` and running `ENFormatBibliography`:

```text
body:          [1]
EN.Layout:     <Style>IEEE</Style>
bibliography:  fully rewritten in IEEE format
```

The in-text citation form, the bibliography entry form, and the ordering all changed, while
the `ADDIN EN.CITE` fields still carried their traveling library payload and the same
`rec-number` values afterwards. Whether EndNote rewrote those fields in place or rebuilt
them was not measured and does not matter here; what matters is the measured outcome, which
is why `set-style` takes no manifest, never re-renders markers, and never re-queries the
library.

Re-run `validate` after a style switch. Gate 7 compares the `EN.Layout` style against the
manifest style case-insensitively, so a switch that silently failed shows up as a validation
error instead of as a wrong-looking document.

## Why an unknown style name hangs the run

A style name that does not resolve to an installed `.ens` file makes EndNote open a modal
dialog asking a human to choose a style. That has three consequences, all measured:

- `Application.Run` never returns. The macro call is blocked on the dialog's message loop.
- `Application.DisplayAlerts = 0` does not help. It suppresses Word's own alerts, not
  dialogs raised by an add-in.
- The only recovery is the watchdog: enumerate `WINWORD.EXE` before and after the call and
  terminate the process this job created. See `troubleshooting.md`.

This is the same failure class as a `rec_number` that does not exist in the library, which
raises a "select matching reference" dialog.

The defence is a pre-check, not a timeout. Both hard pre-conditions are enforced **before
Word is ever opened**:

1. Every `rec_number` in the manifest exists in the target library, verified by reading the
   library's SQLite store.
2. The requested style name resolves to an installed `.ens` file.

`render` performs the style check as step 3 and fails hard with the nearest matching style
names in the error message. Never let an unverified style string reach
`ENFormatBibliography`. A timeout plus a killed Word process is a worse outcome than a
refusal that costs nothing.

## Discovering installed styles

| Surface | Call |
|---|---|
| Bridge | `-Action list-styles` optionally with `-Style <substring>` |
| MCP tool | `list_endnote_styles` |
| Bridge helper | `Get-InstalledStyles`, `Resolve-StyleName` |

Behaviour:

- Enumerates `*.ens` in the installed styles directory and, when it exists, the personal
  styles directory. A missing personal styles directory is normal and is not an error.
- Returns sorted unique base names plus a count.
- `-Style` filters case-insensitively as a substring match.
- Results are capped at 200 entries. When the cap trims the list the response carries
  `truncated: true`, so always filter rather than paging through everything.

Directory discovery is dynamic. The installed styles directory is found by scanning
`HKLM:\SOFTWARE\WOW6432Node\ISI ResearchSoft\EndNote` and the
`C:\Program Files (x86)\EndNote *` directories, newest version first. The major version is
never hardcoded, because a machine may have EndNote 20, 21 or a later release, and more than
one at a time. The reference machine reported 7448 installed `.ens` files, which is why an
unfiltered listing is useless to an agent and the filter argument exists.

Name matching rules for the manifest `style` value:

- The value is the `.ens` base name without the extension. `APA 7th` resolves to
  `APA 7th.ens`.
- Comparison is case-insensitive, both for resolution and for validation gate 7.
- Spaces, digits and ordinal suffixes are part of the name. `APA 7th` and `APA7` are
  different styles, and only one of them may exist on a given machine.
- A style the user authored or edited lives in the personal styles directory and appears in
  the same flat list. The list is deduplicated by base name, so a personal style that reuses
  the name of an installed one shows up once. Which of the two files EndNote then loads was
  not measured, so rename a personal style rather than shadowing an installed name.

Always call `list_endnote_styles` with a filter before writing a style into a manifest. Do
not rely on a style name being present just because it is a well known citation format.

## Reading `EN.Layout` offline

Both document variables persist inside the DOCX package at `word/settings.xml`:

```xml
<w:docVar w:name="EN.Layout" w:val="&lt;ENLayout&gt;&lt;Style&gt;IEEE&lt;/Style&gt;...&lt;/ENLayout&gt;"/>
<w:docVar w:name="EN.Libraries" w:val="&lt;Libraries&gt;...&lt;/Libraries&gt;"/>
```

The XML is stored XML-escaped inside the `w:val` attribute, so it must be unescaped once
before it can be parsed. This is what makes the offline structural tests possible: the test
suite can assert on the style of a rendered document without opening Word or having EndNote
installed. Use the same parser functions in offline tests and in live tests so the two
cannot drift apart.

Offline reading is also the fastest way to answer "what style is this document in" for a
document you did not produce. It costs one unzip and no COM call.

## Validation coupling

Two of the validation gates depend on this document:

- Gate 7: the `EN.Layout` document variable exists and its style equals the manifest style,
  compared case-insensitively.
- Gate 4: when `bibliography.expected` is true, exactly one `ADDIN EN.REFLIST` field exists
  and its result text is non-empty after trimming.

A style switch that produced an empty bibliography fails gate 4 even though gate 7 passes.
That combination usually means `ENFormatBibliography` was interrupted, not that the style is
wrong. Re-run `update` and validate again before touching the manifest.

## IEEE DOI template contract

IEEE is the Agent default. The style's Bibliography Templates, not the Word document and not
the manifest, own the punctuation and field placement. Every reference type that can be cited
must emit the literal `doi: https://doi.org/` followed by the EndNote `DOI` field. The library
field itself must contain only the bare DOI, for example `10.1109/5.771073`.

The three primary templates are:

```text
Journal Article
Author, "Title," |(in Language), |Journal, |Type of Article |vol. Volume, no. Issue, Part/Supplement, p.^pp. Pages, Date Year|, Art no. Article Number|, doi: https://doi.org/DOI|. Original Publication|.

Conference Paper
Author, "Title," presented at the Conference Name, Conference Location, Date, Year|, Type|, Paper No.|, doi: https://doi.org/DOI|. [Online]. Available: URL|.

Web Page
Author. "Title." Publisher. doi: https://doi.org/DOI. URL (accessed Access Date).
```

In EndNote's editor, use the `Link Adjacent Text` operator around optional DOI phrases so
the literal disappears when the field is empty. This project nevertheless blocks any cited
record with an empty DOI before Word opens. The style editor must use `Save As`; installed
styles are not overwritten. Save the result in the personal Styles folder and confirm its
base name through `list_endnote_styles`.

The validator reads the visible `ADDIN EN.REFLIST` result. It fails when any manifest DOI is
not present in the exact form `doi: https://doi.org/<DOI>` or when it sees a doubled resolver
such as `https://doi.org/https://doi.org/<DOI>`. This makes a bad template or a prefixed DOI
field a hard failure instead of a cosmetic defect.

## Default Format and Author (Year)

Each manifest citation may declare one of two display modes:

| Manifest value | EndNote behavior under IEEE | Use when |
|---|---|---|
| `default` | Output style default, normally `[1]` | The reference supports a claim or a merged group. |
| `author-year` | `ENAuthorYearFormat`, normally `Author et al. [1]` | The author is a grammatical subject or named actor in the sentence. |

`author-year` is restricted to one target. Place its marker where the complete author phrase
belongs, for example `{{CITE:smith}} demonstrated the effect.` Do not write
`Smith {{CITE:smith}} demonstrated the effect`, because EndNote itself supplies `Smith` and
the result would duplicate the name. For a normal evidence citation, write
`The effect was reproducible {{CITE:smith}}.` and use `default`.

The bridge first creates every genuine `ADDIN EN.CITE` field through
`ENFormatBibliography`, then selects only the requested native fields by document order and
calls `ENAuthorYearFormat`. It never types the author or citation number. Word stays hidden,
and an EndNote process started by the bridge also stays hidden.
