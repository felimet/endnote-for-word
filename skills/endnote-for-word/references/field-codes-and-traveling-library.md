# Field codes and the traveling library

Everything in this file was read back from a document produced by
`ENFormatBibliography` during the M0 spike. It describes what a correct output looks like,
which is what validation asserts and what the offline OOXML reader parses.

## Contents

1. Field inventory of a formatted document
2. Anatomy of `ADDIN EN.CITE`
3. The nested `ADDIN EN.CITE.DATA` case
4. `ADDIN EN.REFLIST`
5. `HYPERLINK` fields
6. Citation counting rule
7. Document variables: `EN.Layout` and `EN.Libraries`
8. Reading the same facts offline from the DOCX package

## Field inventory of a formatted document

The spike document contained two paragraphs:

```text
Spike test paragraph one {Thomsen, 2023 #1} end of sentence.
Second paragraph with a merged citation {Magrin, 2023 #2; Bernhard, 2020 #3} done.
```

After `ENFormatBibliography` the document held four fields:

| Field | Code | Origin |
|---:|---|---|
| 1 | ` ADDIN EN.CITE <EndNote>...</EndNote>` | The single citation, payload inline. |
| 2 | ` ADDIN EN.CITE ` with a nested field inside it | The merged citation, payload too large to stay inline. |
| 3 | ` ADDIN EN.CITE.DATA ` | The nested payload belonging to field 2. |
| 4 | ` ADDIN EN.REFLIST ` | The bibliography, at end of document. |

Every one of these is a Word add-in field, `wdFieldAddin`, field type 81. Word reports the
code with a leading space, so any prefix test must trim first.

Two citations produced three EN.CITE family fields. That asymmetry is the reason the
counting rule in section 6 exists.

## Anatomy of `ADDIN EN.CITE`

Verified field 1 code, as a single unbroken run of text (the trailing `...` marks content
elided in the spike record, not a literal ellipsis in the document):

```text
 ADDIN EN.CITE <EndNote><Cite><Author>Thomsen</Author><Year>2023</Year><RecNum>1</RecNum><DisplayText>(Thomsen, Shearer et al. 2023)</DisplayText><record><rec-number>1</rec-number><foreign-keys><key app="EN" db-id="EXAMPLEDBID" timestamp="1765295406">1</key></foreign-keys><ref-type name="Journal Article">17</ref-type><contributors>...</contributors>...</record></Cite></EndNote>
```

Reformatted for reading only. The real code has no line breaks:

```xml
<EndNote>
  <Cite>
    <Author>Thomsen</Author>
    <Year>2023</Year>
    <RecNum>1</RecNum>
    <DisplayText>(Thomsen, Shearer et al. 2023)</DisplayText>
    <record>
      <rec-number>1</rec-number>
      <foreign-keys>
        <key app="EN" db-id="EXAMPLEDBID" timestamp="1765295406">1</key>
      </foreign-keys>
      <ref-type name="Journal Article">17</ref-type>
      <contributors>...</contributors>
    </record>
  </Cite>
</EndNote>
```

Reading of the parts:

| Element | Meaning |
|---|---|
| `<Cite>` | One cited reference. A merged citation carries several `<Cite>` elements in one field. |
| `<Author>`, `<Year>`, `<RecNum>` | The resolved form of what the temporary citation asked for. `<RecNum>` is the library record id. |
| `<DisplayText>` | The formatted text Word shows as the field result. It is regenerated on every reformat, so it changes when the style changes. |
| `<record>` | The traveling library payload: a full copy of the bibliographic record embedded in the document. |
| `<rec-number>` | The record id inside the payload. Validation gate 5 matches on `<rec-number>N</rec-number>`. |
| `<foreign-keys><key app="EN" db-id="..." ...>` | Which library the record came from. `EXAMPLEDBID` above is a sanitised value, real documents carry the library's own db-id. |
| `<ref-type name="Journal Article">17</ref-type>` | Reference type, name and numeric code together. |
| `<contributors>` | Author and editor lists. |

The traveling library is why an `ADDIN EN.CITE` field is worth producing at all. It makes
the document self describing: a colleague without your library still gets correct
citations and can import the records. A hand written field code without a `<record>`
payload is not a real citation, which is why fabricating field codes is forbidden.

Extra elements appear according to the temporary citation that produced the field:
`<Pages>45-46</Pages>` for the `@` cited pages form, and a `<Suffix>` element for the
trailing suffix form. See `temporary-citations.md`.

## The nested `ADDIN EN.CITE.DATA` case

When the traveling library payload is large, EndNote does not keep it inside the parent
field code. It moves the payload into a separate nested field:

```text
Field 2:  ADDIN EN.CITE <nested field> ADDIN EN.CITE.DATA <nested>
Field 3:  ADDIN EN.CITE.DATA
```

The parent field remains the citation. The `ADDIN EN.CITE.DATA` field is its payload
carrier and is never a citation in its own right. Consequences that the code must respect:

- Counting: subtract the `ADDIN EN.CITE.DATA` fields. See section 6.
- Validation gate 5: when a manifest record's `<rec-number>` is not found in the parent
  field's code, it must be looked for in the nested field's own code text. Searching only
  parent fields produces false failures on exactly the documents that are hardest to
  reproduce.
- Predictability: the spike observed the split on the merged citation and not on the single
  citation, but the build spec states only that the split happens when the payload is
  large. It gives no size threshold, so never assert on how many
  `ADDIN EN.CITE.DATA` fields a document should contain. Treat the count as informational
  and report it under `counts.citation_data_fields`.

## `ADDIN EN.REFLIST`

The bibliography is one `ADDIN EN.REFLIST` field placed at the end of the document. It is
generated by `ENFormatBibliography` and regenerated on every reformat and style change.

Validation gate 4, active when the manifest says `bibliography.expected` is true:

- Exactly one `ADDIN EN.REFLIST` field exists.
- Its `Result.Text` is non empty after trimming.

A bibliography typed as ordinary paragraphs is not a bibliography for the purposes of this
skill, no matter how correct the text looks.

## `HYPERLINK` fields

The bibliography may contain `HYPERLINK` fields for DOI links. They are Word fields inside
the reference list, and they are not citations. Count them separately and report them
under `counts.hyperlink_fields` so that a document with many DOI links cannot be mistaken
for a document with many citations.

## Citation counting rule

The same rule applies to the COM path and the offline OOXML path:

```
citation_count = count(codes starting with "ADDIN EN.CITE")
               - count(codes starting with "ADDIN EN.CITE.DATA")
```

Applying it correctly:

- Trim the leading whitespace from each field code before the prefix test, because Word
  reports codes as ` ADDIN EN.CITE ...`.
- `ADDIN EN.CITE.DATA` also starts with `ADDIN EN.CITE`, which is exactly why it is
  subtracted rather than excluded by a smarter pattern. Keep the rule as written; a
  reworded version that tries to match `ADDIN EN.CITE ` with a trailing space is fragile
  against the nested field syntax.
- `ADDIN EN.REFLIST` and `HYPERLINK` never enter this arithmetic.
- A merged citation is one field, not one per target. A manifest entry with three targets
  contributes 1 to `citation_count`.

Worked on the spike document: 3 fields start with `ADDIN EN.CITE` (fields 1, 2 and 3), 1
of them starts with `ADDIN EN.CITE.DATA` (field 3), so `citation_count = 3 - 1 = 2`, which
matches the two temporary citations that were written.

## Document variables: `EN.Layout` and `EN.Libraries`

EndNote stores per document state in Word document variables, not in a settings file. Two
matter here.

### `EN.Layout`

Holds the output style and the bibliography layout. Verified full value:

```xml
<ENLayout><Style>Annotated</Style><LeftDelim>{</LeftDelim><RightDelim>}</RightDelim><FontName>Times New Roman</FontName><FontSize>12</FontSize><ReflistTitle></ReflistTitle><StartingRefnum>1</StartingRefnum><FirstLineIndent>0</FirstLineIndent><HangingIndent>720</HangingIndent><LineSpacing>0</LineSpacing><SpaceAfter>0</SpaceAfter><HyperlinksEnabled>0</HyperlinksEnabled><HyperlinksVisible>0</HyperlinksVisible><EnableBibliographyCategories>0</EnableBibliographyCategories></ENLayout>
```

`<Style>` is the output style name and is the only element the bridge writes. Style
switching is fully programmatic and needs no dialog:

1. Read the `EN.Layout` document variable, or create it when it does not exist.
2. Replace the contents of `<Style>...</Style>` with the target style name.
3. Call `ENFormatBibliography`.

Verified round trip: changing `Annotated` to `IEEE` converted body text from
`(Thomsen, Shearer et al. 2023)` to `[1]` and rewrote the whole bibliography in IEEE
format, and `EN.Layout` then read back `<Style>IEEE</Style>`.

Validation gate 7 compares this style against the manifest `style`, case insensitively.

The other elements record delimiters, bibliography font, indents and hyperlink behaviour.
The build spec records their verified values but does not describe writing them, so v1
leaves them exactly as EndNote wrote them.

### `EN.Libraries`

Records provenance: which library, and which record ids were used.

```xml
<Libraries><item db-id="EXAMPLEDBID">ExampleLibrary<record-ids><item>1</item><item>2</item><item>3</item></record-ids></item></Libraries>
```

Structure: one `<item>` per library, carrying the `db-id` attribute, the library name as
its text, and a `<record-ids>` list of the record ids cited from it. `EXAMPLEDBID` and
`ExampleLibrary` above are sanitised values.

Validation gate 8 requires this variable to exist and to contain every manifest
`rec_number` inside `<record-ids>`. It is the proof that the citations came from the
library the manifest named, rather than from some other library that happened to be open.

## Reading the same facts offline from the DOCX package

A DOCX is a ZIP. Both document variables are persisted by Word in `word/settings.xml`:

```xml
<w:docVars>
  <w:docVar w:name="EN.Layout" w:val="&lt;ENLayout&gt;&lt;Style&gt;IEEE&lt;/Style&gt;..."/>
  <w:docVar w:name="EN.Libraries" w:val="&lt;Libraries&gt;..."/>
</w:docVars>
```

The `w:val` attribute holds the XML shown above as an XML attribute value, so the angle
brackets arrive escaped and must be unescaped before the inner XML is parsed.

Field codes are persisted in `word/document.xml` as `w:instrText` runs. The same counting
rule applies to the concatenated `instrText` content of each field.

This offline path needs neither Word nor EndNote, which makes it the right tool for tests
and for inspecting a document someone sent you. It cannot create or reformat citations;
only CWYW can do that.
