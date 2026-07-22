# Adding a new reference to the library

## Contents

1. The rule: EndNote owns every library write
2. Why Direct Export is the supported automation path
3. The four steps
4. Step 1: check the reference is not already there
5. Step 2: generate the RIS file
6. Step 3: import through EndNote Direct Export
7. Step 4: re-query for the new rec-number
8. Failure cases specific to import
9. What is deliberately not attempted

## The rule: EndNote owns every library write

The `.enl` library and its `.Data` directory are user data of the highest value in this
whole pipeline. Losing a Word document costs an afternoon. Corrupting a library that took
years to build costs the research.

The safety rule is about **who performs the write**, not whether import is automated:

- No code path in this project opens `sdb.eni` or another library database file for writing.
- The SQLite reader copies `<Library>.Data\sdb\sdb.eni` to a temporary file first, opens the
  **copy** with `mode=ro`, and deletes the copy in a `finally` block. The original is never
  the file that gets opened.
- `import_reference_file` opens the exact target `.enl` in EndNote and passes the `.ris`
  file to EndNote's native Direct Export path. EndNote itself assigns record numbers and
  updates every related library structure.
- When a target library does not yet exist, the tool can clone a **library with zero stored
  references, including trash**, as an empty template. EndNote must be fully closed during
  this operation. It copies the `.enl` and `.Data` pair together,
  verifies the new target is still empty, and only then invokes Direct Export.

If an agent is asked to add a paper, it must search first, export RIS when needed, call
`import_reference_file`, then search again to obtain EndNote-assigned `rec_number` values.
Never stop merely because the library is empty; an empty library is a valid import target.

## Why Direct Export is the supported automation path

Three independent reasons, none of them a matter of taste:

**The earlier measurement missed EndNote Direct Export.** EndNote's official help states
that `.RIS` is a Direct Export file type and that EndNote recognizes its import filter and
imports it into the open library. The local end-to-end spike confirmed the route: open the
target `.enl`, invoke `EndNote.exe` with the RIS path, then observe the live-reference count
increase and re-query the assigned record numbers.

Official EndNote documentation: [Enabling Direct Export](https://docs.endnote.com/docs/endnote/2025/v1/windows/en/content/appendices/enabling_direct_export.htm)
and [Direct Export providers and formats](https://docs.endnote.com/docs/endnote/2025/v1/windows/en/content/appendices/direct_export_providers.htm).

**The CWYW macro surface is still irrelevant to library import.** The complete macro inventory extracted from the
CWYW template contains `ENExportWordCitations`, `ENExportWordMasterList` and
`ENExportTravelingLibrary`. Every one of them moves data out of Word or out of a library.
There is no import macro, no add-reference macro, and no library-write macro of any kind.
Direct Export belongs to the EndNote desktop executable, not to Word CWYW.

**Direct SQLite writes are unsafe here.** `sdb.eni` is a real SQLite 3 database, but its
schema declares custom collations `ENCI_Base` and `ENCIN_zh_TW` and a custom SQL function
`EN_MAKE_SORT_KEY` that only EndNote implements. The reader works around them by registering
no-op collations, which is fine for a read-only query and catastrophic for a write: any
index or sort key written under a no-op collation is wrong for EndNote's own comparison
rules. Add the sync state, the search index, and the other tables in the `.Data` directory
that a bare `INSERT` would leave untouched, and a hand-written record is a corruption event
waiting for the next library rebuild.

## The four steps

```text
1. search_endnote_library        confirm the reference is absent
2. export_reference_import_file  write a UTF-8 RIS file when needed
3. import_reference_file         EndNote Direct Export performs the import
4. search_endnote_library        re-query to learn the new rec_number
   precheck_endnote_manifest     confirm the manifest is renderable
```

All four steps are tool calls. Do not skip step 4: `imported_count` confirms that the
library count changed, while the fresh DOI search proves which records EndNote assigned.

## Step 1: check the reference is not already there

Search before generating anything:

```console
python scripts\endnote_library.py search --library "<LIBRARY>.enl" --doi 10.3168/jds.2022-22812
python scripts\endnote_library.py search --library "<LIBRARY>.enl" --title "lameness detection" --limit 10
```

Or the MCP tool `search_endnote_library` with the same arguments.

Search matches case-insensitively across `author`, `title`, `secondary_title`,
`electronic_resource_number` (the DOI column), `keywords` and `abstract`. `--year` is an
exact match. Records in the trash are excluded, because the reader always filters
`trash_state = 0`.

Why this step is mandatory: importing a reference that already exists creates a **second**
record with a **different** `rec_number`. Both records then look correct in a search result,
citations that point at the older one and citations that point at the newer one format
identically, and the duplicate only becomes visible much later when the bibliography shows
the same work twice. Deduplicating a library after the fact is manual work for the user.
Five seconds of searching prevents it.

## Step 2: generate the RIS file

```console
python scripts\endnote_library.py export-ris --output "<OUT>\new-references.ris" --references "<IN>\references.json"
```

MCP tool: `export_reference_import_file`. It runs in-process, not as a subprocess, and it is
annotated as a write tool because it creates a file on disk. It writes only to the path you
give it.

Input is a JSON array of reference objects. Every record requires a DOI. Empty input is rejected rather than silently
producing an empty file:

```json
[
  {
    "type": "JOUR",
    "authors": ["Thomsen, Peter T.", "Shearer, Jane K."],
    "year": "2023",
    "title": "Lameness detection in dairy cattle",
    "journal": "Journal of Dairy Science",
    "volume": "106",
    "issue": "5",
    "pages": "3401-3412",
    "doi": "10.3168/jds.2022-22812",
    "url": "https://doi.org/10.3168/jds.2022-22812",
    "abstract": "Short abstract text.",
    "keywords": ["dairy cattle", "lameness", "precision livestock farming"],
    "publisher": ""
  }
]
```

Output constraints that the exporter guarantees:

- UTF-8 encoding.
- CRLF line endings, which is what the RIS format expects and what EndNote's importer is
  least likely to argue with.
- `TY` is the first tag of every record.
- `ER  - ` is the last line of every record.

Field intent, one row per input key:

| Input key | RIS tag | Notes |
|---|---|---|
| `type` | `TY` | Reference type code, for example `JOUR`, `BOOK`, `CHAP`, `CONF`, `THES`. A value the exporter does not recognise falls back to `JOUR` with a warning on stderr, so read the emitted `TY` line before importing a book or a thesis. |
| `authors[]` | `AU` | One tag per author, `Last, First` form, in author order. |
| `year` | `PY` | Four digit year. |
| `title` | `TI` | Primary title. |
| `journal` | `T2` | Secondary title. Maps to the `secondary_title` column after import. |
| `volume` | `VL` | |
| `issue` | `IS` | |
| `pages` | `SP`, `EP` | A range such as `3401-3412` is split into `SP  - 3401` and `EP  - 3412`. A value with no range separator is written as a single `SP`. |
| `doi` | `DO` | Required. Resolver URLs and repeated resolver prefixes are normalised to one bare DOI before export. |
| `url` | `UR` | |
| `abstract` | `AB` | |
| `keywords[]` | `KW` | One tag per keyword. |
| `publisher` | `PB` | |

The emitted file is plain text. Open it and read it before handing it to a user. The
exporter is the authority on the exact tags it writes; this table states the intent of each
field, and a two second look at the file confirms the reality.

## Step 3: import through EndNote Direct Export

Call `import_reference_file` with absolute paths:

```json
{
  "library_path": "C:\\Research\\Project.enl",
  "ris_path": "C:\\Research\\endnote-for-word-references.ris"
}
```

The bridge opens that exact library even when EndNote already has another library open,
confirms the absolute target through EndNote's recent-library registry value rather than a
possibly ambiguous window title, passes the RIS file to `EndNote.exe`, and polls the
read-only counter until the live count increases or the timeout expires. It reports
`live_references_before`,
`live_references_after`, and `imported_count`.

To create a new target first, also provide `empty_library_template_path`. EndNote must be
fully closed and the template must contain zero stored references, including trash. The
tool refuses a non-empty template and refuses to overwrite either an existing target `.enl`
or an orphaned target `.Data` path:

```json
{
  "library_path": "C:\\Research\\NewProject.enl",
  "ris_path": "C:\\Research\\endnote-for-word-references.ris",
  "empty_library_template_path": "C:\\Documents\\Empty EndNote Library.enl"
}
```

Direct Export recognizes RIS and chooses its filter without driving the File > Import
dialog, so no `SendKeys`, localized control labels, or modal UI automation is involved.
If `imported_count` is zero, the tool reports `ok: false`; still search by DOI before retrying,
because every record may already exist even though no new row was observed. If the DOI is
absent, EndNote rejected the file or discarded it according to its current settings.

## Step 4: re-query for the new rec-number

The `rec_number` used everywhere in this toolchain is the SQLite `refs.id` primary key.
EndNote assigns it during import. It is not predictable and it is not necessarily one more
than the previous highest id, because deletions, trash state and library maintenance all
leave gaps.

Never compute it, never assume it, always re-query:

```console
python scripts\endnote_library.py search --library "<LIBRARY>.enl" --doi 10.3168/jds.2022-22812
python scripts\endnote_library.py get --library "<LIBRARY>.enl" --rec-number 412
```

A returned record has this shape, with `abstract_head` truncated to the first 300 characters
so that a search never dumps whole abstracts into a conversation:

```json
{"rec_number": 412, "reference_type": 17, "authors": ["Thomsen, Peter T."],
 "first_author_surname": "Thomsen", "year": "2023", "title": "...", "journal": "...",
 "volume": "106", "issue": "5", "pages": "3401-3412", "doi": "...", "url": "",
 "abstract_head": "..."}
```

Then write the manifest entry using `rec_number`, `first_author_surname`, `year`, and the
bare `doi` exactly
as returned, and run the pre-check before rendering:

```console
python scripts\endnote_library.py precheck --library "<LIBRARY>.enl" --manifest "<PATH>\manifest.json"
```

`precheck` verifies, for every target, that the record exists, is not in the trash, that its
first author surname matches the manifest `author` case-insensitively and accent-tolerantly,
that its year and bare DOI match, and that the DOI resolves through doi.org. It reports every problem it finds, not just the first one, so a
manifest with three bad targets takes one round trip to fix rather than three.

This is the step that prevents the "select matching reference" modal dialog described in
`troubleshooting.md`. A `rec_number` that does not exist in the library hangs Word until the
watchdog kills it.

## Failure cases specific to import

| Symptom | Cause | Fix |
|---|---|---|
| Export refuses to run | The references array is empty | Supply at least one reference object. An empty RIS file has no legitimate use here. |
| Export refuses to run with a DOI error | The DOI is empty or malformed | Supply a valid DOI. A DOI URL is accepted and normalised, but the RIS `DO` tag is always bare. |
| Export refuses to run and names an index | That reference has neither a `title` nor any `authors`, so nothing identifies it after import | Fill in at least one of the two before exporting. |
| `imported_count` is 0 | Records already exist, EndNote discarded duplicates, or Direct Export rejected the file | Search by DOI first. If absent, confirm every record starts with `TY`, regenerate the RIS, and retry once. |
| New target cannot be created | EndNote is running, the template is non-empty, or its `.Data` pair is missing | Fully close EndNote, use a verified zero-record template, and keep the `.enl` + `.Data` pair together. |
| RIS is rejected before EndNote opens | The file is not valid UTF-8 or contains no `TY  - ` record | Regenerate it with `export_reference_import_file`, or transcode a trusted external RIS to UTF-8 and retry. |
| The new record is not found by search | The import went into a different library | Confirm which library was open during the import, and that it is the same path the manifest names. |
| Two records match the search | The reference already existed and step 1 was skipped | Have the user merge or delete the duplicate in EndNote, then re-query. |
| `precheck` reports an author mismatch | The manifest carries a full name or a second author instead of the first author surname | Use `first_author_surname` from the search result verbatim. |
| `precheck` reports `library_doi_not_canonical` | The EndNote DOI field contains `doi:` or a resolver URL | Edit the EndNote record so its DOI field contains only the bare DOI, then run precheck again. |
| `precheck` reports `doi_not_resolvable` | doi.org has no successful handle for the DOI | Correct the DOI from an authoritative source. Never cite the record while this gate fails. |
| A newly imported record cannot be read | EndNote is mid-write and holds the SQLite file | Let the import finish, then re-run the query. The reader copies the file, so a partially written source produces a read error rather than bad data. |

## What is deliberately not attempted

- Writing to `sdb.eni` directly, under any circumstance.
- Driving the EndNote import dialog with synthetic keystrokes or UI automation. Direct Export needs no dialog.
- Creating a blank library database schema from scratch. New targets are cloned only while EndNote is fully closed, from a verified empty EndNote library so the installed EndNote version owns the schema.
- Repairing a library or emptying the trash.
- Editing an existing reference. The same read-only rule applies: correcting a record is a
  human action in EndNote, after which the tool re-queries and continues.
- Using `ENEditLibraryReferences` from the CWYW macro set. It hands control to the EndNote
  UI and is not a programmatic write path.

Everything in this list is a deliberate boundary. Import is safe here because EndNote,
not this project, performs the database write.
