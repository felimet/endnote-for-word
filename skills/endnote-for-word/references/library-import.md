# Adding a new reference to the library

## Contents

1. The rule: this tool never writes to a library
2. Why the workflow is semi-automatic
3. The four steps
4. Step 1: check the reference is not already there
5. Step 2: generate the RIS file
6. Step 3: import through the EndNote desktop UI
7. Step 4: re-query for the new rec-number
8. Failure cases specific to import
9. What is deliberately not attempted

## The rule: this tool never writes to a library

The `.enl` library and its `.Data` directory are user data of the highest value in this
whole pipeline. Losing a Word document costs an afternoon. Corrupting a library that took
years to build costs the research.

So the rule is absolute and has no override switch, no environment variable, and no
"advanced" mode:

- No code path in this project opens a library file for writing.
- The SQLite reader copies `<Library>.Data\sdb\sdb.eni` to a temporary file first, opens the
  **copy** with `mode=ro`, and deletes the copy in a `finally` block. The original is never
  the file that gets opened.
- The bridge never passes a library path to Word for anything except telling EndNote which
  library to have open.
- Adding a reference is therefore a human action performed in the EndNote desktop
  application. This tool prepares the input for that action and verifies the result
  afterwards.

If an agent is asked to "add this paper to my library", the honest answer is that it can
produce a ready-to-import file and tell the user exactly which two clicks to perform, then
confirm the outcome. It cannot perform the insertion itself.

## Why the workflow is semi-automatic

Three independent reasons, none of them a matter of taste:

**No supported write interface was found.** The measurement pass over this installation
turned up no programmatic path that EndNote itself supports for adding a record: not in the
CWYW macro surface, not in the COM add-in, and not in anything the desktop application
exposes to automation. If a vendor API exists outside what was measured here, this project
still does not target it, and it does not substitute an undocumented channel for one.

**The CWYW macro surface is export-only.** The complete macro inventory extracted from the
CWYW template contains `ENExportWordCitations`, `ENExportWordMasterList` and
`ENExportTravelingLibrary`. Every one of them moves data out of Word or out of a library.
There is no import macro, no add-reference macro, and no library-write macro of any kind.
The absence is structural, not an oversight in the inventory.

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
2. export_reference_import_file  write a UTF-8 RIS file
3. EndNote desktop, File > Import   human performs the import
4. search_endnote_library        re-query to learn the new rec_number
   precheck_endnote_manifest     confirm the manifest is renderable
```

Steps 1, 2 and 4 are tool calls. Step 3 is the human. Do not skip step 4: it is the only
verification that the import happened at all, because nothing in this toolchain can observe
the EndNote UI.

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

Input is a JSON array of reference objects. Empty input is rejected rather than silently
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
| `doi` | `DO` | Lands in `electronic_resource_number` after import. |
| `url` | `UR` | |
| `abstract` | `AB` | |
| `keywords[]` | `KW` | One tag per keyword. |
| `publisher` | `PB` | |

The emitted file is plain text. Open it and read it before handing it to a user. The
exporter is the authority on the exact tags it writes; this table states the intent of each
field, and a two second look at the file confirms the reality.

## Step 3: import through the EndNote desktop UI

Hand the user the file path and these instructions. The spike never automated this dialog,
and this project does not attempt to drive it with synthetic keystrokes, because dialog
layouts and control labels change between EndNote builds and interface languages. The
labels below are what the operator confirms on screen, not values this tool sets.

1. Open the target library in EndNote. It must be the same library the manifest points at.
2. Choose `File` then `Import` then `File...`.
3. Select the generated `.ris` file.
4. Set the import option to the Reference Manager RIS filter.
5. Choose a duplicate handling option. Prefer the option that discards duplicates, given
   that step 1 already established the reference should be new.
6. Set the text translation to Unicode UTF-8, matching the file the exporter wrote. This is
   the setting that decides whether non-ASCII author names arrive intact or mangled.
7. Run the import and read the count EndNote reports.

If EndNote reports zero imported references, do not proceed to step 4 and do not guess a
`rec_number`. Read the file, fix the RIS, regenerate it and import again.

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

Then write the manifest entry using `rec_number`, `first_author_surname` and `year` exactly
as returned, and run the pre-check before rendering:

```console
python scripts\endnote_library.py precheck --library "<LIBRARY>.enl" --manifest "<PATH>\manifest.json"
```

`precheck` verifies, for every target, that the record exists, is not in the trash, that its
first author surname matches the manifest `author` case-insensitively and accent-tolerantly,
and that its year matches. It reports every problem it finds, not just the first one, so a
manifest with three bad targets takes one round trip to fix rather than three.

This is the step that prevents the "select matching reference" modal dialog described in
`troubleshooting.md`. A `rec_number` that does not exist in the library hangs Word until the
watchdog kills it.

## Failure cases specific to import

| Symptom | Cause | Fix |
|---|---|---|
| Export refuses to run | The references array is empty | Supply at least one reference object. An empty RIS file has no legitimate use here. |
| Export refuses to run and names an index | That reference has neither a `title` nor any `authors`, so nothing identifies it after import | Fill in at least one of the two before exporting. |
| EndNote imports 0 references | Wrong import filter, or a malformed first tag | Confirm the filter is the Reference Manager RIS filter and that every record starts with `TY`. |
| Author names arrive mangled | Text translation was not set to Unicode UTF-8 | Delete the imported records, set UTF-8, import again. |
| The new record is not found by search | The import went into a different library | Confirm which library was open during the import, and that it is the same path the manifest names. |
| Two records match the search | The reference already existed and step 1 was skipped | Have the user merge or delete the duplicate in EndNote, then re-query. |
| `precheck` reports an author mismatch | The manifest carries a full name or a second author instead of the first author surname | Use `first_author_surname` from the search result verbatim. |
| A newly imported record cannot be read | EndNote is mid-write and holds the SQLite file | Let the import finish, then re-run the query. The reader copies the file, so a partially written source produces a read error rather than bad data. |

## What is deliberately not attempted

- Writing to `sdb.eni` directly, under any circumstance.
- Driving the EndNote import dialog with synthetic keystrokes or UI automation.
- Creating a library, repairing a library, or emptying the trash.
- Editing an existing reference. The same read-only rule applies: correcting a record is a
  human action in EndNote, after which the tool re-queries and continues.
- Using `ENEditLibraryReferences` from the CWYW macro set. It hands control to the EndNote
  UI and is not a programmatic write path.

Everything in this list is a deliberate boundary, not a gap waiting to be filled in a later
version. The value of the library being read-only comes from it being read-only without
exceptions.
