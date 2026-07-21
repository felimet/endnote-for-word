# EndNote for Word

[繁體中文](README-zhTW.md)

An installable AI Agent skill, Codex/Claude plugin, and MCP server that inserts genuine EndNote Cite While You Write (CWYW) citations and bibliographies into Microsoft Word documents, driven by a reviewable JSON manifest.

## What this tool does

- Replaces `{{CITE:id}}` markers in a DOCX with real CWYW citations, so every citation becomes an `ADDIN EN.CITE` field carrying an EndNote traveling library, and the bibliography becomes an `ADDIN EN.REFLIST` field.
- Never types citation text, never fabricates field codes in OOXML, and never passes off a plain-text reference list as a real bibliography.
- Reads your `.enl` library read-only so an agent can search for references and confirm every record number before Word is ever opened.
- Pre-checks the manifest against the library and against the installed output styles, because an unknown record number or an unknown style name makes EndNote open a modal dialog that blocks automation until the process is killed.
- Switches the document output style (for example APA 7th to IEEE) programmatically and rebuilds the bibliography, with no dialog.
- Validates the finished document structurally and reports counts, style, and the libraries recorded in the document.
- Exports a RIS file so references that are not yet in your library can be imported in one step.

## Requirements

- Windows 10 or 11.
- Microsoft Word desktop with COM automation.
- Desktop **EndNote** for Windows with the **Cite While You Write** Word add-in installed and enabled. This is a hard prerequisite: without the CWYW add-in there is no supported way to produce native citation fields, and the tool refuses to fake them.
- Python 3 on `PATH`.
- PowerShell 7 as `pwsh.exe`, or Windows PowerShell 5.1 as `powershell.exe`.

This project is developed and verified against **EndNote 21** with the 64-bit CWYW add-in (`EndNote.WordAddin.Connect`) and its global Word template `EndNote Cwyw.dotm`. The EndNote executable, the styles directory, and the CWYW template are discovered from the registry and from the installation directories at run time, newest version first, so no version number is hardcoded.

EndNote and Microsoft Office are proprietary products and are not distributed by this repository. The EndNote **web** or **online** plugin is not a substitute for the desktop CWYW add-in.

## Terminal compatibility

The bridge is a single PowerShell script written for both PowerShell 7 and Windows PowerShell 5.1. Four terminals are supported.

| Active terminal | Required action |
|---|---|
| PowerShell 7+ | Run the commands directly with `pwsh.exe`. |
| Windows PowerShell 5.1 | Supported. Run with `powershell.exe`, but prefer `pwsh.exe` when it is installed. |
| Bash on Windows, including Git Bash | Invoke Windows `pwsh.exe`, and fall back to `powershell.exe`. |
| CMD | Invoke `pwsh.exe` with the same arguments, and fall back to `powershell.exe`. |

From WSL Bash, invoke the Windows-side executable. Linux `pwsh` cannot automate Windows Word COM. If neither `pwsh.exe` nor `powershell.exe` is available, stop and install PowerShell rather than silently downgrading. See the [PowerShell installation guide](https://learn.microsoft.com/powershell/scripting/install/installing-powershell-on-windows).

The following command works from PowerShell 7, Windows PowerShell 5.1, Git Bash, and CMD:

```console
pwsh.exe -NoProfile -ExecutionPolicy Bypass -File scripts/endnote-word.ps1 -Action probe
```

## Install with an AI agent

If you want to use EndNote for Word in Claude Code, Claude Desktop, Codex, or ChatGPT Desktop, paste the following prompt. The agent will configure it for the current environment:

```text
Install or upgrade the EndNote for Word toolkit from https://github.com/felimet/endnote-for-word. Detect my available terminal and use PowerShell 7, Windows PowerShell 5.1, Bash on Windows including Git Bash, or CMD; prefer pwsh.exe when it exists and fall back to powershell.exe. From WSL Bash, invoke the Windows-side executable rather than Linux pwsh. If neither is available, stop and tell me to install PowerShell using https://learn.microsoft.com/powershell/scripting/install/installing-powershell-on-windows. Verify Microsoft Word desktop plus desktop EndNote with the Cite While You Write add-in, install the portable skill, register the local stdio MCP server, run the probe and the repository tests, preserve existing agent configuration, and report every changed file. Do not claim success unless the probe reports cwyw_macro_callable true and validation returns ok: true.
```

Then tell your agent which DOCX to edit, which `.enl` library to cite from, and which output style to use.

Platform-specific skill locations and MCP configuration are documented in the [installation matrix](skills/endnote-for-word/references/installation-matrix.md).

## Add the skill and MCP to each agent

Replace `<REPO_ROOT>` with the absolute local checkout path. Preserve existing MCP entries when editing configuration files.

### Codex

Copy `skills/endnote-for-word` to `%USERPROFILE%\.codex\skills\endnote-for-word` or `%USERPROFILE%\.agents\skills\endnote-for-word`, then register the server:

```console
codex mcp add endnote-for-word -- pwsh.exe -NoLogo -NoProfile -NonInteractive -ExecutionPolicy Bypass -File "<REPO_ROOT>\scripts\run-mcp.ps1"
codex mcp get endnote-for-word
```

The release also contains `.codex-plugin/plugin.json`, `.codex-mcp.json`, and `dist/endnote-for-word-plugin.zip` for plugin-aware deployment.

### Claude Code

Copy `skills/endnote-for-word` to `%USERPROFILE%\.claude\skills\endnote-for-word`, then register the server:

```console
claude mcp add --scope user endnote-for-word -- pwsh.exe -NoLogo -NoProfile -NonInteractive -ExecutionPolicy Bypass -File "<REPO_ROOT>\scripts\run-mcp.ps1"
claude mcp get endnote-for-word
```

The release also contains `.claude-plugin/plugin.json`, `.mcp.json`, and the combined plugin package. Restart Claude Code after installation.

### Claude Desktop

Upload `dist/endnote-for-word.skill` from **Customize > Skills**. Then merge this server into `%APPDATA%\Claude\claude_desktop_config.json` and restart Claude Desktop:

```json
{
  "mcpServers": {
    "endnote-for-word": {
      "command": "pwsh.exe",
      "args": ["-NoLogo", "-NoProfile", "-NonInteractive", "-ExecutionPolicy", "Bypass", "-File", "<REPO_ROOT>\\scripts\\run-mcp.ps1"]
    }
  }
}
```

See Anthropic's [skill upload instructions](https://support.claude.com/en/articles/12512180-use-skills-in-claude) and [local MCP setup guide](https://modelcontextprotocol.io/docs/develop/connect-local-servers).

### ChatGPT Desktop

If **Plugins** is available for the account or workspace, install or enable the packaged capability through the supported [ChatGPT plugin workflow](https://help.openai.com/en/articles/20001256). ChatGPT Desktop does not automatically discover this local repository. OpenAI's current full-MCP documentation applies to ChatGPT web; do not assume the desktop app has the same developer-mode surface unless it is visible for the account.

ChatGPT cannot connect directly to the bundled local stdio MCP command. It requires a remote MCP endpoint or a [Secure MCP Tunnel](https://help.openai.com/en/articles/12584461-developer-mode-and-full-mcp-connectors-in-chatgpt-beta), with Developer mode or Apps enabled as allowed by the account or workspace. Because EndNote and Word automation must execute on the interactive Windows desktop, the endpoint or tunnel must route execution to that Windows host. This repository currently ships only the local stdio server.

## Workflow

1. **Probe.** Confirm Word COM, the EndNote executable, the CWYW add-in registration, the loaded CWYW template, and that a CWYW macro is callable. The macro check runs on a throwaway empty document so it cannot touch your files.
2. **Choose a style.** List the installed `.ens` output styles and pick an exact name. A style name that does not resolve to an installed file makes EndNote open a modal dialog and hangs the run.
3. **Find the references.** Search the `.enl` library read-only and collect the record numbers, first-author surnames, and years.
4. **Write the manifest.** One entry per citation, each with a unique marker that occurs exactly once in the document.
5. **Pre-check.** Verify every record number exists in the library, is not in the trash, and that the surname and year in the manifest match the stored record. A record number that does not exist makes EndNote open a "select matching reference" dialog and hangs the run.
6. **Render.** The bridge copies the input to the output path, sets the document's `EN.Layout` style variable, replaces each marker with EndNote temporary citation text such as `{Thomsen, 2023 #1}`, runs the CWYW `ENFormatBibliography` macro, waits for the expected field count, and saves atomically.
7. **Validate.** Re-open the result read-only and run every gate below.

Prepare markers in your DOCX first, for example:

```text
Automated lameness scoring is now standard {{CITE:thomsen2023}}.
Two independent trials agree {{CITE:merged}}.
```

## Manifest contract

See [`examples/example-manifest.json`](examples/example-manifest.json) for a complete file. Its `library.path` is a placeholder, so point it at your own `.enl` library before running it, otherwise the manifest fails the `library.path` existence check.

```jsonc
{
  "schema_version": 1,
  "library": { "path": "C:\\EndNoteLibraries\\ExampleLibrary.enl" },
  "style": "APA 7th",
  "citations": [
    {
      "id": "thomsen2023",
      "marker": "{{CITE:thomsen2023}}",
      "targets": [ { "rec_number": 1, "author": "Thomsen", "year": 2023 } ],
      "pages": "45-46"
    }
  ],
  "bibliography": { "expected": true }
}
```

| Rule | Requirement |
|---|---|
| `schema_version` | Must be `1`. |
| `citations` | Non-empty array. Unknown top-level and per-citation keys are rejected. |
| `id` | Matches `^[A-Za-z][A-Za-z0-9_-]*$` and is globally unique. |
| `marker` | Globally unique, and must occur exactly once in the document. |
| `targets` | Non-empty. Each target has an integer `rec_number` greater than zero, a non-empty `author` surname, and an integer `year`. |
| `pages` | Optional. Matches `^[0-9]+(\s*-\s*[0-9]+)?$`. Allowed only when `targets` has exactly one entry. |
| `suffix` | Optional free text such as `", p. 99"`. Allowed only when `targets` has exactly one entry. |
| `library.path` | Must exist and have a sibling `<name>.Data` directory. |
| `style` | Non-empty string that resolves to an installed `.ens` style. |

Each citation becomes one temporary citation, so a merged entry with several targets produces one field, not several:

```text
{Thomsen, 2023 #1}
{Thomsen, 2023 #1@45-46}
{Magrin, 2023 #2; Bernhard, 2020 #3}
```

`pages` and `suffix` apply to the last target only, which is why they are restricted to single-target citations.

## Validation gates

`validate` fails the run with `ok: false` and exit code `3` when any gate fails.

1. No `{{CITE:` marker text remains anywhere in the document.
2. No residual temporary citation text remains, matched with `\{[^{}]*#\d+[^{}]*\}` over the document text.
3. The native citation field count equals the manifest citation count. Counting rule: fields whose code starts with `ADDIN EN.CITE`, minus fields whose code starts with `ADDIN EN.CITE.DATA`, because a large traveling library payload is moved into a nested data field.
4. When `bibliography.expected` is true, exactly one `ADDIN EN.REFLIST` field exists and its result text is not empty.
5. Every manifest record number appears as `<rec-number>N</rec-number>` in some citation field's traveling library XML, and the matching author surname appears in the same field.
6. The document text contains no `Error!`, `Reference source not found`, `{Citation}`, or `{Bibliography}` placeholder strings.
7. The `EN.Layout` document variable exists and its style matches the manifest style, compared case-insensitively.
8. The `EN.Libraries` document variable exists and lists every manifest record number inside `<record-ids>`.

## Version 1 limitations

| Limitation | Reason |
|---|---|
| No citation prefixes | The backslash prefix syntax, for example `{see \Thomsen, 2023 #1}`, opens a modal EndNote dialog that blocks automation. The manifest rejects `prefix` outright. |
| No excluded authors or years | `{\Thomsen, 2023 #1}` and the `ENExcludeAuthor` family are dialog-driven and hang the run. Use `suffix` text instead when the style permits it. |
| The library is read-only | The tool never writes to `.enl` files or their `.Data` directories. It copies the internal database to a temporary file, reads the copy, and deletes it. |
| Adding references is semi-automatic | New references are exported to a RIS file that you import into EndNote yourself, through **File > Import**. The tool does not write into your library. |
| Bibliography position is the CWYW default | EndNote places the reference list where Cite While You Write puts it, normally at the end of the document. Move it in EndNote or in Word if you need it elsewhere; this tool does not relocate it. |
| Instant formatting is not toggled | The CWYW toggle macro has no read-back, so calling it blindly would flip an unknown state. Change it in the EndNote ribbon if you need to. |
| Converting citations to plain text is manual | The CWYW macro behind it, `ENRemoveFieldCodes`, raises a modal confirmation dialog and never returns under automation, so there is no `remove-field-codes` action and no MCP tool for it. Do it yourself from the Word ribbon with **EndNote > Convert Citations and Bibliography > Convert to Plain Text**, which produces a new copy and leaves the original field codes intact. The result is one-way: EndNote can no longer update the citations in the converted copy. |

## Verify the installation

```console
pwsh.exe -NoProfile -ExecutionPolicy Bypass -File scripts/endnote-word.ps1 -Action probe
pwsh.exe -NoProfile -ExecutionPolicy Bypass -File scripts/endnote-word.ps1 -Action list-styles -Style IEEE
pwsh.exe -NoProfile -ExecutionPolicy Bypass -File tests/run-tests.ps1
```

The bridge preserves source files, writes through a temporary sibling DOCX, and refuses to replace an existing output unless `-Overwrite` is explicit. Live integration tests are opt-in and need a real library path.

## Repository layout

| Path | Purpose |
|---|---|
| `skills/endnote-for-word/` | Cross-agent skill, references, and launcher |
| `scripts/` | Word and EndNote automation bridge, library reader, MCP server, and packager |
| `config/defaults.json` | Default output style, temporary citation delimiters, and validation gates |
| `examples/` | Reference manifest |
| `tests/` | Static, MCP protocol, offline structure, library, and live integration checks |
| `evals/` | Evaluation cases |

## Packaging

```console
python scripts/package_plugin.py
```

The command creates `dist/endnote-for-word-plugin.zip` and its SHA-256 file. The standalone skill package is `dist/endnote-for-word.skill`.

## License

[MIT](LICENSE)
