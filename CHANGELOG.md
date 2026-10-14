# Changelog

All notable changes to this project are documented in this file. The format follows
[Keep a Changelog](https://keepachangelog.com/en/1.1.0/), and the project adheres to
[Semantic Versioning](https://semver.org/spec/v2.0.0.html).

## [1.5.1] - 2026-10-14

### Fixed

- Opening Word by hand no longer reopens the working copies of jobs the server had to end. Word
  records every open document under `Resiliency\DocumentRecovery` and reopens each recorded one
  in its own window at the next start when its Word was killed. The bridge now names every
  document its Word opens or saves as (`WORD_DOCUMENT=` on stderr), and after a cancelled or
  timed-out job the server waits for that Word to exit and removes only the entries the run added
  for those documents; the count is reported as `word_recovery_entries_removed`.

## [1.5.0] - 2026-10-14

### Added

- Background jobs. `render_endnote_word_document`, `update_endnote_word_fields` and
  `set_endnote_citation_style` return a `job_id` at once (`background: false` keeps the old
  blocking call). `get_endnote_job_status` reports the state, any lasting dialog under
  `attention` and the result; `cancel_endnote_job` drops a queued job or ends a running one.
- Every action that drives EndNote runs through one queue, and a named mutex
  (`Local\EndNoteForWordAgentBridge`) serialises jobs across every server of the Windows session,
  such as Claude Code and Codex side by side.
- Each job opens its manifest library beside the libraries the user has open and closes only that
  library afterwards. EndNote's recent-library list is restored, so a job does not change the
  library EndNote opens next.
- EndNote's start-up news notices (`EndNote Password Expiration Policy Change`,
  `EndNote 2025 is here!`), which block every CWYW request while visible, are answered with
  Cancel and listed under `notices_dismissed`. No preference changes and no other dialog is
  answered, except the CWYW record chooser below.
- An EndNote the server started closes after `ENDNOTE_IDLE_CLOSE_SECONDS` (default 120) without a
  job, and is ended when it ignores every close request. Bridges and server-started EndNote sit in
  a kill-on-close job object, so they end with the server. A user's own EndNote is never ended.
- An EndNote the server started becomes the user's once the user opens a library in it or shows
  its main window, because EndNote runs one instance per user: the server then never closes or
  ends it, and it outlives the server.

### Changed

- A CWYW `Select Matching Reference` chooser, raised when another open library has overlapping
  record numbers, is answered with Cancel and the job fails with `dialogs_cancelled` and advice.
  Ending the hidden Word at the chooser instead left EndNote unable to exit for minutes.
- `ENDNOTE_READY_TIMEOUT_MS` also bounds the wait for the job's library window.

## [1.4.0] - 2026-10-14

### Added

- A DOCX without a named library gets one beside it. `import-ris` (`import_reference_file`)
  accepts the DOCX as `-InputPath` (`document_path`) instead of `-LibraryPath`, targets
  `<document stem>.enl` in the same folder, creates it on first use and reuses it afterwards.
  A DOCX that already carries `ADDIN EN.` fields is refused, because its citations belong to an
  existing library.
- `configure-defaults -EmptyLibraryTemplatePath` (`configure_endnote_defaults`
  `empty_library_template_path`) registers a zero-record library once. A missing import target is
  cloned from it when no template is passed, and re-running `configure-defaults` keeps it.
  `config/defaults.json` gains `library.empty_template_path`.
- Plugin marketplaces for both clients, so `claude plugin marketplace add felimet/endnote-for-word`
  and `codex plugin marketplace add felimet/endnote-for-word` install the skill and the MCP
  server together, and the clients' update commands upgrade them.
- `SECURITY.md` with private vulnerability reporting, Dependabot updates for GitHub Actions, and a
  plugin security scan workflow pinned to `hashgraph-online/ai-plugin-scanner-action` v1.2.786.

### Changed

- `import_reference_file` no longer requires `library_path`; it requires `ris_path` plus either
  `library_path` or `document_path`.
- Every GitHub Action in CI is pinned to a full commit SHA.
- `requirements.txt` is removed. It declared no package, because the Python side uses only the
  standard library.
- Both plugin manifests declare `homepage` and `repository`; the Codex manifest keeps three
  default prompts, the number Codex shows.

### Fixed

- The MCP server never started as a Codex plugin. Codex reads `.codex-plugin/plugin.json` as a
  legacy manifest: it expands no `${PLUGIN_ROOT}`, exports no plugin-root variable, and starts the
  server in the session folder, so the shared `.mcp.json` launcher looked for `run-mcp.ps1` in the
  user's project. The Codex manifest now points at `.codex-mcp.json`, which runs
  `python scripts/mcp_server.py` with `cwd` set to the installed plugin. The server then leaves that
  folder, because Windows cannot replace a folder a process works in; started anywhere else it
  stays, so relative paths keep their base under Claude Code.
- Codex passes only a whitelist of environment variables, which drops `COMMONPROGRAMFILES*`, `OS`
  and every `ENDNOTE_*` setting, so the probe could not find the CWYW template. `.codex-mcp.json`
  forwards them by name and sets a 1860 s tool limit above the server's 1800 s render watchdog.

## [1.3.0] - 2026-10-13

### Added

- PowerShell 7 auto-install. When the bridge runs under Windows PowerShell 5.1, no `pwsh.exe`
  exists and no host is pinned, it installs PowerShell 7 for the current user with
  `winget install --id Microsoft.PowerShell --scope user --installer-type msix`, which needs no
  administrator rights. winget runs silently with `--accept-package-agreements
  --accept-source-agreements`, so the winget source and package agreements are accepted on the
  user's behalf, and the notice says so.
  - Read-only calls (`probe`, `list-styles`, `validate`) never install. Each one returns a
    `pending` notice, or `unavailable` when winget is missing, until the first other call.
  - The first other call (`render`, `update`, `set-style`, `unformat`, `configure-defaults` or
    `import-ris`) starts the install in the background and keeps running under 5.1.
  - From then on each state change (installing, succeeded, failed, unavailable) is reported once
    as `powershell_notice`, and the agent is told to relay every notice to the user.
  - One attempt is made per user. State and the winget log live in
    `%LOCALAPPDATA%\EndNoteForWordAgent`.
  - Set `ENDNOTE_WORD_AUTO_INSTALL_PWSH=0`, or pin a host with `ENDNOTE_WORD_POWERSHELL`, to
    opt out.
- The MCP server and `invoke-bridge.ps1` resolve the PowerShell host on every call and also find
  PowerShell 7 at its per-user app execution alias, so a newly installed PowerShell 7 takes over
  without a restart. A `powershell_notice` also survives an MCP timeout.
- `scripts/package_plugin.py` also builds the standalone skill package
  `dist/endnote-for-word.skill` with its SHA-256 file, and CI uploads it. The Claude Desktop
  instructions now use it.
- Troubleshooting entries for a `SaveAs2` call that never returns and for the Word recovery state
  that a forced kill leaves behind.

### Changed

- The shipped MCP launch command (`.mcp.json`, `.codex-mcp.json`, and the registration commands
  in the READMEs and the installation reference) is `powershell.exe`, which every supported
  Windows has, so the server starts before PowerShell 7 exists. The bridge still prefers
  PowerShell 7 when it is installed.
- `scripts/run-mcp.ps1` only checks that a PowerShell host exists. It no longer exports the host
  it found as `ENDNOTE_WORD_POWERSHELL`, which now pins the host only when set explicitly.
- Manifests, agent defaults and RIS files are decoded as strict UTF-8 with an optional byte order
  mark. Input that is not valid UTF-8, such as UTF-16 with a byte order mark or Latin-1, is
  rejected with "could not be read as UTF-8 text" instead of being parsed into corrupted author
  names. Most cp950 and cp936 text fails the same way, but legacy bytes that happen to form valid
  UTF-8 cannot be detected. The DOI and library precheck reports the same message.
- Resolved paths in JSON results are plain, full file system paths.

### Fixed

- Under Windows PowerShell 5.1, Word's `SaveAs2` stalled when given an uncast `Join-Path` path
  (measured on Word 16.0.20430 with EndNote 21.5), so the save step that `render`, `update`,
  `set-style` and `unformat` share hung until the watchdog killed Word (observed with
  `render`). The bridge now passes the save path as a plain string. Found while investigating
  [#1](https://github.com/felimet/endnote-for-word/pull/1) by @G-C-M-G.
- Manifests with non-ASCII text failed to parse under Windows PowerShell 5.1 on a non-UTF-8 code
  page. Based on [#1](https://github.com/felimet/endnote-for-word/pull/1) by @G-C-M-G.
- The bridge deadlocked when a child process wrote more than 4 KB to stderr.
- UNC and `\\?\` paths reached Python, Word and .NET in the provider-qualified
  `Microsoft.PowerShell.Core\FileSystem::` form that none of them can open.
- The `unformat` "never in place" guard could be bypassed through an 8.3 short-name alias of the
  input, or with the same UNC path on both sides.
- The MCP render timeout ignored the citation count of a manifest saved with a byte order mark.
- A precheck failure without a problem list no longer prefixes the error with an empty entry.

## [1.2.0] - 2026-07-22

State of `main` when this changelog was started.

[1.5.1]: https://github.com/felimet/endnote-for-word/releases/tag/v1.5.1
[1.5.0]: https://github.com/felimet/endnote-for-word/releases/tag/v1.5.0
[1.4.0]: https://github.com/felimet/endnote-for-word/releases/tag/v1.4.0
[1.3.0]: https://github.com/felimet/endnote-for-word/releases/tag/v1.3.0
[1.2.0]: https://github.com/felimet/endnote-for-word/tree/69a5bb2
