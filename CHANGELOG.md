# Changelog

All notable changes to this project are documented in this file. The format follows
[Keep a Changelog](https://keepachangelog.com/en/1.1.0/), and the project adheres to
[Semantic Versioning](https://semver.org/spec/v2.0.0.html).

## [Unreleased]

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

### Changed

- `import_reference_file` no longer requires `library_path`; it requires `ris_path` plus either
  `library_path` or `document_path`.

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

[Unreleased]: https://github.com/felimet/endnote-for-word/compare/v1.3.0...HEAD
[1.3.0]: https://github.com/felimet/endnote-for-word/releases/tag/v1.3.0
[1.2.0]: https://github.com/felimet/endnote-for-word/tree/69a5bb2
