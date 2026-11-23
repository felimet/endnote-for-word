# endnote-for-word

Claude Code and Codex plugin whose stdio MCP server (stdlib-only Python) drives Microsoft Word and EndNote 21 through COM to insert genuine Cite While You Write citations; it runs only on Windows with Word, EndNote 21, Python 3 and PowerShell 7 or 5.1, in an unlocked interactive desktop session.

Project facts and conventions in this file take precedence over global instructions; the global "risk" rules are never relaxed by this file.

## Commands

- Verify (must pass before declaring work done): `pwsh ./tests/run-tests.ps1` (offline: Python unittest discovery plus a parse check of every .ps1 under pwsh.exe and powershell.exe; about one minute)
- Package (CI runs it after the tests): `python ./scripts/package_plugin.py` (writes `dist/endnote-for-word-plugin.zip`, `dist/endnote-for-word.skill` and their `.sha256` sidecars)
- One test module: `python -m unittest discover -s tests -t tests -p "test_endnote_session.py"`
- Live run against the real Word and EndNote: `pwsh ./tests/run-tests.ps1 -IncludeLive` (not run when this file was written; see Boundaries)

## Conventions

- Every .ps1 must parse and run under Windows PowerShell 5.1 and PowerShell 7: no ternary, `??`, `?.`, `clean {}`, `-Parallel` or `Join-String`. `tests/test_static.py` rejects any `?` outside string literals in .ps1 code. Keep .ps1 files ASCII, because 5.1 reads a BOM-less file in the ANSI code page.
- Python under `scripts/` uses the standard library only; the plugin must start on a bare CPython.
- `scripts/endnote-word.ps1` runs an action as soon as it is loaded. Offline tests lift single functions out of it through the PowerShell AST (`tests/test_bridge_runtime.py`), so put new decision logic in a self-contained function that takes plain values.
- Bridge contract: exactly one compact JSON line, last on stdout; logs go to stderr; exit codes 0 ok, 1 error, 2 probe not ready, 3 postcondition failed.
- Comments and docstrings that say "measured" record EndNote or Word behaviour observed on real installs (dialogs, macros, window titles). Do not simplify such a workaround without measuring again.
- The version string lives in `.claude-plugin/plugin.json`, `.codex-plugin/plugin.json`, `skills/endnote-for-word/SKILL.md`, `scripts/mcp_server.py`, `tests/test_mcp_protocol.py`, `tests/test_static.py` and `CHANGELOG.md`; bump them together.
- Prose files (.md, .json, .py, .ps1, .yaml, .txt, .ris) use hyphens, never em or en dashes, and carry no bracketed TODO or PLACEHOLDER markers or personal home-directory paths; `tests/test_static.py` enforces this.
- CHANGELOG follows Keep a Changelog. Commit subjects follow `type(scope): summary`, for example `fix(bridge): ...`, `feat(mcp): ...`, `chore(release): 1.5.2`.

## Boundaries

- Never write an EndNote library (`.enl`, `<name>.Data/sdb/sdb.eni`) from code. The only library write is `import-ris`, which hands the RIS file to EndNote Direct Export for exactly the requested library.
- The installed plugin copy in the Claude Code plugin cache is not this repository; never edit it. Change this repository and reinstall.
- Ask before running `-IncludeLive` or any bridge action other than probe, list-styles or validate on this machine: they open the user's own Word and EndNote and can write documents.
- `dist/` holds build output; do not commit it.

## References

- Release history: `CHANGELOG.md`. Security policy: `SECURITY.md`.
- Manifest schema, bridge command lines and validation gates: `skills/endnote-for-word/references/workflow-and-schema.md`.
- Known EndNote and Word failure modes: `skills/endnote-for-word/references/troubleshooting.md`.
- Agent scenarios: `evals/evals.json`.
