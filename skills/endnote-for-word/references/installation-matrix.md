# Installation matrix

## Contents

1. Prerequisites
2. Terminal selection
3. Concrete command lines per terminal
4. WSL
5. Claude Code
6. Claude Desktop
7. Codex
8. ChatGPT
9. Environment variables
10. Verification after installation

Set `<REPO_ROOT>` to the absolute path of the checked out repository before using any
command below. No command in this document contains a machine specific path.

## Prerequisites

- Windows 10 or Windows 11. The bridge drives Word through COM in the signed in interactive
  desktop session, so there is no headless or server side variant of this workflow.
- Microsoft Word desktop with COM automation available as `Word.Application`.
- EndNote desktop with its Cite While You Write Word integration installed. Everything here
  was measured against EndNote 21. The bridge discovers the executable and the styles
  directory dynamically and never hardcodes a major version, so another installed release is
  found automatically, but it is untested: run the probe and a fixture render before
  trusting one. The CWYW macro template `EndNote Cwyw.dotm` loads itself as a Word global
  template as soon as a document is open. Nothing needs to be added to Word by hand, and
  nothing should be: calling `Application.AddIns.Add` on that template double-loads it.
- At least one EndNote library, consisting of a `<name>.enl` file and its sibling
  `<name>.Data` directory. Both must exist. A `.enl` file on its own is a stub whose first
  four bytes are `Eds1`; the references live in the `.Data` directory.
- Python 3 on `PATH` as `python` or `python.exe`. The Python side of this project is
  standard library only, so no virtual environment or package install is required for it to
  run.
- PowerShell. Either PowerShell 7 (`pwsh.exe`) or Windows PowerShell 5.1
  (`powershell.exe`) is sufficient. The bridge is written to run unmodified on both, and the
  test runner parse checks every `.ps1` under both hosts when both are present.

If no PowerShell host is available, stop and ask the user to install PowerShell 7 from
<https://learn.microsoft.com/powershell/scripting/install/installing-powershell-on-windows>.
Never silently downgrade to some other execution path.

## Terminal selection

Unlike a PowerShell 7 only bridge, this project supports Windows PowerShell 5.1. That
changes the decision table: 5.1 is a supported runtime, not a host you must escape from.

| Environment | Decision |
|---|---|
| PowerShell 7 or later | Run the bridge directly in the active host. |
| Windows PowerShell 5.1 | Supported. Run directly. Prefer `pwsh.exe` when it is also installed, because 7 is the better tested of the two. |
| Bash on Windows, including Git Bash | Invoke the Windows PowerShell executable. Try `pwsh.exe` first, then `powershell.exe`. |
| CMD | Same as Git Bash. Try `pwsh.exe` first, then `powershell.exe`. |
| WSL Bash | Invoke the Windows side executable through interop. A Linux `pwsh` cannot drive Windows Office COM. |
| Neither `pwsh.exe` nor `powershell.exe` | Stop and prompt the user to install PowerShell 7. Do not attempt a workaround. |

Git Bash and CMD are terminal hosts, not runtimes. They do not replace PowerShell; they
launch it. The MCP server applies the same order automatically: it honours
`ENDNOTE_WORD_POWERSHELL` when set, otherwise it takes the first of `pwsh.exe` and
`powershell.exe` that resolves on `PATH`, and it raises an error naming the install link
above when neither exists. It resolves again on every call, and it also finds PowerShell 7
at its per-user app execution alias, so a PowerShell 7 installed while the server runs
takes over from the next call.

The shipped MCP launch command is `powershell.exe`, which every supported Windows has, so the
server starts on a machine without PowerShell 7; the bridge still runs under `pwsh.exe`
whenever it exists.

When the bridge runs under Windows PowerShell 5.1, no `pwsh.exe` exists, and
`ENDNOTE_WORD_POWERSHELL` is not set, it installs PowerShell 7 for the current user:

1. The read-only actions (`probe`, `list-styles`, `validate`) never install anything. They
   return a `pending` notice, on every such call, saying that the next call that is not
   read-only will install it. A session that starts with `probe`, as the workflow does,
   therefore announces it before anything changes.
2. The first other action starts `winget install --id Microsoft.PowerShell --scope user
   --installer-type msix --no-upgrade --silent`, accepting the winget source and package
   agreements on the user's behalf. winget runs detached, so that call returns as usual under
   5.1, and later calls switch to PowerShell 7 once it is installed.
3. Each state change (pending, installing, succeeded, failed, winget unavailable) is reported
   as `powershell_notice` in the JSON result, and on stderr after a `POWERSHELL_NOTICE=`
   marker so the MCP server can still return it when the call times out. Relay it to the user.

The attempt is recorded in `%LOCALAPPDATA%\EndNoteForWordAgent\pwsh-install.json`, with
winget's log beside it, and a finished or failed attempt is never retried on its own. Delete
that file to allow one more attempt. To opt out, set `ENDNOTE_WORD_AUTO_INSTALL_PWSH=0` in the
environment of the MCP server entry (or of the shell that runs the bridge). Remove PowerShell 7
again with `winget uninstall --id Microsoft.PowerShell`. If the bridge misbehaves under the new
host, set `ENDNOTE_WORD_POWERSHELL=powershell.exe` to return to Windows PowerShell 5.1.

Console encoding on a Traditional Chinese Windows installation defaults to cp950, which
mangles non-ASCII output. Every entry point in this project forces UTF-8 on its own output,
guarded in a try/catch because 5.1 can throw when stdout is redirected. If the text you are
reading in a terminal still looks wrong, the terminal itself is the problem, not the bridge.
See `troubleshooting.md` for the per-terminal fix.

## Concrete command lines per terminal

Every example runs the read-only probe. Substitute a different `-Action` once the probe
reports ready.

**PowerShell 7**

```powershell
pwsh.exe -NoLogo -NoProfile -ExecutionPolicy Bypass -File "<REPO_ROOT>\scripts\endnote-word.ps1" -Action probe
pwsh.exe -NoLogo -NoProfile -ExecutionPolicy Bypass -File "<REPO_ROOT>\scripts\endnote-word.ps1" -Action probe -LibraryPath "<LIBRARY>.enl"
pwsh.exe -NoLogo -NoProfile -ExecutionPolicy Bypass -File "<REPO_ROOT>\scripts\endnote-word.ps1" -Action list-styles -Style "IEEE"
```

**Windows PowerShell 5.1**

```powershell
powershell.exe -NoLogo -NoProfile -ExecutionPolicy Bypass -File "<REPO_ROOT>\scripts\endnote-word.ps1" -Action probe
```

**Bash or Git Bash**

```bash
REPO_ROOT="/d/path/to/endnote-for-word"
pwsh.exe -NoLogo -NoProfile -ExecutionPolicy Bypass -File "$REPO_ROOT/scripts/endnote-word.ps1" -Action probe \
  || powershell.exe -NoLogo -NoProfile -ExecutionPolicy Bypass -File "$REPO_ROOT/scripts/endnote-word.ps1" -Action probe
```

Git Bash accepts forward slash paths and converts them for the Windows executable. Quote
every path that contains a space.

**CMD**

```bat
set REPO_ROOT=D:\path\to\endnote-for-word
chcp 65001 > nul
pwsh.exe -NoLogo -NoProfile -ExecutionPolicy Bypass -File "%REPO_ROOT%\scripts\endnote-word.ps1" -Action probe
```

If `pwsh.exe` is not present, replace it with `powershell.exe` in the same line. `chcp 65001`
makes CMD render the UTF-8 the bridge emits.

The bridge writes exactly one line of compact JSON to stdout as its last stdout line, and
sends all logging to stderr. Capture stdout when you want the result, and read stderr when
you want to know what happened. Exit codes are `0` for success, `1` for an unhandled error,
`2` for a probe that reports not ready, and `3` when validation fails or an RIS import
produces no observable increase in live references.

## WSL

WSL can drive this workflow only by calling back into Windows. The Word and EndNote COM
servers live in the Windows session; a Linux build of PowerShell in the WSL filesystem
cannot reach them.

```bash
REPO_ROOT_WIN="$(wslpath -w "$PWD")"
pwsh.exe -NoLogo -NoProfile -ExecutionPolicy Bypass -File "${REPO_ROOT_WIN}\\scripts\\endnote-word.ps1" -Action probe
```

If `pwsh.exe` does not resolve, Windows interop is either disabled or the executable is not
on the Windows `PATH`. Call it by full path, for example
`"/mnt/c/Program Files/PowerShell/7/pwsh.exe"`, and convert every repository path with
`wslpath -w` before passing it to the bridge. Paths under `/home` inside the WSL filesystem
are not reliably reachable from Word, so keep the documents and the library on a Windows
drive.

## Claude Code

Copy the skill directory:

```console
%USERPROFILE%\.claude\skills\endnote-for-word
```

Register and inspect the MCP server:

```console
claude mcp add --scope user endnote-for-word -- powershell.exe -NoLogo -NoProfile -NonInteractive -ExecutionPolicy Bypass -File "<REPO_ROOT>\scripts\run-mcp.ps1"
claude mcp get endnote-for-word
```

The repository also ships `.claude-plugin/plugin.json` and `.mcp.json` for a managed plugin
deployment. A local folder is not a marketplace entry, so the portable skill directory plus
an explicit `claude mcp add` remains the deterministic path. Restart Claude Code after
installing or changing a skill.

## Claude Desktop

Claude Desktop needs both layers, the skill and the MCP server.

1. Upload the skill through **Customize > Skills**. That uploader takes a ZIP of the skill
   folder itself, meaning the folder that contains `SKILL.md`. Build it with:

   ```console
   python scripts\package_plugin.py
   ```

   and upload `dist\endnote-for-word.skill`, which holds only `skills\endnote-for-word`. The
   same command writes `dist\endnote-for-word-plugin.zip`, the whole repository for plugin
   aware hosts, and a `.sha256` sidecar for each archive. See
   <https://support.claude.com/en/articles/12512180-use-skills-in-claude>.
2. Merge the server entry below into `%APPDATA%\Claude\claude_desktop_config.json`. Merge,
   do not overwrite: that file usually already contains other servers.

```json
{
  "mcpServers": {
    "endnote-for-word": {
      "command": "powershell.exe",
      "args": [
        "-NoLogo",
        "-NoProfile",
        "-NonInteractive",
        "-ExecutionPolicy",
        "Bypass",
        "-File",
        "<REPO_ROOT>\\scripts\\run-mcp.ps1"
      ]
    }
  }
}
```

The bridge still runs under PowerShell 7 whenever it is installed. Restart Claude Desktop,
call `probe_endnote_word`, and run a fixture render before touching a real manuscript.

## Codex

Copy the skill directory to either supported skill root:

```console
%USERPROFILE%\.codex\skills\endnote-for-word
%USERPROFILE%\.agents\skills\endnote-for-word
```

Register and inspect the MCP server:

```console
codex mcp add endnote-for-word -- powershell.exe -NoLogo -NoProfile -NonInteractive -ExecutionPolicy Bypass -File "<REPO_ROOT>\scripts\run-mcp.ps1"
codex mcp get endnote-for-word
```

The repository ships `.codex-plugin/plugin.json` and `.codex-mcp.json` for plugin aware
deployment workflows. Restart Codex after changing installed skills or plugins.

## ChatGPT

ChatGPT cannot connect directly to this repository's local stdio command. It needs a remote
MCP endpoint or OpenAI's Secure MCP Tunnel, with Developer mode or Apps enabled as far as
the account or workspace permits. See
<https://help.openai.com/en/articles/12584461-developer-mode-and-full-mcp-connectors-in-chatgpt-beta>.

Two consequences follow, and neither of them can be worked around by editing a config file:

- Word and EndNote COM must execute in the signed in interactive Windows desktop, so any
  remote endpoint or tunnel has to route tool execution back to that specific Windows host.
  A cloud hosted MCP server with no path back to the desktop cannot render a document.
- This repository ships the local stdio server only. Copying `.mcp.json` somewhere does not
  constitute a completed ChatGPT installation, and it must not be reported as one.

If the account has the plugin surface available, install or enable the packaged capability
through the supported ChatGPT plugin workflow described in
<https://help.openai.com/en/articles/20001256>. ChatGPT Desktop does not discover a local
repository or a Codex skill directory on its own.

## Environment variables

| Variable | Default | Effect |
|---|---|---|
| `ENDNOTE_WORD_POWERSHELL` | unset | Absolute path to the PowerShell executable the MCP server should launch. Overrides the `pwsh.exe` then `powershell.exe` search, which otherwise runs on every call. |
| `ENDNOTE_WORD_AUTO_INSTALL_PWSH` | on | Under Windows PowerShell 5.1 without `pwsh.exe`, the first call other than `probe`, `list-styles` or `validate` installs PowerShell 7 per user with winget (MSIX, no administrator rights) and reports it through `powershell_notice`; read-only calls only announce it. Set to `0` to opt out. A set `ENDNOTE_WORD_POWERSHELL` also disables it. |
| `ENDNOTE_WORD_PYTHON` | unset | Absolute path to the Python interpreter `run-mcp.ps1` uses to host the MCP server. Overrides the `python.exe`, `python3.exe`, `python`, `python3` search on `PATH`. |
| `ENDNOTE_WORD_TIMEOUT_SECONDS` | unset | Explicit MCP watchdog timeout. Without it, calls use 240 seconds; render grows to `max(240, 60 + 8 * citation_count)`, capped at 1800 seconds. On expiry the server terminates only processes the call created. |
| `ENDNOTE_AUTOMATION_TIMEOUT_MS` | `60000` | Bridge side limit for polling until the expected citation field count is reached after `ENFormatBibliography`. |
| `ENDNOTE_READY_TIMEOUT_MS` | `90000` | Bridge side limit for waiting until a started EndNote process reports a main window. |
| `ENDNOTE_LIBRARY_LOG_LEVEL` | `INFO` | Logging level of `endnote_library.py`. Set to `DEBUG` to see the SQLite statements it issues. Logs go to stderr only, never to stdout. |
| `ENDNOTE_LIVE_TEST` | unset | Set to `1` to enable the live integration test. |
| `ENDNOTE_TEST_LIBRARY` | unset | Library path used by the live integration test. Required when `ENDNOTE_LIVE_TEST=1`. |

Raise the timeouts on a slow machine or a very large document rather than editing the
scripts. Every one of these values is a named constant in the code with the environment
variable as its documented override.

## Verification after installation

1. List the skill and the MCP tool set in the host, and confirm all thirteen tools appear.
2. Call `probe_endnote_word` with no arguments. Require `ok: true`. Check in particular
   `cwyw_template_loaded` and `cwyw_macro_callable`, which are the two fields that prove the
   CWYW integration actually responds rather than merely being installed.
3. Call `probe_endnote_word` again with the real library path and confirm `library_exists`
   and `library_data_exists` are both true.
4. Call `list_endnote_styles` with a filter and confirm the style you intend to use is in
   the returned list, spelled exactly as returned.
5. Call `search_endnote_library` and confirm it returns records with plausible
   `rec_number`, `first_author_surname` and `year` values.
6. Render a small fixture document, then run `validate_endnote_word_document` and require
   `ok: true`.
7. Open the rendered DOCX in Word. Confirm the in-text citations are grey shaded fields,
   that the EndNote tab reports them, and that the reference list at the end updates when
   fields are updated. A citation that is plain text is a failed render regardless of how
   correct it looks.
