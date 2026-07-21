#!/usr/bin/env python3
"""Create a reproducible ZIP of the complete cross-agent endnote-for-word plugin."""

from __future__ import annotations

import hashlib
import os
import stat
import sys
import zipfile
from pathlib import Path


# Build artefacts, caches and the scratch workspace the skill uses for intermediate DOCX
# files never belong in a distributable archive.
EXCLUDED_DIRS = {
    ".codex-test-temp",
    ".git",
    "__pycache__",
    ".pytest_cache",
    "dist",
    "endnote-for-word-workspace",
}
EXCLUDED_SUFFIXES = {".pyc", ".pyo"}
EXCLUDED_PREFIX_PATHS = (("evals", "results"),)
ARCHIVE_TIMESTAMP = (1980, 1, 1, 0, 0, 0)
ARCHIVE_FILE_MODE = (stat.S_IFREG | 0o644) << 16

PLUGIN_NAME = "endnote-for-word"
DEFAULT_ARCHIVE_NAME = f"{PLUGIN_NAME}-plugin.zip"


def should_include(path: Path, root: Path) -> bool:
    relative = path.relative_to(root)
    if any(part in EXCLUDED_DIRS for part in relative.parts) or path.suffix in EXCLUDED_SUFFIXES:
        return False
    return all(relative.parts[: len(prefix)] != prefix for prefix in EXCLUDED_PREFIX_PATHS)


def package(root: Path, output: Path) -> str:
    root = root.resolve()
    output = output.resolve()
    sidecar = output.with_suffix(output.suffix + ".sha256")
    generated_outputs = {output, sidecar}
    required = [
        root / ".codex-plugin" / "plugin.json",
        root / ".claude-plugin" / "plugin.json",
        root / ".mcp.json",
        root / ".codex-mcp.json",
        root / "scripts" / "endnote-word.ps1",
        root / "scripts" / "run-mcp.ps1",
        root / "scripts" / "endnote_library.py",
        root / "scripts" / "mcp_server.py",
        root / "skills" / PLUGIN_NAME / "SKILL.md",
    ]
    missing = [str(path) for path in required if not path.is_file()]
    if missing:
        raise FileNotFoundError("Missing required plugin files: " + ", ".join(missing))
    output.parent.mkdir(parents=True, exist_ok=True)
    with zipfile.ZipFile(output, "w", compression=zipfile.ZIP_DEFLATED, compresslevel=9) as archive:
        for directory, directories, filenames in os.walk(root):
            parent = Path(directory)
            directories[:] = [
                name for name in sorted(directories) if should_include(parent / name, root)
            ]
            for name in sorted(filenames):
                path = parent / name
                if not should_include(path, root) or path.resolve() in generated_outputs:
                    continue
                info = zipfile.ZipInfo(
                    (Path(PLUGIN_NAME) / path.relative_to(root)).as_posix(),
                    date_time=ARCHIVE_TIMESTAMP,
                )
                info.create_system = 3
                info.external_attr = ARCHIVE_FILE_MODE
                archive.writestr(
                    info,
                    path.read_bytes(),
                    compress_type=zipfile.ZIP_DEFLATED,
                    compresslevel=9,
                )
    digest = hashlib.sha256(output.read_bytes()).hexdigest()
    sidecar.write_text(f"{digest}  {output.name}\n", encoding="ascii")
    return digest


def main() -> int:
    root = Path(sys.argv[1]) if len(sys.argv) > 1 else Path(__file__).resolve().parents[1]
    output = Path(sys.argv[2]) if len(sys.argv) > 2 else root / "dist" / DEFAULT_ARCHIVE_NAME
    try:
        digest = package(root, output)
    except (FileNotFoundError, OSError, zipfile.BadZipFile) as exc:
        print(f"Packaging failed: {exc}", file=sys.stderr)
        return 1
    print(f"Created: {output.resolve()}")
    print(f"SHA256: {digest}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
