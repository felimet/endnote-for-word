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
DEFAULT_SKILL_ARCHIVE_NAME = f"{PLUGIN_NAME}.skill"


def should_include(path: Path, root: Path) -> bool:
    relative = path.relative_to(root)
    if any(part in EXCLUDED_DIRS for part in relative.parts) or path.suffix in EXCLUDED_SUFFIXES:
        return False
    return all(relative.parts[: len(prefix)] != prefix for prefix in EXCLUDED_PREFIX_PATHS)


def package(root: Path, output: Path) -> str:
    """Archive the whole plugin checkout under a single endnote-for-word/ folder."""
    root = root.resolve()
    required = [
        root / ".codex-plugin" / "plugin.json",
        root / ".claude-plugin" / "plugin.json",
        root / ".mcp.json",
        root / ".codex-mcp.json",
        root / "scripts" / "endnote-word.ps1",
        root / "scripts" / "endnote_cleanup.ps1",
        root / "scripts" / "run-mcp.ps1",
        root / "scripts" / "endnote_library.py",
        root / "scripts" / "mcp_server.py",
        root / "skills" / PLUGIN_NAME / "SKILL.md",
    ]
    return _archive_tree(root, output, required, "plugin")


def package_skill(root: Path, output: Path) -> str:
    """Archive only the portable skill folder, the layout a .skill upload expects."""
    source = root.resolve() / "skills" / PLUGIN_NAME
    return _archive_tree(source, output, [source / "SKILL.md"], "skill")


def _archive_tree(source: Path, output: Path, required: list[Path], kind: str) -> str:
    output = output.resolve()
    sidecar = output.with_suffix(output.suffix + ".sha256")
    generated_outputs = {output, sidecar}
    missing = [str(path) for path in required if not path.is_file()]
    if missing:
        raise FileNotFoundError(f"Missing required {kind} files: " + ", ".join(missing))
    output.parent.mkdir(parents=True, exist_ok=True)
    with zipfile.ZipFile(output, "w", compression=zipfile.ZIP_DEFLATED, compresslevel=9) as archive:
        for directory, directories, filenames in os.walk(source):
            parent = Path(directory)
            directories[:] = [
                name for name in sorted(directories) if should_include(parent / name, source)
            ]
            for name in sorted(filenames):
                path = parent / name
                if not should_include(path, source) or path.resolve() in generated_outputs:
                    continue
                info = zipfile.ZipInfo(
                    (Path(PLUGIN_NAME) / path.relative_to(source)).as_posix(),
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
    skill_output = output.parent / DEFAULT_SKILL_ARCHIVE_NAME
    try:
        archives = ((output, package(root, output)), (skill_output, package_skill(root, skill_output)))
    except (FileNotFoundError, OSError, zipfile.BadZipFile) as exc:
        print(f"Packaging failed: {exc}", file=sys.stderr)
        return 1
    for path, digest in archives:
        print(f"Created: {path.resolve()}")
        print(f"SHA256: {digest}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
