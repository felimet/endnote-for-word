"""Regression checks for deterministic plugin packaging."""

from __future__ import annotations

import os
import sys
import tempfile
import unittest
import zipfile
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
SCRIPTS = ROOT / "scripts"
if str(SCRIPTS) not in sys.path:
    sys.path.insert(0, str(SCRIPTS))

import package_plugin


REQUIRED_FILES = (
    ".codex-plugin/plugin.json",
    ".claude-plugin/plugin.json",
    ".mcp.json",
    ".codex-mcp.json",
    "scripts/endnote-word.ps1",
    "scripts/endnote_cleanup.ps1",
    "scripts/run-mcp.ps1",
    "scripts/endnote_library.py",
    "scripts/mcp_server.py",
    "skills/endnote-for-word/SKILL.md",
)


def build_checkout(root: Path, timestamp: int) -> None:
    for relative in (*REQUIRED_FILES, "README.md", ".git/private", ".codex-test-temp/junk"):
        path = root / relative
        path.parent.mkdir(parents=True, exist_ok=True)
        path.write_text(relative, encoding="utf-8")
        os.utime(path, (timestamp, timestamp))


class PackagePluginTests(unittest.TestCase):
    def test_archive_is_independent_of_checkout_path_and_mtime(self) -> None:
        with tempfile.TemporaryDirectory(prefix="endnote-package-") as directory:
            workspace = Path(directory)
            first_root = workspace / "first-checkout"
            second_root = workspace / "renamed-checkout"
            build_checkout(first_root, 1_700_000_000)
            build_checkout(second_root, 1_800_000_000)

            first_archive = workspace / "first.zip"
            second_archive = workspace / "second.zip"
            first_digest = package_plugin.package(first_root, first_archive)
            second_digest = package_plugin.package(second_root, second_archive)

            self.assertEqual(first_digest, second_digest)
            nested_archive = first_root / "custom-output.zip"
            self.assertEqual(
                package_plugin.package(first_root, nested_archive),
                package_plugin.package(first_root, nested_archive),
            )
            with zipfile.ZipFile(first_archive) as archive:
                names = archive.namelist()
                self.assertTrue(names)
                self.assertTrue(all(name.startswith("endnote-for-word/") for name in names))
                self.assertFalse(any(".git" in name for name in names))
                self.assertFalse(any(".codex-test-temp" in name for name in names))
                self.assertIsNone(archive.testzip())

    def test_skill_archive_holds_only_the_skill_folder(self) -> None:
        with tempfile.TemporaryDirectory(prefix="endnote-package-") as directory:
            workspace = Path(directory)
            root = workspace / "checkout"
            build_checkout(root, 1_700_000_000)
            reference = root / "skills" / "endnote-for-word" / "references" / "notes.md"
            reference.parent.mkdir(parents=True)
            reference.write_text("notes", encoding="utf-8")

            output = workspace / "endnote-for-word.skill"
            digest = package_plugin.package_skill(root, output)

            self.assertEqual(digest, package_plugin.package_skill(root, output))
            sidecar = output.with_name(output.name + ".sha256")
            self.assertEqual(sidecar.read_text(encoding="ascii"), f"{digest}  {output.name}\n")
            with zipfile.ZipFile(output) as archive:
                self.assertEqual(
                    sorted(archive.namelist()),
                    ["endnote-for-word/SKILL.md", "endnote-for-word/references/notes.md"],
                )

    def test_skill_archive_requires_skill_md(self) -> None:
        with tempfile.TemporaryDirectory(prefix="endnote-package-") as directory:
            root = Path(directory) / "checkout"
            (root / "skills" / "endnote-for-word").mkdir(parents=True)
            with self.assertRaises(FileNotFoundError):
                package_plugin.package_skill(root, Path(directory) / "out.skill")


if __name__ == "__main__":
    unittest.main()
