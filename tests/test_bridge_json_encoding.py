"""Offline regression tests for how the bridge decodes its JSON inputs.

Windows PowerShell 5.1 decodes Get-Content with the ANSI code page unless an encoding is
given, so a BOM-less UTF-8 manifest with a non-ASCII author either failed to parse or was
silently corrupted. These tests drive the real bridge with -Action render. The manifest names
a library that does not exist, so the run stops inside Assert-Manifest, before Word or
EndNote can start, and the error echoes the decoded library path back verbatim.
"""

from __future__ import annotations

import json
import os
import shutil
import subprocess
import tempfile
import unittest
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
BRIDGE = ROOT / "scripts" / "endnote-word.ps1"
FIXTURE_MANIFEST = ROOT / "tests" / "fixtures" / "single-manifest.json"

POWERSHELL_HOSTS = ("powershell.exe", "pwsh.exe")
BRIDGE_TIMEOUT_SECONDS = 120
NON_ASCII_LIBRARY_DIRECTORY = "文獻 Müller 張偉"
LATIN1_LIBRARY_DIRECTORY = "Müller"
UTF8_BOM = b"\xef\xbb\xbf"
# The bridge reads %APPDATA%\EndNoteForWordAgent\defaults.json when it exists.
USER_DEFAULTS_RELATIVE_PATH = Path("EndNoteForWordAgent") / "defaults.json"


def available_hosts() -> list[str]:
    return [path for path in (shutil.which(name) for name in POWERSHELL_HOSTS) if path]


class BridgeJsonEncodingTests(unittest.TestCase):
    @classmethod
    def setUpClass(cls) -> None:
        cls.hosts = available_hosts()
        if not cls.hosts:
            raise unittest.SkipTest("No Windows PowerShell or PowerShell 7 host on PATH.")
        cls._directory = tempfile.TemporaryDirectory(prefix="endnote-encoding-")
        cls.workspace = Path(cls._directory.name)
        cls.input_document = cls.workspace / "input.docx"
        cls.input_document.write_bytes(b"")

    @classmethod
    def tearDownClass(cls) -> None:
        cls._directory.cleanup()

    def manifest_document(self, library_directory: str) -> tuple[dict, str]:
        manifest = json.loads(FIXTURE_MANIFEST.read_text(encoding="utf-8"))
        library_path = str(self.workspace / library_directory / "Library.enl")
        manifest["library"]["path"] = library_path
        return manifest, library_path

    def write_bytes(self, name: str, payload: bytes) -> Path:
        path = self.workspace / name
        path.parent.mkdir(parents=True, exist_ok=True)
        path.write_bytes(payload)
        return path

    def run_render(self, host: str, manifest_path: Path, environment: dict | None = None) -> dict:
        # Never let a test start the real PowerShell 7 auto-install on a machine without pwsh.
        environment = {**(environment or os.environ), "ENDNOTE_WORD_AUTO_INSTALL_PWSH": "0"}
        completed = subprocess.run(
            [
                host,
                "-NoLogo",
                "-NoProfile",
                "-NonInteractive",
                "-ExecutionPolicy",
                "Bypass",
                "-File",
                str(BRIDGE),
                "-Action",
                "render",
                "-InputPath",
                str(self.input_document),
                "-OutputPath",
                str(self.workspace / "output" / "output.docx"),
                "-ManifestPath",
                str(manifest_path),
            ],
            stdin=subprocess.DEVNULL,
            capture_output=True,
            text=True,
            encoding="utf-8",
            errors="replace",
            timeout=BRIDGE_TIMEOUT_SECONDS,
            check=False,
            env=environment,
        )
        lines = [line for line in completed.stdout.splitlines() if line.strip()]
        self.assertTrue(lines, f"The bridge printed no JSON result. stderr:\n{completed.stderr}")
        self.assertEqual(completed.returncode, 1, completed.stdout)
        result = json.loads(lines[-1])
        self.assertFalse(result["ok"])
        return result

    def test_utf8_manifests_round_trip_non_ascii_text(self) -> None:
        manifest, library_path = self.manifest_document(NON_ASCII_LIBRARY_DIRECTORY)
        text = json.dumps(manifest, ensure_ascii=False)
        for encoding in ("utf-8", "utf-8-sig"):
            manifest_path = self.write_bytes(f"manifest-{encoding}.json", text.encode(encoding))
            for host in self.hosts:
                with self.subTest(encoding=encoding, host=Path(host).name):
                    result = self.run_render(host, manifest_path)
                    self.assertEqual(result["error"], f"library.path does not exist: {library_path}")

    def test_manifests_that_are_not_utf8_are_rejected_not_corrupted(self) -> None:
        manifest, _ = self.manifest_document(LATIN1_LIBRARY_DIRECTORY)
        text = json.dumps(manifest, ensure_ascii=False)
        payloads = {
            "latin-1": text.encode("latin-1"),
            # A UTF-8 BOM must not switch the reader to a lenient decoder.
            "bom-latin-1": UTF8_BOM + text.encode("latin-1"),
            # The precheck rejects UTF-16, so the bridge rejects it too.
            "utf-16": text.encode("utf-16"),
        }
        for name, payload in payloads.items():
            manifest_path = self.write_bytes(f"manifest-{name}.json", payload)
            for host in self.hosts:
                with self.subTest(variant=name, host=Path(host).name):
                    result = self.run_render(host, manifest_path)
                    self.assertTrue(
                        result["error"].startswith("ManifestPath could not be read as UTF-8 text"),
                        result["error"],
                    )

    def test_agent_defaults_that_are_not_utf8_are_rejected(self) -> None:
        # A manifest without a style makes Assert-Manifest read the user defaults file.
        manifest, _ = self.manifest_document(LATIN1_LIBRARY_DIRECTORY)
        del manifest["style"]
        manifest_path = self.write_bytes(
            "manifest-without-style.json", json.dumps(manifest).encode("utf-8")
        )
        application_data = self.workspace / "appdata"
        self.write_bytes(
            str(Path("appdata") / USER_DEFAULTS_RELATIVE_PATH),
            json.dumps({"style": {"default_output_style": "Müller"}}, ensure_ascii=False).encode("latin-1"),
        )
        environment = {**os.environ, "APPDATA": str(application_data)}
        for host in self.hosts:
            with self.subTest(host=Path(host).name):
                result = self.run_render(host, manifest_path, environment)
                self.assertTrue(
                    result["error"].startswith("The agent defaults file could not be read as UTF-8 text"),
                    result["error"],
                )


if __name__ == "__main__":
    unittest.main(verbosity=2)
