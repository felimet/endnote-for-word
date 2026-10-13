"""Offline tests for the EndNote session helpers the MCP job worker uses.

Window enumeration needs a desktop, so only the decisions built on it are tested here: which
dialog counts as a dismissable news notice, which window belongs to a library, and how the
recent-library list is restored. The restore test writes to its own HKCU test key, never to
EndNote's.
"""

from __future__ import annotations

import os
import sys
import unittest
from pathlib import Path

SCRIPTS = Path(__file__).resolve().parents[1] / "scripts"
if str(SCRIPTS) not in sys.path:
    sys.path.insert(0, str(SCRIPTS))

TEST_REGISTRY_KEY = r"Software\EndNoteForWordAgentTests\Recent Libraries"


@unittest.skipUnless(os.name == "nt", "endnote_session uses user32 and winreg.")
class EndNoteSessionTests(unittest.TestCase):
    @classmethod
    def setUpClass(cls) -> None:
        import endnote_session

        cls.session = endnote_session

    def test_only_the_three_button_endnote_notice_is_dismissable(self) -> None:
        buttons = {"Learn More...", "Cancel", "Ignore This"}
        self.assertTrue(self.session.is_notice("EndNote 2025 is here!", buttons))
        self.assertTrue(self.session.is_notice("EndNote Password Expiration Policy Change", buttons))
        # The CWYW record chooser is no notice; the job worker cancels it on its own path.
        self.assertFalse(self.session.is_notice(
            "EndNote 21 Select Matching Reference", {"&Find", "&Insert", "Ign&ore", "Ignore &All", "Cancel", "&Help"}
        ))
        self.assertFalse(self.session.is_notice("Microsoft Word", buttons))
        self.assertFalse(self.session.is_notice("EndNote 2025 is here!", {"OK"}))

    def test_library_windows_are_matched_by_file_name(self) -> None:
        library = Path(r"C:\Research\Draft Paper.enl")
        self.assertTrue(self.session.library_title_matches("EndNote 21 - Draft Paper.enl", library))
        self.assertTrue(self.session.library_title_matches("EndNote 21 - draft paper.ENL", library))
        self.assertTrue(self.session.library_title_matches("Draft Paper.enl", library))  # after a CWYW job
        self.assertFalse(self.session.library_title_matches("EndNote 21 - Other Draft Paper.enl", library))
        self.assertFalse(self.session.library_title_matches("EndNote 21", library))

    def test_recent_library_list_is_restored_exactly(self) -> None:
        import winreg

        original_key = self.session.RECENT_LIBRARIES_KEY
        self.session.RECENT_LIBRARIES_KEY = TEST_REGISTRY_KEY
        self.addCleanup(setattr, self.session, "RECENT_LIBRARIES_KEY", original_key)
        self.addCleanup(self.delete_test_key)
        snapshot = {"File1": r"C:\Research\Mine.enl", "File2": r"D:\Shared\Group.enl"}
        self.session.restore_recent_libraries(snapshot)
        self.assertEqual(self.session.recent_libraries(), snapshot)
        # A job's library pushed to the front and an extra entry appended, as EndNote does.
        with winreg.CreateKey(winreg.HKEY_CURRENT_USER, TEST_REGISTRY_KEY) as key:
            winreg.SetValueEx(key, "File1", 0, winreg.REG_SZ, r"C:\Jobs\Draft.enl")
            winreg.SetValueEx(key, "File2", 0, winreg.REG_SZ, r"C:\Research\Mine.enl")
            winreg.SetValueEx(key, "File3", 0, winreg.REG_SZ, r"D:\Shared\Group.enl")
        self.session.restore_recent_libraries(snapshot)
        self.assertEqual(self.session.recent_libraries(), snapshot)

    @staticmethod
    def delete_test_key() -> None:
        import winreg

        for path in (TEST_REGISTRY_KEY, TEST_REGISTRY_KEY.rsplit("\\", 1)[0]):
            try:
                winreg.DeleteKey(winreg.HKEY_CURRENT_USER, path)
            except FileNotFoundError:
                pass


if __name__ == "__main__":
    unittest.main(verbosity=2)
