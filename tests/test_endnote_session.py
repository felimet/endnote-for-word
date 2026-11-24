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
TEST_RECOVERY_KEY = r"Software\EndNoteForWordAgentTests\DocumentRecovery"


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

    def test_only_listed_marketing_notices_with_the_measured_buttons_are_notices(self) -> None:
        notice = self.session.is_notice
        measured = "Holiday Sale: Save 25% when you upgrade to the new EndNote!"
        buttons = {"Learn More...", "Cancel", "Ignore This"}
        self.assertTrue(notice(measured, buttons))
        self.assertTrue(notice("Spring offer: upgrade to the new EndNote today", buttons))
        self.assertTrue(notice("HOLIDAY SALE", buttons))
        # A listed title with any other button set stays with attention.
        self.assertFalse(notice("Spring offer: upgrade to the new EndNote today", {"Learn More...", "Cancel"}))
        self.assertFalse(notice("HOLIDAY SALE", {"Learn More...", "Close"}))
        self.assertFalse(notice(measured, {"Learn More...", "Ignore This"}))
        self.assertFalse(notice(measured, {"Buy Now", "Sign In"}))
        # Unlisted dialogs stay with attention, whatever their buttons.
        self.assertFalse(notice("Sign in to EndNote", {"Sign In", "Cancel"}))
        self.assertFalse(notice("Microsoft Word", buttons))

    def test_marketing_notice_is_dismissed_and_other_dialogs_stay_reported(self) -> None:
        window = self.session.Window
        windows = [
            window(1, 7, "#32770", "Holiday Sale: Save 25% when you upgrade to the new EndNote!", True, False),
            window(2, 7, "#32770", "EndNote 2025 is here!", True, False),
            window(3, 7, "#32770", "EndNote 21 Select Matching Reference", True, False),
            # A listed title without the measured button set is not cancelled.
            window(4, 7, "#32770", "Spring offer: upgrade to the new EndNote today", True, False),
        ]
        buttons = {
            1: [(1, "Learn More...", 11), (2, "Cancel", 12), (3, "Ignore This", 13)],
            2: [(1, "Learn More...", 21), (2, "Cancel", 22), (3, "Ignore This", 23)],
            3: [(1, "&Find", 31), (2, "Cancel", 32)],
            4: [(1, "Learn More...", 41), (2, "Cancel", 42)],
        }
        clicked: list[int] = []

        class User32:
            @staticmethod
            def SendMessageW(handle, message, wparam, lparam):
                clicked.append(handle)

        for name, value in (("windows_of", lambda pids: windows), ("buttons_of", lambda handle: buttons[handle]),
                            ("_user32", User32)):
            self.addCleanup(setattr, self.session, name, getattr(self.session, name))
            setattr(self.session, name, value)
        self.assertEqual(
            self.session.dismiss_notices({7}),
            ["Holiday Sale: Save 25% when you upgrade to the new EndNote!", "EndNote 2025 is here!"],
        )
        self.assertEqual(clicked, [12, 22])
        self.assertEqual(
            [dialog["title"] for dialog in self.session.blocking_dialogs({7})],
            ["EndNote 21 Select Matching Reference", "Spring offer: upgrade to the new EndNote today"],
        )

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

    def test_only_new_recovery_entries_of_the_job_documents_are_removed(self) -> None:
        import winreg

        original_key = self.session.WORD_RECOVERY_KEY
        self.session.WORD_RECOVERY_KEY = TEST_RECOVERY_KEY
        self.addCleanup(setattr, self.session, "WORD_RECOVERY_KEY", original_key)
        self.addCleanup(self.delete_test_key, TEST_RECOVERY_KEY)
        work = r"C:\Jobs\.draft.0123456789abcdef0123456789abcdef.work.docx"
        saved = "D:\\論文\\.草稿.fedcba9876543210fedcba9876543210.tmp.docx"
        entries = {
            "177B8E80": work,                      # the job's working copy: removed
            "21C4A000": saved.upper(),             # the job's SaveAs2 target, other case: removed
            "0A0B0C0D": r"C:\Research\Mine.docx",  # the user's own document: kept
            "5E5E5E5E": work,                      # same path, but present before the job: kept
        }
        for name, path in entries.items():
            # The layout measured on Word 16.0: three 32-bit fields, then the UTF-16 path.
            blob = bytes.fromhex("04000000bc430000d8000000") + path.encode("utf-16-le") + b"\x00\x00"
            with winreg.CreateKey(winreg.HKEY_CURRENT_USER, f"{TEST_RECOVERY_KEY}\\{name}") as key:
                winreg.SetValueEx(key, name, 0, winreg.REG_BINARY, blob)
        removed = self.session.clear_document_recovery({work, saved}, keep={"5E5E5E5E"})
        self.assertEqual(sorted(removed), ["177B8E80", "21C4A000"])
        self.assertEqual(self.session.document_recovery_entries(), {"0A0B0C0D", "5E5E5E5E"})
        self.assertEqual(self.session.clear_document_recovery(set(), keep=set()), [])

    @staticmethod
    def delete_test_key(key: str = TEST_REGISTRY_KEY) -> None:
        import winreg

        parent = key.rsplit("\\", 1)[0]
        try:
            with winreg.OpenKey(winreg.HKEY_CURRENT_USER, key) as handle:
                names = [winreg.EnumKey(handle, index) for index in range(winreg.QueryInfoKey(handle)[0])]
            for name in names:
                winreg.DeleteKey(winreg.HKEY_CURRENT_USER, f"{key}\\{name}")
        except FileNotFoundError:
            pass
        for path in (key, parent):
            try:
                winreg.DeleteKey(winreg.HKEY_CURRENT_USER, path)
            except FileNotFoundError:
                pass
            except PermissionError:
                break  # the parent still holds the other test's key


if __name__ == "__main__":
    unittest.main(verbosity=2)
