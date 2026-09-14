from pathlib import Path
import unittest


class RepositorySafetyTest(unittest.TestCase):
    def test_repository_has_required_files_and_no_private_artifacts(self):
        root = Path(".")
        self.assertTrue((root / ".gitignore").is_file())
        self.assertTrue((root / "LICENSE").is_file())
        self.assertFalse(list(root.rglob("秋招个人信息总表.xlsx")))
        self.assertFalse(list(root.rglob("*.docx")))
