"""Exercise a skill-only installation without relying on repository resources."""

import json
from pathlib import Path
import shutil
import subprocess
import sys
import tempfile
import unittest

ROOT = Path(__file__).resolve().parents[1]


class InstalledSkillTest(unittest.TestCase):
    def test_skill_only_install_initializes_validates_and_reads_profile(self):
        with tempfile.TemporaryDirectory() as temporary:
            folder = Path(temporary)
            skill = folder / ".codex" / "skills" / "recruitment-form-autofill"
            shutil.copytree(ROOT / "skill/recruitment-form-autofill", skill, ignore=shutil.ignore_patterns("__pycache__"))
            scripts = skill / "scripts"
            workbook = folder / "user-data" / "profile.xlsx"
            config = folder / "user-data" / "profile-config.json"
            commands = [
                ("init_profile.py", "--workbook", str(workbook), "--config", str(config)),
                ("validate_profile.py", "--config", str(config)),
                ("read_profile.py", "--config", str(config), "inventory"),
            ]
            for command in commands:
                result = subprocess.run([sys.executable, "-B", str(scripts / command[0]), *command[1:]], cwd=folder, capture_output=True, text=True, encoding="utf-8")
                self.assertEqual(result.returncode, 0, result.stderr)
            self.assertIn("姓名", json.loads(result.stdout)["fields"])
            self.assertTrue(workbook.is_file())
