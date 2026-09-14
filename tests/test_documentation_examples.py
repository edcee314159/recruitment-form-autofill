import json
from pathlib import Path
import subprocess
import sys
import tempfile
import unittest


ROOT = Path(__file__).resolve().parents[1]
SCRIPTS = ROOT / "skill" / "recruitment-form-autofill" / "scripts"


class DocumentationExamplesTest(unittest.TestCase):
    def test_quick_start_initializes_and_validates_a_temporary_profile(self):
        with tempfile.TemporaryDirectory() as temporary:
            folder = Path(temporary)
            workbook = folder / "profile.xlsx"
            config = folder / "profile-config.json"
            init = subprocess.run(
                [sys.executable, str(SCRIPTS / "init_profile.py"), "--workbook", str(workbook), "--config", str(config)],
                cwd=ROOT,
                capture_output=True,
                text=True,
                check=False,
            )
            self.assertEqual(init.returncode, 0, init.stderr)
            validator = subprocess.run(
                [sys.executable, str(SCRIPTS / "validate_profile.py"), "--config", str(config)],
                cwd=ROOT,
                capture_output=True,
                text=True,
                check=False,
            )
            self.assertEqual(validator.returncode, 0, validator.stderr)
            self.assertEqual(json.loads(validator.stdout), {"issues": []})
