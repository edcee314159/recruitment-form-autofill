import contextlib
import io
import json
import os
from unittest.mock import patch

from tests.profile_fixtures import WorkbookTestCase
import read_profile


class PublicReaderTest(WorkbookTestCase):
    def setUp(self):
        super().setUp()
        guard = patch.object(read_profile, "DEFAULT_DATABASE", self.directory / "unused.xlsx", create=True)
        guard.start()
        self.addCleanup(guard.stop)

    def test_detects_public_and_reads_row_two_as_fields(self):
        self.public()
        self.assertTrue(callable(getattr(read_profile, "detect_schema", None)))
        self.assertEqual(read_profile.detect_schema(self.path), "public")
        result = read_profile.load_bundle(self.path, ["姓名"], [])
        self.assertEqual([r["value"] for r in result["general"]], ["测试候选人（虚构）"])
        self.assertEqual(read_profile.load_inventory(self.path)["fields"][0], "姓名")

    def test_preserves_each_language_credential_score_pair(self):
        self.public(rows=[
            ["语言能力", "CET-4", None, None, "2024-06", None, 550],
            ["语言能力", "CET-6", None, None, "2024-12", None, 520],
            ["语言能力", "IELTS", None, None, "2025-01-03", None, 6.5],
            ["项目经历", "虚构项目"],
        ])
        result = read_profile.load_bundle(self.path, [], ["语言能力"])
        self.assertEqual([(r["organization"], r["result"]) for r in result["experiences"]], [("CET-4", 550), ("CET-6", 520), ("IELTS", 6.5)])

    def test_legacy_detection_and_existing_read_contract(self):
        self.legacy()
        self.assertTrue(callable(getattr(read_profile, "detect_schema", None)))
        self.assertEqual(read_profile.detect_schema(self.path), "legacy")
        self.assertEqual(read_profile.load_general_profile(self.path)[0]["value"], "测试候选人（虚构）")
        self.assertEqual(len(read_profile.load_experience_records(self.path)), 3)

    def test_normal_cli_uses_local_config_for_all_commands(self):
        self.public(rows=[["语言能力", "IELTS", None, None, None, None, 6.5]])
        config = self.directory / "profile.json"
        config.write_text(json.dumps({"workbook_path": str(self.path)}), encoding="utf-8")
        commands = (["inventory"], ["get", "--field", "姓名"], ["experiences", "--category", "语言能力"], ["bundle", "--field", "姓名", "--category", "语言能力"])
        for command in commands:
            with self.subTest(command=command), patch.dict(os.environ, {"RECRUITMENT_PROFILE_CONFIG": str(config)}), contextlib.redirect_stdout(io.StringIO()) as output, contextlib.redirect_stderr(io.StringIO()):
                self.assertEqual(read_profile.main(command), 0)
                self.assertIsInstance(json.loads(output.getvalue()), dict)

    def test_config_failure_is_redacted(self):
        with patch.dict(os.environ, {"RECRUITMENT_PROFILE_CONFIG": str(self.directory / "PRIVATE-MARKER.json")}), contextlib.redirect_stderr(io.StringIO()) as error:
            self.assertEqual(read_profile.main(["inventory"]), 2)
        self.assertNotIn("PRIVATE-MARKER", error.getvalue())
