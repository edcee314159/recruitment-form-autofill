import importlib
import json
import sys
import tempfile
import unittest
from dataclasses import FrozenInstanceError
from pathlib import Path


ROOT = Path(__file__).resolve().parents[1]
SCRIPTS = ROOT / "skill" / "recruitment-form-autofill" / "scripts"
sys.path.insert(0, str(SCRIPTS))


def profile_config_module():
    try:
        return importlib.import_module("profile_config")
    except ModuleNotFoundError as exc:
        raise AssertionError("profile_config module must be provided") from exc


class ProfileConfigTest(unittest.TestCase):
    def test_init_profile_copies_template_and_writes_resolved_config(self):
        """Catches initialization that omits either the workbook copy or usable config."""
        profile_config = profile_config_module()
        with tempfile.TemporaryDirectory() as temporary_directory:
            temporary = Path(temporary_directory)
            template = temporary / "blank-template.xlsx"
            template.write_bytes(b"synthetic blank workbook")
            workbook = temporary / "candidate" / "profile.xlsx"
            config = temporary / "settings" / "profile.json"

            profile_config.init_profile(workbook, None, config, template_path=template)

            self.assertEqual(workbook.read_bytes(), b"synthetic blank workbook")
            self.assertEqual(
                profile_config.load_config(config).workbook_path,
                workbook.resolve(),
            )

    def test_init_profile_does_not_overwrite_an_existing_workbook(self):
        """Catches initialization that replaces a candidate-maintained workbook."""
        profile_config = profile_config_module()
        with tempfile.TemporaryDirectory() as temporary_directory:
            temporary = Path(temporary_directory)
            template = temporary / "blank-template.xlsx"
            template.write_bytes(b"synthetic blank workbook")
            workbook = temporary / "candidate" / "profile.xlsx"
            workbook.parent.mkdir()
            workbook.write_bytes(b"candidate-maintained workbook")
            config = temporary / "settings" / "profile.json"

            profile_config.init_profile(workbook, None, config, template_path=template)

            self.assertEqual(workbook.read_bytes(), b"candidate-maintained workbook")

    def test_init_profile_creates_attachment_directory_only_when_requested(self):
        """Catches initialization that skips a requested attachment directory."""
        profile_config = profile_config_module()
        with tempfile.TemporaryDirectory() as temporary_directory:
            temporary = Path(temporary_directory)
            template = temporary / "blank-template.xlsx"
            template.write_bytes(b"synthetic blank workbook")
            workbook = temporary / "candidate" / "profile.xlsx"
            attachment_directory = temporary / "candidate" / "attachments"
            config = temporary / "settings" / "profile.json"

            profile_config.init_profile(
                workbook,
                attachment_directory,
                config,
                template_path=template,
            )

            loaded = profile_config.load_config(config)
            self.assertTrue(attachment_directory.is_dir())
            self.assertEqual(loaded.attachment_dir, attachment_directory.resolve())

    def test_load_config_rejects_a_missing_workbook(self):
        """Catches configuration loading that accepts a stale workbook location."""
        profile_config = profile_config_module()
        with tempfile.TemporaryDirectory() as temporary_directory:
            temporary = Path(temporary_directory)
            workbook = temporary / "profile.xlsx"
            workbook.write_bytes(b"synthetic workbook")
            config = temporary / "profile.json"
            profile_config.write_config(
                profile_config.ProfileConfig(workbook_path=workbook), config
            )
            workbook.unlink()

            with self.assertRaisesRegex(profile_config.ProfileConfigError, "workbook not found"):
                profile_config.load_config(config)

    def test_config_inside_repository_requires_explicit_test_override(self):
        """Catches accidental storage of local profile config in the public repository."""
        profile_config = profile_config_module()
        with tempfile.TemporaryDirectory() as temporary_directory:
            temporary = Path(temporary_directory)
            repository_root = temporary / "public-repository"
            repository_root.mkdir()
            template = temporary / "blank-template.xlsx"
            template.write_bytes(b"synthetic blank workbook")
            workbook = temporary / "candidate" / "profile.xlsx"
            config = repository_root / "profile.json"

            with self.assertRaises(profile_config.ProfileConfigError):
                profile_config.init_profile(
                    workbook,
                    None,
                    config,
                    template_path=template,
                    repository_root=repository_root,
                )

            profile_config.init_profile(
                workbook,
                None,
                config,
                template_path=template,
                repository_root=repository_root,
                allow_repository_config=True,
            )
            self.assertTrue(config.is_file())

    def test_profile_config_is_immutable_and_normalizes_paths(self):
        """Catches mutable or unnormalized paths that make persisted config ambiguous."""
        profile_config = profile_config_module()
        with tempfile.TemporaryDirectory() as temporary_directory:
            temporary = Path(temporary_directory)
            workbook = temporary / "candidate" / "profile.xlsx"
            attachment_directory = temporary / "candidate" / "attachments"

            config = profile_config.ProfileConfig(workbook, attachment_directory)

            self.assertEqual(config.workbook_path, workbook.resolve())
            self.assertEqual(config.attachment_dir, attachment_directory.resolve())
            with self.assertRaises(FrozenInstanceError):
                config.workbook_path = temporary / "other.xlsx"

    def test_config_is_utf8_json(self):
        """Catches config writing that cannot safely preserve non-ASCII local paths."""
        profile_config = profile_config_module()
        with tempfile.TemporaryDirectory() as temporary_directory:
            temporary = Path(temporary_directory)
            workbook = temporary / "候选人资料.xlsx"
            workbook.write_bytes(b"synthetic workbook")
            config = temporary / "配置.json"

            profile_config.write_config(
                profile_config.ProfileConfig(workbook_path=workbook), config
            )

            stored = config.read_text(encoding="utf-8")
            self.assertIn("候选人资料.xlsx", stored)
            self.assertEqual(
                json.loads(stored)["workbook_path"], str(workbook.resolve())
            )


if __name__ == "__main__":
    unittest.main()
