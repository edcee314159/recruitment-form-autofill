import importlib
import importlib.util
from openpyxl import load_workbook

from tests.profile_fixtures import WorkbookTestCase, ROOT
import read_profile


class MigrationTest(WorkbookTestCase):
    def migration(self):
        self.assertIsNotNone(importlib.util.find_spec("migrate_profile"), "migration must exist")
        return importlib.import_module("migrate_profile")

    def test_migration_preserves_source_and_copies_only_public_fields(self):
        self.legacy()
        original = self.path.read_bytes()
        destination = self.directory / "public.xlsx"
        warnings = self.migration().migrate_legacy(self.path, destination)
        self.assertEqual(self.path.read_bytes(), original)
        self.assertEqual(read_profile.detect_schema(destination), "public")
        general = read_profile.load_bundle(destination, ["姓名", "手机号码"], [])
        self.assertEqual([r["value"] for r in general["general"]], ["测试候选人（虚构）", "13800000000"])
        experiences = read_profile.load_experience_records(destination)
        self.assertEqual([(r["organization"], r["result"]) for r in experiences], [("CET-4", 550), ("CET-6", 520), ("IELTS", 6.5)])
        workbook = load_workbook(destination)
        self.assertEqual(workbook.sheetnames, ["使用说明", "个人资料", "经历清单", "附件清单"])
        values = repr([tuple(row) for sheet in workbook for row in sheet.iter_rows(values_only=True)])
        workbook.close()
        self.assertNotIn("虚构保留内容", values)
        self.assertNotIn("虚构来源", values)
        self.assertEqual([(w.sheet, w.row, w.code) for w in warnings], [("资料总表", 11, "unmapped_field")])
        self.assertNotIn("虚构保留内容", repr(warnings))

    def test_existing_destination_and_source_alias_are_never_overwritten(self):
        self.legacy()
        module = self.migration()
        original = self.path.read_bytes()
        with self.assertRaises(read_profile.ProfileError):
            module.migrate_legacy(self.path, self.path)
        destination = self.directory / "existing.xlsx"
        destination.write_bytes(b"keep existing content")
        with self.assertRaises(read_profile.ProfileError):
            module.migrate_legacy(self.path, destination)
        self.assertEqual(destination.read_bytes(), b"keep existing content")
        self.assertEqual(self.path.read_bytes(), original)

    def test_public_input_is_rejected_without_creating_output(self):
        self.public()
        destination = self.directory / "public.xlsx"
        with self.assertRaises(read_profile.ProfileError):
            self.migration().migrate_legacy(self.path, destination)
        self.assertFalse(destination.exists())

    def test_migration_rejects_repository_destination(self):
        self.legacy()
        with self.assertRaises(read_profile.ProfileError):
            self.migration().migrate_legacy(self.path, ROOT / "private-migration-test.xlsx")
        self.assertFalse((ROOT / "private-migration-test.xlsx").exists())

    def test_literal_formula_like_text_is_not_turned_into_a_formula(self):
        self.legacy()
        workbook = load_workbook(self.path)
        cell = workbook["资料总表"]["D9"]
        cell.value = "=FICTIONAL(1)"
        cell.data_type = "s"
        workbook.save(self.path)
        workbook.close()
        destination = self.directory / "public.xlsx"
        self.migration().migrate_legacy(self.path, destination)
        workbook = load_workbook(destination, data_only=False)
        self.assertEqual(workbook["个人资料"]["A2"].value, "=FICTIONAL(1)")
        self.assertEqual(workbook["个人资料"]["A2"].data_type, "s")
        workbook.close()
