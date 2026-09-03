import io
import tempfile
import unittest
from pathlib import Path

from openpyxl import Workbook

import app as mail_app


HEADERS = ["公司", "联系人", "职位", "国家/地区", "产品", "匹配理由", "邮箱", "电话", "官网"]


def workbook_bytes():
    workbook = Workbook()
    sheet = workbook.active
    sheet.title = "客户线索"
    sheet.append(HEADERS)
    sheet.append(
        [
            "Example Fasteners",
            "Amy Lee",
            "采购经理",
            "美国",
            "紧固件",
            "采购岗位与产品匹配",
            "amy@example.com sales@example.com AMY@example.com",
            "+1 555 0100",
            "https://example.com",
        ]
    )
    sheet.append(["No Email Hardware", "", "", "英国", "门窗五金", "产品匹配", "", "+44 20 0000", ""])
    output = io.BytesIO()
    workbook.save(output)
    workbook.close()
    return output.getvalue()


class LeadParserTests(unittest.TestCase):
    def test_splits_space_and_newline_separated_emails(self):
        self.assertEqual(
            mail_app.split_lead_emails(
                "one@example.com two@example.com\nthree@example.com"
            ),
            ["one@example.com", "two@example.com", "three@example.com"],
        )

    def test_preserves_display_name_and_removes_duplicates(self):
        self.assertEqual(
            mail_app.split_lead_emails(
                "Sales Team <Sales@Example.com>; sales@example.com"
            ),
            ["sales@example.com"],
        )

    def test_parses_template_and_splits_multiple_emails(self):
        parsed = mail_app.parse_lead_upload("leads.xlsx", workbook_bytes())

        self.assertEqual(parsed["sheet_name"], "客户线索")
        self.assertEqual(len(parsed["records"]), 2)
        self.assertEqual(
            parsed["records"][0]["emails"], ["amy@example.com", "sales@example.com"]
        )
        self.assertEqual(parsed["records"][0]["source_row"], 2)

    def test_rejects_unknown_spreadsheet_shape(self):
        workbook = Workbook()
        workbook.active.append(["未知字段", "其他字段"])
        output = io.BytesIO()
        workbook.save(output)
        workbook.close()

        with self.assertRaisesRegex(ValueError, "没有工作表符合"):
            mail_app.parse_lead_upload("unknown.xlsx", output.getvalue())


class LeadImportApiTests(unittest.TestCase):
    def setUp(self):
        self.temp_dir = tempfile.TemporaryDirectory()
        self.original_data_dir = mail_app.DATA_DIR
        self.original_db_path = mail_app.DB_PATH
        self.original_upload_dir = mail_app.UPLOAD_DIR
        mail_app.DATA_DIR = Path(self.temp_dir.name)
        mail_app.DB_PATH = mail_app.DATA_DIR / "mail.db"
        mail_app.UPLOAD_DIR = mail_app.DATA_DIR / "attachments"
        mail_app.app.config.update(TESTING=True, SECRET_KEY="test-secret")
        mail_app.init_db()
        self.client = mail_app.app.test_client()
        with self.client.session_transaction() as session:
            session["authenticated"] = True

    def tearDown(self):
        mail_app.DATA_DIR = self.original_data_dir
        mail_app.DB_PATH = self.original_db_path
        mail_app.UPLOAD_DIR = self.original_upload_dir
        self.temp_dir.cleanup()

    def upload(self):
        return self.client.post(
            "/api/lead-imports",
            data={"file": (io.BytesIO(workbook_bytes()), "客户线索.xlsx")},
            content_type="multipart/form-data",
        )

    def test_persists_statistics_and_marks_repeat_imports(self):
        first = self.upload()
        self.assertEqual(first.status_code, 201)
        self.assertEqual(first.json["total_rows"], 2)
        self.assertEqual(first.json["email_rows"], 1)
        self.assertEqual(first.json["duplicate_rows"], 0)
        self.assertEqual(first.json["filename"], "客户线索.xlsx")
        self.assertEqual(len(first.json["records"]), 2)

        second = self.upload()
        self.assertEqual(second.status_code, 201)
        self.assertEqual(second.json["duplicate_rows"], 2)

        history = self.client.get("/api/lead-imports")
        self.assertEqual(history.status_code, 200)
        self.assertEqual(len(history.json), 2)

    def test_adds_selected_leads_as_deduplicated_contacts(self):
        imported = self.upload()
        import_id = imported.json["id"]
        record_ids = [record["id"] for record in imported.json["records"]]

        first = self.client.post(
            f"/api/lead-imports/{import_id}/contacts",
            json={"record_ids": record_ids},
        )
        self.assertEqual(first.status_code, 200)
        self.assertEqual(first.json["created"], 2)
        self.assertEqual(first.json["existing"], 0)
        self.assertEqual(first.json["skipped"], 1)

        contacts = self.client.get("/api/contacts").json
        self.assertEqual({contact["email"] for contact in contacts}, {"amy@example.com", "sales@example.com"})
        amy = next(contact for contact in contacts if contact["email"] == "amy@example.com")
        self.assertEqual((amy["first_name"], amy["last_name"]), ("Amy", "Lee"))
        self.assertEqual(amy["company"], "Example Fasteners")
        self.assertIn("采购经理", amy["notes"])

        second = self.client.post(
            f"/api/lead-imports/{import_id}/contacts",
            json={"record_ids": record_ids},
        )
        self.assertEqual(second.json["created"], 0)
        self.assertEqual(second.json["existing"], 2)
        self.assertEqual(second.json["skipped"], 1)


if __name__ == "__main__":
    unittest.main()
