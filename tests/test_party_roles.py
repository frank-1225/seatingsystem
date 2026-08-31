import io
import os
import sqlite3
import tempfile
import unittest

import app as seating_app


class PartyRoleTests(unittest.TestCase):
    def setUp(self):
        self.temp_dir = tempfile.TemporaryDirectory()
        self.original_db_file = seating_app.DB_FILE
        seating_app.DB_FILE = os.path.join(self.temp_dir.name, "party-role-test.db")
        seating_app.init_db()
        seating_app.app.config.update(TESTING=True)
        self.client = seating_app.app.test_client()
        meeting = self.client.post("/api/meetings", json={"name": "主宾测试会"}).get_json()
        self.meeting_id = meeting["id"]

    def tearDown(self):
        seating_app.DB_FILE = self.original_db_file
        self.temp_dir.cleanup()

    def add_attendee(self, name, party_role="unassigned"):
        response = self.client.post(
            f"/api/meetings/{self.meeting_id}/attendees",
            json={
                "name": name,
                "unit": "测试单位",
                "category": "参会人员",
                "phone": "13800138000",
                "party_role": party_role,
                "extra": {},
            },
        )
        self.assertEqual(response.status_code, 200)

    def test_attendee_party_role_can_be_saved_and_bulk_updated(self):
        self.add_attendee("主方一", "host")
        self.add_attendee("宾方一", "guest")
        self.add_attendee("待分类")

        attendees = self.client.get(
            f"/api/meetings/{self.meeting_id}/attendees"
        ).get_json()
        self.assertEqual(
            [attendee["party_role"] for attendee in attendees],
            ["host", "guest", "unassigned"],
        )

        response = self.client.post(
            f"/api/meetings/{self.meeting_id}/attendees/bulk_role",
            json={"ids": [attendees[2]["id"]], "party_role": "下级"},
        )
        self.assertEqual(response.status_code, 200)
        self.assertEqual(response.get_json()["party_role"], "guest")

        updated = self.client.get(
            f"/api/meetings/{self.meeting_id}/attendees"
        ).get_json()
        self.assertEqual(updated[2]["party_role"], "guest")

    def test_csv_import_recognizes_host_and_guest_aliases(self):
        csv_data = "序号,姓名,单位,主宾\n1,甲,主方单位,上级\n1,乙,宾方单位,宾方\n2,丙,其他单位,\n"
        response = self.client.post(
            f"/api/meetings/{self.meeting_id}/upload",
            data={"file": (io.BytesIO(csv_data.encode("utf-8-sig")), "roles.csv")},
            content_type="multipart/form-data",
        )
        self.assertEqual(response.status_code, 200, response.get_json(silent=True))

        attendees = self.client.get(
            f"/api/meetings/{self.meeting_id}/attendees"
        ).get_json()
        roles_by_name = {attendee["name"]: attendee["party_role"] for attendee in attendees}
        self.assertEqual(
            roles_by_name,
            {"甲": "host", "乙": "guest", "丙": "unassigned"},
        )

    def test_existing_database_is_migrated_with_party_role_column(self):
        legacy_db = os.path.join(self.temp_dir.name, "legacy.db")
        connection = sqlite3.connect(legacy_db)
        connection.execute(
            "CREATE TABLE attendees (id INTEGER PRIMARY KEY, meeting_id INTEGER, "
            "sort_order INTEGER, name TEXT, unit TEXT, category TEXT, phone TEXT, extra JSON)"
        )
        connection.commit()
        connection.close()

        seating_app.DB_FILE = legacy_db
        seating_app.init_db()
        connection = sqlite3.connect(legacy_db)
        columns = {row[1] for row in connection.execute("PRAGMA table_info(attendees)")}
        connection.close()
        self.assertIn("party_role", columns)

    def test_page_exposes_party_classification_and_auto_assignment(self):
        html = self.client.get("/").get_data(as_text=True)
        self.assertIn("先分主宾，再分别排序", html)
        self.assertIn("主宾自动对坐（推荐）", html)
        self.assertIn("主宾分序自动排座（推荐）", html)
        self.assertIn("buildRoundtablePartyPlan", html)
        self.assertIn("sortLongtableSeatKeys", html)


if __name__ == "__main__":
    unittest.main()
