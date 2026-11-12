from __future__ import annotations

import base64
import io
import json
import sys
import unittest
from pathlib import Path

import openpyxl
from pptx import Presentation

ROOT_DIR = Path(__file__).resolve().parent.parent
if str(ROOT_DIR) not in sys.path:
    sys.path.insert(0, str(ROOT_DIR))

from api.index import app  # noqa: E402

TEMPLATE_PATH = Path("/Users/jisanjomadder/Downloads/Monthly SEO Visibility Analysis - [property name] TEMPLATE.pptx")
BRAND_XLSX_PATH = Path("/Users/jisanjomadder/Desktop/Renaissance Shoals Resort & Spa - keyword list.xlsx")
BRAND_PDF_PATH = Path("/Users/jisanjomadder/Desktop/Semrush-SEO_Audit_Report-For_Renaissance_Shoals_Resort___Spa-4th_Nov_2026.pdf")
STANDALONE_XLSX_PATH = Path("/Users/jisanjomadder/Downloads/organic-positions-2026-11-12T09_05_25Z.xlsx")


class SEOReportAppTests(unittest.TestCase):
    def setUp(self):
        app.config["TESTING"] = True
        self.client = app.test_client()

    def test_health_endpoint(self):
        resp = self.client.get("/health")
        self.assertEqual(resp.status_code, 200)
        data = resp.get_json()
        self.assertEqual(data["status"], "ok")
        self.assertIn("max_file_mb", data)
        self.assertIn("max_request_mb", data)
        self.assertIn("X-Request-Id", resp.headers)

    def test_generate_rejects_missing_files(self):
        resp = self.client.post("/generate", data={}, content_type="multipart/form-data")
        self.assertEqual(resp.status_code, 400)
        self.assertIn("error", resp.get_json())

    def test_generate_rejects_wrong_extension(self):
        resp = self.client.post(
            "/generate",
            data={
                "template": (io.BytesIO(b"PK\x03\x04dummy"), "bad.txt"),
                "xlsx": (io.BytesIO(b"PK\x03\x04dummy"), "data.xlsx"),
            },
            content_type="multipart/form-data",
        )
        self.assertEqual(resp.status_code, 400)
        self.assertIn(".pptx", resp.get_json()["error"])

    def test_generate_rejects_invalid_magic_bytes(self):
        resp = self.client.post(
            "/generate",
            data={
                "template": (io.BytesIO(b"NOT_A_ZIP_FILE"), "template.pptx"),
                "xlsx": (io.BytesIO(b"PK\x03\x04dummy"), "data.xlsx"),
            },
            content_type="multipart/form-data",
        )
        self.assertEqual(resp.status_code, 400)
        self.assertIn("valid", resp.get_json()["error"].lower())

    def test_generate_rejects_ssrf_target_url(self):
        resp = self.client.post(
            "/generate",
            data={
                "template": (io.BytesIO(b"PK\x03\x04dummy"), "template.pptx"),
                "xlsx": (io.BytesIO(b"PK\x03\x04dummy"), "data.xlsx"),
                "target_url": "http://127.0.0.1/admin",
            },
            content_type="multipart/form-data",
        )
        self.assertEqual(resp.status_code, 400)
        self.assertIn("not allowed", resp.get_json()["error"].lower())

    def test_crawl_rejects_missing_payload_and_localhost(self):
        resp_missing = self.client.post("/crawl", json={})
        self.assertEqual(resp_missing.status_code, 400)

        resp_ssrf = self.client.post("/crawl", json={"target_url": "http://localhost:8080"})
        self.assertEqual(resp_ssrf.status_code, 400)
        self.assertTrue(resp_ssrf.get_json().get("error"))

    @unittest.skipUnless(
        TEMPLATE_PATH.exists() and BRAND_XLSX_PATH.exists() and BRAND_PDF_PATH.exists(),
        "Local brand hotel fixture files not present",
    )
    def test_generate_brand_hotel_end_to_end(self):
        with open(TEMPLATE_PATH, "rb") as tf, open(BRAND_XLSX_PATH, "rb") as xf, open(BRAND_PDF_PATH, "rb") as pf:
            resp = self.client.post(
                "/generate",
                data={
                    "template": (tf, TEMPLATE_PATH.name),
                    "xlsx": (xf, BRAND_XLSX_PATH.name),
                    "pdf": (pf, BRAND_PDF_PATH.name),
                },
                content_type="multipart/form-data",
            )
        self.assertEqual(resp.status_code, 200)
        self.assertIn(".pptx", resp.headers.get("Content-Disposition", ""))
        summary_b64 = resp.headers.get("X-Report-Summary", "")
        self.assertTrue(summary_b64)
        summary = json.loads(base64.b64decode(summary_b64).decode("utf-8"))
        self.assertEqual(summary["slide_count"], 4)
        self.assertTrue(summary["is_brand_hosted"])
        self.assertIn("rank_buckets", summary)

        prs = Presentation(io.BytesIO(resp.data))
        # Brand hotel removes Slide 5 (SEO Errors Report) -> 4 slides remain
        self.assertEqual(len(prs.slides), 4)

        # Slide 2 Row 1 should not contain leftover extra paragraph 'al'
        slide2_tables = [sh.table for sh in prs.slides[1].shapes if sh.has_table]
        self.assertTrue(slide2_tables)
        row1_cell0 = slide2_tables[0].cell(1, 0)
        self.assertEqual(len(row1_cell0.text_frame.paragraphs), 1)
        self.assertNotIn("\nal", row1_cell0.text.lower())

    @unittest.skipUnless(
        TEMPLATE_PATH.exists() and STANDALONE_XLSX_PATH.exists(),
        "Local standalone hotel fixture files not present",
    )
    def test_generate_standalone_hotel_end_to_end(self):
        # Simulate browser-side slimWorkbookIfNeeded (raw export is 6.57 MB > 5 MB limit)
        wb_in = openpyxl.load_workbook(STANDALONE_XLSX_PATH, read_only=True, data_only=True)
        wb_out = openpyxl.Workbook()
        ws_out = wb_out.active
        ws_in = wb_in.active
        ws_out.title = ws_in.title
        for idx, row in enumerate(ws_in.iter_rows(values_only=True)):
            if idx > 500:
                break
            ws_out.append(list(row))
        wb_in.close()
        slim_buf = io.BytesIO()
        wb_out.save(slim_buf)
        slim_buf.seek(0)

        with open(TEMPLATE_PATH, "rb") as tf:
            resp = self.client.post(
                "/generate",
                data={
                    "template": (tf, TEMPLATE_PATH.name),
                    "xlsx": (slim_buf, STANDALONE_XLSX_PATH.name),
                    "target_url": "https://thenines.com/",
                },
                content_type="multipart/form-data",
            )
        self.assertEqual(resp.status_code, 200)
        prs = Presentation(io.BytesIO(resp.data))
        # Stand-alone hotel keeps Slide 5 -> 5 slides
        self.assertEqual(len(prs.slides), 5)

        # Slide 3 rankings must all be within 1..40
        slide3_tables = [sh.table for sh in prs.slides[2].shapes if sh.has_table]
        self.assertTrue(slide3_tables)
        for row_idx in range(1, len(slide3_tables[0].rows)):
            rank_text = slide3_tables[0].cell(row_idx, 1).text.strip()
            if rank_text:
                rank_val = int(rank_text)
                self.assertGreaterEqual(rank_val, 1)
                self.assertLessEqual(rank_val, 40)

        # Slide 5 must not contain "NOTE FOR SEO TEAM"
        slide5_text = "\n".join(sh.text for sh in prs.slides[4].shapes if sh.has_text_frame)
        self.assertNotIn("NOTE FOR SEO TEAM", slide5_text.upper())


if __name__ == "__main__":
    unittest.main()
