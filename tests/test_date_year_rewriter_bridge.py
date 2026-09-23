import importlib.util
import unittest
from pathlib import Path
from types import SimpleNamespace


BRIDGE = Path(__file__).resolve().parents[1] / "scripts" / "date_year_rewriter_bridge.py"
SPEC = importlib.util.spec_from_file_location("date_year_rewriter_bridge", BRIDGE)
bridge = importlib.util.module_from_spec(SPEC)
SPEC.loader.exec_module(bridge)


class DateYearEvidenceTests(unittest.TestCase):
    def test_real_batch_ocr_text_keeps_actual_dates(self):
        for value in ("2023-04-19 16:06:45", "十一月2022", "2022-11"):
            with self.subTest(value=value):
                spans = bridge.date_year_spans(value, 2025)
                self.assertEqual(len(spans), 1)
                self.assertIn(value[spans[0][0]:spans[0][1]], ("2023", "2022"))

    def test_real_batch_ocr_noise_is_not_a_date(self):
        for value in ("0002 李强", "联季2020", "2000-", "04-19 16:06:45", "1916", "2026-02-30", "2024-02-29", "2000-3000", "2026-04-19 16:06:45", "2033-11-17", "十一月2028", "2025-11"):
            with self.subTest(value=value):
                self.assertEqual(bridge.date_year_spans(value, 2025), [])

    def test_year_1916_is_allowed_when_attached_to_valid_date(self):
        self.assertEqual(bridge.date_year_spans("1916-04-19", 2025), [(0, 4)])

    def test_original_ocr_boxes_are_required(self):
        result = SimpleNamespace(
            boxes=[[(0, 0), (100, 0), (100, 20), (0, 20)], [(0, 22), (100, 22), (100, 42), (0, 42)]],
            txts=["2023-04-19 16:06:45", "1916"],
        )
        fields = bridge.ocr_date_fields(result, 2025)
        self.assertEqual(len(fields), 1)
        self.assertEqual(fields[0][0], "2023-04-19 16:06:45")
        self.assertEqual(fields[0][2], (0, 4))


if __name__ == "__main__":
    unittest.main()
