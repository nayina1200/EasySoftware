import contextlib
import importlib.util
import io
import json
import os
import tempfile
import unittest
import zipfile
from pathlib import Path
from types import SimpleNamespace
from unittest.mock import patch


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

    def test_ocr_date_separators_with_spaces_or_apostrophes_are_valid(self):
        samples = {
            "2026 /12/30": "2026-12-30",
            "2026 12/31": "2026-12-31",
            "2026 '01/01": "2026-01-01",
        }
        for text, expected_anchor in samples.items():
            with self.subTest(text=text):
                spans = bridge.date_year_spans(text, 2026, include_target_year=True)
                self.assertEqual(len(spans), 1)
                self.assertEqual(text[spans[0][0]:spans[0][1]], "2026")
                self.assertIn((expected_anchor, bridge.date.fromisoformat(expected_anchor)),
                              bridge._valid_2026_date_anchors(text))

        box = [(0, 0), (100, 0), (100, 20), (0, 20)]
        ocr_result = SimpleNamespace(boxes=[box] * len(samples), txts=list(samples))
        fields, _diagnostics = bridge.adaptive_ocr_date_fields(
            SimpleNamespace(), SimpleNamespace(), object(), ocr_result,
            2026, include_target_year=True,
        )
        self.assertEqual([field[0] for field in fields], list(samples))

    def test_real_batch_ocr_noise_is_not_a_date(self):
        for value in ("0002 李强", "联季2020", "2000-", "04-19 16:06:45", "1916", "C1911-1", "CCGT_1_GT_1", "2026-02-30", "2024-02-29", "2000-3000", "2026-04-19 16:06:45", "2033-11-17", "十一月2028", "2025-11"):
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

    def test_adaptive_ocr_does_not_replace_primary_geometry_with_wider_box(self):
        primary_text = "2021年06月01日09:22:01"
        result = SimpleNamespace(
            boxes=[[(7, 90), (216, 90), (216, 110), (7, 110)]],
            txts=[primary_text],
        )
        engine = SimpleNamespace(
            detect_ui_dates=lambda *_args, **_kwargs: [
                SimpleNamespace(
                    text=primary_text + "星期二",
                    rect=(18, 88, 263, 110),
                    confidence=0.99,
                )
            ]
        )

        fields, diagnostics = bridge.adaptive_ocr_date_fields(
            SimpleNamespace(), engine, object(), result, 2025
        )

        self.assertEqual(fields, [(primary_text, (7, 90, 217, 111), (0, 4))])
        self.assertTrue(diagnostics["used"])

    def test_text_bounds_exclude_adjacent_colored_ui_icon(self):
        image = [[[255, 255, 255] for _ in range(70)] for _ in range(20)]
        for y in range(3, 17):
            for x in range(3, 16):
                image[y][x] = [210, 120, 40]  # Colored clock icon before the date.
        for y in range(5, 15):
            for x in range(25, 56):
                image[y][x] = [20, 20, 20]  # Date text ink.

        bounds = bridge.text_ink_bounds(
            image, (2, 1, 65, 19),
            background=(255, 255, 255), text_color=(20, 20, 20),
        )

        self.assertEqual(bounds, (25, 56))

    def test_background_restore_clears_year_pixels_not_flat_neighboring_ink(self):
        image = [[[255, 255, 255] for _ in range(32)] for _ in range(9)]
        for y in range(2, 7):
            for x in range(8, 20):
                image[y][x] = [70, 70, 70]  # Old year glyphs inside the cleanup box.
            for x in range(23, 32):
                image[y][x] = [70, 70, 70]  # Flat neighboring stroke, like 年's crossbar.
        output = [[pixel[:] for pixel in row] for row in image]

        bridge.restore_year_background(
            None, output, image, (8, 2, 20, 7),
            background_color=(255, 255, 255), text_color=(70, 70, 70),
        )

        self.assertTrue(all(
            pixel == [255, 255, 255]
            for row in output[2:7]
            for pixel in row[8:20]
        ))
        self.assertTrue(all(
            pixel == [70, 70, 70]
            for row in output[2:7]
            for pixel in row[23:32]
        ))

    def test_short_year_month_requires_crop_ocr_confirmation(self):
        image = [[[255, 255, 255] for _ in range(32)] for _ in range(24)]
        module = SimpleNamespace(
            cv2=SimpleNamespace(INTER_CUBIC=1, resize=lambda crop, *_args, **_kwargs: crop)
        )
        engine = SimpleNamespace(
            ocr_lock=contextlib.nullcontext(),
            ocr=lambda _crop: SimpleNamespace(txts=["CCGT_1_GT_1"]),
        )

        confirmed = bridge.corroborate_year_month_field(
            module, engine, image, "1911-3", (3, 3, 25, 20), 2025
        )

        self.assertFalse(confirmed)

        engine.ocr = lambda _crop: SimpleNamespace(txts=["2018-08"])
        confirmed = bridge.corroborate_year_month_field(
            module, engine, image, "2018-08", (3, 3, 25, 20), 2025
        )
        self.assertTrue(confirmed)


class DateYearPolicyTests(unittest.TestCase):
    def test_projects_recognized_source_dates_to_candidate_year(self):
        expected = [
            ("2026-12-22", bridge.date(2026, 12, 22)),
        ]
        self.assertEqual(
            bridge._date_anchors_for_year_field("2024年12月22日", (0, 4), 2026),
            expected,
        )
        self.assertEqual(
            bridge._date_anchors_for_year_field("2025/1/14", (0, 4), 2026),
            [("2026-01-14", bridge.date(2026, 1, 14))],
        )

    def test_future_projected_date_rolls_whole_document_back_to_2025(self):
        effective, trigger, reason = bridge._select_effective_target_year(
            2025,
            observed_2026_dates=[("2026-04-01", bridge.date(2026, 4, 1))],
            projected_2026_dates=[
                ("2026-04-01", bridge.date(2026, 4, 1)),
                ("2026-12-22", bridge.date(2026, 12, 22)),
            ],
            today=bridge.date(2026, 9, 26),
        )

        self.assertEqual(effective, 2025)
        self.assertEqual(trigger, "2026-12-22")
        self.assertIn("2026-12-22", reason)

    def test_candidate_year_is_kept_when_all_projected_dates_are_not_future(self):
        effective, trigger, reason = bridge._select_effective_target_year(
            2025,
            observed_2026_dates=[("2026-04-01", bridge.date(2026, 4, 1))],
            projected_2026_dates=[
                ("2026-01-14", bridge.date(2026, 1, 14)),
                ("2026-04-01", bridge.date(2026, 4, 1)),
            ],
            today=bridge.date(2026, 9, 26),
        )

        self.assertEqual((effective, trigger, reason), (2026, "2026-04-01", None))

    def test_explicit_2026_target_still_rolls_back_when_it_would_be_future(self):
        effective, _trigger, reason = bridge._select_effective_target_year(
            2026,
            observed_2026_dates=[],
            projected_2026_dates=[("2026-12-22", bridge.date(2026, 12, 22))],
            today=bridge.date(2026, 9, 26),
        )

        self.assertEqual(effective, 2025)
        self.assertIsNotNone(reason)

    def test_rewrite_docx_uses_document_wide_future_date_fallback(self):
        received_target_years = []
        recognized_fields = [
            ("2024年12月22日", (0, 0, 100, 20), (0, 4)),
            ("2025/1/14", (0, 22, 100, 42), (0, 4)),
            ("2026-04-01", (0, 44, 100, 64), (0, 4)),
        ]

        class FakeImage:
            def __enter__(self):
                return self

            def __exit__(self, *_args):
                return False

            def convert(self, *_args):
                return self

        class FakeArray:
            def __getitem__(self, _key):
                return self

        image_data = FakeArray()
        module = SimpleNamespace(
            np=SimpleNamespace(
                uint8=object(),
                asarray=lambda *_args, **_kwargs: image_data,
                array_equal=lambda *_args: False,
            ),
            cv2=SimpleNamespace(
                COLOR_RGB2BGR=1,
                COLOR_BGR2RGB=2,
                cvtColor=lambda image, *_args: image,
            ),
            Image=SimpleNamespace(open=lambda *_args: FakeImage()),
            ImageOps=SimpleNamespace(exif_transpose=lambda image: image),
        )
        engine = SimpleNamespace(ocr_lock=contextlib.nullcontext(), ocr=lambda _image: None)

        def replace_fields(_module, _engine, image, _fields, target_year):
            received_target_years.append(target_year)
            return image, [], []

        fixed_clock = SimpleNamespace(
            now=lambda: SimpleNamespace(
                astimezone=lambda: SimpleNamespace(date=lambda: bridge.date(2026, 9, 26))
            )
        )
        with tempfile.TemporaryDirectory() as temp_dir:
            docx = Path(temp_dir) / "future_projection.docx"
            with zipfile.ZipFile(docx, "w") as archive:
                archive.writestr("word/media/image1.png", b"test image")

            with (
                patch.object(bridge, "load_engine", return_value=(module, engine)),
                patch.object(
                    bridge,
                    "adaptive_ocr_date_fields",
                    return_value=(recognized_fields, {"used": False}),
                ),
                patch.object(bridge, "replace_year_fields", side_effect=replace_fields),
                patch.object(bridge, "encode_like_source", return_value=b"processed image"),
                patch.object(bridge, "BaseDateTime", fixed_clock),
                patch.dict(os.environ, {"EASYSOFTWARE_PERF_TRACE": "1"}),
            ):
                result = bridge.rewrite_docx(Path(temp_dir), docx, 2025)

        self.assertEqual(result["effective_target_year"], 2025)
        self.assertIn("2026-12-22", result["rollback_reason"])
        self.assertEqual(received_target_years, [2025])
        self.assertEqual(result["performance"]["ocr_calls"], 1)
        self.assertEqual(result["performance"]["primary_image_ocr_calls"], 1)
        self.assertEqual(result["performance"]["adaptive_or_local_ocr_calls"], 0)

    def test_batch_reuses_one_engine_and_isolates_document_failure(self):
        shared_engine = SimpleNamespace()
        runtime = (SimpleNamespace(), shared_engine)
        real_rewrite = bridge.rewrite_docx
        with tempfile.TemporaryDirectory() as temp_dir:
            root = Path(temp_dir)
            good = root / "good.docx"
            second_good = root / "second_good.docx"
            bad = root / "bad.docx"
            for path in (good, second_good):
                with zipfile.ZipFile(path, "w") as archive:
                    archive.writestr("word/document.xml", b"valid")
            bad.write_bytes(b"not a zip")
            request_text = "".join(
                json.dumps({"op": "rewrite", "path": str(path)}) + "\n"
                for path in (bad, good, second_good)
            ) + '{"op":"close"}\n'
            output = io.StringIO()
            with (
                patch.object(bridge, "load_engine", return_value=runtime) as load_engine,
                patch.object(bridge, "rewrite_docx", wraps=real_rewrite) as rewrite,
                patch.object(bridge.sys, "stdin", io.StringIO(request_text)),
                patch.object(bridge.sys, "stdout", output),
            ):
                status = bridge.rewrite_docx_stream(root, 2025)
            responses = [json.loads(line) for line in output.getvalue().splitlines()]
            results = responses[1:]

            self.assertEqual(status, 0)
            self.assertTrue(responses[0]["ready"])
            self.assertEqual(load_engine.call_count, 1)
            self.assertEqual(len(results), 3)
            self.assertEqual(results[0]["status"], "FAILED")
            self.assertFalse(results[0]["changed"])
            self.assertFalse(bad.with_suffix(bad.suffix + ".date-year.tmp").exists())
            self.assertEqual(results[1]["status"], "UNCHANGED")
            self.assertEqual(results[2]["status"], "UNCHANGED")
            self.assertTrue(good.is_file() and second_good.is_file())
            self.assertEqual(rewrite.call_count, 3)
            self.assertTrue(all(call.kwargs["runtime"] is runtime for call in rewrite.call_args_list))

    def test_batch_docx_output_matches_single_document_path(self):
        runtime = (SimpleNamespace(), SimpleNamespace())
        with tempfile.TemporaryDirectory() as temp_dir:
            root = Path(temp_dir)
            single = root / "single.docx"
            batched = root / "batched.docx"
            for path in (single, batched):
                with zipfile.ZipFile(path, "w") as archive:
                    archive.writestr("word/document.xml", b"same document")
            single_result = bridge.rewrite_docx(root, single, 2025, runtime=runtime)
            output = io.StringIO()
            with (
                patch.object(bridge, "load_engine", return_value=runtime),
                patch.object(bridge.sys, "stdin", io.StringIO(
                    json.dumps({"op": "rewrite", "path": str(batched)}) + "\n" + json.dumps({"op": "close"}) + "\n"
                )),
                patch.object(bridge.sys, "stdout", output),
            ):
                bridge.rewrite_docx_stream(root, 2025)
            responses = [json.loads(line) for line in output.getvalue().splitlines()]
            batch_result = responses[1]

            self.assertEqual(single_result["status"], batch_result["status"])
            self.assertEqual(single_result["changed"], batch_result["changed"])
            with zipfile.ZipFile(single) as single_zip, zipfile.ZipFile(batched) as batch_zip:
                self.assertEqual(single_zip.namelist(), batch_zip.namelist())
                self.assertEqual(
                    {name: single_zip.read(name) for name in single_zip.namelist()},
                    {name: batch_zip.read(name) for name in batch_zip.namelist()},
                )

    def test_legacy_single_document_cli_call_remains_supported(self):
        runtime = (SimpleNamespace(), SimpleNamespace())
        with tempfile.TemporaryDirectory() as temp_dir:
            root = Path(temp_dir)
            docx = root / "legacy.docx"
            with zipfile.ZipFile(docx, "w") as archive:
                archive.writestr("word/document.xml", b"same document")
            output = io.StringIO()
            with (
                patch.object(bridge, "load_engine", return_value=runtime),
                patch.object(bridge.sys, "argv", [str(BRIDGE), str(root), str(docx), "2025"]),
                patch.object(bridge.sys, "stdout", output),
            ):
                status = bridge.main()
            self.assertEqual(status, 0)
            self.assertEqual(json.loads(output.getvalue())["status"], "UNCHANGED")


if __name__ == "__main__":
    unittest.main()
