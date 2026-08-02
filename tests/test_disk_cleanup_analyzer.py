import contextlib
import io
import os
import sys
import tempfile
import time
import unittest
from pathlib import Path
from unittest import mock


PROJECT_ROOT = os.path.dirname(os.path.dirname(__file__))
SRC_DIR = os.path.join(PROJECT_ROOT, "src")
if SRC_DIR not in sys.path:
    sys.path.insert(0, SRC_DIR)

from tools import disk_cleanup_analyzer as analyzer


class DiskCleanupAnalyzerTest(unittest.TestCase):
    def test_scan_path_applies_age_cutoff_and_skips_symlink(self):
        with tempfile.TemporaryDirectory() as temp_dir:
            root = Path(temp_dir)
            old_file = root / "old.tmp"
            new_file = root / "new.tmp"
            old_file.write_bytes(b"o" * 10)
            new_file.write_bytes(b"n" * 20)
            now = time.time()
            os.utime(old_file, (now - 20 * 86400, now - 20 * 86400))

            result = analyzer.scan_path(root, cutoff_timestamp=now - 7 * 86400)

            self.assertEqual(result.total_bytes, 30)
            self.assertEqual(result.reclaimable_bytes, 10)
            self.assertEqual(result.file_count, 2)
            self.assertEqual(result.reclaimable_file_count, 1)

    def test_candidates_separate_reclaimable_from_occupied_reference(self):
        with tempfile.TemporaryDirectory() as temp_dir:
            root = Path(temp_dir)
            update_cache = root / "Windows" / "SoftwareDistribution" / "Download"
            update_cache.mkdir(parents=True)
            (update_cache / "update.bin").write_bytes(b"x" * 123)

            specs = analyzer.build_candidate_specs(root)
            results = analyzer.scan_candidates(specs, workers=1, now=time.time())
            item = next(value for value in results if "更新下载缓存" in value["name"])

            self.assertEqual(item["occupied_bytes"], 123)
            self.assertEqual(item["reclaimable_bytes"], 0)
            self.assertFalse(item["reclaimable"])

    def test_analyze_drive_finds_profile_temp_without_full_scan(self):
        with tempfile.TemporaryDirectory() as temp_dir:
            root = Path(temp_dir)
            user_temp = root / "Users" / "alice" / "AppData" / "Local" / "Temp"
            user_temp.mkdir(parents=True)
            old_file = user_temp / "old.tmp"
            old_file.write_bytes(b"x" * 64)
            now = time.time()
            os.utime(old_file, (now - 10 * 86400, now - 10 * 86400))

            report = analyzer.analyze_drive(root, full_scan=False, workers=1, now=now)

            self.assertEqual(report["reclaimable_summary"][analyzer.LEVEL_RECOMMENDED], 64)
            self.assertEqual(report["overview"], [])
            self.assertTrue(report["policy"]["scan_read_only"])

    def test_format_bytes(self):
        self.assertEqual(analyzer.format_bytes(0), "0 B")
        self.assertEqual(analyzer.format_bytes(1024), "1.00 KiB")

    def test_progress_reports_files_bytes_and_completed_item(self):
        with tempfile.TemporaryDirectory() as temp_dir:
            root = Path(temp_dir)
            for index in range(300):
                (root / f"{index}.tmp").write_bytes(b"x")
            output = io.StringIO()

            with analyzer.ScanProgress(stream=output, refresh_interval=60) as progress:
                progress.begin_phase("测试扫描", 1)
                progress.item_started(root)
                analyzer.scan_path(root, progress=progress)
                progress.item_finished(root)
                progress.finish_phase()

            rendered = output.getvalue()
            self.assertIn("测试扫描", rendered)
            self.assertIn("1/1 项", rendered)
            self.assertIn("已扫描 300 个文件", rendered)

    def test_default_is_quick_and_full_requires_flag(self):
        parser = analyzer.build_parser()
        self.assertFalse(parser.parse_args([]).full)
        self.assertTrue(parser.parse_args(["--full"]).full)
        self.assertTrue(parser.parse_args(["--skip-analyze"]).skip_analyze)

    def test_cleanup_permanently_deletes_only_old_recommended_files(self):
        with tempfile.TemporaryDirectory() as temp_dir:
            root = Path(temp_dir)
            temp_root = root / "Windows" / "Temp"
            nested = temp_root / "empty-after-clean"
            nested.mkdir(parents=True)
            old_file = nested / "old.tmp"
            new_file = temp_root / "new.tmp"
            old_file.write_bytes(b"o" * 11)
            new_file.write_bytes(b"n" * 13)
            now = time.time()
            os.utime(old_file, (now - 10 * 86400, now - 10 * 86400))
            specs = analyzer.build_candidate_specs(root, temp_age_days=7)

            result = analyzer.clean_recommended_files(root, specs, workers=2, now=now)

            self.assertFalse(old_file.exists())
            self.assertFalse(nested.exists())
            self.assertTrue(new_file.exists())
            self.assertTrue(temp_root.exists())
            self.assertEqual(result.deleted_file_count, 1)
            self.assertEqual(result.deleted_bytes, 11)

    def test_cleanup_deletes_only_pip_and_npm_cache_files_older_than_30_days(self):
        with tempfile.TemporaryDirectory() as temp_dir:
            root = Path(temp_dir)
            local = root / "Users" / "alice" / "AppData" / "Local"
            pip_cache = local / "pip" / "Cache"
            npm_cache = local / "npm-cache"
            pip_cache.mkdir(parents=True)
            npm_cache.mkdir(parents=True)
            old_pip = pip_cache / "old.whl"
            new_pip = pip_cache / "new.whl"
            old_npm = npm_cache / "old.tgz"
            new_npm = npm_cache / "new.tgz"
            for path, content in (
                (old_pip, b"p" * 17),
                (new_pip, b"P" * 19),
                (old_npm, b"n" * 23),
                (new_npm, b"N" * 29),
            ):
                path.write_bytes(content)
            now = time.time()
            old_time = now - 31 * 86400
            os.utime(old_pip, (old_time, old_time))
            os.utime(old_npm, (old_time, old_time))
            specs = analyzer.build_candidate_specs(root, cache_age_days=30)

            result = analyzer.clean_recommended_files(root, specs, workers=2, now=now)

            self.assertFalse(old_pip.exists())
            self.assertFalse(old_npm.exists())
            self.assertTrue(new_pip.exists())
            self.assertTrue(new_npm.exists())
            self.assertEqual(result.deleted_file_count, 2)
            self.assertEqual(result.deleted_bytes, 40)

    def test_cleanup_confirmation_is_case_insensitive_and_exact(self):
        report = {
            "reclaimable_summary": {analyzer.LEVEL_RECOMMENDED: 10},
            "candidates": [
                {
                    "level": analyzer.LEVEL_RECOMMENDED,
                    "name": "旧临时文件",
                    "reclaimable_bytes": 10,
                    "reclaimable_file_count": 1,
                }
            ],
        }
        output = io.StringIO()

        self.assertTrue(
            analyzer._prompt_for_cleanup(
                report, stream=output, input_stream=io.StringIO("cLeAn\n")
            )
        )
        self.assertFalse(
            analyzer._prompt_for_cleanup(
                report, stream=io.StringIO(), input_stream=io.StringIO("clean all\n")
            )
        )

    def test_skip_analyze_cleans_without_running_analysis(self):
        with tempfile.TemporaryDirectory() as temp_dir:
            root = Path(temp_dir)
            temp_root = root / "Windows" / "Temp"
            temp_root.mkdir(parents=True)
            old_file = temp_root / "old.tmp"
            old_file.write_bytes(b"old")
            old_time = time.time() - 10 * 86400
            os.utime(old_file, (old_time, old_time))

            with contextlib.redirect_stdout(io.StringIO()), \
                 mock.patch.object(sys, "stdin", io.StringIO("clean\n")), \
                 mock.patch.object(
                     analyzer, "analyze_drive", side_effect=AssertionError("must not analyze")
                 ):
                exit_code = analyzer.main(
                    ["--drive", str(root), "--skip-analyze", "--no-progress"]
                )

            self.assertEqual(exit_code, 0)
            self.assertFalse(old_file.exists())

    def test_skip_analyze_rejects_conflicting_modes(self):
        with tempfile.TemporaryDirectory() as temp_dir, \
             contextlib.redirect_stderr(io.StringIO()):
            self.assertEqual(
                analyzer.main(["--drive", temp_dir, "--skip-analyze", "--full"]),
                2,
            )


if __name__ == "__main__":
    unittest.main()
