from __future__ import annotations

import logging
import subprocess
import tempfile
import unittest
from pathlib import Path
from unittest.mock import patch

from app.services.rpa_launcher import RpaLaunchError, launch_bat


class RpaLauncherTests(unittest.TestCase):
    def setUp(self) -> None:
        self.temp = tempfile.TemporaryDirectory()
        self.root = Path(self.temp.name)
        self.folder = self.root / "folder with spaces"
        self.folder.mkdir()
        self.bat = self.folder / "run input information.bat"
        self.bat.write_text("@echo off\n", encoding="utf-8")
        self.logger = logging.getLogger("rpa-launcher-test")

    def tearDown(self) -> None:
        self.temp.cleanup()

    @patch("app.services.rpa_launcher.subprocess.Popen")
    def test_launch_existing_bat_with_spaces(self, popen) -> None:
        process = popen.return_value
        process.wait.return_value = 0

        result = launch_bat(
            str(self.bat),
            logger=self.logger,
            description="nhập thông tin",
            missing_message="missing",
        )

        self.assertIs(result, process)
        popen.assert_called_once()
        args, kwargs = popen.call_args
        self.assertEqual(args[0], ["cmd.exe", "/d", "/c", str(self.bat.resolve())])
        self.assertEqual(kwargs["cwd"], str(self.folder.resolve()))
        self.assertEqual(kwargs["env"]["PYTHONUTF8"], "1")
        self.assertEqual(kwargs["env"]["PYTHONIOENCODING"], "utf-8")
        self.assertEqual(kwargs["stderr"], subprocess.STDOUT)
        self.assertIn("stdout", kwargs)

    @patch("app.services.rpa_launcher.subprocess.Popen")
    def test_running_bat_is_returned_without_waiting_for_completion(self, popen) -> None:
        process = popen.return_value
        process.wait.side_effect = subprocess.TimeoutExpired("cmd.exe", 2)

        result = launch_bat(
            str(self.bat),
            logger=self.logger,
            description="nhập thông tin",
            missing_message="missing",
        )

        self.assertIs(result, process)

    @patch("app.services.rpa_launcher.subprocess.Popen")
    def test_fast_bat_error_is_reported(self, popen) -> None:
        process = popen.return_value
        process.wait.return_value = 7

        with self.assertRaises(RpaLaunchError) as ctx:
            launch_bat(
                str(self.bat),
                logger=self.logger,
                description="nhập thông tin",
                missing_message="missing",
            )

        self.assertIn("mã lỗi 7", str(ctx.exception))

    def test_missing_bat_is_reported(self) -> None:
        with self.assertRaises(RpaLaunchError) as ctx:
            launch_bat(
                str(self.folder / "missing.bat"),
                logger=self.logger,
                description="nhập thông tin",
                missing_message="Không tìm thấy file chạy RPA nhập thông tin.",
            )

        self.assertIn("Không tìm thấy file chạy RPA nhập thông tin", str(ctx.exception))

    @patch("app.services.rpa_launcher.subprocess.Popen", side_effect=OSError("boom"))
    def test_process_error_is_reported(self, _popen) -> None:
        with self.assertRaises(RpaLaunchError) as ctx:
            launch_bat(
                str(self.bat),
                logger=self.logger,
                description="nhập thông tin",
                missing_message="missing",
            )

        self.assertIn("Lỗi khi gọi tiến trình BAT", str(ctx.exception))


if __name__ == "__main__":
    unittest.main()
