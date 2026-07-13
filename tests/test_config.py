import json
import os
import tempfile
import unittest
from pathlib import Path

from app.config import AppConfig, DAILY_TRACKING_FILENAME


def _same_path(left: str, right: str) -> bool:
    return os.path.normcase(os.path.abspath(left)) == os.path.normcase(os.path.abspath(right))


class AppConfigSingleOutputTests(unittest.TestCase):
    def test_load_migrates_legacy_split_folders_to_single_output(self) -> None:
        with tempfile.TemporaryDirectory() as tmp:
            root = Path(tmp)
            config_path = root / "Config" / "settings.json"
            config_path.parent.mkdir()
            config_path.write_text(
                json.dumps(
                    {
                        "app_root": str(root),
                        "download_folder": str(root / "Downloads"),
                        "output_folder": str(root / "Outputs"),
                        "daily_tracking_file": str(root / "Daily" / DAILY_TRACKING_FILENAME),
                    },
                    ensure_ascii=False,
                ),
                encoding="utf-8",
            )

            config = AppConfig.load(str(config_path))

            expected_output = str(root / "output")
            self.assertTrue(_same_path(config.output_folder, expected_output))
            self.assertTrue(
                _same_path(config.daily_tracking_file, str(root / "output" / DAILY_TRACKING_FILENAME))
            )

            saved = json.loads(config_path.read_text(encoding="utf-8"))
            self.assertNotIn("download_folder", saved)
            self.assertTrue(_same_path(saved["output_folder"], expected_output))

    def test_ensure_folders_creates_only_single_output_data_folder(self) -> None:
        with tempfile.TemporaryDirectory() as tmp:
            root = Path(tmp)
            config = AppConfig(
                {
                    "app_root": str(root),
                    "download_folder": str(root / "Downloads"),
                    "output_folder": str(root / "Outputs"),
                    "daily_tracking_file": str(root / "Daily" / DAILY_TRACKING_FILENAME),
                },
                str(root / "Config" / "settings.json"),
            )

            config.ensure_folders()

            self.assertTrue((root / "output").is_dir())
            self.assertFalse((root / "Downloads").exists())
            self.assertFalse((root / "Outputs").exists())
            self.assertFalse((root / "Daily").exists())


if __name__ == "__main__":
    unittest.main()
