from __future__ import annotations

import json
import tempfile
import unittest
import urllib.parse
from pathlib import Path
from unittest.mock import patch

from scripts import pad_launcher


class PadLauncherTests(unittest.TestCase):
    def setUp(self) -> None:
        self.temp = tempfile.TemporaryDirectory()
        self.root = Path(self.temp.name)
        (self.root / "Config").mkdir()
        (self.root / "runtime").mkdir()
        (self.root / "scripts").mkdir()
        (self.root / "runtime" / "rpa_input_selection.json").write_text("{}", encoding="utf-8")
        (self.root / "scripts" / "rpa_excel_helper.py").write_text("# helper\n", encoding="utf-8")
        (self.root / "Config" / "pad_flows.json").write_text(
            json.dumps(
                {
                    "create_new": {
                        "workflow_name": "Test - tạo dòng quyết toán mới",
                        "workflow_id": "",
                    },
                    "input_information": {
                        "workflow_name": "Test - nhập thông tin",
                        "workflow_id": "",
                    },
                },
                ensure_ascii=False,
            ),
            encoding="utf-8",
        )

    def tearDown(self) -> None:
        self.temp.cleanup()

    def _query(self, url: str) -> dict[str, list[str]]:
        return urllib.parse.parse_qs(urllib.parse.urlsplit(url).query)

    def test_create_new_url_preserves_vietnamese_flow_name(self) -> None:
        config = pad_launcher.load_flow_config(self.root, "create_new")
        url = pad_launcher.build_run_url(config, {})

        self.assertEqual(
            self._query(url)["workflowName"],
            ["Test - tạo dòng quyết toán mới"],
        )

    def test_input_information_url_includes_selection_paths(self) -> None:
        config = pad_launcher.load_flow_config(self.root, "input_information")
        args = pad_launcher.input_arguments(self.root, "input_information", config)
        url = pad_launcher.build_run_url(config, args)
        encoded_args = self._query(url)["inputArguments"][0]
        payload = json.loads(encoded_args)

        self.assertEqual(self._query(url)["workflowName"], ["Test - nhập thông tin"])
        self.assertEqual(payload["SelectionJsonPath"], str(self.root / "runtime" / "rpa_input_selection.json"))
        self.assertEqual(payload["HelperScriptPath"], str(self.root / "scripts" / "rpa_excel_helper.py"))
        self.assertEqual(payload["ProjectRoot"], str(self.root))

    def test_input_information_can_disable_url_input_arguments(self) -> None:
        config = pad_launcher.load_flow_config(self.root, "input_information")
        config["input_arguments_enabled"] = False

        args = pad_launcher.input_arguments(self.root, "input_information", config)
        url = pad_launcher.build_run_url(config, args)

        self.assertEqual(args, {})
        self.assertNotIn("inputArguments", self._query(url))

    def test_run_resolves_local_workflow_name_to_id(self) -> None:
        workflow_id = "11111111-2222-3333-4444-555555555555"

        with (
            patch("scripts.pad_launcher.known_local_flows", return_value={workflow_id: "Test - nhập thông tin"}),
            patch("scripts.pad_launcher.launch") as launch,
        ):
            self.assertEqual(pad_launcher.run(self.root, "input_information"), 0)

        url = launch.call_args.args[0]
        query = self._query(url)
        self.assertEqual(query["workflowId"], [workflow_id])
        self.assertNotIn("workflowName", query)

    def test_run_reports_unknown_workflow_name_when_local_flows_exist(self) -> None:
        with patch(
            "scripts.pad_launcher.known_local_flows",
            return_value={"99999999-2222-3333-4444-555555555555": "Other flow"},
        ):
            with self.assertRaises(pad_launcher.PadLauncherError) as ctx:
                pad_launcher.run(self.root, "input_information")

        self.assertIn("Chưa tìm thấy PAD flow", str(ctx.exception))
        self.assertIn("Other flow", str(ctx.exception))


if __name__ == "__main__":
    unittest.main()
