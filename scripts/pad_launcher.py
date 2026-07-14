"""Launch Power Automate Desktop flows from a UTF-8 safe config."""

from __future__ import annotations

import argparse
import json
import os
import subprocess
import sys
import urllib.parse
from datetime import datetime
from pathlib import Path
from typing import Any, Dict, Optional


DEFAULT_FLOWS: Dict[str, Dict[str, Any]] = {
    "create_new": {
        "workflow_name": "Test - tạo dòng quyết toán mới",
        "workflow_id": "",
        "environment_id": "",
        "autologin": False,
    },
    "input_information": {
        "workflow_name": "Test - nhập thông tin",
        "workflow_id": "",
        "environment_id": "",
        "autologin": False,
        "input_argument_names": {
            "selection_json": "SelectionJsonPath",
            "helper_script": "HelperScriptPath",
            "project_root": "ProjectRoot",
        },
    },
    "input_expense": {
        "workflow_name": "Test - nhập khoản chi",
        "workflow_id": "",
        "environment_id": "",
        "autologin": False,
        "input_argument_names": {
            "selection_json": "SelectionJsonPath",
            "helper_script": "HelperScriptPath",
            "project_root": "ProjectRoot",
        },
    },
}


class PadLauncherError(RuntimeError):
    pass


def _now() -> str:
    return datetime.now().strftime("%Y-%m-%d %H:%M:%S")


def _log(project_root: Path, message: str) -> None:
    logs_dir = project_root / "Logs"
    logs_dir.mkdir(parents=True, exist_ok=True)
    with (logs_dir / "pad_launcher.log").open("a", encoding="utf-8") as f:
        f.write(f"{_now()} {message}\n")


def _deep_merge(base: Dict[str, Any], override: Dict[str, Any]) -> Dict[str, Any]:
    result = dict(base)
    for key, value in override.items():
        if isinstance(value, dict) and isinstance(result.get(key), dict):
            result[key] = _deep_merge(result[key], value)
        else:
            result[key] = value
    return result


def load_flow_config(project_root: Path, flow_key: str) -> Dict[str, Any]:
    flows = dict(DEFAULT_FLOWS)
    config_path = project_root / "Config" / "pad_flows.json"
    if config_path.exists():
        try:
            user_config = json.loads(config_path.read_text(encoding="utf-8-sig"))
        except (json.JSONDecodeError, OSError) as exc:
            raise PadLauncherError(f"Không đọc được Config/pad_flows.json: {exc}") from exc
        for key, value in (user_config or {}).items():
            if isinstance(value, dict):
                flows[key] = _deep_merge(flows.get(key, {}), value)

    if flow_key not in flows:
        raise PadLauncherError(f"Không có cấu hình PAD flow: {flow_key}")

    config = flows[flow_key]
    if not config.get("workflow_id") and not config.get("workflow_name"):
        raise PadLauncherError(
            f"PAD flow '{flow_key}' chưa có workflow_id hoặc workflow_name."
        )
    return config


def input_arguments(project_root: Path, flow_key: str, config: Dict[str, Any]) -> Dict[str, str]:
    if flow_key not in {"input_information", "input_expense"}:
        return {}
    if config.get("input_arguments_enabled") is False:
        return {}

    selection_json = project_root / "runtime" / "rpa_input_selection.json"
    helper_script = project_root / "scripts" / "rpa_excel_helper.py"
    if not selection_json.is_file():
        raise PadLauncherError(
            f"Thiếu file runtime/rpa_input_selection.json: {selection_json}"
        )
    if not helper_script.is_file():
        raise PadLauncherError(f"Thiếu helper script: {helper_script}")

    names = config.get("input_argument_names") or {}
    result: Dict[str, str] = {}
    if names.get("selection_json"):
        result[str(names["selection_json"])] = str(selection_json)
    if names.get("helper_script"):
        result[str(names["helper_script"])] = str(helper_script)
    if names.get("project_root"):
        result[str(names["project_root"])] = str(project_root)
    return result


def build_run_url(config: Dict[str, Any], arguments: Dict[str, str]) -> str:
    params = []
    environment_id = str(config.get("environment_id") or "").strip()
    workflow_id = str(config.get("workflow_id") or "").strip()
    workflow_name = str(config.get("workflow_name") or "").strip()

    if environment_id:
        params.append(("environmentId", environment_id))
    if workflow_id:
        params.append(("workflowId", workflow_id))
    else:
        params.append(("workflowName", workflow_name))
    if arguments:
        params.append(
            (
                "inputArguments",
                json.dumps(arguments, ensure_ascii=False, separators=(",", ":")),
            )
        )
    if bool(config.get("autologin")):
        params.append(("autologin", "true"))

    return "ms-powerautomate:/console/flow/run?" + urllib.parse.urlencode(params)


def _format_known_flows(known: Dict[str, str]) -> str:
    if not known:
        return "(không tìm thấy flow PAD nào trong lịch sử local)"
    return "; ".join(f"{name} [{workflow_id}]" for workflow_id, name in sorted(known.items(), key=lambda item: item[1]))


def resolve_flow_config(
    project_root: Path,
    flow_key: str,
    config: Dict[str, Any],
    known: Dict[str, str],
) -> Dict[str, Any]:
    """Prefer workflowId and fail fast when a configured name is clearly unknown."""
    workflow_id = str(config.get("workflow_id") or "").strip()
    workflow_name = str(config.get("workflow_name") or "").strip()

    if workflow_id:
        if known and workflow_id not in known:
            _log(project_root, f"WARNING workflow_id chưa thấy trong lịch sử local: {workflow_id}")
        return dict(config)

    if not workflow_name:
        raise PadLauncherError(f"PAD flow '{flow_key}' chưa có workflow_id hoặc workflow_name.")

    matches = [known_id for known_id, known_name in known.items() if known_name == workflow_name]
    if len(matches) == 1:
        resolved = dict(config)
        resolved["workflow_id"] = matches[0]
        _log(project_root, f"Resolved workflow_name '{workflow_name}' to workflow_id={matches[0]}")
        return resolved

    if len(matches) > 1:
        raise PadLauncherError(
            f"PAD flow '{flow_key}' có nhiều workflow trùng tên '{workflow_name}'. "
            "Vui lòng điền workflow_id trong Config/pad_flows.json."
        )

    if known:
        raise PadLauncherError(
            f"Chưa tìm thấy PAD flow '{workflow_name}' cho '{flow_key}' trong lịch sử local. "
            "Hãy kiểm tra lại workflow_name/workflow_id trong Config/pad_flows.json. "
            f"Các flow local hiện có: {_format_known_flows(known)}"
        )

    _log(project_root, f"WARNING không đọc được lịch sử PAD local để kiểm tra workflow_name: {workflow_name}")
    return dict(config)


def _program_files_candidates() -> list[Path]:
    result = []
    for env_name in ("ProgramFiles(x86)", "ProgramFiles"):
        root = os.environ.get(env_name)
        if root:
            result.append(
                Path(root)
                / "Power Automate Desktop"
                / "dotnet"
                / "PAD.Console.Host.exe"
            )
    return result


def _running_pad_console_path() -> Optional[Path]:
    try:
        completed = subprocess.run(
            [
                "powershell",
                "-NoProfile",
                "-ExecutionPolicy",
                "Bypass",
                "-Command",
                "(Get-Process PAD.Console.Host -ErrorAction SilentlyContinue | "
                "Select-Object -First 1 -ExpandProperty Path)",
            ],
            capture_output=True,
            text=True,
            encoding="utf-8",
            errors="replace",
            timeout=5,
            check=False,
        )
    except (OSError, subprocess.TimeoutExpired):
        return None
    path = completed.stdout.strip()
    if not path:
        return None
    candidate = Path(path)
    return candidate if candidate.is_file() else None


def _windows_apps_candidates() -> list[Path]:
    root = Path(os.environ.get("ProgramW6432") or r"C:\Program Files") / "WindowsApps"
    try:
        return sorted(
            root.glob("Microsoft.PowerAutomateDesktop_*__8wekyb3d8bbwe/dotnet/PAD.Console.Host.exe"),
            reverse=True,
        )
    except OSError:
        return []


def find_pad_console() -> Optional[Path]:
    candidates = _program_files_candidates()
    running = _running_pad_console_path()
    if running is not None:
        candidates.append(running)
    candidates.extend(_windows_apps_candidates())
    for candidate in candidates:
        if candidate.is_file():
            return candidate
    return None


def known_local_flows() -> Dict[str, str]:
    base = Path.home() / "AppData" / "Local" / "Microsoft" / "Power Automate Desktop" / "Console" / "Scripts"
    result: Dict[str, str] = {}
    if not base.is_dir():
        return result
    for path in base.rglob("RunDefinition.json"):
        try:
            data = json.loads(path.read_text(encoding="utf-8-sig"))
        except (json.JSONDecodeError, OSError):
            continue
        workflow = data.get("workflow") or {}
        workflow_id = workflow.get("workflowId")
        name = workflow.get("name")
        if workflow_id and name:
            result[str(workflow_id)] = str(name)
    return result


def launch(url: str, project_root: Path) -> None:
    console = find_pad_console()
    if console is not None:
        _log(project_root, f"PAD console: {console}")
        try:
            subprocess.Popen([str(console), url])
            return
        except OSError as exc:
            _log(project_root, f"WARNING cannot start PAD console directly: {exc}")

    _log(project_root, "Using ms-powerautomate protocol.")
    if hasattr(os, "startfile"):
        os.startfile(url)  # type: ignore[attr-defined]
        return
    subprocess.Popen(["xdg-open", url])


def run(project_root: Path, flow_key: str, *, dry_run: bool = False) -> int:
    project_root = project_root.resolve()
    config = load_flow_config(project_root, flow_key)
    known = known_local_flows()
    config = resolve_flow_config(project_root, flow_key, config, known)
    args = input_arguments(project_root, flow_key, config)
    url = build_run_url(config, args)

    workflow_name = str(config.get("workflow_name") or "")
    workflow_id = str(config.get("workflow_id") or "")

    _log(project_root, f"Launching flow_key={flow_key} workflow_id={workflow_id or '-'} workflow_name={workflow_name or '-'}")
    _log(project_root, f"URL={url}")
    if dry_run:
        print(url)
        return 0
    launch(url, project_root)
    return 0


def main(argv: Optional[list[str]] = None) -> int:
    parser = argparse.ArgumentParser(description="Launch a PAD flow.")
    parser.add_argument("--project-root", default=str(Path(__file__).resolve().parents[1]))
    parser.add_argument("--flow", required=True, choices=sorted(DEFAULT_FLOWS))
    parser.add_argument("--dry-run", action="store_true")
    args = parser.parse_args(argv)

    try:
        return run(Path(args.project_root), args.flow, dry_run=args.dry_run)
    except PadLauncherError as exc:
        project_root = Path(args.project_root).resolve()
        _log(project_root, f"ERROR {exc}")
        print(f"[ERROR] {exc}", file=sys.stderr)
        return 1


if __name__ == "__main__":
    raise SystemExit(main())
