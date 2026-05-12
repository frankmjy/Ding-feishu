from __future__ import annotations

import argparse
import json
import os
import subprocess
import sys
import threading
import time
import uuid
from dataclasses import dataclass, field
from datetime import datetime
from http import HTTPStatus
from http.server import BaseHTTPRequestHandler, ThreadingHTTPServer
from pathlib import Path
from typing import Any
from urllib.parse import parse_qs, urlparse

from sync_dingtalk_to_feishu import SyncError, load_settings, read_excel_rows


ROOT = Path(__file__).resolve().parent
DEFAULT_PORT = 8111
TASKS_FILE = ROOT / "sync_tasks.json"

DEFAULT_TASKS = {
    "drill": {
        "name": "演练问题同步",
        "env": ".env",
        "line": "LINE 01",
        "description": "钉钉演练问题下载到本地，再写入飞书演练管理多维表。",
    },
    "aliyun_problem": {
        "name": "EA118问题表模块",
        "env": ".env.aliyun_problem",
        "line": "LINE 02",
        "description": "钉钉问题 Excel 无附件导出，再同步到飞书多维表【阿里问题登记簿】。",
    },
}


@dataclass
class TaskState:
    id: str | None = None
    action: str | None = None
    status: str = "idle"
    command: list[str] = field(default_factory=list)
    started_at: str | None = None
    ended_at: str | None = None
    returncode: int | None = None
    process_pid: int | None = None
    module_id: str | None = None
    stop_requested: bool = False
    logs: list[str] = field(default_factory=list)


STATE = TaskState()
STATE_LOCK = threading.Lock()


class ArgsForSettings:
    def __init__(self, env: str = ".env") -> None:
        self.env = env
        self.excel = None
        self.sheet = None
        self.header_row = None
        self.mode = None
        self.create_missing_fields = None
        self.skip_download = False
        self.download_only = False
        self.dry_run = False


def now_iso() -> str:
    return datetime.now().strftime("%Y-%m-%d %H:%M:%S")


def append_log(line: str) -> None:
    with STATE_LOCK:
        STATE.logs.append(line.rstrip("\n"))
        STATE.logs = STATE.logs[-800:]


def state_snapshot() -> dict[str, Any]:
    with STATE_LOCK:
        return {
            "id": STATE.id,
            "action": STATE.action,
            "status": STATE.status,
            "command": STATE.command,
            "started_at": STATE.started_at,
            "ended_at": STATE.ended_at,
            "returncode": STATE.returncode,
            "process_pid": STATE.process_pid,
            "module_id": STATE.module_id,
            "can_stop": STATE.status == "running",
            "logs": STATE.logs,
        }


def load_task_configs() -> dict[str, dict[str, str]]:
    if not TASKS_FILE.exists():
        return dict(DEFAULT_TASKS)
    try:
        data = json.loads(TASKS_FILE.read_text(encoding="utf-8"))
    except Exception:
        return dict(DEFAULT_TASKS)
    raw_tasks = data.get("tasks") if isinstance(data, dict) else data
    if not isinstance(raw_tasks, dict):
        return dict(DEFAULT_TASKS)
    tasks: dict[str, dict[str, str]] = {}
    for key, value in raw_tasks.items():
        if not isinstance(value, dict):
            continue
        env = str(value.get("env") or "").strip()
        if not env:
            continue
        tasks[str(key)] = {
            "name": str(value.get("name") or key),
            "env": env,
            "line": str(value.get("line") or ""),
            "description": str(value.get("description") or ""),
        }
    return tasks or dict(DEFAULT_TASKS)


def task_config(module_id: str | None) -> tuple[str, dict[str, str]]:
    tasks = load_task_configs()
    if module_id and module_id in tasks:
        return module_id, tasks[module_id]
    if "aliyun_problem" in tasks:
        return "aliyun_problem", tasks["aliyun_problem"]
    first_key = next(iter(tasks))
    return first_key, tasks[first_key]


def masked(value: str, keep: int = 6) -> str:
    if not value:
        return ""
    if len(value) <= keep * 2:
        return value[:2] + "***"
    return f"{value[:keep]}...{value[-keep:]}"


def read_status(module_id: str | None = None) -> dict[str, Any]:
    selected_id, selected_task = task_config(module_id)
    settings = load_settings(ArgsForSettings(selected_task["env"]))
    excel_info: dict[str, Any] = {
        "path": str(settings.local_excel_path),
        "exists": settings.local_excel_path.exists(),
    }

    if settings.local_excel_path.exists():
        try:
            headers, rows = read_excel_rows(settings.local_excel_path, settings.sheet_name, settings.header_row)
            excel_info.update({"headers": headers, "row_count": len(rows), "error": None})
        except Exception as exc:  # Status should stay readable even if the workbook is odd.
            excel_info.update({"headers": [], "row_count": None, "error": str(exc)})

    return {
        "app": {
            "name": "钉钉文档同步飞书",
            "port": DEFAULT_PORT,
            "cwd": str(ROOT),
        },
        "module": {
            "id": selected_id,
            "name": selected_task["name"],
            "line": selected_task.get("line", ""),
            "description": selected_task.get("description", ""),
            "env": selected_task["env"],
            "sheet": settings.sheet_name or "第一个工作表",
            "mode": settings.sync_mode,
            "mapping": str(settings.field_mapping_file) if settings.field_mapping_file else "",
        },
        "modules": [
            {"id": key, **value}
            for key, value in load_task_configs().items()
        ],
        "targets": {
            "dingtalk_doc_url": settings.dingtalk_doc_url,
            "dingtalk_doc_title": settings.dingtalk_doc_title,
            "feishu_bitable_url": settings.feishu_bitable_url,
            "feishu_app_id": masked(settings.feishu_app_id),
            "feishu_app_token": masked(settings.feishu_bitable_app_token or ""),
        },
        "excel": excel_info,
        "task": state_snapshot(),
    }


def command_for_action(action: str, module_id: str | None = None) -> list[str]:
    _, selected_task = task_config(module_id)
    base = [sys.executable, "-u", "-X", "utf8", str(ROOT / "sync_dingtalk_to_feishu.py"), "--env", selected_task["env"]]
    actions = {
        "full": [],
        "sync-local": ["--skip-download"],
        "download-only": ["--download-only"],
        "dry-run": ["--skip-download", "--dry-run"],
    }
    if action not in actions:
        raise ValueError(f"Unsupported action: {action}")
    return base + actions[action]


def run_subprocess(action: str, command: list[str], task_id: str) -> None:
    env = os.environ.copy()
    env["PYTHONUTF8"] = "1"
    append_log(f"[dashboard] {now_iso()} task {task_id} started: {action}")

    try:
        process = subprocess.Popen(
            command,
            cwd=str(ROOT),
            stdout=subprocess.PIPE,
            stderr=subprocess.STDOUT,
            text=True,
            encoding="utf-8",
            errors="replace",
            env=env,
        )
        with STATE_LOCK:
            if STATE.id == task_id and STATE.status in {"running", "stopping"}:
                STATE.process_pid = process.pid
                should_stop = STATE.stop_requested
            else:
                should_stop = False
        if should_stop:
            terminate_process_tree(process.pid)
        assert process.stdout is not None
        for line in process.stdout:
            append_log(line)
        returncode = process.wait()
    except Exception as exc:
        returncode = 1
        append_log(f"[dashboard] failed to run task: {exc}")

    with STATE_LOCK:
        if STATE.id == task_id:
            stopped = STATE.stop_requested
            STATE.returncode = returncode
            STATE.ended_at = now_iso()
            STATE.status = "stopped" if stopped else ("success" if returncode == 0 else "error")
            STATE.process_pid = None
            STATE.stop_requested = False
    append_log(f"[dashboard] {now_iso()} task {task_id} ended with code {returncode}")


def start_task(action: str, module_id: str | None = None) -> dict[str, Any]:
    selected_id, _ = task_config(module_id)
    command = command_for_action(action, selected_id)
    with STATE_LOCK:
        if STATE.status in {"running", "stopping"}:
            raise RuntimeError("已有同步任务正在运行，请等待完成或先停止当前任务。")
        STATE.id = uuid.uuid4().hex[:10]
        STATE.action = action
        STATE.module_id = selected_id
        STATE.status = "running"
        STATE.command = command
        STATE.started_at = now_iso()
        STATE.ended_at = None
        STATE.returncode = None
        STATE.process_pid = None
        STATE.stop_requested = False
        STATE.logs = []
        task_id = STATE.id

    thread = threading.Thread(target=run_subprocess, args=(action, command, task_id), daemon=True)
    thread.start()
    return state_snapshot()


def terminate_process_tree(pid: int) -> None:
    if os.name == "nt":
        subprocess.run(
            ["taskkill", "/PID", str(pid), "/T", "/F"],
            cwd=str(ROOT),
            stdout=subprocess.DEVNULL,
            stderr=subprocess.DEVNULL,
            check=False,
        )
        cleanup_automation_browser_processes()
        return

    try:
        os.kill(pid, 15)
    except OSError:
        return


def ps_single_quote(value: str) -> str:
    return "'" + value.replace("'", "''") + "'"


def cleanup_automation_browser_processes() -> None:
    if os.name != "nt":
        return
    profile_marker = str((ROOT / ".browser" / "dingtalk").resolve())
    script = (
        f"$marker = {ps_single_quote(profile_marker)}; "
        "$names = @('msedge.exe','chrome.exe','msedgewebview2.exe'); "
        "Get-CimInstance Win32_Process | "
        "Where-Object { $_.CommandLine -and $_.CommandLine.Contains($marker) -and $names -contains $_.Name } | "
        "ForEach-Object { Stop-Process -Id $_.ProcessId -Force -ErrorAction SilentlyContinue }"
    )
    subprocess.run(
        ["powershell", "-NoProfile", "-ExecutionPolicy", "Bypass", "-Command", script],
        cwd=str(ROOT),
        stdout=subprocess.DEVNULL,
        stderr=subprocess.DEVNULL,
        check=False,
    )


def stop_task() -> dict[str, Any]:
    with STATE_LOCK:
        if STATE.status not in {"running", "stopping"}:
            raise RuntimeError("当前没有正在运行的任务。")
        pid = STATE.process_pid
        STATE.stop_requested = True
        STATE.status = "stopping"
        task_id = STATE.id

    append_log(f"[dashboard] {now_iso()} stop requested for task {task_id}.")
    if pid is not None:
        terminate_process_tree(pid)
    else:
        append_log("[dashboard] task process is still starting; it will stop as soon as the process is ready.")
        cleanup_automation_browser_processes()
    return state_snapshot()


def json_bytes(payload: Any) -> bytes:
    return json.dumps(payload, ensure_ascii=False, indent=2).encode("utf-8")


HTML = r"""<!doctype html>
<html lang="zh-CN">
<head>
  <meta charset="utf-8" />
  <meta name="viewport" content="width=device-width, initial-scale=1" />
  <title>钉钉文档同步飞书 | 本地操作台</title>
  <style>
    @import url('https://fonts.googleapis.com/css2?family=IBM+Plex+Mono:wght@400;500;600&family=Noto+Serif+SC:wght@500;700;900&family=Space+Grotesk:wght@400;500;700&display=swap');

    :root {
      --paper: oklch(0.975 0.018 88);
      --paper-strong: oklch(0.94 0.028 83);
      --ink: oklch(0.19 0.035 173);
      --muted: oklch(0.48 0.025 160);
      --line: oklch(0.84 0.026 93);
      --green: oklch(0.42 0.095 166);
      --green-2: oklch(0.55 0.12 160);
      --gold: oklch(0.76 0.115 82);
      --warn: oklch(0.62 0.16 42);
      --danger: oklch(0.55 0.18 24);
      --card: oklch(0.99 0.012 88 / 0.86);
      --shadow: 0 24px 70px oklch(0.23 0.045 165 / 0.12);
      --serif: "Noto Serif SC", serif;
      --sans: "Space Grotesk", "Noto Serif SC", sans-serif;
      --mono: "IBM Plex Mono", monospace;
    }

    * { box-sizing: border-box; }

    body {
      margin: 0;
      min-height: 100vh;
      color: var(--ink);
      font-family: var(--sans);
      background:
        radial-gradient(circle at 15% 4%, oklch(0.84 0.10 92 / 0.38), transparent 30rem),
        radial-gradient(circle at 92% 10%, oklch(0.72 0.08 166 / 0.20), transparent 34rem),
        linear-gradient(135deg, var(--paper), oklch(0.96 0.012 107));
      overflow-x: hidden;
    }

    body::before {
      content: "";
      position: fixed;
      inset: 0;
      pointer-events: none;
      opacity: 0.34;
      background-image:
        linear-gradient(to right, oklch(0.56 0.025 120 / 0.18) 1px, transparent 1px),
        linear-gradient(to bottom, oklch(0.56 0.025 120 / 0.14) 1px, transparent 1px);
      background-size: 44px 44px;
      mask-image: linear-gradient(to bottom, black, transparent 82%);
    }

    button, input { font: inherit; }

    .shell {
      width: min(1440px, calc(100% - 40px));
      margin: 0 auto;
      padding: 28px 0 48px;
      animation: rise 560ms ease-out both;
    }

    .topbar {
      display: flex;
      align-items: center;
      justify-content: space-between;
      gap: 20px;
      padding: 14px 0 30px;
    }

    .brand {
      display: flex;
      align-items: center;
      gap: 14px;
    }

    .mark {
      width: 46px;
      height: 46px;
      border: 1px solid var(--ink);
      border-radius: 16px;
      display: grid;
      place-items: center;
      background: linear-gradient(145deg, var(--ink), oklch(0.28 0.06 165));
      color: var(--paper);
      box-shadow: 8px 8px 0 oklch(0.26 0.04 160 / 0.12);
      font-family: var(--mono);
      font-weight: 600;
      letter-spacing: -0.05em;
    }

    .brand small {
      display: block;
      margin-top: 2px;
      color: var(--muted);
      font-family: var(--mono);
      font-size: 12px;
      letter-spacing: 0.12em;
      text-transform: uppercase;
    }

    .brand strong {
      display: block;
      font-family: var(--serif);
      font-size: clamp(18px, 2vw, 24px);
      letter-spacing: -0.03em;
    }

    .status-pill {
      min-width: 148px;
      border: 1px solid var(--line);
      border-radius: 999px;
      padding: 10px 16px;
      background: oklch(0.99 0.01 90 / 0.7);
      box-shadow: 0 1px 0 white inset;
      display: inline-flex;
      align-items: center;
      justify-content: center;
      gap: 8px;
      color: var(--muted);
      font-size: 13px;
    }

    .dot {
      width: 9px;
      height: 9px;
      border-radius: 999px;
      background: var(--green-2);
      box-shadow: 0 0 0 6px oklch(0.55 0.12 160 / 0.12);
    }

    .dot.running { animation: pulse 900ms ease-in-out infinite; background: var(--gold); }
    .dot.error { background: var(--danger); box-shadow: 0 0 0 6px oklch(0.55 0.18 24 / 0.12); }
    .dot.stopped { background: var(--warn); box-shadow: 0 0 0 6px oklch(0.62 0.16 42 / 0.12); }

    .hero {
      border: 1px solid var(--line);
      border-radius: 30px;
      background: var(--card);
      box-shadow: var(--shadow);
      padding: clamp(18px, 3vw, 34px);
      position: relative;
      overflow: hidden;
    }

    .hero::after {
      content: "";
      position: absolute;
      right: -10%;
      top: -22%;
      width: 46%;
      aspect-ratio: 1;
      border-radius: 50%;
      border: 1px solid oklch(0.43 0.08 168 / 0.25);
      background: radial-gradient(circle, oklch(0.76 0.11 82 / 0.28), transparent 58%);
    }

    .hero-grid {
      position: relative;
      z-index: 1;
      display: grid;
      grid-template-columns: 1fr;
      gap: 18px;
      align-items: stretch;
    }

    .eyebrow {
      margin: 0 0 18px;
      color: var(--green);
      font-family: var(--mono);
      font-size: 12px;
      letter-spacing: 0.16em;
      text-transform: uppercase;
    }

    h1 {
      margin: 0;
      font-family: var(--serif);
      font-size: clamp(42px, 6vw, 84px);
      line-height: 0.95;
      letter-spacing: -0.07em;
      max-width: 760px;
    }

    .hero-copy {
      margin: 22px 0 0;
      max-width: 640px;
      color: var(--muted);
      font-size: clamp(15px, 1.4vw, 18px);
      line-height: 1.75;
    }

    .route {
      display: grid;
      gap: 10px;
      margin-top: 18px;
      font-family: var(--mono);
      font-size: 12px;
      color: var(--muted);
    }

    .route-row {
      display: grid;
      grid-template-columns: 90px 1fr;
      gap: 14px;
      align-items: center;
    }

    .route-label {
      color: var(--ink);
      font-weight: 600;
    }

    .route-value {
      overflow: hidden;
      text-overflow: ellipsis;
      white-space: nowrap;
      border-bottom: 1px dotted oklch(0.48 0.025 160 / 0.35);
      padding-bottom: 5px;
    }

    .module-card {
      align-self: stretch;
      border: 1px solid oklch(0.70 0.055 165 / 0.65);
      border-radius: 28px;
      padding: 24px;
      background:
        linear-gradient(145deg, oklch(0.99 0.012 90 / 0.86), oklch(0.94 0.03 92 / 0.76)),
        linear-gradient(90deg, transparent, oklch(0.75 0.12 83 / 0.18));
    }

    .summary-card {
      min-height: 0;
    }

    .module-head {
      display: flex;
      justify-content: space-between;
      gap: 16px;
      align-items: flex-start;
      border-bottom: 1px solid var(--line);
      padding-bottom: 18px;
    }

    .module-head h2 {
      margin: 0;
      font-family: var(--serif);
      font-size: clamp(24px, 2.6vw, 36px);
      letter-spacing: -0.05em;
    }

    .tag {
      display: inline-flex;
      align-items: center;
      border: 1px solid var(--ink);
      border-radius: 999px;
      padding: 6px 10px;
      font-family: var(--mono);
      font-size: 11px;
      text-transform: uppercase;
      background: var(--ink);
      color: var(--paper);
      white-space: nowrap;
    }

    .metrics {
      display: grid;
      grid-template-columns: repeat(3, 1fr);
      gap: 12px;
      margin-top: 18px;
    }

    .metric {
      border: 1px solid var(--line);
      border-radius: 18px;
      padding: 14px;
      background: oklch(1 0 0 / 0.42);
    }

    .metric span {
      display: block;
      color: var(--muted);
      font-size: 12px;
      margin-bottom: 6px;
    }

    .metric strong {
      font-family: var(--mono);
      font-size: clamp(18px, 2vw, 26px);
    }

    .main-grid {
      display: grid;
      grid-template-columns: minmax(280px, 0.78fr) minmax(420px, 1.22fr);
      gap: 20px;
      margin-top: 18px;
    }

    .panel {
      border: 1px solid var(--line);
      border-radius: 28px;
      background: oklch(0.99 0.012 90 / 0.74);
      box-shadow: 0 14px 40px oklch(0.23 0.045 165 / 0.08);
      padding: 22px;
      min-width: 0;
    }

    .panel h3 {
      margin: 0 0 14px;
      font-family: var(--serif);
      font-size: 24px;
      letter-spacing: -0.04em;
    }

    .module-list {
      display: grid;
      gap: 12px;
    }

    .module-item {
      border: 1px solid var(--line);
      border-radius: 20px;
      padding: 16px;
      background: oklch(1 0 0 / 0.52);
    }

    .module-item.active {
      border-color: oklch(0.45 0.09 164 / 0.72);
      background: linear-gradient(135deg, oklch(0.96 0.036 145 / 0.72), oklch(0.99 0.012 90 / 0.8));
    }

    .module-item.disabled {
      opacity: 0.55;
      filter: grayscale(0.15);
    }

    .module-item strong {
      display: block;
      margin-bottom: 6px;
      font-family: var(--serif);
      font-size: 18px;
    }

    .module-item p {
      margin: 0;
      color: var(--muted);
      font-size: 13px;
      line-height: 1.55;
    }

    .steps {
      display: grid;
      gap: 12px;
      margin-top: 18px;
    }

    .step {
      display: grid;
      grid-template-columns: 24px 1fr;
      gap: 10px;
      align-items: start;
      color: var(--muted);
      font-size: 13px;
      line-height: 1.45;
    }

    .step i {
      width: 22px;
      height: 22px;
      display: grid;
      place-items: center;
      border: 1px solid var(--line);
      border-radius: 999px;
      color: var(--green);
      font-style: normal;
      font-family: var(--mono);
      font-size: 11px;
      background: var(--paper);
    }

    .actions {
      display: grid;
      grid-template-columns: repeat(5, minmax(0, 1fr));
      gap: 12px;
      margin: 18px 0 16px;
    }

    .btn {
      min-height: 54px;
      border: 1px solid var(--ink);
      border-radius: 18px;
      cursor: pointer;
      padding: 13px 14px;
      background: oklch(1 0 0 / 0.64);
      color: var(--ink);
      display: inline-flex;
      align-items: center;
      justify-content: center;
      gap: 8px;
      transition: transform 160ms ease, box-shadow 160ms ease, background 160ms ease;
    }

    .btn:hover {
      transform: translateY(-2px);
      box-shadow: 0 14px 24px oklch(0.23 0.045 165 / 0.12);
    }

    .btn:active { transform: translateY(0) scale(0.99); }
    .btn:disabled { cursor: not-allowed; opacity: 0.54; transform: none; box-shadow: none; }

    .btn.primary {
      background: var(--ink);
      color: var(--paper);
      box-shadow: 7px 7px 0 oklch(0.76 0.115 82 / 0.34);
    }

    .btn.warn {
      border-color: var(--warn);
      color: oklch(0.38 0.11 42);
      background: oklch(0.96 0.055 72 / 0.7);
    }

    .btn.danger {
      border-color: var(--danger);
      color: oklch(0.36 0.13 24);
      background: oklch(0.96 0.035 28 / 0.72);
    }

    .btn.danger:not(:disabled) {
      box-shadow: 7px 7px 0 oklch(0.55 0.18 24 / 0.18);
    }

    .meta-grid {
      display: grid;
      grid-template-columns: repeat(2, minmax(0, 1fr));
      gap: 12px;
      margin-top: 16px;
    }

    .meta {
      border: 1px solid var(--line);
      border-radius: 18px;
      padding: 14px;
      min-width: 0;
      background: oklch(1 0 0 / 0.4);
    }

    .meta span {
      display: block;
      color: var(--muted);
      font-size: 12px;
      margin-bottom: 8px;
    }

    .meta code {
      font-family: var(--mono);
      font-size: 12px;
      overflow: hidden;
      text-overflow: ellipsis;
      white-space: nowrap;
      display: block;
    }

    .meta select {
      width: 100%;
      border: 1px solid var(--line);
      border-radius: 10px;
      background: rgba(255, 255, 255, 0.72);
      color: var(--ink);
      font: inherit;
      min-height: 36px;
      padding: 6px 8px;
    }

    .logbox {
      height: 380px;
      border: 1px solid oklch(0.24 0.04 165 / 0.28);
      border-radius: 22px;
      background:
        linear-gradient(180deg, oklch(0.17 0.035 165), oklch(0.11 0.026 165));
      color: oklch(0.91 0.026 150);
      padding: 18px;
      overflow: auto;
      font-family: var(--mono);
      font-size: 12px;
      line-height: 1.65;
      white-space: pre-wrap;
      box-shadow: 0 12px 34px oklch(0.18 0.04 165 / 0.18) inset;
    }

    .empty-log {
      color: oklch(0.74 0.035 150);
    }

    .footer-note {
      margin: 18px 0 0;
      color: var(--muted);
      font-size: 12px;
      line-height: 1.6;
    }

    @keyframes rise {
      from { opacity: 0; transform: translateY(18px); }
      to { opacity: 1; transform: translateY(0); }
    }

    @keyframes pulse {
      0%, 100% { box-shadow: 0 0 0 5px oklch(0.76 0.115 82 / 0.12); }
      50% { box-shadow: 0 0 0 12px oklch(0.76 0.115 82 / 0.04); }
    }

    @media (max-width: 980px) {
      .hero-grid,
      .main-grid {
        grid-template-columns: 1fr;
      }
      .actions {
        grid-template-columns: repeat(2, minmax(0, 1fr));
      }
    }

    @media (max-width: 620px) {
      .shell { width: min(100% - 24px, 1440px); padding-top: 14px; }
      .topbar { align-items: flex-start; flex-direction: column; }
      .hero { border-radius: 24px; padding: 18px; }
      .metrics,
      .meta-grid,
      .actions {
        grid-template-columns: 1fr;
      }
      .logbox { height: 320px; }
    }
  </style>
</head>
<body>
  <div class="shell">
    <header class="topbar">
      <div class="brand" aria-label="应用名称">
        <div class="mark">D→F</div>
        <div>
          <strong>钉钉文档同步飞书</strong>
          <small>Local Sync Console · Port 8111</small>
        </div>
      </div>
      <div class="status-pill" id="taskPill"><span class="dot" id="taskDot"></span><span id="taskText">待命</span></div>
    </header>

    <main>
      <section class="hero" aria-labelledby="moduleName">
        <div class="hero-grid">
          <aside class="module-card summary-card" aria-label="当前模块摘要">
            <div class="module-head">
              <div>
                <h2 id="moduleName">表格记录同步</h2>
                <p class="footer-note">安全策略：只新增缺失记录；已有相同唯一键时跳过，不覆盖飞书原记录。</p>
              </div>
              <span class="tag">MODULE 01</span>
            </div>
            <div class="metrics">
              <div class="metric"><span>Excel 行数</span><strong id="rowCount">--</strong></div>
              <div class="metric"><span>工作表</span><strong id="sheetName">--</strong></div>
              <div class="metric"><span>同步模式</span><strong id="syncMode">--</strong></div>
            </div>
            <div class="route" aria-label="同步链路">
              <div class="route-row"><span class="route-label">SOURCE</span><span class="route-value" id="dingUrl">读取中...</span></div>
              <div class="route-row"><span class="route-label">DOC</span><span class="route-value" id="dingTitle">读取中...</span></div>
              <div class="route-row"><span class="route-label">TARGET</span><span class="route-value" id="feiUrl">读取中...</span></div>
            </div>
          </aside>
        </div>
      </section>

      <section class="main-grid" aria-label="操作区域">
        <aside class="panel">
          <h3>模块</h3>
          <div class="module-list">
            <article class="module-item active">
              <strong>表格记录同步</strong>
              <p>当前已实现。按配置的唯一键检查飞书多维表，只登记缺失记录。</p>
            </article>
            <article class="module-item disabled">
              <strong>更多表格模块</strong>
              <p>待接入。适合做矩阵表展开、排班结构化或多工作表同步。</p>
            </article>
            <article class="module-item disabled">
              <strong>评分/补字段模块</strong>
              <p>待接入。适合处理多层表头、评分维度或补字段场景。</p>
            </article>
          </div>

          <div class="steps" aria-label="流程步骤">
            <div class="step"><i>1</i><span>读取本机钉钉登录态，必要时打开浏览器导出 Excel。</span></div>
            <div class="step"><i>2</i><span>解析配置的工作表，修正钉钉导出维度异常。</span></div>
            <div class="step"><i>3</i><span>按唯一键字段匹配飞书记录；已存在就跳过，不重复登记。</span></div>
            <div class="step"><i>4</i><span>保留目标表公式、AI 字段、关联字段等后续加工字段。</span></div>
          </div>
        </aside>

        <section class="panel">
          <h3>操作</h3>
          <div class="meta-grid" aria-label="链路选择" style="margin-bottom: 14px;">
            <label class="meta" for="moduleSelect"><span>当前链路</span><select id="moduleSelect"></select></label>
            <div class="meta"><span>配置文件</span><code id="envPath">--</code></div>
          </div>
          <div class="actions">
            <button class="btn primary" data-action="full">下载并同步</button>
            <button class="btn" data-action="sync-local">仅同步本地 Excel</button>
            <button class="btn" data-action="dry-run">预览解析</button>
            <button class="btn warn" data-action="download-only">只下载 Excel</button>
            <button class="btn danger" id="stopTask" type="button" disabled>停止任务</button>
          </div>

          <div class="meta-grid" aria-label="配置摘要">
            <div class="meta"><span>本地 Excel</span><code id="excelPath">--</code></div>
            <div class="meta"><span>字段映射</span><code id="mappingPath">--</code></div>
            <div class="meta"><span>飞书应用</span><code id="appId">--</code></div>
            <div class="meta"><span>Base Token</span><code id="appToken">--</code></div>
          </div>

          <p class="footer-note">
            如果选择“下载并同步”，钉钉会使用独立的自动化浏览器；若出现登录页，请在该窗口完成一次登录，随后会重新打开配置的目标文档并继续导出。
          </p>

          <h3 style="margin-top: 22px;">运行日志</h3>
          <div class="logbox" id="logs"><span class="empty-log">还没有任务日志。选择上方按钮开始。</span></div>
        </section>
      </section>
    </main>
  </div>

  <script>
    const $ = (id) => document.getElementById(id);
    const buttons = [...document.querySelectorAll('[data-action]')];
    const stopButton = $('stopTask');
    const moduleSelect = $('moduleSelect');

    function taskLabel(status) {
      return {
        idle: '待命',
        running: '运行中',
        stopping: '停止中',
        stopped: '已停止',
        success: '完成',
        error: '失败'
      }[status] || status;
    }

    function renderTask(task) {
      const dot = $('taskDot');
      dot.className = 'dot';
      if (task.status === 'running') dot.classList.add('running');
      if (task.status === 'stopping') dot.classList.add('running');
      if (task.status === 'stopped') dot.classList.add('stopped');
      if (task.status === 'error') dot.classList.add('error');
      $('taskText').textContent = taskLabel(task.status);
      const isActive = task.status === 'running' || task.status === 'stopping';
      buttons.forEach((btn) => btn.disabled = isActive);
      stopButton.disabled = task.status !== 'running' || task.can_stop === false;
      const lines = task.logs || [];
      $('logs').textContent = lines.length ? lines.join('\n') : '还没有任务日志。选择上方按钮开始。';
      $('logs').scrollTop = $('logs').scrollHeight;
    }

    function renderStatus(data) {
      if (moduleSelect && data.modules) {
        const selected = data.module.id;
        if (!moduleSelect.options.length || moduleSelect.dataset.loaded !== '1') {
          moduleSelect.innerHTML = '';
          data.modules.forEach((item) => {
            const option = document.createElement('option');
            option.value = item.id;
            option.textContent = item.name;
            moduleSelect.appendChild(option);
          });
          moduleSelect.dataset.loaded = '1';
        }
        moduleSelect.value = selected;
      }
      $('dingUrl').textContent = data.targets.dingtalk_doc_url || '--';
      $('dingTitle').textContent = data.targets.dingtalk_doc_title || '--';
      $('feiUrl').textContent = data.targets.feishu_bitable_url || '--';
      $('moduleName').textContent = data.module.name || '表格记录同步';
      $('envPath').textContent = data.module.env || '--';
      $('rowCount').textContent = data.excel.row_count ?? '--';
      $('sheetName').textContent = data.module.sheet || '--';
      $('syncMode').textContent = data.module.mode || '--';
      $('excelPath').textContent = data.excel.path || '--';
      $('mappingPath').textContent = data.module.mapping || '--';
      $('appId').textContent = data.targets.feishu_app_id || '--';
      $('appToken').textContent = data.targets.feishu_app_token || '--';
      renderTask(data.task || {});
    }

    async function refreshStatus() {
      const moduleId = moduleSelect?.value || '';
      const response = await fetch('/api/status?module=' + encodeURIComponent(moduleId), { cache: 'no-store' });
      renderStatus(await response.json());
    }

    async function refreshTask() {
      const response = await fetch('/api/task', { cache: 'no-store' });
      renderTask(await response.json());
    }

    async function runAction(action) {
      const response = await fetch('/api/run', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({ action, module: moduleSelect?.value || undefined })
      });
      if (!response.ok) {
        const error = await response.json().catch(() => ({ error: response.statusText }));
        alert(error.error || '启动任务失败');
        return;
      }
      renderTask(await response.json());
    }

    async function stopCurrentTask() {
      stopButton.disabled = true;
      const response = await fetch('/api/stop', { method: 'POST' });
      if (!response.ok) {
        const error = await response.json().catch(() => ({ error: response.statusText }));
        alert(error.error || '停止任务失败');
        return;
      }
      renderTask(await response.json());
    }

    buttons.forEach((button) => {
      button.addEventListener('click', () => runAction(button.dataset.action));
    });
    stopButton.addEventListener('click', stopCurrentTask);
    moduleSelect?.addEventListener('change', () => refreshStatus().catch(console.error));

    refreshStatus().catch(console.error);
    setInterval(refreshTask, 1400);
    setInterval(refreshStatus, 10000);
  </script>
</body>
</html>
"""


class DashboardHandler(BaseHTTPRequestHandler):
    server_version = "DingFeiDashboard/1.0"

    def log_message(self, format: str, *args: Any) -> None:  # Keep console quiet.
        return

    def send_json(self, payload: Any, status: HTTPStatus = HTTPStatus.OK) -> None:
        body = json_bytes(payload)
        self.send_response(status)
        self.send_header("Content-Type", "application/json; charset=utf-8")
        self.send_header("Content-Length", str(len(body)))
        self.end_headers()
        self.wfile.write(body)

    def send_html(self) -> None:
        body = HTML.encode("utf-8")
        self.send_response(HTTPStatus.OK)
        self.send_header("Content-Type", "text/html; charset=utf-8")
        self.send_header("Content-Length", str(len(body)))
        self.end_headers()
        self.wfile.write(body)

    def read_body_json(self) -> dict[str, Any]:
        length = int(self.headers.get("Content-Length", "0"))
        if not length:
            return {}
        raw = self.rfile.read(length)
        return json.loads(raw.decode("utf-8"))

    def do_GET(self) -> None:
        parsed = urlparse(self.path)
        path = parsed.path
        query = parse_qs(parsed.query)
        try:
            if path == "/":
                self.send_html()
            elif path == "/api/status":
                module_id = (query.get("module") or [""])[0] or None
                self.send_json(read_status(module_id))
            elif path == "/api/task":
                self.send_json(state_snapshot())
            else:
                self.send_json({"error": "Not found"}, HTTPStatus.NOT_FOUND)
        except SyncError as exc:
            self.send_json({"error": str(exc)}, HTTPStatus.BAD_REQUEST)
        except Exception as exc:
            self.send_json({"error": str(exc)}, HTTPStatus.INTERNAL_SERVER_ERROR)

    def do_POST(self) -> None:
        path = urlparse(self.path).path
        try:
            if path == "/api/run":
                payload = self.read_body_json()
                action = str(payload.get("action", "sync-local"))
                module_id = str(payload.get("module") or "") or None
                self.send_json(start_task(action, module_id))
                return
            if path == "/api/stop":
                self.send_json(stop_task())
                return
            self.send_json({"error": "Not found"}, HTTPStatus.NOT_FOUND)
        except ValueError as exc:
            self.send_json({"error": str(exc)}, HTTPStatus.BAD_REQUEST)
        except RuntimeError as exc:
            self.send_json({"error": str(exc)}, HTTPStatus.CONFLICT)
        except Exception as exc:
            self.send_json({"error": str(exc)}, HTTPStatus.INTERNAL_SERVER_ERROR)


def parse_cli() -> argparse.Namespace:
    parser = argparse.ArgumentParser(description="Local dashboard for DingTalk to Feishu sync.")
    parser.add_argument("--host", default="127.0.0.1")
    parser.add_argument("--port", type=int, default=DEFAULT_PORT)
    return parser.parse_args()


def main() -> int:
    args = parse_cli()
    server = ThreadingHTTPServer((args.host, args.port), DashboardHandler)
    print(f"Dashboard running at http://{args.host}:{args.port}")
    try:
        server.serve_forever()
    except KeyboardInterrupt:
        print("\nDashboard stopped.")
    finally:
        server.server_close()
    return 0


if __name__ == "__main__":
    raise SystemExit(main())

