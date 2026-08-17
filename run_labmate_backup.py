#!/usr/bin/env python3
"""Python launcher for the validated Labmate PowerShell automation.

Shows every PowerShell log line live, saves the same output to a launcher log,
and requests Windows administrator elevation for restore mode.
"""

from __future__ import annotations

import argparse
import ctypes
import json
import os
from pathlib import Path
import subprocess
import sys
from datetime import datetime


def is_admin() -> bool:
    try:
        return bool(ctypes.windll.shell32.IsUserAnAdmin())
    except Exception:
        return False


def quote_windows_arg(value: str) -> str:
    if not value or any(ch in value for ch in ' \t"'):
        return '"' + value.replace('"', '\\"') + '"'
    return value


def request_elevation(argv: list[str]) -> int:
    parameters = " ".join(quote_windows_arg(arg) for arg in argv)
    result = ctypes.windll.shell32.ShellExecuteW(
        None, "runas", sys.executable, parameters, None, 1
    )
    if result <= 32:
        raise RuntimeError(f"Administrator elevation failed (ShellExecute code {result}).")
    return 0


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(
        description="Restore the newest Labmate .bak and compare SQL Server with the configured target MySQL."
    )
    parser.add_argument("--compare-only", action="store_true", help="Skip SQL restore.")
    parser.add_argument("--force", action="store_true", help="Skip typed RESTORE confirmation.")
    parser.add_argument(
        "--only-table",
        action="append",
        default=[],
        help="Compare only one configured table; repeat for more tables.",
    )
    parser.add_argument("--no-pause", action="store_true", help="Do not pause before closing.")
    parser.add_argument("--elevated", action="store_true", help=argparse.SUPPRESS)
    return parser.parse_args()


def main() -> int:
    args = parse_args()
    base_dir = Path(__file__).resolve().parent
    ps_script = base_dir / "Run-LabmateBackup.ps1"
    config_path = base_dir / "labmate-config.json"
    if not ps_script.exists() or not config_path.exists():
        raise FileNotFoundError("PowerShell engine or configuration file is missing.")

    if not args.compare_only and not is_admin():
        if args.elevated:
            raise PermissionError("Administrator elevation was requested but is not active.")
        elevated_args = [str(Path(__file__).resolve()), *sys.argv[1:], "--elevated"]
        print("Administrator permission is required for SQL restore. Opening UAC prompt...", flush=True)
        return request_elevation(elevated_args)

    with config_path.open("r", encoding="utf-8") as handle:
        config = json.load(handle)
    report_dir = Path(config["ReportFolder"])
    report_dir.mkdir(parents=True, exist_ok=True)
    launcher_log = report_dir / f"Python_Launcher_{datetime.now():%Y%m%d_%H%M%S}.log"

    command = [
        r"C:\Windows\System32\WindowsPowerShell\v1.0\powershell.exe",
        "-NoProfile",
        "-ExecutionPolicy",
        "Bypass",
        "-File",
        str(ps_script),
        "-ConfigPath",
        str(config_path),
    ]
    if args.compare_only:
        command.append("-CompareOnly")
    if args.force:
        command.append("-Force")
    if args.only_table:
        command.extend(["-OnlyTable", ",".join(args.only_table)])

    print("=" * 78)
    print("LABMATE BACKUP AUTOMATION - PYTHON LIVE LOG")
    print(f"Mode: {'COMPARE ONLY' if args.compare_only else 'RESTORE + COMPARE'}")
    print(f"Launcher log: {launcher_log}")
    print("=" * 78, flush=True)

    with launcher_log.open("w", encoding="utf-8", newline="") as log_handle:
        log_handle.write(f"Command: {subprocess.list2cmdline(command)}\n\n")
        log_handle.flush()
        child_env = os.environ.copy()
        # Codex's bundled runtime can inject a PowerShell module path that is
        # incompatible with Windows PowerShell. Let Windows PowerShell rebuild
        # its native module path so DPAPI credential decryption works.
        child_env.pop("PSModulePath", None)
        child_env.pop("PSMODULEPATH", None)
        process = subprocess.Popen(
            command,
            cwd=str(base_dir),
            env=child_env,
            stdin=None,
            stdout=subprocess.PIPE,
            stderr=subprocess.STDOUT,
            text=True,
            encoding="utf-8",
            errors="replace",
            bufsize=1,
        )
        assert process.stdout is not None
        for line in process.stdout:
            print(line, end="", flush=True)
            log_handle.write(line)
            log_handle.flush()
        return_code = process.wait()

    print("=" * 78)
    if return_code == 0:
        print("SUCCESS - automation completed.")
    else:
        print(f"FAILED - automation exited with code {return_code}.")
    print(f"Saved launcher log: {launcher_log}")
    print("=" * 78, flush=True)

    if not args.no_pause:
        try:
            input("Press Enter to close...")
        except EOFError:
            pass
    return return_code


if __name__ == "__main__":
    try:
        raise SystemExit(main())
    except KeyboardInterrupt:
        print("\nCancelled by user.", flush=True)
        raise SystemExit(130)
    except Exception as exc:
        print(f"ERROR: {exc}", flush=True)
        try:
            input("Press Enter to close...")
        except EOFError:
            pass
        raise SystemExit(1)
