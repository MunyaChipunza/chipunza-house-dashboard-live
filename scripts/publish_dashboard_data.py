from __future__ import annotations

import argparse
import os
import shutil
import subprocess
import sys
from pathlib import Path


SCRIPT_DIR = Path(__file__).resolve().parent
BUNDLE_DIR = SCRIPT_DIR.parent
sys.path.insert(0, str(SCRIPT_DIR))

from refresh_dashboard_data import DEFAULT_OUTPUT, DEFAULT_WORKBOOK, refresh_dashboard_data  # noqa: E402


CREATE_NO_WINDOW = getattr(subprocess, "CREATE_NO_WINDOW", 0)


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(description="Refresh dashboard_data.json and optionally push it to GitHub.")
    parser.add_argument("--workbook", help="Optional local workbook path.")
    parser.add_argument("--workbook-url", help="Optional public workbook URL for cloud refreshes.")
    parser.add_argument("--output", default=str(DEFAULT_OUTPUT), help="Output path for dashboard JSON.")
    parser.add_argument("--commit-message", default="Refresh Chipunza dashboard data", help="Git commit message.")
    return parser.parse_args()


def git_executable() -> str:
    path = shutil.which("git")
    if path:
        return path
    for candidate in (Path("C:/Program Files/Git/cmd/git.exe"), Path("C:/Program Files/Git/bin/git.exe")):
        if candidate.exists():
            return str(candidate)
    raise FileNotFoundError("Git executable not found.")


def run_git(*args: str, check: bool = True) -> subprocess.CompletedProcess[str]:
    startupinfo = None
    if os.name == "nt":
        startupinfo = subprocess.STARTUPINFO()
        startupinfo.dwFlags |= subprocess.STARTF_USESHOWWINDOW
        startupinfo.wShowWindow = 0
    return subprocess.run(
        [git_executable(), "-C", str(BUNDLE_DIR), *args],
        check=check,
        text=True,
        capture_output=True,
        creationflags=CREATE_NO_WINDOW,
        startupinfo=startupinfo,
    )


def find_default_workbook() -> Path | None:
    return DEFAULT_WORKBOOK if DEFAULT_WORKBOOK.exists() else None


def is_git_repo() -> bool:
    result = run_git("rev-parse", "--is-inside-work-tree", check=False)
    return result.returncode == 0 and result.stdout.strip() == "true"


def has_origin() -> bool:
    result = run_git("remote", "get-url", "origin", check=False)
    return result.returncode == 0 and bool(result.stdout.strip())


def ensure_identity() -> None:
    name = run_git("config", "--get", "user.name", check=False)
    email = run_git("config", "--get", "user.email", check=False)
    if name.returncode == 0 and email.returncode == 0 and name.stdout.strip() and email.stdout.strip():
        return
    run_git("config", "user.name", "Chipunza Dashboard Sync")
    run_git("config", "user.email", "chipunza-dashboard-sync@local")


def sync_repo() -> None:
    fetch = run_git("fetch", "origin", "main", check=False)
    if fetch.returncode != 0:
        raise RuntimeError(fetch.stderr.strip() or "Could not fetch origin/main.")
    rebase = run_git("rebase", "origin/main", check=False)
    if rebase.returncode != 0:
        raise RuntimeError(rebase.stderr.strip() or rebase.stdout.strip() or "Could not rebase onto origin/main.")


def has_changes(output_path: Path) -> bool:
    rel = output_path.relative_to(BUNDLE_DIR)
    status = run_git("status", "--short", "--", str(rel), check=False)
    return bool(status.stdout.strip())


def push_dashboard(workbook_path: Path | None, workbook_url: str | None, output_path: Path, commit_message: str) -> bool:
    refresh_dashboard_data(workbook=str(workbook_path) if workbook_path else None, workbook_url=workbook_url, output=output_path)
    if not is_git_repo():
        print("Dashboard data refreshed locally. No Git repository detected.")
        return False
    if not has_origin():
        print("Dashboard data refreshed locally. No origin remote is configured.")
        return False
    sync_repo()
    refresh_dashboard_data(workbook=str(workbook_path) if workbook_path else None, workbook_url=workbook_url, output=output_path)
    if not has_changes(output_path):
        print("Dashboard data is already up to date.")
        return False
    ensure_identity()
    rel = output_path.relative_to(BUNDLE_DIR)
    run_git("add", "--", str(rel))
    run_git("commit", "-m", commit_message)
    run_git("push", "-u", "origin", "main")
    print("Dashboard data refreshed and pushed.")
    return True


def main() -> None:
    args = parse_args()
    workbook_path = Path(args.workbook).expanduser().resolve() if args.workbook else find_default_workbook()
    output_path = Path(args.output).expanduser()
    if not output_path.is_absolute():
        output_path = (BUNDLE_DIR / output_path).resolve()
    push_dashboard(workbook_path=workbook_path, workbook_url=args.workbook_url, output_path=output_path, commit_message=args.commit_message)


if __name__ == "__main__":
    main()
