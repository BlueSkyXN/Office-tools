#!/usr/bin/env python3
"""Package a selected docx-toolbox desktop app with PyInstaller."""

from __future__ import annotations

import argparse
import platform
import shutil
import subprocess
import sys
import zipfile
from pathlib import Path


APP_CONFIG: dict[str, dict[str, object]] = {
    "pyside6": {
        "entry": Path("pyside6/app/main.py"),
        "hidden_imports": [],
    },
    "pyqt6": {
        "entry": Path("pyqt6/app/main.py"),
        "hidden_imports": [],
    },
    "tk": {
        "entry": Path("tk/app/main.py"),
        "hidden_imports": [],
    },
    "flet": {
        "entry": Path("flet/app/main.py"),
        "hidden_imports": ["flet"],
    },
    "pywebview": {
        "entry": Path("pywebview/backend/app.py"),
        "hidden_imports": ["webview"],
    },
}


def _data_separator() -> str:
    return ";" if sys.platform.startswith("win") else ":"


def _platform_label() -> str:
    if sys.platform.startswith("darwin"):
        return "macos"
    if sys.platform.startswith("win"):
        return "windows"
    return "linux"


def _arch_label() -> str:
    machine = platform.machine().lower()
    if machine in {"x86_64", "amd64"}:
        return "x64"
    if machine in {"arm64", "aarch64"}:
        return "arm64"
    return machine or "unknown"


def _zip_path(source: Path, target_zip: Path) -> None:
    target_zip.parent.mkdir(parents=True, exist_ok=True)
    if target_zip.exists():
        target_zip.unlink()

    with zipfile.ZipFile(target_zip, "w", compression=zipfile.ZIP_DEFLATED) as zf:
        if source.is_file():
            zf.write(source, source.name)
            return

        for item in sorted(source.rglob("*")):
            if item.is_file():
                arcname = str(Path(source.name) / item.relative_to(source))
                zf.write(item, arcname)


def package_app(app: str, root: Path, output_dir: Path, version: str) -> Path:
    if app not in APP_CONFIG:
        raise ValueError(f"Unsupported app: {app}")

    config = APP_CONFIG[app]
    entry = root / config["entry"]
    if not entry.is_file():
        raise FileNotFoundError(f"Entry script not found: {entry}")

    sep = _data_separator()
    bundle_name = f"docx-toolbox-{app}"

    build_root = root / "build" / "pyinstaller" / app
    dist_root = root / "dist" / "pyinstaller" / app
    spec_root = root / "build" / "spec" / app
    shutil.rmtree(build_root, ignore_errors=True)
    shutil.rmtree(dist_root, ignore_errors=True)
    spec_root.mkdir(parents=True, exist_ok=True)
    output_dir.mkdir(parents=True, exist_ok=True)

    add_data: list[tuple[Path, str]] = [
        (root / "references", "references"),
    ]
    if app == "pywebview":
        frontend_dist = root / "pywebview" / "frontend" / "dist"
        index_html = frontend_dist / "index.html"
        if not index_html.is_file():
            raise FileNotFoundError(
                "pywebview frontend build not found. Run `npm ci && npm run build` in "
                "pywebview/frontend first."
            )
        add_data.append((frontend_dist, "pywebview/frontend/dist"))

    cmd = [
        sys.executable,
        "-m",
        "PyInstaller",
        "--noconfirm",
        "--clean",
        "--windowed",
        "--name",
        bundle_name,
        "--paths",
        str(root),
        "--workpath",
        str(build_root),
        "--distpath",
        str(dist_root),
        "--specpath",
        str(spec_root),
    ]

    for src, dest in add_data:
        cmd.extend(["--add-data", f"{src}{sep}{dest}"])

    for hidden in config["hidden_imports"]:
        cmd.extend(["--hidden-import", str(hidden)])

    cmd.append(str(entry))
    subprocess.check_call(cmd, cwd=root)

    bundle_dir = dist_root / bundle_name
    bundle_file = dist_root / f"{bundle_name}.exe"
    if bundle_dir.exists():
        source = bundle_dir
    elif bundle_file.exists():
        source = bundle_file
    else:
        raise FileNotFoundError(
            f"PyInstaller output not found for {app}: {bundle_dir} / {bundle_file}"
        )

    artifact_name = f"{app}-{_platform_label()}-{_arch_label()}-{version}.zip"
    artifact_path = output_dir / artifact_name
    _zip_path(source, artifact_path)
    return artifact_path


def main() -> int:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--app", required=True, choices=sorted(APP_CONFIG))
    parser.add_argument("--version", default="dev")
    parser.add_argument("--output-dir", default="artifacts")
    args = parser.parse_args()

    root = Path(__file__).resolve().parent.parent
    output_dir = root / args.output_dir
    artifact = package_app(args.app, root=root, output_dir=output_dir, version=args.version)
    print(artifact)
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
