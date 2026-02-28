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
}
BUNDLE_MODES = ("auto", "onefile", "onedir")


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


def _resolve_bundle_mode(bundle_mode: str) -> str:
    if bundle_mode != "auto":
        return bundle_mode
    # Prefer .app UX on macOS by default (double-click without Terminal popup).
    if sys.platform.startswith("darwin"):
        return "onedir"
    return "onefile"


def _zip_path(source: Path, target_zip: Path) -> None:
    target_zip.parent.mkdir(parents=True, exist_ok=True)
    if target_zip.exists():
        target_zip.unlink()

    # macOS .app bundles rely on symlink-heavy .framework layouts.
    # Use `ditto` to preserve Finder-compatible bundle structure.
    if sys.platform.startswith("darwin"):
        subprocess.check_call(
            [
                "ditto",
                "-c",
                "-k",
                "--sequesterRsrc",
                "--keepParent",
                str(source),
                str(target_zip),
            ]
        )
        return

    with zipfile.ZipFile(target_zip, "w", compression=zipfile.ZIP_DEFLATED) as zf:
        if source.is_file():
            zf.write(source, source.name)
            return

        for item in sorted(source.rglob("*")):
            if item.is_file():
                arcname = str(Path(source.name) / item.relative_to(source))
                zf.write(item, arcname)


def _resolve_bundle_output(dist_root: Path, bundle_name: str) -> Path:
    # On macOS onedir+windowed can produce both "{name}" and "{name}.app".
    # Prefer the GUI bundle when available.
    if sys.platform.startswith("darwin"):
        candidates = [
            dist_root / f"{bundle_name}.app",
            dist_root / bundle_name,
            dist_root / f"{bundle_name}.exe",
        ]
    else:
        candidates = [
            dist_root / bundle_name,
            dist_root / f"{bundle_name}.exe",
            dist_root / f"{bundle_name}.app",
        ]
    for candidate in candidates:
        if candidate.exists():
            return candidate
    raise FileNotFoundError(
        "PyInstaller output not found: "
        + " / ".join(str(candidate) for candidate in candidates)
    )


def package_app(
    app: str,
    root: Path,
    output_dir: Path,
    version: str,
    bundle_mode: str = "auto",
) -> Path:
    if app not in APP_CONFIG:
        raise ValueError(f"Unsupported app: {app}")
    if bundle_mode not in BUNDLE_MODES:
        raise ValueError(f"Unsupported bundle mode: {bundle_mode}")
    effective_bundle_mode = _resolve_bundle_mode(bundle_mode)

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

    # macOS onefile + --windowed is deprecated (PyInstaller 7.0 will block it).
    # Keep onefile as console-style binary on macOS; use --windowed for onedir to
    # produce a Finder-friendly .app bundle.
    use_windowed = not (
        sys.platform.startswith("darwin") and effective_bundle_mode == "onefile"
    )

    cmd = [
        sys.executable,
        "-m",
        "PyInstaller",
        "--noconfirm",
        "--clean",
        *(["--windowed"] if use_windowed else []),
        f"--{effective_bundle_mode}",
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

    source = _resolve_bundle_output(dist_root, bundle_name)

    artifact_name = f"{app}-{_platform_label()}-{_arch_label()}-{version}.zip"
    artifact_path = output_dir / artifact_name
    _zip_path(source, artifact_path)
    return artifact_path


def main() -> int:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--app", required=True, choices=sorted(APP_CONFIG))
    parser.add_argument("--version", default="dev")
    parser.add_argument("--output-dir", default="artifacts")
    parser.add_argument(
        "--bundle-mode",
        choices=BUNDLE_MODES,
        default="auto",
        help=(
            "PyInstaller bundle mode (default: auto = onedir on macOS, onefile on "
            "other platforms)."
        ),
    )
    args = parser.parse_args()

    root = Path(__file__).resolve().parent.parent
    output_dir = root / args.output_dir
    artifact = package_app(
        args.app,
        root=root,
        output_dir=output_dir,
        version=args.version,
        bundle_mode=args.bundle_mode,
    )
    print(artifact)
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
