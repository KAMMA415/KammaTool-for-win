"""Build an architecture-specific portable Windows ZIP."""

import re
import subprocess
import sys
import sysconfig
from pathlib import Path
from tempfile import TemporaryDirectory
from textwrap import dedent
from zipfile import ZIP_DEFLATED, ZipFile

from version import APP_VERSION

ROOT = Path(__file__).resolve().parent
ASSETS = (
    "Dongdong tool 2.ico",
    "icon_rename.png",
    "icon_pdf.png",
    "icon_merge.png",
    "icon_organize.png",
)


def main():
    architecture = {"win-amd64": "x64", "win-arm64": "arm64"}.get(sysconfig.get_platform())
    if architecture is None:
        raise SystemExit("Use an x64 or ARM64 Windows Python interpreter.")
    if not re.fullmatch(r"\d+\.\d+\.\d+(?:-(?:alpha|beta|rc)\.\d+)?", APP_VERSION):
        raise SystemExit("Set a release version such as 2.1.0 or 2.2.0-beta.1 in version.py.")

    build_root = (ROOT / "build").resolve()
    dist_dir = (ROOT / "dist").resolve()
    if build_root.parent != ROOT or dist_dir.parent != ROOT:
        raise SystemExit("The build and dist directories must be inside the project.")
    build_root.mkdir(exist_ok=True)
    dist_dir.mkdir(exist_ok=True)
    archive = dist_dir / f"KammaTool-v{APP_VERSION}-windows-{architecture}-portable.zip"

    with TemporaryDirectory(
        prefix=f"windows-{architecture}-", dir=build_root, ignore_cleanup_errors=True
    ) as temporary_dir:
        build_dir = Path(temporary_dir)
        bundle_dir = build_dir / "bundle" / "KammaTool"
        diagnostics_hook = build_dir / "smoke_diagnostics.py"
        diagnostics_hook.write_text(dedent("""\
            import faulthandler
            import sys

            if sys.argv[1:] == ["--smoke-test"]:
                sys.stdout = sys.stderr = open("smoke-test.log", "w", encoding="utf-8", buffering=1)
                print("Starting packaged Python runtime.")
                faulthandler.dump_traceback_later(25, file=sys.stderr)
            """), encoding="utf-8")
        command = [
            sys.executable, "-m", "PyInstaller", "--windowed", "--onedir",
            "--name", "KammaTool", "--icon", str(ROOT / ASSETS[0]),
            "--collect-data", "ttkbootstrap",
            "--runtime-hook", str(diagnostics_hook),
            "--distpath", str(bundle_dir.parent),
            "--workpath", str(build_dir / "work"),
            "--specpath", str(build_dir),
        ]
        for asset in ASSETS:
            command.extend(("--add-data", f"{ROOT / asset}:."))
        command.append(str(ROOT / "main.py"))
        subprocess.run(command, cwd=ROOT, check=True)
        executable = bundle_dir / "KammaTool.exe"
        if not executable.is_file():
            raise SystemExit("PyInstaller did not produce KammaTool.exe.")
        smoke_log = bundle_dir / "smoke-test.log"
        try:
            subprocess.run([str(executable), "--smoke-test"], cwd=bundle_dir, check=True, timeout=45)
        except (subprocess.CalledProcessError, subprocess.TimeoutExpired):
            if smoke_log.is_file():
                print(smoke_log.read_text(encoding="utf-8", errors="replace"))
            raise
        finally:
            smoke_log.unlink(missing_ok=True)
        print(f"Packaged startup and file rename passed ({architecture}).")

        temporary_archive = build_dir / archive.name
        with ZipFile(temporary_archive, "w", compression=ZIP_DEFLATED) as output:
            for file in sorted(bundle_dir.rglob("*")):
                if file.is_file():
                    output.write(file, file.relative_to(bundle_dir.parent))
            output.write(ROOT / "USAGE.md", "KammaTool/USAGE.md")
        temporary_archive.replace(archive)
    print(archive)


if __name__ == "__main__":
    main()
