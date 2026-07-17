"""Resolve the app version for display, kept in sync with GitHub Releases.

The build workflow writes a `VERSION` file (the git tag it just built,
e.g. "v1.3") next to the executable, bundled the same way as the docx
templates. Running from source has no such file, so fall back to `git
describe`, and finally to "dev" if this isn't even a git checkout.
"""
import os
import subprocess
import sys


def resource_path(rel):
    base = getattr(sys, "_MEIPASS", os.path.dirname(os.path.abspath(__file__)))
    return os.path.join(base, rel)


def app_version():
    bundled = resource_path("VERSION")
    if os.path.exists(bundled):
        try:
            with open(bundled, encoding="utf-8") as fh:
                text = fh.read().strip()
            if text:
                return text
        except OSError:
            pass
    try:
        out = subprocess.run(
            ["git", "describe", "--tags", "--always"],
            cwd=os.path.dirname(os.path.abspath(__file__)),
            capture_output=True, text=True, timeout=2)
        text = out.stdout.strip()
        if out.returncode == 0 and text:
            return text
    except (OSError, subprocess.SubprocessError):
        pass
    return "dev"
