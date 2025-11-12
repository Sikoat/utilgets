"""
Launcher script for util.exe during the transition.

- Can be started with pythonw.exe (no console) or python.exe from a console.
- Starts util.exe located in the same folder as this script and exits immediately.
- If an error occurs and no console is present, shows a Windows message box.
"""
from __future__ import annotations

import os
import sys
import ctypes


def _has_console() -> bool:
    """Return True if this process has a console window attached."""
    try:
        return bool(ctypes.windll.kernel32.GetConsoleWindow())  # type: ignore[attr-defined]
    except Exception:
        # Best-effort; if anything goes wrong, assume have a console
        return True


def _notify_error(message: str) -> None:
    """Report an error via stderr if a console exists, else show a message box."""
    if _has_console():
        try:
            print(message, file=sys.stderr, flush=True)
        except Exception:
            pass
    else:
        try:
            # MB_ICONERROR (0x10) | MB_OK (0x0)
            ctypes.windll.user32.MessageBoxW(  # type: ignore[attr-defined]
                None, message, "util_temp", 0x10
            )
        except Exception:
            # Last resort: ignore, there's nowhere to display it
            pass


def main() -> int:
    script_dir = os.path.dirname(os.path.abspath(__file__))
    exe_path = os.path.join(script_dir, "util.exe")

    if not os.path.isfile(exe_path):
        _notify_error(f"util.exe was not found next to this script:\n{exe_path}")
        return 1

    try:
        # Match typical .bat behavior: run from the script's directory
        os.chdir(script_dir)

        # Prefer ShellExecute-style launch which returns immediately and lets
        # Windows decide the appropriate subsystem/console behavior.
        try:
            os.startfile(exe_path)  # type: ignore[attr-defined]
        except AttributeError:
            # Fallback if os.startfile isn't available (non-Windows Python)
            import subprocess

            subprocess.Popen([exe_path], cwd=script_dir)
    except Exception as exc:
        _notify_error(f"Failed to launch util.exe:\n{exc}")
        return 1

    # Exit immediately after spawning util.exe
    return 0


if __name__ == "__main__":
    sys.exit(main())
