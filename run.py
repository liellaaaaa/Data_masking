#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Data Masking Tool - Python Launcher
Double-click this file to run. It will:
  1. Find a Python interpreter
  2. Create a local .venv on first run
  3. Auto-install dependencies if missing
  4. Launch the Streamlit app
Any error keeps the window open so you can read it.
"""
import os
import sys
import subprocess

BASE = os.path.dirname(os.path.abspath(__file__))
LOG = os.path.join(BASE, "launch.log")
VENV_PY = os.path.join(BASE, ".venv", "Scripts", "python.exe")
REQ = os.path.join(BASE, "requirements.txt")
APP = os.path.join(BASE, "mask_tool.py")


def log(msg):
    try:
        with open(LOG, "a", encoding="utf-8") as f:
            f.write(msg + "\n")
    except Exception:
        pass
    print(msg)


def find_python():
    """Return a usable base python executable path."""
    candidates = []
    # sys.executable is the python running this launcher
    if sys.version_info >= (3, 9):
        candidates.append(sys.executable)
    # try the 'py' launcher
    try:
        r = subprocess.run(["py", "-3", "--version"], capture_output=True, text=True)
        if r.returncode == 0:
            candidates.insert(0, "py -3")
    except Exception:
        pass
    # try bare 'python'
    try:
        r = subprocess.run(["python", "--version"], capture_output=True, text=True)
        if r.returncode == 0:
            candidates.append("python")
    except Exception:
        pass
    for c in candidates:
        return c
    return None


def run(cmd):
    log("[CMD] " + " ".join(cmd) if isinstance(cmd, list) else cmd)
    return subprocess.run(cmd, shell=isinstance(cmd, str))


def main():
    log("=" * 50)
    log("Data Masking Tool - Python Launcher")
    log("=" * 50)

    py = find_python()
    if not py:
        log("[ERROR] Python not found. Install Python 3.9+ and check 'Add to PATH'.")
        log("        Download: https://www.python.org/downloads/")
        return

    log("[OK] Python: " + py)

    # 1) Create venv if missing
    if not os.path.exists(VENV_PY):
        log("[INFO] First run: creating virtual environment .venv ...")
        code = run([py, "-m", "venv", os.path.join(BASE, ".venv")]).returncode
        if code != 0:
            log("[ERROR] Failed to create virtual environment. Check Python install.")
            return

    # 2) Install deps if streamlit missing
    try:
        subprocess.run([VENV_PY, "-m", "streamlit", "--version"],
                       capture_output=True, check=True)
    except subprocess.CalledProcessError:
        log("[INFO] Installing dependencies (first run, please wait) ...")
        run([VENV_PY, "-m", "pip", "install", "--upgrade", "pip"])
        code = run([VENV_PY, "-m", "pip", "install", "-r", REQ]).returncode
        if code != 0:
            log("[ERROR] Dependency install failed. Check network connection.")
            return

    # 3) Launch
    log("[INFO] Starting, open http://localhost:8501 in your browser")
    log("[INFO] To stop, close this window")
    run([VENV_PY, "-m", "streamlit", "run", APP, "--server.headless", "true"])


if __name__ == "__main__":
    try:
        main()
    except Exception as e:
        log("[ERROR] " + str(e))
    finally:
        input("Press Enter to exit...")
