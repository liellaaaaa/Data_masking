from __future__ import print_function
import os
import sys
import subprocess

venv_py = os.path.join(
    os.path.dirname(os.path.abspath(__file__)), ".venv", "Scripts", "python.exe"
)
if not os.path.isfile(venv_py):
    sys.exit(2)
try:
    out = subprocess.check_output(
        [venv_py, "-c", "import sys; print(1 if sys.version_info >= (3, 9) else 0)"]
    )
    ok = int(out.decode("ascii", "ignore").strip().splitlines()[-1])
    sys.exit(0 if ok else 1)
except Exception:
    sys.exit(1)
