from __future__ import print_function
import os
import sys
import subprocess
import re


def ver_of(exe):
    try:
        out = subprocess.check_output(
            [exe, "-c", "import sys; print('%d.%d' % sys.version_info[:2])"]
        )
        s = out.decode("ascii", "ignore").strip().splitlines()[-1]
        parts = s.split(".")
        return (int(parts[0]), int(parts[1]))
    except Exception:
        return None


candidates = []

# 1) PATH
for d in os.environ.get("PATH", "").split(os.pathsep):
    p = os.path.join(d, "python.exe")
    if os.path.isfile(p):
        candidates.append(p)

# 2) common install locations
roots = []
la = os.path.expandvars("%LOCALAPPDATA%")
if la:
    roots.append(os.path.join(la, "Programs", "Python"))
for r in [r"C:\Python", os.path.expandvars("%ProgramFiles%"),
          os.path.expandvars("%ProgramFiles(x86)%")]:
    if r:
        roots.append(r)
for root in roots:
    if not os.path.isdir(root):
        continue
    for name in os.listdir(root):
        low = name.lower()
        if low.startswith("python"):
            p = os.path.join(root, name, "python.exe")
            if os.path.isfile(p):
                candidates.append(p)

# 3) py launcher: resolve highest 3.x
try:
    out = subprocess.check_output(["py", "-0"]).decode("ascii", "ignore")
    best = None
    for line in out.splitlines():
        m = re.search(r"(\d+)\.(\d+)", line)
        if m:
            v = (int(m.group(1)), int(m.group(2)))
            if v[0] == 3 and (best is None or v > best[0]):
                best = (v, line)
    if best:
        tag = "%d.%d" % best[0]
        try:
            exe = subprocess.check_output(
                ["py", "-" + tag, "-c", "import sys; print(sys.executable)"]
            ).decode("ascii", "ignore").strip()
            if os.path.isfile(exe):
                candidates.append(exe)
        except Exception:
            pass
except Exception:
    pass

best = None
best_ver = (0, 0)
seen = set()
for c in candidates:
    if c in seen:
        continue
    seen.add(c)
    v = ver_of(c)
    if v and v >= (3, 9) and v > best_ver:
        best_ver = v
        best = c

if best:
    print(best)
else:
    print("NONE")
