import os
import re
import zipfile
import rarfile
import shutil
from pathlib import Path

MAIN_RAR    = "backup_20250527.rar"
WORKDIR     = "work"
OUTDIR      = "output"
ZIP_PATTERN = re.compile(r'^(?P<prefix>.+)-(?P<date>\d{8})\.zip$', re.IGNORECASE)

Path(WORKDIR).mkdir(exist_ok=True)
Path(OUTDIR).mkdir(exist_ok=True)

# 1) Extract RAR
with rarfile.RarFile(MAIN_RAR) as rf:
    for member in rf.infolist():
        try:
            rf.extract(member, WORKDIR)
        except rarfile.BadRarFile:
            # skip bad entries
            continue

# 2) Find all ZIPs under WORKDIR (recursive)
zip_paths = []
for root, _, files in os.walk(WORKDIR):
    for f in files:
        if f.lower().endswith(".zip"):
            zip_paths.append(os.path.join(root, f))

print(f"Found {len(zip_paths)} zip files in all subfolders")

# 3) Group by prefix and pick the latest date
groups = {}
for fullpath in zip_paths:
    fname = os.path.basename(fullpath)
    m = ZIP_PATTERN.match(fname)
    if not m:
        print(f"Skipping non-matching: {fname}")
        continue
    prefix, date_str = m.group("prefix"), m.group("date")
    date_num = int(date_str)
    prev = groups.get(prefix)
    if prev is None or date_num > prev[0]:
        groups[prefix] = (date_num, fullpath)

# 4) Delete older zips
for fullpath in zip_paths:
    fname = os.path.basename(fullpath)
    m = ZIP_PATTERN.match(fname)
    if m:
        prefix = m.group("prefix")
        # if this fullpath is not the one we decided to keep:
        if fullpath != groups[prefix][1]:
            os.remove(fullpath)

# 5) Extract each kept zip into its own OUTDIR folder
for prefix, (date_num, keep_path) in groups.items():
    fname = os.path.basename(keep_path)
    folder = f"{prefix}-{date_num}"
    dest  = os.path.join(OUTDIR, folder)
    Path(dest).mkdir(exist_ok=True)
    print(f"Extracting {fname} → {dest}/")
    try:
        with zipfile.ZipFile(keep_path, "r") as zf:
            zf.extractall(dest)
    except zipfile.BadZipFile:
        print(f"⚠️ Skipping invalid ZIP file: {keep_path}")


print("✅ Done!  Extracted only the latest zip per prefix.")
