"""
Regenerate karyotype_assets.py from source image files in assets/.
Run this whenever source assets change:
    python3 generate_assets_py.py
"""
import base64
import os

ASSETS_DIR = os.path.join(os.path.dirname(os.path.abspath(__file__)), "assets")
OUT_FILE   = os.path.join(os.path.dirname(os.path.abspath(__file__)), "karyotype_assets.py")

FILES = [
    ("header.jpg",        "HEADER"),
    ("footer.jpg",        "FOOTER"),
    ("stamp_mc6558.jpg",  "STAMP"),
    ("sign_deepika.jpg",  "SIGN_DEEPIKA"),
    ("sign_teena.jpg",    "SIGN_TEENA"),
    ("sign_suriya.jpg",   "SIGN_SURIYA"),
]

lines = [
    '"""',
    'Auto-generated asset file for Karyotype Report Generator.',
    'Regenerate with:  python3 generate_assets_py.py',
    '"""',
    '',
]

for fname, varname in FILES:
    fpath = os.path.join(ASSETS_DIR, fname)
    if not os.path.isfile(fpath):
        print(f"  WARNING: {fpath} not found — skipping")
        continue
    data = base64.b64encode(open(fpath, "rb").read()).decode()
    kb   = os.path.getsize(fpath) // 1024
    lines.append(f"# Source: {fname}  ({kb} KB)")
    lines.append(f'{varname} = "{data}"')
    lines.append("")
    print(f"  {varname}: {kb} KB  ✓")

with open(OUT_FILE, "w") as f:
    f.write("\n".join(lines))

print(f"\nWrote {OUT_FILE}  ({os.path.getsize(OUT_FILE) // 1024} KB)")
