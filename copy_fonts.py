import os
import shutil

src_dir = r"C:\Windows\Fonts"
dst_dir = os.path.join(os.path.dirname(__file__), "fonts")
os.makedirs(dst_dir, exist_ok=True)

files = ["times.ttf", "timesbd.ttf", "timesi.ttf", "timesbi.ttf"]
copied = 0
for f in files:
    src = os.path.join(src_dir, f)
    dst = os.path.join(dst_dir, f)
    if os.path.exists(src):
        shutil.copy(src, dst)
        print(f"SUCCESS: Copied {f} to fonts/")
        copied += 1
    else:
        print(f"WARNING: Could not find {src}")

if copied > 0:
    print(f"\nFont setup complete! {copied} font file(s) ready in project fonts/ directory.")
