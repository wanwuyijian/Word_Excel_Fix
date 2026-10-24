import io
import sys

sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding="utf-8", errors="replace")
sys.path.insert(0, "src")

from openpyxl import load_workbook

p = "tests/out_v11/样例_内容优化.xlsx"
import glob, os

outs = sorted(glob.glob("tests/out_v11/样例_内容优化_排版优化_*.xlsx"))
print(outs)
for f in outs:
    wb = load_workbook(f)
    ws = wb["清单"]
    imgs = getattr(ws, "_images", [])
    print(f, "images:", len(imgs))
    for im in imgs:
        print("  w/h:", im.width, im.height, "anchor:", type(im.anchor).__name__,
              "ext:", getattr(im.anchor, "ext", None))
    wb.close()
