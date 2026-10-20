"""批量把 docx 渲染成 PDF/PNG 做视觉检查（依赖本机 Word + PyMuPDF）。

用法:
    python build_tools/render_batch.py tests/out --out-dir render_out --filter 发布通知 --pages 2
"""

from __future__ import annotations

import argparse
import glob
import os
import subprocess
import sys

BASE = os.path.abspath(os.path.join(os.path.dirname(__file__), ".."))
PS = os.path.join(BASE, "build_tools", "render_check.ps1")
PDF2PNG = os.path.join(BASE, "build_tools", "pdf2png.py")


def main() -> int:
    ap = argparse.ArgumentParser()
    ap.add_argument("folder")
    ap.add_argument("--out-dir", default=os.path.join(BASE, "render_out"))
    ap.add_argument("--limit", type=int, default=0)
    ap.add_argument("--pages", type=int, default=2)
    ap.add_argument("--filter", default="")
    args = ap.parse_args()

    folder = os.path.abspath(args.folder)
    out_dir = os.path.abspath(args.out_dir)
    os.makedirs(out_dir, exist_ok=True)

    pattern = f"*{args.filter}*" if args.filter else "*"
    cmd = ["powershell", "-NoProfile", "-ExecutionPolicy", "Bypass",
           "-File", PS, "-InDir", folder, "-OutDir", out_dir,
           "-Filter", pattern, "-MaxFiles", str(args.limit or 50)]
    proc = subprocess.run(cmd, capture_output=True, timeout=1800)
    print(proc.stdout.decode("utf-8", "ignore").strip())
    if proc.stderr.strip():
        print(proc.stderr.decode("utf-8", "ignore")[:800])

    pdfs = sorted(glob.glob(os.path.join(out_dir, "*.pdf")))
    for pdf in pdfs:
        subprocess.run([sys.executable, PDF2PNG, pdf, out_dir,
                        "--dpi=110", f"--max-pages={args.pages}"],
                       capture_output=True, timeout=300)
    print(f"PDF {len(pdfs)} 个 -> {out_dir}")
    return 0


if __name__ == "__main__":
    sys.exit(main())
