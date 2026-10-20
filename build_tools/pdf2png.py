"""把 PDF 的每一页渲染成 PNG，便于人工目检排版效果。

用法: python build_tools/pdf2png.py <pdf目录或文件> [输出目录] [--dpi 110] [--max-pages 4]
"""

from __future__ import annotations

import os
import sys

import fitz  # PyMuPDF


def render(pdf: str, out_dir: str, dpi: int = 110, max_pages: int = 4) -> list[str]:
    os.makedirs(out_dir, exist_ok=True)
    stem = os.path.splitext(os.path.basename(pdf))[0]
    out = []
    doc = fitz.open(pdf)
    try:
        for i, page in enumerate(doc):
            if i >= max_pages:
                break
            pix = page.get_pixmap(dpi=dpi)
            path = os.path.join(out_dir, f"{stem}_p{i + 1}.png")
            pix.save(path)
            out.append(path)
    finally:
        doc.close()
    return out


def main() -> None:
    args = [a for a in sys.argv[1:] if not a.startswith("--")]
    dpi = 110
    max_pages = 4
    for a in sys.argv[1:]:
        if a.startswith("--dpi"):
            dpi = int(a.split("=")[-1])
        if a.startswith("--max-pages"):
            max_pages = int(a.split("=")[-1])
    src = args[0] if args else "tests/preview"
    out = args[1] if len(args) > 1 else os.path.join(src, "png")
    pdfs = []
    if os.path.isdir(src):
        pdfs = [os.path.join(src, f) for f in sorted(os.listdir(src)) if f.endswith(".pdf")]
    else:
        pdfs = [src]
    for pdf in pdfs:
        for p in render(pdf, out, dpi, max_pages):
            print(p)


if __name__ == "__main__":
    main()
