"""无人工干预地启动界面、截图、退出（开发自检用）。

用法: python build_tools/gui_shot.py [输出png]
依赖: Pillow
"""

from __future__ import annotations

import os
import sys

SRC = os.path.abspath(os.path.join(os.path.dirname(__file__), "..", "src"))
sys.path.insert(0, SRC)

from wxfixer.utils import logging_setup  # noqa: E402


def shot(path: str = "shot.png", files: list[str] | None = None) -> None:
    logging_setup.setup_logging(console=False)
    from wxfixer.gui.app import App

    app = App()
    if files:
        app._add_paths(files)
    app.update()
    app.update_idletasks()

    from PIL import ImageGrab

    x = app.winfo_rootx()
    y = app.winfo_rooty()
    w = app.winfo_width()
    h = app.winfo_height()
    app.after(400, lambda: None)
    app.update()
    img = ImageGrab.grab(bbox=(x, y, x + w, y + h))
    img.save(path)
    print(f"saved {path} {img.size}")
    app.destroy()


if __name__ == "__main__":
    out = sys.argv[1] if len(sys.argv) > 1 else "shot.png"
    samples = sys.argv[2:] or None
    shot(out, samples)
