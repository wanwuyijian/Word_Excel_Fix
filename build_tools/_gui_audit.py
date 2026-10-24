"""GUI 布局几何审计：不依赖截图，用 Tk 的几何信息检查越界/挤压/截断。

用法：PYTHONPATH=src python build_tools/_gui_audit.py [宽 高 ...]
"""
from __future__ import annotations

import io
import os
import sys
import time

sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding="utf-8", errors="replace")
sys.path.insert(0, os.path.join(os.path.dirname(__file__), "..", "src"))

from wxfixer.gui.app import App


def pump(app, n=20):
    for _ in range(n):
        app.update()
        time.sleep(0.01)


def walk(w, depth=0):
    yield w, depth
    for c in w.winfo_children():
        yield from walk(c, depth + 1)


def audit(app, tag):
    win_w, win_h = app.winfo_width(), app.winfo_height()
    problems = []
    for w, depth in walk(app):
        try:
            if not w.winfo_ismapped():
                continue
            x, y = w.winfo_rootx() - app.winfo_rootx(), w.winfo_rooty() - app.winfo_rooty()
            ww, hh = w.winfo_width(), w.winfo_height()
            rw, rh = w.winfo_reqwidth(), w.winfo_reqheight()
        except Exception:
            continue
        name = f"{w.winfo_class()}:{str(w)[-40:]}"
        if x + ww > win_w + 2 or y + hh > win_h + 2:
            problems.append(f"越界 {name} 位置({x},{y}) 尺寸({ww}x{hh}) 窗口({win_w}x{win_h})")
        if ww + 1 < min(rw, 24):
            problems.append(f"过窄 {name} 实际宽{ww} < 需要{rw}")
        if hh + 1 < min(rh, 30):
            problems.append(f"过矮 {name} 实际高{hh} < 需要{rh}")
    print(f"[{tag}] 窗口 {win_w}x{win_h} 问题 {len(problems)} 项")
    for p in problems[:20]:
        print("   ", p)
    return len(problems)


def main():
    sizes = [(1220, 860), (1040, 760), (1600, 900)]
    total = 0
    for w, h in sizes:
        app = App()
        app.geometry(f"{w}x{h}+20+10")
        pump(app, 30)
        app.toggle_vars["content"].set(True)
        app._on_option_change()
        pump(app, 15)
        total += audit(app, f"内容预警条展开 {w}x{h}")
        app.toggle_vars["content"].set(False)
        app._on_option_change()
        pump(app, 10)
        total += audit(app, f"默认 {w}x{h}")

        # 关键卡片高度检查
        print("    日志区高度:", app.log_frame.winfo_height(),
              "| 进度条高:", app.progress.winfo_height(),
              "| 树高:", app.tree.winfo_height())
        app.destroy()
    print("合计问题:", total)
    return 0 if total == 0 else 1


if __name__ == "__main__":
    sys.exit(main())
