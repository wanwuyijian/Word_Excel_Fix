"""GUI 冒烟 + 截图。

用法：PYTHONPATH=src python build_tools/_t_gui.py [--out shot.png]
"""
from __future__ import annotations

import io
import os
import sys

sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding="utf-8", errors="replace")
sys.path.insert(0, os.path.join(os.path.dirname(__file__), "..", "src"))

from wxfixer.gui.app import App

OUT = os.path.join(os.path.dirname(__file__), "..", "render_out")
os.makedirs(OUT, exist_ok=True)
shot = os.path.join(OUT, "gui_v11.png")


def pump(app, n=25):
    for _ in range(n):
        app.update()
        app.after(20)
        import time

        time.sleep(0.01)


def main():
    app = App()
    app.geometry("1220x860+40+20")
    pump(app, 40)
    print("窗口尺寸:", app.winfo_width(), "x", app.winfo_height())
    print("模板下拉项:", app.tpl_combo.cget("values") if hasattr(app, "tpl_combo") else "(radio)")
    print("模板当前:", app._current_template().name)
    print("选项:", {k: v.get() for k, v in app.toggle_vars.items()})

    # 交互：切换各开关
    app.toggle_vars["content"].set(True)
    app._on_option_change()
    pump(app, 10)
    print("开启内容优化后摘要:", app.opt_summary.cget("text"))
    print("警示条显示:", app._warn_shown)

    app.toggle_vars["layout"].set(False)
    app._on_option_change()
    pump(app, 10)
    print("关闭排版后摘要:", app.opt_summary.cget("text"))
    disabled = [str(cb.cget("text")) for cb in app._style_checks
                if str(cb.cget("state")) == "disabled"]
    print("被禁用的版式项:", disabled)

    app._preset_layout_only()
    pump(app, 10)
    print("只调排版预设:", {k: v.get() for k, v in app.toggle_vars.items()})
    print("预设后摘要:", app.opt_summary.cget("text"))

    # 加文件
    samples = os.path.join(os.path.dirname(__file__), "..", "samples")
    if os.path.isdir(samples):
        paths = [os.path.join(samples, f) for f in os.listdir(samples)
                 if f.endswith((".docx", ".xlsx"))][:2]
        app._add_paths(paths)
        pump(app, 10)
        print("已加文件:", len(app.files), "| 计数标签:", app.file_count_lbl.cget("text"))

    # 切模板
    for tpl in app._current_template().__class__.__mro__[0] and []:
        pass
    from wxfixer.templates import ALL_TEMPLATES

    for t in ALL_TEMPLATES:
        app.tpl_key_var.set(t.key)
        app._show_template_info()
        pump(app, 3)
        assert app._current_template().key == t.key
    print("8 个模板均可切换且说明可刷新")

    try:
        from PIL import ImageGrab

        img = ImageGrab.grab()
        img.save(shot)
        print("截图:", shot, img.size)
    except Exception as exc:  # noqa: BLE001
        print("截图失败:", exc)

    app.destroy()
    return 0


if __name__ == "__main__":
    sys.exit(main())
