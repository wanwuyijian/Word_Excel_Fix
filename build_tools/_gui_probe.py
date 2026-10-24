import io
import os
import sys
import time

sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding="utf-8", errors="replace")
sys.path.insert(0, os.path.join(os.path.dirname(__file__), "..", "src"))

from wxfixer.gui.app import App

app = App()
app.geometry("1220x860+10+10")
for _ in range(40):
    app.update()
    time.sleep(0.01)

def show(w, depth=0, maxdepth=4):
    if depth > maxdepth:
        return
    try:
        print("  " * depth, w.winfo_class(), w.winfo_name(),
              "req:", w.winfo_reqwidth(), "x", w.winfo_reqheight(),
              "actual:", w.winfo_width(), "x", w.winfo_height(),
              "mapped:", w.winfo_ismapped(),
              "mgr:", w.winfo_manager())
    except Exception as e:
        print("  " * depth, "?", e)
    for c in w.winfo_children():
        show(c, depth + 1, maxdepth)

print("=== 右侧 body 第2列 ===")
show(app._body.winfo_children()[1], 0, 3)
app.destroy()
