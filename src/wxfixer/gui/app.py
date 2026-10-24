"""主窗口（v1.1）。

布局：
┌───────────────────────────────────────────────────────────────┐
│ 顶栏：应用名 + 版本徽标 + 一句话定位                            │
├────────────────────────────────┬──────────────────────────────┤
│ ① 选择文件                      │ ③ 排版模板（8 选一 + 说明）   │
│ ② 交互选项（版式 / 内容与素材） │ ④ 输出与执行（进度 + 主按钮） │
├────────────────────────────────┴──────────────────────────────┤
│ 运行日志（可折叠）                                             │
└───────────────────────────────────────────────────────────────┘
"""

from __future__ import annotations

import datetime as _dt
import os
import queue
import subprocess
import sys
import threading
import tkinter as tk
import tkinter.messagebox as mb
import tkinter.filedialog as fd
from tkinter import ttk

from ..core import OptimizeError, OptimizeOptions, UnsupportedFileError, optimize_file
from ..core.runner import SUPPORTED_EXT
from ..templates import ALL_TEMPLATES, get_template
from ..templates.schema import TOGGLE_META
from ..utils import logging_setup, paths
from ..version import APP_NAME, APP_VERSION, BUILD_DATE, GITHUB_URL, full_title
from . import theme

log = logging_setup.operations_logger()

ACCEPT_EXTS = (".docx", ".docm", ".xlsx", ".xlsm")

# TOGGLE_META: (key, 名称, 说明, 组, 默认)
_GROUP_TITLE = {
    "style": ("版式调整", "不改变任何文字，只调整版式与视觉层次"),
    "content": ("内容与素材", "会改动文字或图片尺寸，请自行确认"),
}
_GROUP_ORDER = ("style", "content")


def _base_class():
    """优先使用支持拖放的 Tk 子类；不可用时退回标准 tk.Tk。"""
    try:
        from tkinterdnd2 import Tk as DnDTk

        return DnDTk
    except Exception:
        return tk.Tk


def _bind_dnd(app) -> bool:
    """递归给窗口内所有控件注册拖放，返回是否成功。"""
    from tkinterdnd2 import DND_FILES

    ok = False

    def walk(widget):
        nonlocal ok
        for child in widget.winfo_children():
            try:
                child.drop_target_register(DND_FILES)
                child.dnd_bind("<<Drop>>", app._on_drop)
                ok = True
            except Exception:
                pass
            walk(child)

    walk(app)
    try:
        app.drop_target_register(DND_FILES)
        app.dnd_bind("<<Drop>>", app._on_drop)
        ok = True
    except Exception:
        pass
    return ok


class App(_base_class()):
    def __init__(self) -> None:
        super().__init__()
        self.title(full_title())
        self.configure(bg=theme.BG)
        self.resizable(True, True)   # 支持最大化 / 最小化 / 自由缩放
        self.minsize(1040, 760)
        self._center(1220, 860)

        theme.apply_theme(self)
        try:
            self.iconbitmap(os.path.join(paths.bundle_dir(), "assets", "app.ico"))
        except Exception:
            pass

        self.files: list[str] = []
        self.out_dir_var = tk.StringVar()
        self._auto_out = True
        self.queue: queue.Queue = queue.Queue()
        self.worker: threading.Thread | None = None
        self._running = False
        self._results: list = []
        self._log_visible = tk.BooleanVar(value=True)

        self.tpl_key_var = tk.StringVar(value=ALL_TEMPLATES[0].key)
        self.toggle_vars: dict[str, tk.BooleanVar] = {
            meta[0]: tk.BooleanVar(value=meta[4]) for meta in TOGGLE_META
        }
        self._style_checks: list[ttk.Checkbutton] = []

        self._build()
        self._enable_dnd()
        self.protocol("WM_DELETE_WINDOW", self._on_close)
        self.after(80, self._poll_queue)
        self.bind("<Configure>", self._on_resize)
        logging_setup.set_crash_notifier(self._on_crash)
        self._log_ui(f"就绪 · 日志目录 {paths.logs_dir()}")
        log.info("界面已启动（v%s）", APP_VERSION)

    # ------------------------------------------------------------------
    # 布局骨架
    # ------------------------------------------------------------------
    def _center(self, w: int, h: int) -> None:
        self.update_idletasks()
        sw, sh = self.winfo_screenwidth(), self.winfo_screenheight()
        w = min(w, max(900, sw - 80))
        h = min(h, max(700, sh - 100))
        self.geometry(f"{w}x{h}+{max(0, (sw - w) // 2)}+{max(0, (sh - h) // 3)}")

    def _build(self) -> None:
        outer = ttk.Frame(self, padding=(18, 14, 18, 12))
        outer.pack(fill="both", expand=True)

        self._build_header(outer)

        # 先把「底部固定区」用 side=bottom 排好，保证窗口变小时日志不会
        # 被中间的可伸缩区域挤成 1 像素。
        self._build_statusbar(outer)
        self._build_log(outer)

        body = ttk.Frame(outer)
        body.pack(fill="both", expand=True, pady=(12, 0))
        self._body = body
        body.columnconfigure(0, weight=1, uniform="col")
        body.columnconfigure(1, weight=1, uniform="col")
        body.rowconfigure(0, weight=1)

        left = ttk.Frame(body)
        left.grid(row=0, column=0, sticky="nsew", padx=(0, 12))
        left.columnconfigure(0, weight=1)
        left.rowconfigure(0, weight=3, minsize=248)
        left.rowconfigure(1, weight=2, minsize=252)

        right = ttk.Frame(body)
        right.grid(row=0, column=1, sticky="nsew")
        right.columnconfigure(0, weight=1)
        right.rowconfigure(0, weight=3, minsize=240)
        right.rowconfigure(1, weight=2, minsize=262)

        self._build_file_card(left)
        self._build_option_card(left)
        self._build_template_card(right)
        self._build_run_card(right)

    def _build_header(self, parent) -> None:
        header = ttk.Frame(parent)
        header.pack(fill="x")

        left = ttk.Frame(header)
        left.pack(side="left")
        row = ttk.Frame(left)
        row.pack(anchor="w")
        ttk.Label(row, text=APP_NAME, style="H1.TLabel").pack(side="left")
        ttk.Label(row, text=f" v{APP_VERSION} ", style="Badge.TLabel").pack(
            side="left", padx=(8, 0), pady=(2, 0))
        ttk.Label(
            left,
            text="一键把 Word / Excel 调成「能直接交出去」的样子：默认只动排版，不动事实",
            style="Muted.TLabel",
        ).pack(anchor="w", pady=(2, 0))

        ttk.Label(header, text="本地运行 · 不联网 · 不上传文件",
                  style="BadgeGray.TLabel").pack(side="right", pady=(6, 0))

        ttk.Separator(parent, orient="horizontal").pack(fill="x", pady=(10, 0))

    def _card(self, parent, row: int, num: str, title: str, subtitle: str,
              pady=(0, 10), padding=14):
        # 外层 1px 描边 + 内层白底，让卡片在浅灰底色上有清晰的边界。
        shell = tk.Frame(parent, bg=theme.BORDER, bd=0, highlightthickness=0)
        shell.grid(row=row, column=0, sticky="nsew", pady=pady)
        card = ttk.Frame(shell, style="Card.TFrame", padding=padding)
        card.pack(fill="both", expand=True, padx=1, pady=1)
        head = ttk.Frame(card, style="Card.TFrame")
        head.pack(fill="x")
        ttk.Label(head, text=num, style="Badge.TLabel").pack(side="left")
        ttk.Label(head, text=title, style="H2.TLabel").pack(side="left", padx=(7, 0))
        if subtitle:
            ttk.Label(card, text=subtitle, style="CardMuted.TLabel").pack(
                anchor="w", pady=(3, 0))
        return card

    # ---- ① 文件 ----
    def _build_file_card(self, parent) -> None:
        card = self._card(parent, 0, "1", "选择文件",
                          "支持 Word（.docx / .docm）与 Excel（.xlsx / .xlsm）")
        self.drop_hint = ttk.Label(card, text="把文件拖到这里，或点击「添加文件」",
                                   style="CardMuted.TLabel")
        self.drop_hint.pack(anchor="w", pady=(2, 8))

        btns = ttk.Frame(card, style="Card.TFrame")
        btns.pack(fill="x")
        ttk.Button(btns, text="添加文件…", command=self.add_files).pack(side="left")
        ttk.Button(btns, text="移除选中", style="Ghost.TButton",
                   command=self.remove_selected).pack(side="left", padx=6)
        ttk.Button(btns, text="清空", style="Ghost.TButton",
                   command=self.clear_files).pack(side="left")
        self.file_count_lbl = ttk.Label(btns, text="尚未添加文件", style="CardMuted.TLabel")
        self.file_count_lbl.pack(side="right")

        tree_wrap = ttk.Frame(card, style="Card.TFrame")
        tree_wrap.pack(fill="both", expand=True, pady=(10, 0))
        cols = ("type", "name", "size")
        self.tree = ttk.Treeview(tree_wrap, columns=cols, show="headings",
                                 selectmode="extended", height=7)
        self.tree.heading("type", text="类型")
        self.tree.heading("name", text="文件名")
        self.tree.heading("size", text="大小")
        self.tree.column("type", width=70, anchor="center", stretch=False)
        self.tree.column("name", width=360, anchor="w")
        self.tree.column("size", width=86, anchor="e", stretch=False)
        vsb = ttk.Scrollbar(tree_wrap, orient="vertical", command=self.tree.yview)
        self.tree.configure(yscrollcommand=vsb.set)
        self.tree.pack(side="left", fill="both", expand=True)
        vsb.pack(side="right", fill="y")
        self.tree.bind("<Double-1>", self._open_selected_file)
        self.tree.bind("<Delete>", lambda e: self.remove_selected())

    # ---- ② 交互选项 ----
    def _build_option_card(self, parent) -> None:
        card = self._card(parent, 1, "2", "交互选项",
                          "勾选 = 允许程序动这一类；不勾 = 该部分保持原样")
        wrap = ttk.Frame(card, style="Card.TFrame")
        wrap.pack(fill="both", expand=True, pady=(8, 0))
        wrap.columnconfigure(0, weight=1, uniform="grp")
        wrap.columnconfigure(1, weight=1, uniform="grp")
        wrap.rowconfigure(0, weight=1)

        groups = {g: [m for m in TOGGLE_META if m[3] == g] for g in _GROUP_ORDER}
        for col, gkey in enumerate(_GROUP_ORDER):
            title, hint = _GROUP_TITLE[gkey]
            box = ttk.Frame(wrap, style="Group.TFrame", padding=(9, 7))
            box.grid(row=0, column=col, sticky="nsew",
                     padx=(0, 5) if col == 0 else (5, 0))
            box.columnconfigure(0, weight=1)
            ttk.Label(box, text=title, style="GroupTitle.TLabel").grid(
                row=0, column=0, columnspan=2, sticky="w")
            ttk.Label(box, text=hint, style="GroupMuted.TLabel",
                      wraplength=250, justify="left").grid(
                row=1, column=0, columnspan=2, sticky="w", pady=(1, 5))

            items = groups[gkey]
            two_col = len(items) > 2
            for i, (key, name, _desc, _g, _d) in enumerate(items):
                cb = ttk.Checkbutton(
                    box, text=name, variable=self.toggle_vars[key],
                    command=self._on_option_change,
                    style="Group.TCheckbutton" if gkey == "style" else "Card.TCheckbutton")
                r = 2 + (i // 2 if two_col else i)
                c = (i % 2) if two_col else 0
                cb.grid(row=r, column=c, sticky="w", pady=1)
                if gkey == "style":
                    self._style_checks.append(cb)

        self.content_warn = ttk.Frame(card, style="Warning.TFrame", padding=(10, 6))
        ttk.Label(self.content_warn, text="注意",
                  style="WarningTitle.TLabel").pack(side="left", padx=(0, 6))
        self.content_warn_lbl = ttk.Label(
            self.content_warn,
            text="已允许改动文字：程序只删赘语、合并重复标点、清理多余空格，不改写事实。",
            style="Warning.TLabel", justify="left")
        self.content_warn_lbl.pack(side="left", fill="x", expand=True)

        quick = ttk.Frame(card, style="Card.TFrame")
        quick.pack(fill="x", pady=(8, 0))
        ttk.Button(quick, text="全部允许", style="Ghost.TButton",
                   command=lambda: self._set_toggles(True)).pack(side="left")
        ttk.Button(quick, text="全部不改", style="Ghost.TButton",
                   command=lambda: self._set_toggles(False)).pack(side="left", padx=6)
        ttk.Button(quick, text="只调排版", style="Ghost.TButton",
                   command=self._preset_layout_only).pack(side="left")
        self.opt_summary = ttk.Label(quick, text="", style="CardMuted.TLabel")
        self.opt_summary.pack(side="right")
        self._quick_row = quick
        self._warn_shown = False
        self._on_option_change()

    # ---- ③ 模板 ----
    def _build_template_card(self, parent) -> None:
        card = self._card(parent, 0, "3", "排版模板",
                          "按文档用途选一套版式规范；不确定就用第一项")
        grid = ttk.Frame(card, style="Card.TFrame")
        grid.pack(fill="x", pady=(8, 0))
        grid.columnconfigure(0, weight=1)
        grid.columnconfigure(1, weight=1)
        for i, tpl in enumerate(ALL_TEMPLATES):
            ttk.Radiobutton(
                grid, text=tpl.name, value=tpl.key, variable=self.tpl_key_var,
                command=self._show_template_info, style="Card.TCheckbutton",
            ).grid(row=i // 2, column=i % 2, sticky="w", pady=3, padx=(0, 10))

        ttk.Separator(card, orient="horizontal").pack(fill="x", pady=(10, 8))
        info = ttk.Frame(card, style="Card.TFrame")
        info.pack(fill="both", expand=True)
        self.tpl_desc = tk.Text(info, height=8, wrap="word", bd=0, relief="flat",
                                bg="#F8FAFD", fg=theme.TEXT, font=theme.ui_font(9),
                                padx=10, pady=8, cursor="arrow", spacing1=1,
                                spacing3=2)
        self.tpl_desc.pack(fill="both", expand=True)
        self._show_template_info()

    # ---- ④ 输出与执行 ----
    def _build_run_card(self, parent) -> None:
        card = self._card(parent, 1, "4", "输出与执行",
                          "结果另存为新文件，源文件只读不改")
        row = ttk.Frame(card, style="Card.TFrame")
        row.pack(fill="x", pady=(8, 0))
        self.out_entry = ttk.Entry(row, textvariable=self.out_dir_var)
        self.out_entry.pack(side="left", fill="x", expand=True)
        ttk.Button(row, text="浏览…", style="Ghost.TButton",
                   command=self.choose_out_dir).pack(side="left", padx=(6, 0))
        ttk.Button(row, text="打开", style="Ghost.TButton",
                   command=self.open_out_dir).pack(side="left", padx=(6, 0))

        self.progress = ttk.Progressbar(card, style="Blue.Horizontal.TProgressbar",
                                        maximum=100, value=0)
        self.progress.pack(fill="x", pady=(14, 5))
        self.progress_txt = ttk.Label(card, text="等待开始", style="CardMuted.TLabel")
        self.progress_txt.pack(anchor="w")

        action = ttk.Frame(card, style="Card.TFrame")
        action.pack(fill="x", pady=(12, 0))
        self.run_btn = ttk.Button(action, text="开始优化", style="Primary.TButton",
                                  command=self.start)
        self.run_btn.pack(side="left")
        ttk.Button(action, text="打开日志", style="Link.TButton",
                   command=self.open_logs).pack(side="right")
        ttk.Button(action, text="复制日志", style="Link.TButton",
                   command=self._copy_log).pack(side="right", padx=(0, 8))

    # ---- 日志与状态栏 ----
    def _build_log(self, parent) -> None:
        bar = ttk.Frame(parent)
        bar.pack(side="bottom", fill="x", pady=(8, 0))
        ttk.Checkbutton(bar, text="显示运行日志", variable=self._log_visible,
                        command=self._toggle_log).pack(side="left")
        ttk.Button(bar, text="打开输出目录", style="Link.TButton",
                   command=self.open_out_dir).pack(side="right")

        self.log_frame = ttk.Frame(parent, style="Card.TFrame", padding=(1, 1))
        self.log_frame.pack(side="bottom", fill="x", pady=(4, 0))
        self.log_text = tk.Text(self.log_frame, height=6, wrap="word", bd=0, relief="flat",
                                bg="#0F172A", fg="#D6E2F5", insertbackground="#D6E2F5",
                                font=theme.mono_font(9), padx=10, pady=6)
        self.log_text.pack(side="left", fill="both", expand=True)
        lsb = ttk.Scrollbar(self.log_frame, orient="vertical", command=self.log_text.yview)
        self.log_text.configure(yscrollcommand=lsb.set)
        lsb.pack(side="right", fill="y")
        self.log_text.tag_configure("ok", foreground="#6EE7B7")
        self.log_text.tag_configure("err", foreground="#FCA5A5")
        self.log_text.tag_configure("warn", foreground="#FCD34D")
        self.log_text.tag_configure("muted", foreground="#94A3B8")

    def _build_statusbar(self, parent) -> None:
        bar = ttk.Frame(parent)
        bar.pack(side="bottom", fill="x", pady=(6, 0))
        ttk.Label(bar, text=f"版本 {APP_VERSION} · 构建 {BUILD_DATE}",
                  style="Muted.TLabel").pack(side="left")
        link = ttk.Label(bar, text=GITHUB_URL, style="Muted.TLabel",
                         foreground=theme.PRIMARY, cursor="hand2")
        link.pack(side="right")
        link.bind("<Button-1>", lambda e: self._open_url(GITHUB_URL))
        self.status_lbl = ttk.Label(bar, text="", style="Muted.TLabel")
        self.status_lbl.pack(side="right", padx=16)

    # ------------------------------------------------------------------
    # 选项联动
    # ------------------------------------------------------------------
    def _current_template(self):
        return get_template(self.tpl_key_var.get())

    def _show_template_info(self) -> None:
        tpl = self._current_template()
        lines = [tpl.description, ""]
        lines += [f"• {tip}" for tip in tpl.tips]
        self.tpl_desc.configure(state="normal")
        self.tpl_desc.delete("1.0", "end")
        self.tpl_desc.insert("1.0", "\n".join(lines))
        self.tpl_desc.configure(state="disabled")

    def _set_toggles(self, value: bool) -> None:
        for v in self.toggle_vars.values():
            v.set(value)
        self._on_option_change()

    def _preset_layout_only(self) -> None:
        """只调排版：版式全开，内容与图片不动。"""
        for key, var in self.toggle_vars.items():
            var.set(key in ("layout", "font_family", "font_size",
                            "font_color", "background", "page_break"))
        self._on_option_change()

    def _wrap_warn(self) -> None:
        """警示条文字按可用宽度自动换行（窗口缩放时同步）。"""
        try:
            w = max(180, self.content_warn.winfo_width() - 70)
            self.content_warn_lbl.configure(wraplength=w)
        except Exception:
            pass

    def _on_resize(self, event) -> None:
        if event.widget is self and self._warn_shown:
            self._wrap_warn()

    def _on_option_change(self) -> None:
        layout_on = self.toggle_vars["layout"].get()
        for cb in self._style_checks:
            if str(cb.cget("text")) == "自动排版":
                continue
            try:
                cb.configure(state="normal" if layout_on else "disabled")
            except Exception:
                pass

        if self.toggle_vars["content"].get():
            if not self._warn_shown:
                self.content_warn.pack(fill="x", pady=(8, 0),
                                       before=self._quick_row)
                self._warn_shown = True
                self.after_idle(self._wrap_warn)
        elif self._warn_shown:
            self.content_warn.pack_forget()
            self._warn_shown = False

        parts = []
        if layout_on:
            n = sum(1 for k in ("font_family", "font_size", "font_color",
                                "background", "page_break") if self.toggle_vars[k].get())
            parts.append(f"版式 {n}/5")
        else:
            parts.append("不改版式")
        if self.toggle_vars["content"].get():
            parts.append("精简文字")
        if self.toggle_vars["image_size"].get():
            parts.append("适配图片")
        self.opt_summary.configure(text="当前：" + " · ".join(parts))

    # ------------------------------------------------------------------
    # 文件管理
    # ------------------------------------------------------------------
    def add_files(self) -> None:
        patterns = [("Word / Excel 文档", "*.docx *.docm *.xlsx *.xlsm"),
                    ("Word 文档", "*.docx *.docm"),
                    ("Excel 工作簿", "*.xlsx *.xlsm"),
                    ("所有文件", "*.*")]
        chosen = fd.askopenfilenames(title="选择要优化排版的文件", filetypes=patterns)
        if chosen:
            self._add_paths(chosen)

    def _refresh_count(self) -> None:
        n = len(self.files)
        self.file_count_lbl.configure(
            text="尚未添加文件" if not n else f"已选 {n} 个文件")
        self.status_lbl.configure(text="" if not n else f"待处理 {n} 个文件")

    def _add_paths(self, items) -> None:
        added = 0
        for path in items:
            path = os.path.abspath(path)
            if path in self.files or not os.path.isfile(path):
                continue
            ext = os.path.splitext(path)[1].lower()
            if ext not in ACCEPT_EXTS:
                self._log_ui(f"跳过（类型不支持）：{os.path.basename(path)}", "warn")
                continue
            self.files.append(path)
            self.tree.insert("", "end", iid=path,
                             values=(self._type_label(ext), os.path.basename(path),
                                     self._size_label(path)))
            added += 1
        if added:
            self._sync_out_dir()
            self._refresh_count()
            self._log_ui(f"已添加 {added} 个文件，共 {len(self.files)} 个")
            log.info("添加文件 %d 个（累计 %d）", added, len(self.files))

    @staticmethod
    def _type_label(ext: str) -> str:
        return "Word" if ext in (".docx", ".docm") else "Excel"

    @staticmethod
    def _size_label(path: str) -> str:
        try:
            size = os.path.getsize(path)
        except OSError:
            return "?"
        if size > 1024 * 1024:
            return f"{size / 1024 / 1024:.2f} MB"
        return f"{size / 1024:.0f} KB"

    def remove_selected(self) -> None:
        for iid in self.tree.selection():
            if iid in self.files:
                self.files.remove(iid)
            self.tree.delete(iid)
        self._sync_out_dir()
        self._refresh_count()

    def clear_files(self) -> None:
        for iid in self.tree.get_children():
            self.tree.delete(iid)
        self.files.clear()
        self._sync_out_dir()
        self._refresh_count()
        self.progress["value"] = 0
        self.progress_txt.configure(text="等待开始")

    def _open_selected_file(self, _event=None) -> None:
        sel = self.tree.selection()
        if sel:
            try:
                os.startfile(sel[0])  # noqa: S606 - 打开用户自己的文件
            except Exception as exc:
                mb.showwarning("无法打开", str(exc))

    # ------------------------------------------------------------------
    # 输出目录
    # ------------------------------------------------------------------
    def _sync_out_dir(self) -> None:
        if not self._auto_out:
            return
        if self.files:
            self.out_dir_var.set(paths.default_output_dir(self.files[0]))
        else:
            self.out_dir_var.set("")

    def choose_out_dir(self) -> None:
        initial = self.out_dir_var.get() or paths.app_dir()
        chosen = fd.askdirectory(title="选择输出目录", initialdir=initial)
        if chosen:
            self.out_dir_var.set(os.path.normpath(chosen))
            self._auto_out = False
            log.info("输出目录设为 %s", chosen)

    def open_out_dir(self) -> None:
        target = self.out_dir_var.get()
        if not target:
            target = paths.default_output_dir(self.files[0]) if self.files else paths.app_dir()
        os.makedirs(target, exist_ok=True)
        self._open_path(target)

    def open_logs(self) -> None:
        self._open_path(paths.logs_dir())

    def _open_path(self, target: str) -> None:
        try:
            if sys.platform == "win32":
                os.startfile(target)  # noqa: S606
            elif sys.platform == "darwin":
                subprocess.Popen(["open", target])
            else:
                subprocess.Popen(["xdg-open", target])
        except Exception as exc:
            mb.showwarning("无法打开", f"{target}\n{exc}")

    def _open_url(self, url: str) -> None:
        try:
            import webbrowser

            webbrowser.open(url)
        except Exception:
            pass

    # ------------------------------------------------------------------
    # 拖放
    # ------------------------------------------------------------------
    def _enable_dnd(self) -> None:
        try:
            if _bind_dnd(self):
                log.info("拖放功能已启用")
                self.drop_hint.configure(text="把文件拖到这里，或点击「添加文件」")
            else:
                raise RuntimeError("no widget accepted DnD")
        except Exception:
            log.info("拖放功能不可用，改用【添加文件】按钮")
            self._log_ui("提示：拖放不可用，请用「添加文件…」按钮选择", "muted")

    def _on_drop(self, event) -> None:
        try:
            items = self.tk.splitlist(event.data)
        except Exception:
            items = [event.data]
        self._add_paths(items)

    # ------------------------------------------------------------------
    # 执行
    # ------------------------------------------------------------------
    def start(self) -> None:
        if self._running:
            return
        if not self.files:
            mb.showinfo("请先选择文件", "请先添加至少一个 Word 或 Excel 文件。")
            return
        toggles = self._toggles_obj()
        if not toggles.any_change:
            mb.showinfo("没有可执行的修改",
                        "当前所有选项都处于「不改动」状态，请至少勾选一项。")
            return
        out_dir = self.out_dir_var.get().strip()
        if not out_dir:
            out_dir = paths.default_output_dir(self.files[0])
            self.out_dir_var.set(out_dir)
        tpl = self._current_template()
        options = OptimizeOptions(template_key=tpl.key, toggles=toggles)
        self._running = True
        self._results = []
        self.run_btn.configure(state="disabled", text="正在处理…")
        self.progress["value"] = 0
        self._log_ui("─" * 62, "muted")
        log.info("开始批量优化：%d 个文件 | 模板=%s | 开关=%s",
                 len(self.files), tpl.name, toggles.to_dict())

        files = list(self.files)
        self.worker = threading.Thread(
            target=self._run_worker, args=(files, tpl, options, out_dir), daemon=True)
        self.worker.start()

    def _toggles_obj(self):
        from ..templates.schema import ToggleOptions

        return ToggleOptions(**{k: v.get() for k, v in self.toggle_vars.items()})

    def _run_worker(self, files, tpl, options, out_dir) -> None:
        total = len(files)
        ok = 0
        failed = 0
        for idx, path in enumerate(files):
            base = idx / total * 100
            span = 100 / total

            def progress(pct, msg, _base=base, _span=span, _path=path):
                self.queue.put(("progress", _base + _span * pct / 100,
                                f"{os.path.basename(_path)} · {msg}"))

            self.queue.put(("log", f"[{idx + 1}/{total}] {os.path.basename(path)}", ""))
            try:
                result = optimize_file(path, tpl, options, out_dir, progress=progress)
                self.queue.put(("result", result, None))
                ok += 1
            except (UnsupportedFileError, OptimizeError) as exc:
                failed += 1
                self.queue.put(("error", path, str(exc)))
            except Exception as exc:  # 未预期异常：写崩溃日志但不中断整批
                failed += 1
                logging_setup.write_crash_log(
                    type(exc), exc, exc.__traceback__,
                    f"处理文件: {path}")
                self.queue.put(("error", path, f"{type(exc).__name__}: {exc}"))
        self.queue.put(("done", ok, failed))

    def _poll_queue(self) -> None:
        try:
            while True:
                kind, *payload = self.queue.get_nowait()
                if kind == "progress":
                    pct, msg = payload
                    self.progress["value"] = pct
                    self.progress_txt.configure(text=msg)
                elif kind == "log":
                    self._log_ui(payload[0])
                elif kind == "result":
                    self._on_result(payload[0])
                elif kind == "error":
                    path, msg = payload
                    self._log_ui(f"  × {os.path.basename(path)}：{msg}", "err")
                    log.error("优化失败 %s：%s", path, msg)
                elif kind == "done":
                    self._on_done(*payload)
        except queue.Empty:
            pass
        self.after(80, self._poll_queue)

    def _on_result(self, result) -> None:
        self._results.append(result)
        stat_txt = "，".join(f"{k} {v}" for k, v in result.stats.items() if v)
        self._log_ui(f"  ✓ {os.path.basename(result.target)}  [{result.template_name}]"
                     f" {result.elapsed:.2f}s", "ok")
        if stat_txt:
            self._log_ui(f"      {stat_txt}", "muted")
        for warn in result.warnings:
            self._log_ui(f"      ! {warn}", "warn")

    def _on_done(self, ok: int, failed: int) -> None:
        self._running = False
        self.run_btn.configure(state="normal", text="开始优化")
        self.progress["value"] = 100
        if failed == 0:
            self.progress_txt.configure(text=f"全部完成：成功 {ok} 个")
            self.status_lbl.configure(text=f"成功 {ok} 个", foreground=theme.SUCCESS)
            self._log_ui(f"完成：成功 {ok} 个。", "ok")
        else:
            self.progress_txt.configure(text=f"完成：成功 {ok} 个，失败 {failed} 个")
            self.status_lbl.configure(text=f"成功 {ok} / 失败 {failed}",
                                      foreground=theme.DANGER)
            self._log_ui(f"完成：成功 {ok} 个，失败 {failed} 个，详见上方错误行。", "warn")
        log.info("批量完成：成功 %d，失败 %d", ok, failed)
        if ok and failed == 0:
            if mb.askyesno("优化完成", f"已成功处理 {ok} 个文件。\n\n是否打开输出目录？"):
                self.open_out_dir()

    # ------------------------------------------------------------------
    # 日志显示
    # ------------------------------------------------------------------
    def _log_ui(self, text: str, tag: str = "") -> None:
        stamp = _dt.datetime.now().strftime("%H:%M:%S")
        self.log_text.configure(state="normal")
        self.log_text.insert("end", f"{stamp}  {text}\n", tag)
        self.log_text.see("end")
        self.log_text.configure(state="disabled")

    def _toggle_log(self) -> None:
        if self._log_visible.get():
            self.log_frame.pack(fill="x", pady=(4, 0))
        else:
            self.log_frame.pack_forget()

    def _copy_log(self) -> None:
        self.clipboard_clear()
        self.clipboard_append(self.log_text.get("1.0", "end").strip())
        self.status_lbl.configure(text="日志已复制")

    # ------------------------------------------------------------------
    def _on_crash(self, summary: str, detail) -> None:
        """全局未捕获异常：弹窗提示并指引日志位置。"""
        try:
            self._log_ui(f"程序异常：{summary}", "err")
            mb.showerror(
                "出现异常",
                f"程序遇到未处理的错误：\n{summary}\n\n"
                f"详细信息已写入日志：\n{os.path.join(paths.logs_dir(), 'crash.log')}\n\n"
                "请把该文件反馈给开发者。")
        except Exception:
            pass

    def _on_close(self) -> None:
        if self._running and not mb.askyesno("正在处理", "仍有文件在处理中，确定要退出吗？"):
            return
        log.info("界面关闭")
        self.destroy()


def main() -> None:
    app = App()
    app.mainloop()
