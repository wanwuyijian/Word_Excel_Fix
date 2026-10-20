"""主窗口。

布局：
┌───────────────────────────────────────────────────────┐
│ 标题栏（应用名 + 版本）                                 │
├──────────────────────────┬────────────────────────────┤
│ ① 文件（拖入/选择/列表） │ ③ 排版模板（8 选一 + 说明） │
│ ② 交互选项（5 个开关）   │ ④ 输出与进度                │
├──────────────────────────┴────────────────────────────┤
│ 运行日志（可折叠）                                      │
└───────────────────────────────────────────────────────┘
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
from ..utils import logging_setup, paths
from ..version import APP_NAME, APP_VERSION, BUILD_DATE, GITHUB_URL, full_title
from . import theme

log = logging_setup.operations_logger()

ACCEPT_EXTS = (".docx", ".docm", ".xlsx", ".xlsm")


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
        self.geometry("1180x820")
        self.minsize(1000, 720)
        self.configure(bg=theme.BG)
        # 最大化 / 最小化：系统标题栏按钮即可，另外保证窗口可自由缩放
        self.resizable(True, True)
        self._center()

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

        self.tpl_key_var = tk.StringVar(value=ALL_TEMPLATES[0].name)
        self.toggle_vars = {
            "font_family": tk.BooleanVar(value=True),
            "font_size": tk.BooleanVar(value=True),
            "font_color": tk.BooleanVar(value=True),
            "background": tk.BooleanVar(value=True),
            "page_break": tk.BooleanVar(value=True),
        }

        self._build()
        self._enable_dnd()
        self.protocol("WM_DELETE_WINDOW", self._on_close)
        self.after(80, self._poll_queue)
        logging_setup.set_crash_notifier(self._on_crash)
        self._log_ui(f"就绪 · 日志目录 {paths.logs_dir()}")
        log.info("界面已启动")

    # ------------------------------------------------------------------
    # 布局
    # ------------------------------------------------------------------
    def _center(self) -> None:
        self.update_idletasks()
        w, h = 1180, 820
        sw, sh = self.winfo_screenwidth(), self.winfo_screenheight()
        self.geometry(f"{w}x{h}+{max(0, (sw - w) // 2)}+{max(0, (sh - h) // 3)}")

    def _build(self) -> None:
        outer = ttk.Frame(self, padding=(16, 12, 16, 10))
        outer.pack(fill="both", expand=True)

        header = ttk.Frame(outer)
        header.pack(fill="x", pady=(0, 10))
        ttk.Label(header, text=APP_NAME, style="H1.TLabel").pack(side="left")
        ttk.Label(header, text=f"v{APP_VERSION}", style="Muted.TLabel").pack(
            side="left", padx=(8, 0), pady=(6, 0))
        ttk.Label(header, text="纯排版优化 · 内容零改动", style="Muted.TLabel").pack(
            side="right", pady=(6, 0))

        body = ttk.Frame(outer)
        body.pack(fill="both", expand=True)
        body.columnconfigure(0, weight=3, uniform="col")
        body.columnconfigure(1, weight=2, uniform="col")
        body.rowconfigure(0, weight=1)

        self._build_left(body)
        self._build_right(body)
        self._build_log(outer)
        self._build_statusbar(outer)

    # ---- 左列 ----
    def _build_left(self, parent) -> None:
        left = ttk.Frame(parent)
        left.grid(row=0, column=0, sticky="nsew", padx=(0, 12))
        left.rowconfigure(1, weight=3)
        left.rowconfigure(2, weight=2)
        left.columnconfigure(0, weight=1)

        # ① 文件
        card = ttk.Frame(left, style="Card.TFrame", padding=14)
        card.grid(row=1, column=0, sticky="nsew", pady=(0, 10))
        ttk.Label(card, text="① 选择文件", style="H2.TLabel").pack(anchor="w")
        self.drop_hint = ttk.Label(
            card, text="支持 Word（.docx/.docm）与 Excel（.xlsx/.xlsm），可多选或直接拖入窗口",
            style="CardMuted.TLabel")
        self.drop_hint.pack(anchor="w", pady=(2, 8))

        btns = ttk.Frame(card, style="Card.TFrame")
        btns.pack(fill="x")
        ttk.Button(btns, text="添加文件…", command=self.add_files).pack(side="left")
        ttk.Button(btns, text="移除选中", command=self.remove_selected).pack(side="left", padx=6)
        ttk.Button(btns, text="清空", command=self.clear_files).pack(side="left")

        tree_wrap = ttk.Frame(card, style="Card.TFrame")
        tree_wrap.pack(fill="both", expand=True, pady=(10, 0))
        cols = ("type", "name", "size")
        self.tree = ttk.Treeview(tree_wrap, columns=cols, show="headings",
                                 selectmode="extended", height=7)
        self.tree.heading("type", text="类型")
        self.tree.heading("name", text="文件名")
        self.tree.heading("size", text="大小")
        self.tree.column("type", width=76, anchor="center", stretch=False)
        self.tree.column("name", width=360, anchor="w")
        self.tree.column("size", width=90, anchor="e", stretch=False)
        vsb = ttk.Scrollbar(tree_wrap, orient="vertical", command=self.tree.yview)
        self.tree.configure(yscrollcommand=vsb.set)
        self.tree.pack(side="left", fill="both", expand=True)
        vsb.pack(side="right", fill="y")
        self.tree.bind("<Double-1>", self._open_selected_file)
        self.tree.bind("<Delete>", lambda e: self.remove_selected())

        # ② 交互选项
        card2 = ttk.Frame(left, style="Card.TFrame", padding=14)
        card2.grid(row=2, column=0, sticky="nsew")
        ttk.Label(card2, text="② 交互选项", style="H2.TLabel").pack(anchor="w")
        ttk.Label(card2, text="取消勾选表示「保持原样，不做这类修改」",
                  style="CardMuted.TLabel").pack(anchor="w", pady=(2, 8))
        grid = ttk.Frame(card2, style="Card.TFrame")
        grid.pack(fill="x")
        items = [
            ("font_family", "修改字体（中英文字体族）"),
            ("font_size", "修改字号"),
            ("font_color", "修改字体颜色"),
            ("background", "修改背景颜色（底纹/填充）"),
            ("page_break", "允许调整分页（标题另起页）"),
        ]
        for i, (key, text) in enumerate(items):
            ttk.Checkbutton(grid, text=text, variable=self.toggle_vars[key],
                            style="TCheckbutton").grid(
                row=i // 2, column=i % 2, sticky="w", padx=(0, 24), pady=4)
        quick = ttk.Frame(card2, style="Card.TFrame")
        quick.pack(fill="x", pady=(8, 0))
        ttk.Button(quick, text="全选", command=lambda: self._set_toggles(True)).pack(side="left")
        ttk.Button(quick, text="全不选", command=lambda: self._set_toggles(False)).pack(
            side="left", padx=6)

    # ---- 右列 ----
    def _build_right(self, parent) -> None:
        right = ttk.Frame(parent)
        right.grid(row=0, column=1, sticky="nsew")
        right.columnconfigure(0, weight=1)
        right.rowconfigure(0, weight=3)
        right.rowconfigure(1, weight=2)

        card = ttk.Frame(right, style="Card.TFrame", padding=14)
        card.grid(row=0, column=0, sticky="nsew", pady=(0, 10))
        ttk.Label(card, text="③ 排版模板", style="H2.TLabel").pack(anchor="w")
        ttk.Label(card, text="自动识别内容结构，仅调整版式与视觉层次",
                  style="CardMuted.TLabel").pack(anchor="w", pady=(2, 8))

        self.tpl_combo = ttk.Combobox(
            card, values=[t.name for t in ALL_TEMPLATES],
            textvariable=self.tpl_key_var, state="readonly")
        self.tpl_combo.pack(fill="x")
        self.tpl_combo.bind("<<ComboboxSelected>>", lambda e: self._show_template_info())

        info = ttk.Frame(card, style="Card.TFrame")
        info.pack(fill="both", expand=True, pady=(10, 0))
        self.tpl_desc = tk.Text(info, height=9, wrap="word", bd=0, relief="flat",
                                bg="#F8FAFD", fg=theme.TEXT, font=theme.ui_font(9),
                                padx=10, pady=8, cursor="arrow")
        self.tpl_desc.pack(fill="both", expand=True)
        self._show_template_info()
        self.tpl_desc.configure(state="disabled")

        card2 = ttk.Frame(right, style="Card.TFrame", padding=14)
        card2.grid(row=1, column=0, sticky="nsew")
        ttk.Label(card2, text="④ 输出与执行", style="H2.TLabel").pack(anchor="w")
        ttk.Label(card2, text="输出为新文件，源文件不会被修改或覆盖",
                  style="CardMuted.TLabel").pack(anchor="w", pady=(2, 8))

        row = ttk.Frame(card2, style="Card.TFrame")
        row.pack(fill="x")
        self.out_entry = ttk.Entry(row, textvariable=self.out_dir_var)
        self.out_entry.pack(side="left", fill="x", expand=True)
        ttk.Button(row, text="浏览…", command=self.choose_out_dir).pack(side="left", padx=(6, 0))
        ttk.Button(row, text="打开", command=self.open_out_dir).pack(side="left", padx=(6, 0))

        self.progress = ttk.Progressbar(card2, style="Blue.Horizontal.TProgressbar",
                                        maximum=100, value=0)
        self.progress.pack(fill="x", pady=(12, 4))
        self.progress_txt = ttk.Label(card2, text="等待开始", style="CardMuted.TLabel")
        self.progress_txt.pack(anchor="w")

        action = ttk.Frame(card2, style="Card.TFrame")
        action.pack(fill="x", pady=(10, 0))
        self.run_btn = ttk.Button(action, text="开始排版优化", style="Primary.TButton",
                                  command=self.start)
        self.run_btn.pack(side="left")
        ttk.Button(action, text="打开日志目录", command=self.open_logs).pack(side="right")

    # ---- 日志与状态栏 ----
    def _build_log(self, parent) -> None:
        bar = ttk.Frame(parent)
        bar.pack(fill="x", pady=(8, 0))
        ttk.Checkbutton(bar, text="显示运行日志", variable=self._log_visible,
                        command=self._toggle_log).pack(side="left")
        ttk.Button(bar, text="打开输出目录", style="Link.TButton",
                   command=self.open_out_dir).pack(side="right")
        ttk.Button(bar, text="复制日志", style="Link.TButton",
                   command=self._copy_log).pack(side="right", padx=(0, 8))

        self.log_frame = ttk.Frame(parent, style="Card.TFrame", padding=(1, 1))
        self.log_frame.pack(fill="x", pady=(4, 0))
        self.log_text = tk.Text(self.log_frame, height=7, wrap="word", bd=0, relief="flat",
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
        bar.pack(fill="x", pady=(6, 0))
        ttk.Label(bar, text=f"版本 {APP_VERSION}（{BUILD_DATE}）", style="Muted.TLabel").pack(side="left")
        link = ttk.Label(bar, text=GITHUB_URL, style="Muted.TLabel",
                         foreground=theme.PRIMARY, cursor="hand2")
        link.pack(side="right")
        link.bind("<Button-1>", lambda e: self._open_url(GITHUB_URL))
        self.status_lbl = ttk.Label(bar, text="", style="Muted.TLabel")
        self.status_lbl.pack(side="right", padx=16)

    # ------------------------------------------------------------------
    # 模板说明
    # ------------------------------------------------------------------
    def _current_template(self):
        name = self.tpl_key_var.get()
        for t in ALL_TEMPLATES:
            if t.name == name:
                return t
        return ALL_TEMPLATES[0]

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
            self._log_ui(f"已添加 {added} 个文件，共 {len(self.files)} 个")
            log.info("添加文件 %d 个（累计 %d）", added, len(self.files))
            self.status_lbl.configure(text=f"待处理 {len(self.files)} 个文件")

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
        self.status_lbl.configure(text=f"待处理 {len(self.files)} 个文件")

    def clear_files(self) -> None:
        for iid in self.tree.get_children():
            self.tree.delete(iid)
        self.files.clear()
        self._sync_out_dir()
        self.progress["value"] = 0
        self.progress_txt.configure(text="等待开始")
        self.status_lbl.configure(text="")

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
                self.drop_hint.configure(text="可拖入文件，或点击「添加文件…」")
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
        out_dir = self.out_dir_var.get().strip()
        if not out_dir:
            out_dir = paths.default_output_dir(self.files[0])
            self.out_dir_var.set(out_dir)
        tpl = self._current_template()
        options = OptimizeOptions(
            template_key=tpl.key,
            toggles=self._toggles_obj(),
        )
        self._running = True
        self._results = []
        self.run_btn.configure(state="disabled", text="正在处理…")
        self.progress["value"] = 0
        self._log_ui("─" * 60, "muted")
        log.info("开始批量优化：%d 个文件 | 模板=%s | 开关=%s",
                 len(self.files), tpl.name, options.toggles.to_dict())

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
                import traceback

                logging_setup.write_crash_log(
                    type(exc), exc, exc.__traceback__,
                    f"处理文件: {path}")
                self.queue.put(("error", path, f"{type(exc).__name__}: {exc}"))
                _ = traceback
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
        self.run_btn.configure(state="normal", text="开始排版优化")
        self.progress["value"] = 100
        if failed == 0:
            self.progress_txt.configure(text=f"全部完成：成功 {ok} 个")
            self.status_lbl.configure(text=f"成功 {ok} 个", foreground=theme.SUCCESS)
            self._log_ui(f"完成：成功 {ok} 个，全部文件未改动原文内容。", "ok")
        else:
            self.progress_txt.configure(text=f"完成：成功 {ok} 个，失败 {failed} 个")
            self.status_lbl.configure(text=f"成功 {ok} / 失败 {failed}", foreground=theme.DANGER)
            self._log_ui(f"完成：成功 {ok} 个，失败 {failed} 个，详见上方错误行。", "warn")
        log.info("批量完成：成功 %d，失败 %d", ok, failed)
        if ok and failed == 0:
            if mb.askyesno("排版完成", f"已成功优化 {ok} 个文件。\n\n是否打开输出目录？"):
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
