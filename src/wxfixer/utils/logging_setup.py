"""日志系统。

三类日志，全部落在 ``<exe目录>/data/logs`` 下：

* ``app.log``            —— 全部运行细节（DEBUG 及以上，按天滚动）
* ``operations.log``     —— 用户操作流水：打开了哪个文件、选了哪个模板、
                            产出了什么文件、耗时多少（便于事后追溯）
* ``crash.log``          —— 未捕获异常 / 崩溃堆栈（threading.excepthook +
                            sys.excepthook + tkinter 回调异常）

同时注册全局异常钩子，保证任何未处理异常都会先写盘再报错。
"""

from __future__ import annotations

import datetime as _dt
import logging
import logging.handlers
import os
import platform
import sys
import threading
import traceback

from . import paths

_LOG_FORMAT = "%(asctime)s | %(levelname)-7s | %(threadName)-12s | %(name)s | %(message)s"
_DATE_FORMAT = "%Y-%m-%d %H:%M:%S"

_INITIALISED = False
_CRASH_ERROR_SINK = None  # 可选回调：GUI 用来弹窗提示


class _SafeRotatingHandler(logging.handlers.RotatingFileHandler):
    """写日志失败时静默降级，绝不因为日志问题让主程序崩溃。"""

    def handleError(self, record):  # noqa: D401 - logging 钩子
        pass


def _make_handler(filename: str, level: int) -> logging.Handler:
    handler = _SafeRotatingHandler(
        os.path.join(paths.logs_dir(), filename),
        maxBytes=5 * 1024 * 1024,
        backupCount=5,
        encoding="utf-8",
    )
    handler.setLevel(level)
    handler.setFormatter(logging.Formatter(_LOG_FORMAT, _DATE_FORMAT))
    return handler


def setup_logging(console: bool = False) -> None:
    """初始化日志。重复调用无副作用。"""
    global _INITIALISED
    if _INITIALISED:
        return

    root = logging.getLogger()
    root.setLevel(logging.DEBUG)
    for old in list(root.handlers):
        root.removeHandler(old)

    root.addHandler(_make_handler("app.log", logging.DEBUG))

    if console:
        for stream in (sys.stdout, sys.stderr):
            if stream is None:
                continue
            try:
                handler = logging.StreamHandler(stream=stream)
            except Exception:
                continue
            handler.setLevel(logging.INFO)
            handler.setFormatter(logging.Formatter("%(levelname)s | %(message)s"))
            root.addHandler(handler)

    # 操作日志独立通道，只记录用户动作
    ops = logging.getLogger("wxfixer.operations")
    ops.setLevel(logging.INFO)
    ops.propagate = True
    ops.addHandler(_make_handler("operations.log", logging.INFO))

    _install_exception_hooks()
    _INITIALISED = True

    log = logging.getLogger("wxfixer.startup")
    log.info("=" * 78)
    log.info(
        "启动 %s | Python %s | %s %s | frozen=%s",
        _app_version(),
        platform.python_version(),
        platform.system(),
        platform.release(),
        getattr(sys, "frozen", False),
    )
    log.info("程序目录: %s", paths.app_dir())
    log.info("日志目录: %s", paths.logs_dir())


def _app_version() -> str:
    try:
        from ..version import APP_VERSION

        return APP_VERSION
    except Exception:  # pragma: no cover - 防御性
        return "?"


def operations_logger() -> logging.Logger:
    return logging.getLogger("wxfixer.operations")


def set_crash_notifier(sink) -> None:
    """注册崩溃时的 GUI 提示回调（可选）。"""
    global _CRASH_ERROR_SINK
    _CRASH_ERROR_SINK = sink


def write_crash_log(exc_type, exc_value, exc_tb, extra: str = "") -> str:
    """把异常堆栈写入 crash.log，返回粗体文本供界面展示。"""
    crash_logger = logging.getLogger("wxfixer.crash")
    if not crash_logger.handlers:
        crash_logger.addHandler(_make_handler("crash.log", logging.DEBUG))

    text = "".join(traceback.format_exception(exc_type, exc_value, exc_tb))
    header = (
        f"\n{'=' * 78}\n"
        f"[崩溃] {_dt.datetime.now().strftime(_DATE_FORMAT)} | "
        f"版本 {_app_version()} | {platform.platform()}\n"
        f"{extra}\n{'-' * 78}\n"
    )
    crash_logger.error("%s%s", header, text)
    return f"{type(exc_value).__name__}: {exc_value}"


def _install_exception_hooks() -> None:
    def _sys_hook(exc_type, exc_value, exc_tb):
        if issubclass(exc_type, KeyboardInterrupt):
            sys.__excepthook__(exc_type, exc_value, exc_tb)
            return
        summary = write_crash_log(exc_type, exc_value, exc_tb, "来源: sys.excepthook")
        if _CRASH_ERROR_SINK:
            try:
                _CRASH_ERROR_SINK(summary, traceback.format_exception(exc_type, exc_value, exc_tb))
            except Exception:  # pragma: no cover - 提示器自身不能再抛
                pass

    def _thread_hook(args):
        if args.exc_type is SystemExit:
            return
        summary = write_crash_log(
            args.exc_type,
            args.exc_value,
            args.exc_traceback,
            f"来源: threading.excepthook | 线程: {args.thread.name if args.thread else '?'}",
        )
        if _CRASH_ERROR_SINK:
            try:
                _CRASH_ERROR_SINK(
                    summary,
                    traceback.format_exception(
                        args.exc_type, args.exc_value, args.exc_traceback
                    ),
                )
            except Exception:  # pragma: no cover
                pass

    sys.excepthook = _sys_hook
    threading.excepthook = _thread_hook


def log_environment() -> None:
    """记录目标机器环境，便于远程排障。"""
    log = logging.getLogger("wxfixer.env")
    try:
        import docx  # noqa: F401
        import openpyxl
        import lxml.etree as etree

        log.info(
            "依赖版本: openpyxl=%s lxml=%s",
            openpyxl.__version__,
            etree.LXML_VERSION,
        )
    except Exception as exc:  # pragma: no cover
        log.warning("依赖版本读取失败: %s", exc)
