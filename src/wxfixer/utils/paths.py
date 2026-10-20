"""运行目录探测。

单文件 exe（PyInstaller onefile）运行时 ``sys.executable`` 指向临时解包目录，
程序退出即被删除。因此"可写数据"（日志、配置文件）必须放在 **exe 所在目录**，
即用 ``sys.argv[0]`` 推导。
"""

from __future__ import annotations

import os
import sys
import tempfile


def is_frozen() -> bool:
    return bool(getattr(sys, "frozen", False))


def app_dir() -> str:
    """exe（或源码入口）所在目录——用户可见、可持久写入。"""
    if is_frozen():
        # onefile 模式下 sys.executable 是临时目录，sys.argv[0] 才是真实 exe 路径
        return os.path.dirname(os.path.abspath(sys.argv[0]))
    # 源码运行：src/wxfixer/utils/paths.py -> 项目根
    here = os.path.dirname(os.path.abspath(__file__))
    return os.path.abspath(os.path.join(here, "..", "..", ".."))


def bundle_dir() -> str:
    """打包进 exe 的只读资源目录。"""
    meipass = getattr(sys, "_MEIPASS", None)
    if meipass:
        return meipass
    return app_dir()


def _ensure_writable(target: str) -> str | None:
    """尝试创建目录并写入探测文件，失败返回 None。"""
    try:
        os.makedirs(target, exist_ok=True)
        probe = os.path.join(target, ".write_probe")
        with open(probe, "w", encoding="utf-8") as fh:
            fh.write("ok")
        os.remove(probe)
        return target
    except OSError:
        return None


def data_dir() -> str:
    """日志/配置根目录，优先放在 exe 同目录；不可写时退回到用户目录。"""
    preferred = os.path.join(app_dir(), "data")
    ok = _ensure_writable(preferred)
    if ok:
        return ok
    fallback = os.path.join(
        os.path.expanduser("~"), ".word_excel_fix", "data"
    )
    ok = _ensure_writable(fallback)
    if ok:
        return ok
    return os.path.join(tempfile.gettempdir(), "word_excel_fix_data")


def logs_dir() -> str:
    path = os.path.join(data_dir(), "logs")
    os.makedirs(path, exist_ok=True)
    return path


def config_dir() -> str:
    path = os.path.join(data_dir(), "config")
    os.makedirs(path, exist_ok=True)
    return path


def default_output_dir(input_path: str) -> str:
    """默认输出目录：输入文件同级的 ``排版优化输出`` 文件夹。"""
    base = os.path.dirname(os.path.abspath(input_path)) if input_path else app_dir()
    return os.path.join(base, "排版优化输出")
