"""程序入口。

命令行参数：
    --selftest        不开界面，跑一遍离线自检（用于打包后验证 exe 是否可用）
    --version         打印版本号
    --console         日志同时输出到控制台
"""

from __future__ import annotations

import os
import sys

# 允许直接以脚本方式运行（源码模式）
_SRC = os.path.dirname(os.path.abspath(__file__))
if _SRC not in sys.path and not getattr(sys, "frozen", False):
    sys.path.insert(0, _SRC)


def _write(text: str) -> None:
    """向控制台输出（无控制台的窗口程序下自动忽略）。"""
    for stream in (sys.stdout, sys.stderr):
        if stream is None:
            continue
        try:
            stream.write(text + "\n")
            stream.flush()
            return
        except Exception:
            continue


def main(argv: list[str] | None = None) -> int:
    argv = list(sys.argv[1:] if argv is None else argv)
    console = "--console" in argv
    argv = [a for a in argv if a != "--console"]

    from wxfixer.utils import logging_setup

    logging_setup.setup_logging(console=console)
    logging_setup.log_environment()

    if "--version" in argv:
        from wxfixer.version import full_title

        _write(full_title())
        return 0

    if "--selftest" in argv:
        from wxfixer.selftest import main as selftest_main

        return selftest_main(argv, writer=_write)

    from wxfixer.gui import main as gui_main

    gui_main()
    return 0


if __name__ == "__main__":
    sys.exit(main())
