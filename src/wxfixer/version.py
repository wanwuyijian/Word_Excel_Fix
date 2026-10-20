"""版本信息唯一来源。

升级版本时只需要修改这里，其余模块（GUI 标题、日志、exe 元数据、关于窗口）
都从本文件读取，避免版本号分散在多处导致不一致。
"""

APP_NAME = "Word/Excel 排版优化工具"
APP_NAME_EN = "Word & Excel Layout Optimizer"
APP_SLUG = "Word_Excel_Fix"

# 语义化版本：主版本.次版本.修订号
APP_VERSION_TUPLE = (1, 0, 0)
APP_VERSION = ".".join(str(x) for x in APP_VERSION_TUPLE)
BUILD_DATE = "2026-10-20"

__version__ = APP_VERSION

# GitHub 仓库
GITHUB_REPO = "wanwuyijian/Word_Excel_Fix"
GITHUB_URL = f"https://github.com/{GITHUB_REPO}"


def full_title() -> str:
    return f"{APP_NAME} v{APP_VERSION}"


def version_info_tuple() -> tuple:
    """用于 PyInstaller / Windows 文件属性的四段式版本号。"""
    major, minor, patch = APP_VERSION_TUPLE
    return (major, minor, patch, 0)
