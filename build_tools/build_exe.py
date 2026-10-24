"""一键打包单文件 exe（Windows 10/11，免安装）。

用法:
    python build_tools/build_exe.py            # 打包
    python build_tools/build_exe.py --clean    # 先清理 build/dist
产物:
    dist/Word_Excel_Layout_Optimizer_v1.1.0.exe
"""

from __future__ import annotations

import os
import shutil
import subprocess
import sys

BASE = os.path.abspath(os.path.join(os.path.dirname(__file__), ".."))
sys.path.insert(0, os.path.join(BASE, "src"))
sys.path.insert(0, os.path.join(BASE, "build_tools"))

from wxfixer.version import APP_SLUG, APP_VERSION, full_title  # noqa: E402

# 文件名用纯 ASCII：GitHub Release 会把中文附件名替换成 "_"，导致下载链接对不上。
EXE_NAME = f"Word_Excel_Layout_Optimizer_v{APP_VERSION}"
SPEC = os.path.join(BASE, "Word_Excel_Fix.spec")


def clean() -> None:
    for name in ("build", "dist"):
        path = os.path.join(BASE, name)
        if os.path.isdir(path):
            shutil.rmtree(path, ignore_errors=True)
            print("removed", path)


def make_spec() -> str:
    """生成 .spec，写入版本资源、图标、tkinterdnd2 的 tkdnd 数据目录。"""
    import make_version_info

    version_file = make_version_info.main().replace("\\", "/")
    icon = os.path.join(BASE, "assets", "app.ico").replace("\\", "/")

    try:
        import tkinterdnd2

        dnd_dir = os.path.dirname(tkinterdnd2.__file__).replace("\\", "/")
        datas = f"[('{dnd_dir}/tkdnd', 'tkinterdnd2/tkdnd')]"
        hidden = "['tkinterdnd2']"
    except Exception:
        datas = "[]"
        hidden = "[]"

    spec = f'''# -*- mode: python ; coding: utf-8 -*-
# 由 build_tools/build_exe.py 自动生成，请勿手工修改
import os
from PyInstaller.utils.hooks import collect_data_files

datas = {datas} + collect_data_files('docx') + []
datas += [('{BASE.replace(chr(92), "/")}/assets', 'assets')]

a = Analysis(
    ['{BASE.replace(chr(92), "/")}/src/main.py'],
    pathex=['{BASE.replace(chr(92), "/")}/src'],
    binaries=[],
    datas=datas,
    hiddenimports={hidden} + ['wxfixer.selftest'],
    hookspath=[],
    runtime_hooks=[],
    excludes=['numpy', 'pandas', 'matplotlib', 'scipy', 'PIL.ImageQt', 'pytest'],
    noarchive=False,
)
pyz = PYZ(a.pure)

exe = EXE(
    pyz,
    a.scripts,
    a.binaries,
    a.datas,
    [],
    name='{EXE_NAME}',
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=False,
    upx_exclude=[],
    runtime_tmpdir=None,
    console=False,
    disable_windowed_traceback=False,
    argv_emulation=False,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
    icon='{icon}',
    version='{version_file}',
)
'''
    with open(SPEC, "w", encoding="utf-8") as fh:
        fh.write(spec)
    print("saved", SPEC)
    return SPEC


def build() -> int:
    spec = make_spec()
    cmd = [sys.executable, "-m", "PyInstaller", "--noconfirm", "--clean", spec]
    print(" ".join(cmd))
    proc = subprocess.run(cmd, cwd=BASE)
    if proc.returncode != 0:
        return proc.returncode

    dist = os.path.join(BASE, "dist")
    exe = os.path.join(dist, EXE_NAME + ".exe")
    if not os.path.isfile(exe):
        print("未找到产物:", exe)
        return 1
    size_mb = os.path.getsize(exe) / 1024 / 1024
    print(f"\n打包完成: {exe}  ({size_mb:.1f} MB)  {full_title()}")

    # 把说明文档一并放进 dist，便于整体拷贝分发
    for name in ("使用说明.md", "LICENSE", "CHANGELOG.md"):
        src = os.path.join(BASE, name)
        if os.path.isfile(src):
            shutil.copy2(src, dist)
    return 0


if __name__ == "__main__":
    if "--clean" in sys.argv:
        clean()
    sys.exit(build())
