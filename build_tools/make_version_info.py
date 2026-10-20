"""生成 PyInstaller 用的版本信息资源文件 build/version_info.txt。"""

from __future__ import annotations

import os
import sys

BASE = os.path.abspath(os.path.join(os.path.dirname(__file__), ".."))
sys.path.insert(0, os.path.join(BASE, "src"))

from wxfixer.version import (  # noqa: E402
    APP_NAME, APP_NAME_EN, APP_VERSION, APP_VERSION_TUPLE,
    GITHUB_URL, full_title,
)

TEMPLATE = """# UTF-8
VSVersionInfo(
  ffi=FixedFileInfo(
    filevers={tuple4},
    prodvers={tuple4},
    mask=0x3f,
    flags=0x0,
    OS=0x40004,
    fileType=0x1,
    subtype=0x0,
    date=(0, 0)
  ),
  kids=[
    StringFileInfo([
      StringTable('080404B0', [
        StringStruct('CompanyName', ''),
        StringStruct('FileDescription', '{desc}'),
        StringStruct('FileVersion', '{version}'),
        StringStruct('InternalName', 'Word_Excel_Fix'),
        StringStruct('LegalCopyright', 'MIT License'),
        StringStruct('OriginalFilename', 'Word_Excel_Fix.exe'),
        StringStruct('ProductName', '{name}'),
        StringStruct('ProductVersion', '{version}'),
        StringStruct('Comments', '{url}')
      ])
    ]),
    VarFileInfo([VarStruct('Translation', [2052, 1200])])
  ]
)
"""


def main() -> str:
    out_dir = os.path.join(BASE, "build")
    os.makedirs(out_dir, exist_ok=True)
    path = os.path.join(out_dir, "version_info.txt")
    text = TEMPLATE.format(
        tuple4=APP_VERSION_TUPLE + (0,),
        desc=f"{APP_NAME} {APP_NAME_EN}",
        version=APP_VERSION,
        name=full_title(),
        url=GITHUB_URL,
    )
    with open(path, "w", encoding="utf-8") as fh:
        fh.write(text)
    print("saved", path)
    return path


if __name__ == "__main__":
    main()
