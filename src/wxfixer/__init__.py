"""Word & Excel 排版优化工具 (Word/Excel Layout Optimizer).

一个纯排版优化工具：只调整字体、字号、颜色、行距、分页、单元格样式等
"表现形式"，绝不增删、改写文档的实际内容。
"""

from .version import (
    APP_NAME,
    APP_NAME_EN,
    APP_VERSION,
    APP_VERSION_TUPLE,
    BUILD_DATE,
    __version__,
)

__all__ = [
    "APP_NAME",
    "APP_NAME_EN",
    "APP_VERSION",
    "APP_VERSION_TUPLE",
    "BUILD_DATE",
    "__version__",
]
