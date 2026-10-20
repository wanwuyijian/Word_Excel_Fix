"""排版方案（模板）数据结构 + 选项开关。

设计原则：模板只描述"长什么样"，不描述"写什么"。
优化器读取 ``LayoutSpec`` 后逐项套用，不做任何文字层面的增删改。
"""

from __future__ import annotations

from dataclasses import dataclass, field, asdict
from typing import Dict, List, Optional


# ---------------------------------------------------------------------------
# 用户可勾选的 5 个交互开关
# ---------------------------------------------------------------------------
@dataclass
class ToggleOptions:
    """交互选项：每一项决定某一类属性是否允许被模板修改。"""

    font_family: bool = True       # 是否修改字体
    font_size: bool = True         # 是否修改字号
    font_color: bool = True        # 是否修改字体颜色
    background: bool = True        # 是否修改背景颜色
    page_break: bool = True        # 是否修改分页（标题前分页等）

    def to_dict(self) -> Dict[str, bool]:
        return asdict(self)

    @classmethod
    def from_dict(cls, data: Optional[dict]) -> "ToggleOptions":
        data = data or {}
        default = cls()
        return cls(
            font_family=bool(data.get("font_family", default.font_family)),
            font_size=bool(data.get("font_size", default.font_size)),
            font_color=bool(data.get("font_color", default.font_color)),
            background=bool(data.get("background", default.background)),
            page_break=bool(data.get("page_break", default.page_break)),
        )


# ---------------------------------------------------------------------------
# 字体设置
# ---------------------------------------------------------------------------
@dataclass
class FontSpec:
    """一套字体设置。字号单位：Word 用磅(pt)，Excel 用磅(pt)。"""

    east_asia: str = "宋体"          # 中文字体
    ascii_font: str = "Times New Roman"  # 西文字体
    size_pt: float = 12.0
    bold: bool = False
    color: Optional[str] = None       # 形如 "000000"，None 表示不改
    underline: bool = False
    # 是否强制覆盖：True = 无条件套用；False = 保留原有个性化设置
    force: bool = False


# ---------------------------------------------------------------------------
# 完整排版方案
# ---------------------------------------------------------------------------
@dataclass
class LayoutSpec:
    # ---- 基本信息 ----
    key: str = "auto"
    name: str = "自动识别"
    category: str = "word"            # "word" | "excel" | "both"
    description: str = ""
    tips: List[str] = field(default_factory=list)

    # ---- Word：页面 ----
    page_margin_cm: Dict[str, float] = field(
        default_factory=lambda: {"top": 2.54, "bottom": 2.54, "left": 3.18, "right": 3.18}
    )
    orientation: Optional[str] = None  # None=保持, "portrait", "landscape"
    header_distance_cm: Optional[float] = None
    footer_distance_cm: Optional[float] = None

    # ---- Word：正文 / 标题 / 各级样式 ----
    body: FontSpec = field(default_factory=lambda: FontSpec(size_pt=12.0))
    line_spacing: float = 1.5          # 行距倍数
    space_before_pt: float = 0.0
    space_after_pt: float = 0.0
    first_line_indent_chars: float = 2.0   # 首行缩进（字符）
    align: str = "justify"             # left / center / right / justify
    # 各级标题：level -> FontSpec（含对齐、段前段后由下面的字典给）
    headings: Dict[int, FontSpec] = field(default_factory=dict)
    heading_space: Dict[int, tuple] = field(default_factory=dict)  # level -> (before, after)
    heading_keep_with_next: bool = True
    heading_page_break_levels: List[int] = field(default_factory=list)  # 哪些级别标题前分页
    heading_numbering: bool = False    # 是否给标题加自动编号（仅改样式层，不改正文）

    # ---- Word：表格 ----
    table_font_size: Optional[float] = None
    table_header_bold: bool = True
    table_header_bg: str = "D9E2F3"
    table_header_color: Optional[str] = None
    table_header_repeat: bool = True
    table_row_cant_split: bool = True
    table_zebra: bool = False
    table_zebra_bg: str = "F2F5FA"
    table_border_color: str = "A6A6A6"
    table_border_size_eighths: int = 4   # 1/8 磅；4 = 0.5pt
    table_autofit_window: bool = True
    table_center: bool = True
    table_caption_center: bool = True

    # ---- Word：列表 ----
    list_indent_cm: float = 0.74
    list_hanging_cm: float = 0.74

    # ---- Word：页眉页脚 ----
    footer_page_number: bool = True
    header_text: Optional[str] = None   # None = 不改
    title_rule: bool = True             # 文档大标题下是否加一条居中的分隔横线

    # ---- Excel：表格外观 ----
    excel_font: FontSpec = field(default_factory=lambda: FontSpec(size_pt=11.0))
    excel_header_font: FontSpec = field(
        default_factory=lambda: FontSpec(size_pt=11.0, bold=True, color="FFFFFF")
    )
    excel_header_bg: str = "4472C4"
    excel_header_fg: str = "FFFFFF"
    excel_body_fg: str = "000000"
    excel_zebra: bool = True
    excel_zebra_bg: str = "EEF3FB"
    excel_border_color: str = "BFBFBF"
    excel_border_style: str = "thin"
    excel_border_outer_style: str = "medium"
    excel_align_h: str = "auto"          # auto: 数字右、文本左、表头居中
    excel_wrap_text: bool = True
    excel_row_height_auto: bool = True
    excel_freeze_header: bool = True
    excel_autofilter: bool = True
    excel_column_fit: int = 1            # 0=不改宽 1=适配内容 2=适配并压缩超宽列
    excel_max_col_width: float = 50.0
    excel_header_repeat: bool = True     # 打印时表头每页重复

    # ---- Excel：页面与打印 ----
    excel_margins_cm: Dict[str, float] = field(
        default_factory=lambda: {"top": 1.91, "bottom": 1.91, "left": 1.78, "right": 1.78,
                                 "header": 0.76, "footer": 0.76}
    )
    excel_orientation: Optional[str] = "auto"   # auto / portrait / landscape
    excel_paper_size: int = 9                   # 9 = A4
    excel_fit_to_width: bool = True
    excel_fit_width_pages: int = 1
    excel_fit_height_pages: int = 0             # 0 = 自动
    excel_print_title_rows: Optional[str] = None  # "1:1"
    excel_outline: bool = True                  # 加细边框
    excel_footer_page_number: bool = True
    excel_tab_color: Optional[str] = None

    # ---- 通用 ----
    title_page: bool = False            # 是否在文首插入分页（不动文档内容）
    preferred_landscape_chars: int = 0  # 内容平均行长超过该值倾向横向

    # 元信息
    version: str = "1.0.0"

    # ------------------------------------------------------------------
    def heading_font(self, level: int) -> Optional[FontSpec]:
        if level in self.headings:
            return self.headings[level]
        # 回落：找最近的更高级别
        for lv in sorted(self.headings):
            if lv <= level:
                return self.headings[lv]
        return None

    def heading_spacing(self, level: int) -> tuple:
        if level in self.heading_space:
            return self.heading_space[level]
        for lv in sorted(self.heading_space, reverse=True):
            if lv <= level:
                return self.heading_space[lv]
        return (12.0, 6.0)


def font(family: str, size: float, bold: bool = False, color: Optional[str] = None,
         ascii_font: str = "Times New Roman") -> FontSpec:
    return FontSpec(
        east_asia=family,
        ascii_font=ascii_font,
        size_pt=size,
        bold=bold,
        color=color,
    )
