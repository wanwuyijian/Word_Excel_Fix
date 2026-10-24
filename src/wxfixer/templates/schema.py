"""排版方案（模板）数据结构 + 选项开关。

设计原则：模板只描述"长什么样"，不描述"写什么"。
优化器读取 ``LayoutSpec`` 后逐项套用，不做任何文字层面的增删改。
"""

from __future__ import annotations

from dataclasses import dataclass, field, asdict
from typing import Dict, List, Optional


# ---------------------------------------------------------------------------
# 用户可勾选的交互开关
#
# 分两组：
#   ① 版式组（layout_gate 之下）——只动"长什么样"
#   ② 内容与素材组——会动文字 / 图片，属于"改文件"的能力，默认更保守
# ---------------------------------------------------------------------------
@dataclass
class ToggleOptions:
    """交互选项：每一项决定某一类属性是否允许被模板修改。"""

    # ---- 版式组 ----
    font_family: bool = True       # 是否修改字体
    font_size: bool = True         # 是否修改字号
    font_color: bool = True        # 是否修改字体颜色
    background: bool = True        # 是否修改背景颜色
    page_break: bool = True        # 是否修改分页（标题前分页等）

    # ---- 版式总开关 ----
    layout: bool = True            # 是否允许自动排版（关闭=完全保持原版式）

    # ---- 内容与素材组 ----
    content: bool = False          # 是否优化文字内容（精简赘语，会改动文字）
    image_size: bool = True        # 是否调整图片尺寸

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
            layout=bool(data.get("layout", default.layout)),
            content=bool(data.get("content", default.content)),
            image_size=bool(data.get("image_size", default.image_size)),
        )

    # ---- 便捷判断 ----
    @property
    def any_style(self) -> bool:
        """版式组里是否至少勾了一项。"""
        return any((self.font_family, self.font_size, self.font_color,
                    self.background, self.page_break))

    @property
    def any_change(self) -> bool:
        """是否存在任何会被执行的改动（版式 / 内容 / 图片）。"""
        return bool(self.layout and self.any_style) or self.content or self.image_size


# 界面与自检共用的字段元信息：key -> (名称, 说明, 所属组, 是否默认勾选)
TOGGLE_META: List[tuple] = [
    ("layout", "自动排版",
     "按模板自动识别层级与版式（总开关）", "style", True),
    ("font_family", "修改字体",
     "中英文字体族统一为模板字体", "style", True),
    ("font_size", "修改字号",
     "标题/正文/表格字号按模板分级", "style", True),
    ("font_color", "修改字体颜色",
     "清除杂色，标题与正文按模板配色", "style", True),
    ("background", "修改背景颜色",
     "表头底纹、斑马纹等填充色", "style", True),
    ("page_break", "调整分页",
     "标题另起页、避免表格跨页断裂", "style", True),
    ("content", "优化文字内容",
     "精简赘语与重复标点，会改动文字", "content", False),
    ("image_size", "调整图片尺寸",
     "过宽/过高的图片等比缩放以适配版面", "content", True),
]


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
    page_size: str = "A4"              # A4 / A3 / Letter / None(保持原样)
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

    # ---- Word：图片（"调整图片尺寸"开关） ----
    # 可用正文宽度 = 页面宽度 - 左右页边距；图片按可用宽度的比例缩放
    image_max_width_ratio: float = 0.92   # 图片最大宽度占可用宽度的比例
    image_max_height_cm: float = 17.5     # 图片最大高度（A4 纵向留出页眉页脚）
    image_min_width_ratio: float = 0.30   # 小于该比例的图片视作图标/小图，不动
    image_allow_upscale: bool = False     # 小图是否放大（默认不放大，避免糊）
    image_center: bool = True             # 独占一段的图片是否居中
    image_caption_gap_pt: float = 4.0     # 图片与紧随其后的题注间距

    # ---- Excel：图片（"调整图片尺寸"开关） ----
    excel_image_max_width_cm: float = 18.0
    excel_image_max_height_cm: float = 12.0

    # ---- 内容优化（"优化文字内容"开关） ----
    content_profile: str = "general"   # general / notice / contract / tender / bid / tech / product
    content_target_saving: float = 0.18  # 期望精简幅度上限，仅用于统计展示
    content_keep_newlines: bool = True
    content_max_repeat_trim: int = 2   # 连续重复标点最多保留个数

    # 元信息
    version: str = "1.1.0"

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
