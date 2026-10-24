"""内置排版模板库（8 套）。

每套模板同时给出 Word 与 Excel 的排版参数，因此用户选一次模板即可同时
适用于两种文件类型。模板之间只在"观感"上不同，不影响任何文字内容。
"""

from __future__ import annotations

from typing import Dict, List

from .schema import FontSpec, LayoutSpec

# 中文字体族
SONG = "宋体"
HEI = "黑体"
FS = "仿宋"
KAI = "楷体"
MSYH = "微软雅黑"
YH = "等线"


def _h(level: int, family: str, size: float, bold: bool = True,
       color: str | None = None, use_hei_for_ascii: bool = False) -> FontSpec:
    return FontSpec(
        east_asia=family,
        ascii_font=("Arial" if use_hei_for_ascii else "Times New Roman"),
        size_pt=size,
        bold=bold,
        color=color,
    )


# ---------------------------------------------------------------------------
# 1. 自动识别（智能排版）
# ---------------------------------------------------------------------------
AUTO = LayoutSpec(
    key="auto",
    name="自动识别内容排版",
    category="both",
    content_profile="general",
    description="智能分析文档结构（标题层级、正文、表格），按通用印刷体规范自动排版，适用于任何文档。",
    tips=["自动识别标题层级并套用黑体标题 + 宋体正文", "正文 1.5 倍行距、首行缩进 2 字符", "表格自动加表头底色并跨页重复表头"],
    page_margin_cm={"top": 2.54, "bottom": 2.54, "left": 3.17, "right": 3.17},
    body=FontSpec(east_asia=SONG, ascii_font="Times New Roman", size_pt=12),
    line_spacing=1.5,
    space_before_pt=0,
    space_after_pt=0,
    first_line_indent_chars=2,
    align="justify",
    headings={1: _h(1, HEI, 18, use_hei_for_ascii=True),
              2: _h(2, HEI, 16, use_hei_for_ascii=True),
              3: _h(3, HEI, 14, use_hei_for_ascii=True),
              4: _h(4, HEI, 12, use_hei_for_ascii=True)},
    heading_space={1: (24, 18), 2: (18, 12), 3: (12, 6), 4: (6, 6)},
    heading_page_break_levels=[1],
    table_font_size=10.5,
    table_header_bg="D9E2F3",
    table_zebra=True,
    excel_font=FontSpec(east_asia=MSYH, ascii_font="Calibri", size_pt=11),
    excel_header_bg="4472C4",
)

# ---------------------------------------------------------------------------
# 2. 发布通知模板（红头文件 / 公文风格）
# ---------------------------------------------------------------------------
NOTICE = LayoutSpec(
    key="notice",
    name="发布通知模板",
    category="both",
    content_profile="notice",
    description="参照党政机关公文（GB/T 9704）与企业内部发文规范：仿宋正文、黑体标题、方正小标宋样式大标题。",
    tips=["页面：上 3.7cm / 下 3.5cm / 左 2.8cm / 右 2.6cm", "正文仿宋_GB2312 三号（16pt），行距约 28 磅（1.9 倍）", "一级标题黑体、二级标题楷体，末页保留署名区"],
    page_margin_cm={"top": 3.7, "bottom": 3.5, "left": 2.8, "right": 2.6},
    header_distance_cm=1.5,
    footer_distance_cm=2.8,
    body=FontSpec(east_asia=FS, ascii_font="Times New Roman", size_pt=16),
    line_spacing=1.9,
    space_before_pt=0,
    space_after_pt=0,
    first_line_indent_chars=2,
    align="justify",
    headings={1: _h(1, SONG, 22, use_hei_for_ascii=True),
              2: _h(2, HEI, 16, use_hei_for_ascii=True),
              3: _h(3, KAI, 16, use_hei_for_ascii=False),
              4: _h(4, FS, 16, bold=True)},
    heading_space={1: (0, 24), 2: (13, 13), 3: (13, 13), 4: (0, 0)},
    heading_page_break_levels=[],
    title_rule=False,
    table_font_size=14,
    table_header_bg="E7E6E6",
    table_zebra=False,
    table_border_color="000000",
    table_border_size_eighths=8,
    footer_page_number=True,
    excel_font=FontSpec(east_asia=FS, ascii_font="Times New Roman", size_pt=12),
    excel_header_font=FontSpec(east_asia=HEI, ascii_font="Arial", size_pt=12, bold=True, color="000000"),
    excel_header_bg="E7E6E6",
    excel_header_fg="000000",
    excel_zebra=False,
    excel_border_color="000000",
    excel_margins_cm={"top": 2.5, "bottom": 2.5, "left": 2.8, "right": 2.6,
                      "header": 1.5, "footer": 1.75},
)

# ---------------------------------------------------------------------------
# 3. 合同模板（法律文书）
# ---------------------------------------------------------------------------
CONTRACT = LayoutSpec(
    key="contract",
    name="合同模板",
    category="both",
    content_profile="contract",
    description="合同/协议类文书：仿宋四号正文、条款标题加粗黑体、宽松行距便于修改批注与签署。",
    tips=["正文仿宋 14pt，行距 1.75 倍，便于手写批注", "条款标题（第一条/1.1）黑体加粗，段前留白", "表格使用细实线、表头浅灰底纹", "页脚自动页码，便于逐页签署"],
    page_margin_cm={"top": 2.54, "bottom": 2.54, "left": 3.17, "right": 3.17},
    body=FontSpec(east_asia=FS, ascii_font="Times New Roman", size_pt=14),
    line_spacing=1.75,
    space_before_pt=0,
    space_after_pt=6,
    first_line_indent_chars=2,
    align="justify",
    headings={1: _h(1, SONG, 22, use_hei_for_ascii=False),
              2: _h(2, HEI, 15, use_hei_for_ascii=True),
              3: _h(3, HEI, 14, use_hei_for_ascii=True),
              4: _h(4, FS, 14, bold=True)},
    heading_space={1: (0, 18), 2: (12, 6), 3: (9, 3), 4: (6, 0)},
    heading_page_break_levels=[],
    table_font_size=12,
    table_header_bg="EDEDED",
    table_zebra=False,
    table_border_color="000000",
    table_border_size_eighths=4,
    footer_page_number=True,
    excel_font=FontSpec(east_asia=FS, ascii_font="Times New Roman", size_pt=11),
    excel_header_font=FontSpec(east_asia=HEI, ascii_font="Arial", size_pt=11, bold=True, color="000000"),
    excel_header_bg="EDEDED",
    excel_header_fg="000000",
    excel_border_color="000000",
)

# ---------------------------------------------------------------------------
# 4. 招标文件排版
# ---------------------------------------------------------------------------
TENDER = LayoutSpec(
    key="tender",
    name="招标文件排版",
    category="both",
    content_profile="tender",
    description="招标/采购文件：结构清晰、编号规范、便于评委快速定位评分要点。",
    tips=["卷一/章标题另起一页，黑体二号", "正文宋体小四，1.5 倍行距", "评分表、清单表加粗表头 + 隔行底纹 + 跨页重复", "大纲级别设置，便于评委在导航窗格跳转"],
    page_margin_cm={"top": 2.54, "bottom": 2.54, "left": 3.0, "right": 3.0},
    body=FontSpec(east_asia=SONG, ascii_font="Times New Roman", size_pt=12, color="000000"),
    line_spacing=1.5,
    space_before_pt=0,
    space_after_pt=0,
    first_line_indent_chars=2,
    align="justify",
    headings={1: _h(1, HEI, 22, use_hei_for_ascii=True),
              2: _h(2, HEI, 16, use_hei_for_ascii=True),
              3: _h(3, HEI, 14, use_hei_for_ascii=True),
              4: _h(4, SONG, 12, bold=True)},
    heading_space={1: (0, 24), 2: (18, 12), 3: (12, 6), 4: (6, 6)},
    heading_page_break_levels=[1],
    table_font_size=10.5,
    table_header_bg="DEEAF6",
    table_zebra=True,
    table_zebra_bg="F2F7FC",
    footer_page_number=True,
    excel_font=FontSpec(east_asia=SONG, ascii_font="Arial", size_pt=11),
    excel_header_bg="4472C4",
    excel_orientation="auto",
)

# ---------------------------------------------------------------------------
# 5. 投标文件排版
# ---------------------------------------------------------------------------
BID = LayoutSpec(
    key="bid",
    name="投标文件排版",
    category="both",
    content_profile="bid",
    description="投标/应答文件：正式、紧凑、页码齐全，评审翻阅效率高；表格是重点。",
    tips=["封面与目录标题居中，章节标题黑体二号并另起一页", "正文宋体小四 12pt、1.5 倍行距（常见评审要求）", "表格：表头深蓝底白字、隔行浅色、禁止跨页断行", "页脚居中页码便于装订核对"],
    page_margin_cm={"top": 2.54, "bottom": 2.54, "left": 3.17, "right": 3.17},
    body=FontSpec(east_asia=SONG, ascii_font="Times New Roman", size_pt=12),
    line_spacing=1.5,
    space_before_pt=0,
    space_after_pt=0,
    first_line_indent_chars=2,
    align="justify",
    headings={1: _h(1, HEI, 22, use_hei_for_ascii=True),
              2: _h(2, HEI, 16, use_hei_for_ascii=True),
              3: _h(3, HEI, 14, use_hei_for_ascii=True),
              4: _h(4, SONG, 12, bold=True)},
    heading_space={1: (0, 24), 2: (18, 10), 3: (12, 6), 4: (6, 6)},
    heading_page_break_levels=[1],
    table_font_size=10.5,
    table_header_bg="1F4E79",
    table_header_color="FFFFFF",
    table_zebra=True,
    table_zebra_bg="EAF1F8",
    table_row_cant_split=True,
    footer_page_number=True,
    excel_font=FontSpec(east_asia=SONG, ascii_font="Arial", size_pt=10.5),
    excel_header_bg="1F4E79",
    excel_header_fg="FFFFFF",
    excel_zebra_bg="EAF1F8",
    excel_header_repeat=True,
)

# ---------------------------------------------------------------------------
# 6. 技术流程模板
# ---------------------------------------------------------------------------
PROCESS = LayoutSpec(
    key="process",
    name="技术流程模板",
    category="both",
    content_profile="tech",
    description="操作手册 / 作业流程 / SOP：步骤感强，编号醒目，适合打印张贴与现场执行。",
    tips=["步骤标题黑体加粗 + 蓝色强调", "正文微软雅黑 11pt、行距 1.4，屏读友好", "步骤间留白，流程图/表格细线框", "层级标题段前分区，便于快速扫读"],
    page_margin_cm={"top": 2.2, "bottom": 2.2, "left": 2.54, "right": 2.54},
    body=FontSpec(east_asia=MSYH, ascii_font="Segoe UI", size_pt=11),
    line_spacing=1.4,
    space_before_pt=0,
    space_after_pt=4,
    first_line_indent_chars=0,
    align="left",
    headings={1: _h(1, MSYH, 20, color="1F4E79", use_hei_for_ascii=True),
              2: _h(2, MSYH, 15, color="2E75B6", use_hei_for_ascii=True),
              3: _h(3, MSYH, 13, color="2E75B6", use_hei_for_ascii=True),
              4: _h(4, MSYH, 11.5, color="000000", use_hei_for_ascii=True)},
    heading_space={1: (18, 12), 2: (14, 6), 3: (10, 4), 4: (6, 2)},
    heading_page_break_levels=[1],
    list_indent_cm=0.6,
    list_hanging_cm=0.6,
    table_font_size=10,
    table_header_bg="2E75B6",
    table_header_color="FFFFFF",
    table_zebra=True,
    table_zebra_bg="EAF2FB",
    table_border_color="9DC3E6",
    footer_page_number=True,
    excel_font=FontSpec(east_asia=MSYH, ascii_font="Segoe UI", size_pt=10.5),
    excel_header_bg="2E75B6",
    excel_header_fg="FFFFFF",
    excel_zebra_bg="EAF2FB",
    excel_border_color="9DC3E6",
)

# ---------------------------------------------------------------------------
# 7. 技术说明模板
# ---------------------------------------------------------------------------
SPEC = LayoutSpec(
    key="spec",
    name="技术说明模板",
    category="both",
    content_profile="tech",
    description="技术方案 / 说明书 / 设计文档：层级分明、编号规范、术语与图表统一，适合评审与归档。",
    tips=["多级标题黑体、按级别递减字号", "正文宋体 12pt、行距 1.5，段前 6 磅便于区分", "图表标题居中加粗、表格细线框隔行底纹", "一级标题另起一页，页脚页码 + 文档名"],
    page_margin_cm={"top": 2.54, "bottom": 2.54, "left": 3.17, "right": 3.17},
    body=FontSpec(east_asia=SONG, ascii_font="Times New Roman", size_pt=12),
    line_spacing=1.5,
    space_before_pt=6,
    space_after_pt=6,
    first_line_indent_chars=0,
    align="justify",
    headings={1: _h(1, HEI, 20, use_hei_for_ascii=True),
              2: _h(2, HEI, 16, use_hei_for_ascii=True),
              3: _h(3, HEI, 14, use_hei_for_ascii=True),
              4: _h(4, HEI, 12, use_hei_for_ascii=True)},
    heading_space={1: (12, 18), 2: (12, 6), 3: (10, 6), 4: (6, 6)},
    heading_page_break_levels=[1],
    table_font_size=10.5,
    table_header_bg="D9E2F3",
    table_zebra=True,
    table_caption_center=True,
    footer_page_number=True,
    excel_font=FontSpec(east_asia=SONG, ascii_font="Calibri", size_pt=11),
    excel_header_bg="8EA9DB",
    excel_header_fg="FFFFFF",
    excel_zebra_bg="EDF2F9",
    excel_border_color="B4C6E7",
)

# ---------------------------------------------------------------------------
# 8. 产品介绍模板（宣传册风格）
# ---------------------------------------------------------------------------
PRODUCT = LayoutSpec(
    key="product",
    name="产品介绍模板",
    category="both",
    content_profile="product",
    description="产品手册 / 宣传资料 / 方案介绍：视觉层次丰富，标题带强调色，适合展示与打印。",
    tips=["主标题大字号 + 品牌蓝，副标题浅灰", "正文微软雅黑 11.5pt、行距 1.5，阅读轻松", "表格表头深色底白字圆角感、隔行浅色", "段落命名色区分区块，打印机下仍保持清晰"],
    page_margin_cm={"top": 2.2, "bottom": 2.2, "left": 2.54, "right": 2.54},
    body=FontSpec(east_asia=MSYH, ascii_font="Segoe UI", size_pt=11.5, color="262626"),
    line_spacing=1.5,
    space_before_pt=0,
    space_after_pt=6,
    first_line_indent_chars=0,
    align="justify",
    headings={1: _h(1, MSYH, 24, color="1F3864", use_hei_for_ascii=True),
              2: _h(2, MSYH, 17, color="2F5597", use_hei_for_ascii=True),
              3: _h(3, MSYH, 14, color="404040", use_hei_for_ascii=True),
              4: _h(4, MSYH, 12, color="404040", use_hei_for_ascii=True)},
    heading_space={1: (24, 12), 2: (18, 8), 3: (12, 6), 4: (8, 4)},
    heading_page_break_levels=[],
    table_font_size=10.5,
    table_header_bg="2F5597",
    table_header_color="FFFFFF",
    table_zebra=True,
    table_zebra_bg="EEF3FB",
    table_border_color="B4C6E7",
    footer_page_number=True,
    excel_font=FontSpec(east_asia=MSYH, ascii_font="Segoe UI", size_pt=10.5, color="262626"),
    excel_header_bg="2F5597",
    excel_header_fg="FFFFFF",
    excel_zebra_bg="EEF3FB",
    excel_border_color="B4C6E7",
)


ALL_TEMPLATES: List[LayoutSpec] = [
    AUTO, NOTICE, CONTRACT, TENDER, BID, PROCESS, SPEC, PRODUCT
]

TEMPLATE_MAP: Dict[str, LayoutSpec] = {t.key: t for t in ALL_TEMPLATES}


def get_template(key: str) -> LayoutSpec:
    """按 key 取模板；未知 key 回落到自动识别。"""
    return TEMPLATE_MAP.get(key, AUTO)


def template_names() -> List[str]:
    return [f"{i + 1}. {t.name}" for i, t in enumerate(ALL_TEMPLATES)]
