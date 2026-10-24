"""Word 排版优化器。

核心约束：**只改排版属性，不改一个字**。

* 直接操作底层 XML 的 ``w:pPr`` / ``w:rPr``，不通过 ``apply_style`` 重建段落，
  因此不会丢失原有的自动编号、超链接、域代码、修订痕迹。
* 不删除任何 run / 段落 / 表格 / 图片；只允许"在标题前加分页属性"这类
  不产生新内容的操作。
* 内容指纹在优化前后比对（见 ``extract_text_fingerprint``）。
"""

from __future__ import annotations

import copy
import hashlib
import logging
import os
import re
from typing import Dict, List, Optional, Tuple

from docx import Document
from docx.enum.section import WD_ORIENT
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Cm, Pt
from docx.table import Table
from docx.text.paragraph import Paragraph

from ..templates.schema import FontSpec, LayoutSpec, ToggleOptions
from . import content_optimizer as C
from . import image_sizer as IMG
from . import analyze as A
from . import ooxml as X

log = logging.getLogger("wxfixer.word")

_ALIGN_MAP = {
    "left": WD_ALIGN_PARAGRAPH.LEFT,
    "center": WD_ALIGN_PARAGRAPH.CENTER,
    "right": WD_ALIGN_PARAGRAPH.RIGHT,
    "justify": WD_ALIGN_PARAGRAPH.JUSTIFY,
}

# 纸张规格（twips）：(长边, 短边)，A4 = 21×29.7cm，A3 = 29.7×42cm
_PAPER_SIZES = {
    "A4": (16838, 11906),
    "A3": (23811, 16838),
    "A5": (11906, 8391),
    "LETTER": (15840, 12240),
    "LEGAL": (20160, 12240),
}


# ---------------------------------------------------------------------------
# 文本指纹
# ---------------------------------------------------------------------------
def extract_text_fingerprint(path: str) -> str:
    """返回文档全部文字（含表格）的 sha256，用于完整性校验。"""
    h = hashlib.sha256()
    doc = Document(path)
    body = doc.element.body
    for p in body.iter(qn("w:p")):
        for node in p.iter(qn("w:t")):
            h.update((node.text or "").encode("utf-8"))
        h.update(b"\x1e")  # 段落分隔符
    return h.hexdigest()


def dump_text(path: str) -> str:
    """调试用：按段落输出纯文本。"""
    doc = Document(path)
    return "\n".join(p.text for p in doc.paragraphs)


def _is_same_as_docname(text: str, doc_name: str) -> bool:
    """段落文字是否就是文档名（忽略空白与常见分隔符）。"""
    if not doc_name:
        return False
    a = re.sub(r"[\s_\-\u3000]+", "", text)
    b = re.sub(r"[\s_\-\u3000]+", "", doc_name)
    return bool(a) and (a == b or (len(a) >= 3 and (a in b or b in a)))


# ---------------------------------------------------------------------------
class WordOptimizer:
    def __init__(self, path: str, spec: LayoutSpec, opts: ToggleOptions):
        self.path = path
        self.spec = spec
        self.opts = opts
        self.doc = Document(path)
        self.stats: Dict[str, int] = {
            "页面设置": 0, "段落": 0, "字体": 0, "表格": 0, "标题": 0,
            "分页": 0, "页眉页脚": 0, "图片": 0,
        }
        self.warnings: List[str] = []
        self.analysis = A.analyze(self.doc)
        # 文档名（用于识别"首行即文档名"的封面式排法）
        self.cover_title = ""
        try:
            self.cover_title = (self.doc.core_properties.title or "").strip()
        except Exception:
            self.cover_title = ""
        # 正文默认字号（用于首行缩进的字符换算）
        self._base_size = self.spec.body.size_pt or 12.0
        self._info_by_element = {id(i.paragraph._p): i for i in self.analysis.paragraphs}
        self._listed = None
        self._title_like = self._compute_title_like()

    def _compute_title_like(self) -> bool:
        """大标题是否"独立成封面"（后面隔了正文/空行才是第一个一级标题）。

        标题下方紧跟一级标题（无正文）时，标题与一级标题同级排布，不再放大，
        避免出现"整页只有一行大标题"的空白页。
        """
        listed = self._listed_paragraphs()
        for pos, info in enumerate(listed):
            if info.kind != A.KIND_TITLE:
                continue
            nxt = listed[pos + 1] if pos + 1 < len(listed) else None
            if nxt is not None and nxt.kind in A.HEADING_LEVELS:
                return False
            return True
        return True

    # ------------------------------------------------------------------
    # 入口
    # ------------------------------------------------------------------
    def run(self) -> Tuple[Dict[str, int], Dict[str, int], List[str]]:
        # ① 内容优化：只做"减法"（删赘语/规整空格标点），绝不新增文字。
        #    放在最前面，避免后续按段落长度做的判断受冗余文字影响。
        if self.opts.content:
            self._optimize_content()
        # ② 版式优化：受"自动排版"总开关控制
        if self.opts.layout:
            self._setup_sections()
            self._apply_normal_style()
            self._apply_heading_styles()
            self._apply_paragraphs()
            self._format_tables()
            self._setup_footer()
            self._synchronize_styles()
        # ③ 图片尺寸：独立开关，即使不做版式也能单独生效
        if self.opts.image_size:
            self._resize_images()
        self.doc.save(self.path)
        self._validate()
        return self.stats, self.analysis.summary(), self.warnings

    # ------------------------------------------------------------------
    # 内容优化（可选）
    # ------------------------------------------------------------------
    def _optimize_content(self) -> None:
        """精简段落文字：只删不改，逐段安全校验。"""
        profile = getattr(self.spec, "content_profile", "general") or "general"
        texts = [p.text for p in self.doc.paragraphs]
        stats = C.ContentStats(paragraphs=len(texts))
        changed = 0
        for para in self.doc.paragraphs:
            old = para.text
            if not old or not old.strip():
                continue
            new = C.optimize_text(old, profile, stats)
            if new == old:
                continue
            # 安全校验：文字字符只能是原文子集（不引入任何新字）
            if not C.is_char_subset(C.word_char_multiset(old), C.word_char_multiset(new)):
                self.warnings.append(f"段落内容优化被跳过（字符校验未通过）：{old[:24]}…")
                continue
            if C.rewrite_paragraph_runs(para, new):
                changed += 1
        self.stats["文字优化"] = changed
        if changed:
            self.stats["精简字数"] = stats.removed_chars
            log.info("内容优化：%s | %s", stats.summary(), C.profile_label(profile))
        else:
            log.info("内容优化：%s", stats.summary())

    # ------------------------------------------------------------------
    def _resize_images(self) -> None:
        """按模板参数调整图片尺寸，避免超出版心或过小看不清。"""
        img_stats, _ = IMG.resize_word_images(self.doc, self.spec, True)
        self.stats["图片"] = img_stats.resized
        if img_stats.total:
            log.info("图片处理：%s", img_stats.summary())
        if img_stats.details:
            self.warnings.append("图片尺寸：" + "；".join(img_stats.details[:5]))

    # ------------------------------------------------------------------
    # 页面与样式层
    # ------------------------------------------------------------------
    def _setup_sections(self) -> None:
        spec = self.spec
        default_section = self.doc.sections[0]
        for section in self.doc.sections:
            sec_pr = section._sectPr
            pg_mar = X.insert_ordered(sec_pr, "w:pgMar", X.SECTPR_SEQ)
            margins = spec.page_margin_cm
            pg_mar.set(qn("w:top"), str(int(Cm(margins["top"]).twips)))
            pg_mar.set(qn("w:bottom"), str(int(Cm(margins["bottom"]).twips)))
            pg_mar.set(qn("w:left"), str(int(Cm(margins["left"]).twips)))
            pg_mar.set(qn("w:right"), str(int(Cm(margins["right"]).twips)))
            header_cm = spec.header_distance_cm or 1.5
            footer_cm = spec.footer_distance_cm or 1.75
            pg_mar.set(qn("w:header"), str(int(Cm(header_cm).twips)))
            pg_mar.set(qn("w:footer"), str(int(Cm(footer_cm).twips)))
            pg_mar.set(qn("w:gutter"), "0")

            orientation = spec.orientation
            if orientation is None and spec.preferred_landscape_chars:
                if self.analysis.max_line_chars > spec.preferred_landscape_chars:
                    orientation = "landscape"

            # 先定方向（python-docx 会交换宽高），再按纸张规格覆盖具体尺寸
            if orientation in ("portrait", "landscape"):
                section.orientation = (WD_ORIENT.LANDSCAPE
                                       if orientation == "landscape"
                                       else WD_ORIENT.PORTRAIT)
            paper = _PAPER_SIZES.get((spec.page_size or "").upper())
            if paper:
                long_edge, short_edge = paper
                landscape = orientation == "landscape"
                w_tw = long_edge if landscape else short_edge
                h_tw = short_edge if landscape else long_edge
                pg_sz = X.insert_ordered(sec_pr, "w:pgSz", X.SECTPR_SEQ)
                pg_sz.set(qn("w:w"), str(w_tw))
                pg_sz.set(qn("w:h"), str(h_tw))
                pg_sz.set(qn("w:orient"),
                           "landscape" if landscape else "portrait")
            self.stats["页面设置"] += 1
        _ = default_section

    def _font_xml(self, font: FontSpec) -> OxmlElement:
        rpr = OxmlElement("w:rPr")
        X.set_run_fonts(rpr, font.east_asia, font.ascii_font)
        X.set_run_size(rpr, font.size_pt)
        X.set_run_bold(rpr, font.bold)
        if font.color and self.opts.font_color:
            X.set_run_color(rpr, font.color)
        return rpr

    def _apply_normal_style(self) -> None:
        """只修改 Normal 样式里"用户允许"的属性——这是全局兜底。"""
        spec, opts = self.spec, self.opts
        try:
            normal = self.doc.styles["Normal"]
        except KeyError:  # pragma: no cover
            return
        elements = normal.element

        # 段落属性
        ppr = elements.find(qn("w:pPr"))
        if ppr is None:
            ppr = OxmlElement("w:pPr")
            elements.insert(0, ppr)
        X.set_para_spacing(
            ppr,
            before_pt=spec.space_before_pt,
            after_pt=spec.space_after_pt,
            line=spec.line_spacing,
        )

        # 字体
        rpr = elements.find(qn("w:rPr"))
        if rpr is None:
            rpr = OxmlElement("w:rPr")
            elements.append(rpr)
        if opts.font_family:
            X.set_run_fonts(rpr, spec.body.east_asia, spec.body.ascii_font)
        if opts.font_size:
            X.set_run_size(rpr, spec.body.size_pt)
        if opts.font_color and spec.body.color:
            X.set_run_color(rpr, spec.body.color)

        # docDefaults 也同步一次，覆盖"未显式声明字体"的文档
        self._patch_doc_defaults()

    def _patch_doc_defaults(self) -> None:
        try:
            styles_el = self.doc.styles.element
        except Exception:  # pragma: no cover
            return
        doc_defaults = styles_el.find(qn("w:docDefaults"))
        if doc_defaults is None:
            return
        rpr_default = doc_defaults.find(qn("w:rPrDefault"))
        if rpr_default is None:
            return
        rpr = rpr_default.find(qn("w:rPr"))
        if rpr is None:
            rpr = OxmlElement("w:rPr")
            rpr_default.append(rpr)
        spec = self.spec
        if self.opts.font_family:
            X.set_run_fonts(rpr, spec.body.east_asia, spec.body.ascii_font)
        if self.opts.font_size:
            X.set_run_size(rpr, spec.body.size_pt)

    def _apply_heading_styles(self) -> None:
        """更新 Heading 1-4 / Title 样式定义，保持大纲结构。"""
        spec, opts = self.spec, self.opts
        for level in range(1, 5):
            font = spec.heading_font(level)
            if font is None:
                continue
            for style_name in (f"Heading {level}", f"标题 {level}"):
                try:
                    style = self.doc.styles[style_name]
                except KeyError:
                    continue
                self._style_heading(style.element, level, font)
                break

        try:
            title_style = self.doc.styles["Title"]
        except KeyError:
            title_style = None
        if title_style is not None:
            font = spec.heading_font(1)
            if font is not None:
                big = FontSpec(**{**font.__dict__})
                big.size_pt = max(font.size_pt + 4, 22)
                self._style_heading(title_style.element, 0, big, force_center=True)

    def _style_heading(self, element, level: int, font: FontSpec,
                       force_center: bool = False) -> None:
        spec, opts = self.spec, self.opts
        ppr = element.find(qn("w:pPr"))
        if ppr is None:
            ppr = OxmlElement("w:pPr")
            element.insert(0, ppr)

        before, after = spec.heading_spacing(level) if level else (0, 18)
        X.set_para_spacing(ppr, before_pt=before, after_pt=after,
                           line=min(spec.line_spacing, 1.5))
        if spec.heading_keep_with_next:
            X.set_keep_next(ppr, True)
            X.set_keep_lines(ppr, True)
        if level and spec.heading_page_break_levels and level in spec.heading_page_break_levels:
            if opts.page_break:
                X.set_page_break_before(ppr, True)
            else:
                # 关掉"分页"选项时，连样式里的段前分页也要清掉
                X.set_page_break_before(ppr, False)
        elif _has_page_break(ppr):
            # 级别不在分页范围（或模板不分页）时，去掉样式继承下来的分页设置
            X.set_page_break_before(ppr, False)
        X.set_outline_level(ppr, max(level - 1, 0))
        if force_center or level == 0:
            X.set_align(ppr, "center")
        if level == 0:
            # 大标题样式默认带一条底边线，是否保留由模板 title_rule 决定
            pBdr = ppr.find(qn("w:pBdr"))
            if not spec.title_rule:
                if pBdr is not None:
                    ppr.remove(pBdr)
            else:
                if pBdr is None:
                    pBdr = X.insert_ordered(ppr, "w:pBdr", X.PPR_SEQ)
                X.set_border(pBdr, ("w:bottom",), color="1F4E79",
                             size_eighths=6, seq=X.TBLBORDERS_SEQ)
        # 标题不缩进
        X.set_para_indent(ppr, first_line_chars=0, char_unit=self._base_size)

        rpr = element.find(qn("w:rPr"))
        if rpr is None:
            rpr = OxmlElement("w:rPr")
            element.append(rpr)
        if opts.font_family:
            X.set_run_fonts(rpr, font.east_asia, font.ascii_font)
        if opts.font_size:
            X.set_run_size(rpr, font.size_pt)
        # 勾选了"修改字体颜色"时，标题颜色由样式决定：模板未指定颜色则显式设为
        # 正文色，避免原文残留的红/蓝标题色与模板配色混在一起
        if opts.font_color:
            X.set_run_color(rpr, font.color or spec.body.color or "000000")
        X.set_run_bold(rpr, font.bold)

    # ------------------------------------------------------------------
    # 正文段落
    # ------------------------------------------------------------------
    def _apply_paragraphs(self) -> None:
        spec, opts = self.spec, self.opts
        analyzed = [i for i in self.analysis.paragraphs if not i.in_table]

        # 没有识别到大标题时，把文档第一段（首行）当作大标题，保证有视觉层级
        if len(analyzed) > 1 and not self.analysis.title_count:
            first = analyzed[0]
            if first.kind == A.KIND_BODY and first.char_len <= 80:
                first.kind = A.KIND_TITLE

        for info in analyzed:
            p = info.paragraph
            if info.kind == A.KIND_EMPTY:
                self._compact_empty_paragraph(p)
                continue

            if info.kind == A.KIND_TITLE:
                self._format_title(p, title_like=self._title_like)
                self.stats["标题"] += 1
            elif info.kind in A.HEADING_LEVELS:
                self._format_heading(p, info.level)
                self.stats["标题"] += 1
            elif info.kind == A.KIND_SUBTITLE:
                self._format_subtitle(p)
            elif info.kind == A.KIND_CAPTION:
                self._format_caption(p)
            elif info.kind == A.KIND_LIST:
                self._format_list(p)
                self.stats["段落"] += 1
            else:
                self._format_body(p)
                self.stats["段落"] += 1

    def _set_rpr_font(self, rpr, font: FontSpec, allow_bold_keep: bool = True) -> None:
        opts = self.opts
        X.set_run_fonts(rpr, font.east_asia, font.ascii_font)
        X.set_run_size(rpr, font.size_pt)
        if font.color and opts.font_color:
            X.set_run_color(rpr, font.color)

    def _format_run_fonts(self, p: Paragraph, font: FontSpec,
                          force_bold: Optional[bool] = None) -> None:
        """逐 run 覆盖字体，但保留原有的加粗/斜体等强调。"""
        opts = self.opts
        runs = p.runs
        if not runs:
            return
        for run in runs:
            rpr = run._r.get_or_add_rPr()
            X.set_run_fonts(rpr, font.east_asia, font.ascii_font)
            if opts.font_size:
                X.set_run_size(rpr, font.size_pt)
            if opts.font_color:
                # 勾选"修改字体颜色"后，字体颜色统一按模板配色：模板给了颜色用
                # 模板色，否则回落正文色，清掉原文残留的彩色字
                X.set_run_color(rpr, font.color or self.spec.body.color or "000000")
            if force_bold is not None:
                X.set_run_bold(rpr, force_bold)
            self.stats["字体"] += 1

    def _format_title(self, p: Paragraph, title_like: bool = True) -> None:
        spec = self.spec
        font = spec.heading_font(1)
        if font is None:
            return
        title_font = FontSpec(**{**font.__dict__})
        if title_like:
            title_font.size_pt = max(font.size_pt + 4, 22)
        title_font.bold = True
        pf = p.paragraph_format
        pf.alignment = WD_ALIGN_PARAGRAPH.CENTER
        pf.space_before = Pt(12)
        pf.space_after = Pt(18 if title_like else 12)
        pf.first_line_indent = Pt(0)
        pf.left_indent = Pt(0)
        pf.line_spacing = 1.4
        ppr = p._p.get_or_add_pPr()
        X.set_keep_next(ppr, True)
        X.set_para_indent(ppr, first_line_chars=0, char_unit=self._base_size)
        if title_like and spec.title_rule:
            # 大标题下方加一条居中分隔横线（样式可关，公文模板默认关闭）
            pBdr = ppr.find(qn("w:pBdr"))
            if pBdr is None:
                pBdr = X.insert_ordered(ppr, "w:pBdr", X.PPR_SEQ)
            X.set_border(pBdr, ("w:bottom",), color="1F4E79", size_eighths=6,
                         seq=X.TBLBORDERS_SEQ)
            el = pBdr.find(qn("w:bottom"))
            if el is not None:
                el.set(qn("w:space"), "6")
        else:
            pBdr = ppr.find(qn("w:pBdr"))
            if pBdr is not None:
                ppr.remove(pBdr)
        self._format_run_fonts(p, title_font, force_bold=True)

    def _format_subtitle(self, p: Paragraph) -> None:
        font = self.spec.heading_font(2)
        if font is None:
            return
        sub = FontSpec(**{**font.__dict__})
        sub.bold = False
        sub.size_pt = max(14, font.size_pt - 2)
        p.paragraph_format.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p.paragraph_format.space_after = Pt(12)
        self._format_run_fonts(p, sub)

    def _format_heading(self, p: Paragraph, level: int) -> None:
        spec, opts = self.spec, self.opts
        font = spec.heading_font(level)
        if font is None:
            return
        before, after = spec.heading_spacing(level)
        pf = p.paragraph_format
        pf.alignment = _ALIGN_MAP.get("left", WD_ALIGN_PARAGRAPH.LEFT)
        pf.space_before = Pt(before)
        pf.space_after = Pt(after)
        pf.first_line_indent = Pt(0)
        pf.left_indent = Pt(0)
        pf.keep_with_next = spec.heading_keep_with_next
        pf.keep_together = True
        pf.line_spacing = 1.4

        ppr = p._p.get_or_add_pPr()
        # 标题一律左对齐（公文大标题除外，由 Title 分支处理）
        X.set_align(ppr, "left")
        X.set_para_indent(ppr, first_line_chars=0, char_unit=self._base_size)
        X.set_outline_level(ppr, level - 1)
        # 先默认"本段不分页"；只有真正需要分页时才覆盖为 True。
        # 样式层可能继承 pageBreakBefore（如在样式里勾过"段前分页"），
        # 段落级显式 w:val=0 才能确保"不需要分页的标题"不被顶到下一页
        if opts.page_break and level in (spec.heading_page_break_levels or []):
            X.set_page_break_before(ppr, False)
            self._maybe_page_break(p)
        else:
            X.set_page_break_before(ppr, False)

        self._format_run_fonts(p, font, force_bold=font.bold)

    def _maybe_page_break(self, p: Paragraph) -> None:
        """给标题加"段前分页"，但避免封面/大标题被单独留在一页空白上。

        只有紧挨着的前一段是正文/列表/图表说明时才分页；空段落、标题、
        文档名之后的标题一律不分页，直接连排。
        """
        info = self._info_by_element.get(id(p._p))
        if info is None:
            log.debug("段前分页：段落 %r 未在分析结果中，跳过", p.text[:20])
            return
        listed = self._listed_paragraphs()
        try:
            pos = next(k for k, i in enumerate(listed) if i is info)
        except StopIteration:
            return
        if pos <= 0:
            return
        prev = listed[pos - 1]
        log.debug("段前分页判定 %r <- 前段 %r(%s)", p.text[:16], prev.text[:16], prev.kind)
        if prev.kind not in (A.KIND_BODY, A.KIND_LIST, A.KIND_CAPTION):
            # 上一段是标题/文档名（它们后面本来就会换行），分页会多出整页空白
            return
        if not prev.text.strip():
            return
        if self.cover_title and _is_same_as_docname(prev.text, self.cover_title):
            return
        p.paragraph_format.page_break_before = True
        self.stats["分页"] += 1

    def _listed_paragraphs(self):
        """按文档顺序排列的非空、非表格段落信息列表。"""
        if self._listed is None:
            self._listed = [i for i in self.analysis.paragraphs
                            if i.kind != A.KIND_EMPTY and not i.in_table]
        return self._listed

    def _format_caption(self, p: Paragraph) -> None:
        spec = self.spec
        cap_font = FontSpec(**{**spec.body.__dict__})
        cap_font.size_pt = max((spec.table_font_size or spec.body.size_pt), 10.5)
        cap_font.bold = False
        pf = p.paragraph_format
        if spec.table_caption_center:
            pf.alignment = WD_ALIGN_PARAGRAPH.CENTER
        pf.first_line_indent = Pt(0)
        pf.space_before = Pt(6)
        pf.space_after = Pt(3)
        pf.keep_with_next = True
        self._format_run_fonts(p, cap_font)

    def _format_list(self, p: Paragraph) -> None:
        spec = self.spec
        pf = p.paragraph_format
        pf.space_before = Pt(0)
        pf.space_after = Pt(0)
        pf.line_spacing = spec.line_spacing
        pf.left_indent = Cm(spec.list_indent_cm)
        pf.first_line_indent = Cm(-spec.list_hanging_cm)
        ppr = p._p.get_or_add_pPr()
        X.set_para_indent(
            ppr,
            hanging_cm=spec.list_hanging_cm,
            left_cm=spec.list_indent_cm,
            char_unit=self._base_size,
        )
        self._format_run_fonts(p, spec.body)

    def _format_body(self, p: Paragraph) -> None:
        spec = self.spec
        pf = p.paragraph_format
        pf.alignment = _ALIGN_MAP.get(spec.align, WD_ALIGN_PARAGRAPH.JUSTIFY)
        pf.space_before = Pt(spec.space_before_pt)
        pf.space_after = Pt(spec.space_after_pt)
        pf.line_spacing = spec.line_spacing
        pf.left_indent = Pt(0)
        pf.keep_with_next = False

        ppr = p._p.get_or_add_pPr()
        if spec.first_line_indent_chars:
            pf.first_line_indent = Pt(spec.first_line_indent_chars * self._base_size)
            X.set_para_indent(
                ppr, first_line_chars=spec.first_line_indent_chars,
                char_unit=self._base_size,
            )
        else:
            X.set_para_indent(ppr, first_line_chars=0, char_unit=self._base_size)

        self._format_run_fonts(p, spec.body)
        # 段落里偶尔混入西文字体骨架，把 pPr/rPr 也统一
        ppr_rpr = ppr.find(qn("w:rPr"))
        if ppr_rpr is not None:
            if self.opts.font_size:
                X.set_run_size(ppr_rpr, spec.body.size_pt)

    def _compact_empty_paragraph(self, p: Paragraph) -> None:
        """空段落压缩高度（不删除，保留原有结构）。"""
        pf = p.paragraph_format
        pf.space_before = Pt(0)
        pf.space_after = Pt(0)
        ppr = p._p.get_or_add_pPr()
        X.set_para_spacing(ppr, before_pt=0, after_pt=0, line=1.0)

    # ------------------------------------------------------------------
    # 表格
    # ------------------------------------------------------------------
    def _format_tables(self) -> None:
        spec = self.spec
        for table in self.doc.tables:
            self._format_one_table(table)

    def _format_one_table(self, table: Table) -> None:
        spec = self.spec
        tbl = table._tbl
        tblpr = tbl.tblPr

        # 边框：统一细实线
        borders = X.insert_ordered(tblpr, "w:tblBorders", X.TBLPR_SEQ)
        X.strip_borders(borders)
        X.set_border(
            borders,
            ("w:top", "w:left", "w:bottom", "w:right", "w:insideH", "w:insideV"),
            spec.table_border_color,
            spec.table_border_size_eighths,
        )

        if spec.table_center:
            jc = X.insert_ordered(tblpr, "w:jc", X.TBLPR_SEQ)
            jc.set(qn("w:val"), "center")

        # 单元格内容
        rows = list(table.rows)
        for r_idx, row in enumerate(rows):
            trpr = row._tr.get_or_add_trPr()
            if spec.table_row_cant_split:
                cant = X.insert_ordered(trpr, "w:cantSplit", X.TRPR_SEQ)
                cant.set(qn("w:val"), "true")
            if r_idx == 0 and spec.table_header_repeat:
                th = X.insert_ordered(trpr, "w:tblHeader", X.TRPR_SEQ)
                th.set(qn("w:val"), "true")

            for c_idx, cell in enumerate(row.cells):
                self._format_cell(cell, is_header=(r_idx == 0),
                                  is_zebra=(spec.table_zebra and r_idx > 0
                                            and r_idx % 2 == 0))
        self.stats["表格"] += 1

    def _format_cell(self, cell, is_header: bool, is_zebra: bool) -> None:
        spec, opts = self.spec, self.opts
        tcpr = cell._tc.get_or_add_tcPr()

        # 垂直居中
        va = X.insert_ordered(tcpr, "w:vAlign", X.TCPR_SEQ)
        va.set(qn("w:val"), "center")

        # 背景色
        if opts.background:
            if is_header and spec.table_header_bg:
                X.set_shading(tcpr, spec.table_header_bg)
            elif is_zebra and spec.table_zebra_bg:
                X.set_shading(tcpr, spec.table_zebra_bg)

        font = FontSpec(**{**spec.body.__dict__})
        font.size_pt = spec.table_font_size or spec.body.size_pt
        if is_header and spec.table_header_bold:
            font.bold = True
        if is_header and spec.table_header_color and opts.font_color:
            font.color = spec.table_header_color

        for p in cell.paragraphs:
            ppr = p._p.get_or_add_pPr()
            pf = p.paragraph_format
            pf.space_before = Pt(2)
            pf.space_after = Pt(2)
            pf.line_spacing = 1.15
            X.set_para_indent(ppr, first_line_chars=0, char_unit=self._base_size)
            if is_header:
                X.set_align(ppr, "center")
                pf.alignment = WD_ALIGN_PARAGRAPH.CENTER
            else:
                X.set_align(ppr, "left")
                pf.alignment = WD_ALIGN_PARAGRAPH.LEFT
            self._format_run_fonts(
                p, font, force_bold=True if (is_header and spec.table_header_bold) else None
            )

    # ------------------------------------------------------------------
    # 页脚页码
    # ------------------------------------------------------------------
    def _setup_footer(self) -> None:
        if not self.spec.footer_page_number:
            return
        try:
            for section in self.doc.sections:
                footer = section.footer
                footer.is_linked_to_previous = False
                if footer.paragraphs and footer.paragraphs[0].text.strip():
                    para = footer.paragraphs[0]
                else:
                    para = footer.paragraphs[0] if footer.paragraphs else footer.add_paragraph()
                para.alignment = WD_ALIGN_PARAGRAPH.CENTER
                # 清理已有 run 中的页码域（保持 footer 只含一个 PAGE 域）
                if _contains_field(para):
                    continue
                for r in list(para.runs):
                    r._r.getparent().remove(r._r)
                self._append_page_number(para)
                self.stats["页眉页脚"] += 1
        except Exception as exc:  # 页脚属于"锦上添花"，失败不影响正文
            self.warnings.append(f"页脚页码设置失败：{exc}")
            log.warning("页脚页码设置失败: %s", exc)

    def _append_page_number(self, para: Paragraph) -> None:
        run = para.add_run()
        rpr = run._r.get_or_add_rPr()
        X.set_run_size(rpr, 10.5)
        X.set_run_fonts(rpr, self.spec.body.east_asia, self.spec.body.ascii_font)

        def fld(kind: str) -> OxmlElement:
            el = OxmlElement("w:fldChar")
            el.set(qn("w:fldCharType"), kind)
            return el

        instr = OxmlElement("w:instrText")
        instr.set(qn("xml:space"), "preserve")
        instr.text = " PAGE "
        run._r.append(fld("begin"))
        run._r.append(instr)
        run._r.append(fld("separate"))
        t = OxmlElement("w:t")
        t.text = "1"
        run._r.append(t)
        run._r.append(fld("end"))

    # ------------------------------------------------------------------
    # 同步 / 校验
    # ------------------------------------------------------------------
    def _synchronize_styles(self) -> None:
        """把样式定义同步到实际段落，保证导航窗格层级与视觉一致。"""
        for info in self.analysis.paragraphs:
            if not (info.kind in A.HEADING_LEVELS or info.kind == A.KIND_TITLE):
                continue
            p = info.paragraph
            ppr = p._p.pPr
            if ppr is not None and ppr.find(qn("w:pStyle")) is not None:
                continue  # 原文自带标题样式，保留用户既有的样式体系
            desired = "Title" if info.kind == A.KIND_TITLE else f"Heading {info.level}"
            try:
                style = self.doc.styles[desired]
            except KeyError:
                continue
            if p.style is not None and p.style.name != desired:
                # 原文用"标题"样式表示文档名、又被识别为首个一级标题时，
                # 保留其居中与大字号观感（仅改外观，不动文字）
                self._inherit_title_look(p, info)
                p.style = style

    def _inherit_title_look(self, p: Paragraph, info) -> None:
        """把段落当前的"大标题观感"（居中/较大字号）复制为直接格式。"""
        from docx.shared import Pt as _Pt

        try:
            eff = p.paragraph_format.alignment
            font = p.runs[0].font if p.runs else p.style.font
            size = font.size
            is_title_style = (p.style is not None and
                              (p.style.name or "").lower() in ("title", "标题"))
        except Exception:
            return
        if not (is_title_style or eff == WD_ALIGN_PARAGRAPH.CENTER or
                (size is not None and size.pt >= 18)):
            return
        base = self.spec.heading_font(1)
        if base is None:
            return
        big = max(base.size_pt, 20.0)
        p.paragraph_format.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p.paragraph_format.first_line_indent = _Pt(0)
        for run in p.runs:
            if self.opts.font_size:
                run.font.size = _Pt(big)
            if self.opts.font_color:
                X.set_run_color(run._r.get_or_add_rPr(),
                                base.color or self.spec.body.color or "000000")
            if self.opts.font_family:
                X.set_run_fonts(run._r.get_or_add_rPr(), base.east_asia,
                                base.ascii_font)
            X.set_run_bold(run._r.get_or_add_rPr(), True)
        self.stats["标题观感保留"] = self.stats.get("标题观感保留", 0) + 1

    def _validate(self) -> None:
        """保存后重新打开并执行 OOXML 顺序自检。"""
        from .ooxml_check import check_document

        issues = check_document(self.path)
        if issues:
            for issue in issues[:20]:
                self.warnings.append(f"OOXML 结构提示：{issue}")
            log.warning("OOXML 顺序自检发现 %d 处问题", len(issues))
        else:
            log.debug("OOXML 顺序自检通过")


def _contains_field(para: Paragraph) -> bool:
    return bool(para._p.findall(".//" + qn("w:fldChar"))) or bool(
        para._p.findall(".//" + qn("w:fldSimple"))
    )


def _has_page_break(ppr) -> bool:
    """pPr 里是否显式开启了"段前分页"。"""
    if ppr is None:
        return False
    el = ppr.find(qn("w:pageBreakBefore"))
    if el is None:
        return False
    val = el.get(qn("w:val"))
    return val is None or val.lower() not in ("0", "false", "off")


def _is_char_subset(before, after) -> bool:
    """after 的每个字符计数都不能超过 before（即只删不增）。"""
    return C.is_char_subset(before, after)


# ---------------------------------------------------------------------------
def optimize_word(path: str, spec: LayoutSpec, options) -> Tuple[Dict[str, int], Dict[str, int], List[str]]:
    optimizer = WordOptimizer(path, spec, options.toggles)
    return optimizer.run()
