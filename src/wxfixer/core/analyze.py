"""文档内容分析：识别标题层级、正文、表格、列表。

只读分析，不修改文档。识别结果用于"自动识别内容排版"以及所有模板的
样式套用决策。
"""

from __future__ import annotations

import re
from dataclasses import dataclass, field
from typing import Dict, List, Optional, Tuple

from docx.document import Document as DocumentObject
from docx.oxml.ns import qn
from docx.table import Table, _Cell
from docx.text.paragraph import Paragraph
from docx.enum.text import WD_ALIGN_PARAGRAPH

# ---------------------------------------------------------------------------
# 段落分类
# ---------------------------------------------------------------------------
KIND_TITLE = "title"       # 文档大标题 / 封面标题
KIND_H1 = "h1"
KIND_H2 = "h2"
KIND_H3 = "h3"
KIND_H4 = "h4"
KIND_SUBTITLE = "subtitle"
KIND_BODY = "body"
KIND_LIST = "list"
KIND_CAPTION = "caption"   # 图 / 表题
KIND_EMPTY = "empty"

HEADING_LEVELS = {KIND_H1: 1, KIND_H2: 2, KIND_H3: 3, KIND_H4: 4}

# 中文标题编号：一、  （一）  1.  1.1  （1）  第X章  第一章
_RE_CN_NUM = re.compile(r"^\s*([一二三四五六七八九十百]+)\s*[、．.]")
_RE_CN_PAREN = re.compile(r"^\s*[（(]\s*([一二三四五六七八九十]+)\s*[）)]")
_RE_ARABIC_DOT = re.compile(r"^\s*(\d{1,2})\s*[、．]")
_RE_ARABIC_PAREN = re.compile(r"^\s*[（(]\s*(\d{1,3})\s*[）)]")
_RE_MULTI = re.compile(r"^\s*(\d{1,2}(?:\.\d{1,2}){1,4})\s*[.、]?\s*\S")
# 单级阿拉伯数字编号："1 概述" / "3、技术指标"
_RE_ARABIC_HEADING = re.compile(r"^\s*(\d{1,2})(?:\s+|\s*[、．.]\s*)(\S.*)$")
_RE_CHAPTER = re.compile(r"^\s*第\s*[一二三四五六七八九十百零\d]+\s*[章节篇部分条]")
_RE_APPENDIX = re.compile(r"^\s*(附\s*[件录表]|附件|附录|附表)")
_RE_CAPTION = re.compile(r"^\s*(图|表|Figure|Table)\s*\d+([-－.．]\d+)*")
_RE_LIST_MARK = re.compile(r"^\s*([•·▪◦●○■□◆◇►▶*\-—–]|\(\d{1,2}\)|（\d{1,2}）)\s*")

_HEADING_MAX_CHARS = 60
_TITLE_MAX_CHARS = 80


@dataclass
class ParaInfo:
    index: int                      # 在正文块序列中的序号
    paragraph: Paragraph
    text: str
    kind: str = KIND_BODY
    level: int = 0                  # 标题级别，0 = 非标题
    numbering: Optional[str] = None  # 识别到的编号前缀 "一、" / "1.1"
    in_table: bool = False
    char_len: int = 0
    has_drawing: bool = False
    has_formula: bool = False


@dataclass
class DocAnalysis:
    paragraphs: List[ParaInfo] = field(default_factory=list)
    tables: List[Table] = field(default_factory=list)
    heading_counts: Dict[int, int] = field(default_factory=dict)
    title_count: int = 0
    body_char_count: int = 0
    max_line_chars: int = 0
    has_conclusion_label: bool = False
    keywords: List[str] = field(default_factory=list)
    section_count: int = 1
    used_numbering_styles: bool = False

    @property
    def heading_count(self) -> int:
        return sum(self.heading_counts.values())

    def summary(self) -> dict:
        return {
            "段落数": len(self.paragraphs),
            "表格数": len(self.tables),
            "一级标题": self.heading_counts.get(1, 0),
            "二级标题": self.heading_counts.get(2, 0),
            "三级标题": self.heading_counts.get(3, 0),
            "大标题": self.title_count,
            "正文字数": self.body_char_count,
            "最长行字符数": self.max_line_chars,
            "节数": self.section_count,
        }


# ---------------------------------------------------------------------------
def iter_block_paragraphs(container) -> List[Paragraph]:
    """按文档顺序返回所有段落（含表格单元格内的段落）。"""
    out: List[Paragraph] = []
    _walk_paragraphs(container, out)
    return out


def _walk_paragraphs(parent, out: List[Paragraph]) -> None:
    if hasattr(parent, "paragraphs"):
        for p in parent.paragraphs:
            out.append(p)
    if hasattr(parent, "tables"):
        for tbl in parent.tables:
            for row in tbl.rows:
                for cell in row.cells:
                    _walk_paragraphs(cell, out)


def _heading_level_from_style(p: Paragraph) -> Optional[int]:
    name = (p.style.name if p.style is not None else "") or ""
    low = name.lower()
    if low in ("title", "标题"):
        return 0  # 文档标题
    if low in ("subtitle", "副标题"):
        return -1
    m = re.match(r"(heading|标题)\s*(\d)", low)
    if m:
        return int(m.group(2))
    return None


def _outline_level(p: Paragraph) -> Optional[int]:
    ppr = p._p.pPr
    if ppr is None:
        return None
    el = ppr.find(qn("w:outlineLvl"))
    if el is None:
        return None
    try:
        val = int(el.get(qn("w:val")))
    except (TypeError, ValueError):
        return None
    return val + 1 if 0 <= val <= 8 else None


def _detect_numbering(text: str) -> Tuple[Optional[str], Optional[int]]:
    """返回 (编号前缀, 推断层级)。

    编号体系做兼容：
    - "一、" 一级；"（一）" 二级；"1." / "1、" 视作二级
    - "1.1" 二级；"1.1.1" 三级
    - 独立一行的 "1 概述" 视作一级
    """
    m = _RE_CHAPTER.match(text)
    if m:
        return m.group(0).strip(), 1
    m = _RE_CN_NUM.match(text)
    if m:
        return m.group(0).strip(), 1
    m = _RE_MULTI.match(text)
    if m:
        depth = m.group(1).count(".") + 1
        return m.group(1), min(depth, 4)
    m = _RE_ARABIC_HEADING.match(text)
    if m and len(text.strip()) <= _HEADING_MAX_CHARS:
        return m.group(1), 1
    m = _RE_CN_PAREN.match(text)
    if m:
        return m.group(0).strip(), 2
    m = _RE_ARABIC_DOT.match(text)
    if m:
        return m.group(0).strip(), 2
    m = _RE_ARABIC_PAREN.match(text)
    if m:
        return m.group(0).strip(), 3
    return None, None


def _is_cn_number(num: Optional[str]) -> bool:
    """"一、" / "十、" 这类中文序号编号。"""
    if not num:
        return False
    return bool(re.match(r"^[一二三四五六七八九十百]+[、.．]", num.strip()))


def _is_arabic_dotted(num: Optional[str]) -> bool:
    """是否为 "1." / "1、" 这类单级阿拉伯编号（层级被推断为 2）。"""
    if not num:
        return False
    return bool(re.match(r"^\d{1,2}[.．、]$", num.strip()) or re.match(r"^\d{1,2}$", num.strip()))


def _is_multi_level(num: Optional[str]) -> bool:
    """是否为 "1.1" / "2.3.1" 这类多级编号。"""
    return bool(num) and bool(re.match(r"^\d{1,2}(?:\.\d{1,2})+$", num.strip()))


def _max_heading_level(result: DocAnalysis) -> int:
    heads = [i for i in result.paragraphs
             if i.kind in (KIND_H1, KIND_H2, KIND_H3, KIND_H4)]
    return max((i.level for i in heads), default=0)


def _demote_excess_levels(result: DocAnalysis) -> None:
    """标题只有两级且级别超过 2 时，整体上提，避免出现"只有三级标题"。"""
    heads = [i for i in result.paragraphs
             if i.kind in (KIND_H1, KIND_H2, KIND_H3, KIND_H4)]
    if not heads:
        return
    levels = sorted({i.level for i in heads})
    if len(levels) <= 2 and levels[0] > 1:
        shift = levels[0] - 1
        for info in heads:
            info.level -= shift
            info.kind = {1: KIND_H1, 2: KIND_H2, 3: KIND_H3, 4: KIND_H4}[info.level]
        counts: Dict[int, int] = {}
        for info in heads:
            counts[info.level] = counts.get(info.level, 0) + 1
        result.heading_counts = counts


def _fix_ambiguous_arabic(result: DocAnalysis) -> None:
    """"1." 编号既可能是一级也可能在"一、"体系下是二级。

    - 文档里同时存在"一、"或"1.1"这类证据 → "1." 视为二级；
    - 否则（例如 1./2./3. 并列成篇）→ 视为一级。
    """
    heads = [i for i in result.paragraphs
             if i.kind in (KIND_H1, KIND_H2, KIND_H3, KIND_H4)]
    if not heads:
        return
    dotted = [i for i in heads if _is_arabic_dotted(i.numbering)]
    if not dotted:
        return
    has_strong_h1 = any(
        _is_cn_number(i.numbering) or _is_multi_level(i.numbering)
        for i in heads if i not in dotted)
    if has_strong_h1:
        return
    changed = False
    for info in dotted:
        if info.level > 1:
            info.level = 1
            info.kind = KIND_H1
            changed = True
    if not changed:
        return
    for info in heads:
        if info in dotted or _is_cn_number(info.numbering):
            continue
        # 其余标题跟随下移，保持相对层级
        info.level = min(info.level, 3)
        info.kind = {1: KIND_H1, 2: KIND_H2, 3: KIND_H3, 4: KIND_H4}[info.level]
    counts: Dict[int, int] = {}
    for info in heads:
        counts[info.level] = counts.get(info.level, 0) + 1
    result.heading_counts = counts


def _fix_numbered_levels(result: DocAnalysis) -> None:
    """一级标题用"一、"时，把"1." / "1.1" 这类编号整体下移一级。

    典型公文/流程写法：一、 目的 / 1. 准备阶段 / 1.1 细节，
    直接按编号深度识别会得到"一、=1 级、1.=2 级、1.1=2 级"，
    下移后变成 1 / 2 / 3，层次与阅读顺序一致。
    """
    heads = [i for i in result.paragraphs
             if i.kind in (KIND_H1, KIND_H2, KIND_H3, KIND_H4)]
    if not heads:
        return
    if not any(_is_cn_number(i.numbering) for i in heads):
        return
    if any(_is_arabic_dotted(i.numbering) for i in heads):
        for info in heads:
            if _is_cn_number(info.numbering):
                continue
            info.level = min(info.level + 1, 4)
            info.kind = {1: KIND_H1, 2: KIND_H2, 3: KIND_H3, 4: KIND_H4}[info.level]
        counts: Dict[int, int] = {}
        for info in heads:
            counts[info.level] = counts.get(info.level, 0) + 1
        result.heading_counts = counts


def _has_drawing(p: Paragraph) -> bool:
    xml = p._p
    if xml.findall(".//" + qn("w:drawing")) or xml.findall(".//" + qn("w:pict")):
        return True
    return False


def _has_formula(p: Paragraph) -> bool:
    return bool(p._p.findall(".//" + qn("m:oMath"))) or bool(
        p._p.findall(".//" + qn("m:oMathPara"))
    )


def _short_and_unnumbered_punct(text: str) -> bool:
    stripped = text.strip()
    if not stripped:
        return False
    if re.search(r"[。；;！？!?]", stripped):
        return False
    if stripped.endswith(("，", ",", "：", ":")) and len(stripped) > 12:
        return False
    return True


def _document_name_hits(text: str, doc_name: str) -> bool:
    """该段落是否就是在重复文档文件名（去掉扩展名后相互包含）。"""
    if not doc_name:
        return False
    a = re.sub(r"[\s_\-\u3000]+", "", text.strip())
    b = re.sub(r"[\s_\-\u3000]+", "", doc_name)
    if not a or not b:
        return False
    return a == b or (len(a) >= 3 and (a in b or b in a))


def _merge_docname_first_line(result: DocAnalysis, doc_name: str) -> None:
    """首行就是文档名、且下一段是一级标题时，把首行与大标题合并展示。

    例：文件名"技术说明.docx"，首行"技术说明"，下一行"1 概述"。
    首行作为大标题会触发"大标题下插分页 / 一级标题另起页"，
    结果整页只剩一行标题；这里把首行降级为一级标题，两行连排。
    """
    heads = [i for i in result.paragraphs if i.kind in HEADING_LEVELS]
    if not heads:
        return
    first_head = heads[0]
    idx = result.paragraphs.index(first_head) if first_head in result.paragraphs else -1
    if idx <= 0:
        return
    prev = result.paragraphs[idx - 1]
    if prev.kind != KIND_TITLE:
        return
    if not _document_name_hits(prev.text, doc_name):
        return
    prev.kind = KIND_H1
    prev.level = 1
    prev.numbering = None
    result.title_count = max(0, result.title_count - 1)
    result.heading_counts[1] = result.heading_counts.get(1, 0) + 1


def analyze(document: DocumentObject) -> DocAnalysis:
    """分析整篇文档结构。"""
    result = DocAnalysis()
    blocks = iter_block_paragraphs(document)
    core_props = getattr(document, "core_properties", None)
    cover_title = ""
    try:
        cover_title = (getattr(core_props, "title", "") or "").strip()
    except Exception:
        cover_title = ""

    # 记录段落是否位于表格内
    in_table_ids = set()
    for tbl in document.tables:
        for row in tbl.rows:
            for cell in row.cells:
                for p in cell.paragraphs:
                    in_table_ids.add(id(p._p))
    result.tables = list(document.tables)

    try:
        result.section_count = max(1, len(document.sections))
    except Exception:
        result.section_count = 1

    first_non_empty_seen = False

    for idx, p in enumerate(blocks):
        text = p.text or ""
        stripped = text.strip()
        info = ParaInfo(index=idx, paragraph=p, text=text,
                        in_table=id(p._p) in in_table_ids,
                        char_len=len(stripped),
                        has_drawing=_has_drawing(p),
                        has_formula=_has_formula(p))

        if not stripped and not info.has_drawing and not info.has_formula:
            info.kind = KIND_EMPTY
            result.paragraphs.append(info)
            continue

        if not first_non_empty_seen and not info.in_table:
            first_non_empty_seen = True

        # 1) 样式给出的标题级别
        level = _heading_level_from_style(p)
        if level is None:
            lv = _outline_level(p)
            if lv is not None and lv <= 4:
                level = lv

        # 2) 编号/格式启发式
        number, num_level = _detect_numbering(stripped)
        if level is None or level == 0:
            if _RE_CAPTION.match(stripped) and info.char_len <= 60 and not stripped.endswith("。"):
                info.kind = KIND_CAPTION
                result.paragraphs.append(info)
                continue
            if num_level is not None and info.char_len <= _HEADING_MAX_CHARS and \
                    _short_and_unnumbered_punct(stripped):
                level = num_level
                info.numbering = number
            elif _RE_APPENDIX.match(stripped) and info.char_len <= _HEADING_MAX_CHARS:
                level = 1
            elif level is None and _looks_like_heading(p, stripped, idx):
                level = 2 if info.char_len <= 25 else 3

        # 首行居中 / 大字号 → 文档大标题（带编号的段落不当作标题）
        if not info.in_table and idx <= 8 and result.title_count == 0 \
                and info.char_len <= _TITLE_MAX_CHARS and num_level is None \
                and _is_title_candidate(p, stripped):
            info.kind = KIND_TITLE
            info.level = 0
            result.title_count += 1
            result.paragraphs.append(info)
            continue

        if level == -1:
            info.kind = KIND_SUBTITLE
            info.level = 0
            result.paragraphs.append(info)
            continue

        if level == 0:
            info.kind = KIND_TITLE
            info.level = 0
            result.title_count += 1
            result.paragraphs.append(info)
            continue

        if level and level >= 1:
            # 大标题：文档前部、居中的一句话
            if idx < 8 and _is_title_candidate(p, stripped) and result.title_count == 0 \
                    and level >= 2 and not info.in_table:
                info.kind = KIND_TITLE
                info.level = 0
                result.title_count += 1
            else:
                lv = min(level, 4)
                info.kind = {1: KIND_H1, 2: KIND_H2, 3: KIND_H3, 4: KIND_H4}[lv]
                info.level = lv
                info.numbering = info.numbering or number
                result.heading_counts[lv] = result.heading_counts.get(lv, 0) + 1
            result.paragraphs.append(info)
            continue

        # 3) 正文 / 列表
        if _RE_LIST_MARK.match(stripped) and info.char_len <= 200:
            info.kind = KIND_LIST
        else:
            info.kind = KIND_BODY
            result.body_char_count += info.char_len
            result.max_line_chars = max(result.max_line_chars, info.char_len)

        if "结" in stripped[:2] and "论" in stripped[:4]:
            result.has_conclusion_label = True

        result.paragraphs.append(info)

    # 无标题文档：把最突出的短段落提升为标题，保证排版有层次
    if result.heading_count == 0:
        _promote_implied_headings(result)

    _normalize_levels(result)
    _fix_ambiguous_arabic(result)
    _fix_numbered_levels(result)
    _demote_excess_levels(result)
    _merge_cover_title(result)
    if cover_title:
        _merge_docname_first_line(result, cover_title)
    return result


def _normalize_levels(result: DocAnalysis) -> None:
    """修正层级：补齐缺失的上级标题，避免出现"只有三级标题、没有二级"。

    例：1 / 1.1 / 1.2 被识别成 1 级和 3 级时，把 1.1 提升为 2 级；
    文档缺少 2 级标题时，3 级标题整体提升为 2 级。
    """
    kinds = (KIND_H1, KIND_H2, KIND_H3, KIND_H4)

    def rebuild() -> None:
        counts: Dict[int, int] = {}
        for info in result.paragraphs:
            if info.kind in kinds:
                counts[info.level] = counts.get(info.level, 0) + 1
        result.heading_counts = counts

    # 1) 没有一级标题 → 整体上移，保证有顶级标题
    present = sorted(lv for lv in (1, 2, 3, 4) if result.heading_counts.get(lv))
    if present and present[0] > 1:
        shift = present[0] - 1
        for info in result.paragraphs:
            if info.kind in kinds:
                info.level = max(1, info.level - shift)
                info.kind = kinds[info.level - 1]
        rebuild()
        present = [lv - shift for lv in present]

    # 2) 缺少二级标题（只有一级和三级）→ 三级提升为二级
    if present and 2 not in present and 3 in present:
        for info in result.paragraphs:
            if info.kind == KIND_H3:
                info.kind = KIND_H2
                info.level = 2
            elif info.kind == KIND_H4:
                info.kind = KIND_H3
                info.level = 3
        rebuild()

    # 3) 缺少三级标题（只有一级和四级）→ 四级提升为三级
    present = sorted(lv for lv in (1, 2, 3, 4) if result.heading_counts.get(lv))
    if present and 3 not in present and 4 in present:
        for info in result.paragraphs:
            if info.kind == KIND_H4:
                info.kind = KIND_H3
                info.level = 3
        rebuild()


def _merge_cover_title(result: DocAnalysis) -> None:
    """文档标题与第一个一级标题之间没有正文时，两段都视为一级标题。

    例："园区网络系统技术说明" + "1 概述"，若把前者当大标题、后者加"段前分页"，
    就会产生"整页只有一行标题"的空白页；改为同级标题后自然连排。
    """
    seq = [i for i in result.paragraphs
           if i.kind != KIND_EMPTY and not i.in_table]
    if len(seq) < 2:
        return
    if seq[0].kind != KIND_TITLE or seq[1].kind != KIND_H1:
        return
    seq[0].kind = KIND_H1
    seq[0].level = 1
    result.title_count = 0
    result.heading_counts[1] = result.heading_counts.get(1, 0) + 1


def _is_title_candidate(p: Paragraph, stripped: str) -> bool:
    if len(stripped) > _TITLE_MAX_CHARS:
        return False
    align = p.alignment if p.alignment is not None else None
    centered = align in (WD_ALIGN_PARAGRAPH.CENTER,)
    size = _max_run_size(p)
    return centered or (size is not None and size >= 18)


def _looks_like_heading(p: Paragraph, stripped: str, idx: int) -> bool:
    if len(stripped) > _HEADING_MAX_CHARS:
        return False
    if stripped.endswith(("。", "；", ";", "，")) and len(stripped) > 15:
        return False
    try:
        has_list = "List" in (p.style.name or "")
    except Exception:
        has_list = False
    if has_list:
        return False
    size = _max_run_size(p)
    if size is None:
        return False
    if size >= 14 and len(stripped) <= 30:
        return True
    if size >= 13 and any(r.bold for r in p.runs[:6] if r.bold) and len(stripped) <= 25:
        return True
    return False


def _max_run_size(p: Paragraph) -> Optional[float]:
    sizes = []
    for r in p.runs[:12]:
        if r.font.size is not None:
            try:
                sizes.append(r.font.size.pt)
            except Exception:
                pass
    return max(sizes) if sizes else None


def _promote_implied_headings(result: DocAnalysis) -> None:
    for info in result.paragraphs:
        if info.in_table or info.kind != KIND_BODY or info.char_len > 30:
            continue
        text = info.text.strip()
        if not text or text.endswith("。"):
            continue
        if info.numbering:
            lv = 4 if text.count(".") >= 2 else 3
            info.kind = {3: KIND_H3, 4: KIND_H4}[lv]
            info.level = lv
            result.heading_counts[lv] = result.heading_counts.get(lv, 0) + 1
        else:
            # 无编号的短段落（如"核心特性""应用场景"）当作三级小标题
            info.kind = KIND_H3
            info.level = 3
            result.heading_counts[3] = result.heading_counts.get(3, 0) + 1
