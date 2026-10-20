"""OOXML 顺序安全写入辅助。

python-docx 只对少数元素提供 ``get_or_add_*``，直接 ``append`` 会破坏
schema 规定的子元素顺序，Word/WPS 打开可能提示"文件已损坏"。
这里提供通用顺序插入函数；顺序表只列常用元素，未列出的元素插在已有
同族元素之后（保守做法，对 Word 无副作用）。
"""

from __future__ import annotations

from lxml import etree
from docx.oxml import OxmlElement
from docx.oxml.ns import qn

W_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
R_NS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"

# ---------------------------------------------------------------------------
# 合法子元素顺序表（摘自 ECMA-376 Part 1）
# ---------------------------------------------------------------------------
TCPR_SEQ = [
    "w:cnfStyle", "w:tcW", "w:gridSpan", "w:hMerge", "w:vMerge", "w:tcBorders",
    "w:shd", "w:noWrap", "w:tcMar", "w:textDirection", "w:tcFitText",
    "w:vAlign", "w:hideMark", "w:headers", "w:cellIns", "w:cellDel",
    "w:cellMerge", "w:tcPrChange",
]

TRPR_SEQ = [
    "w:cnfStyle", "w:divId", "w:gridBefore", "w:gridAfter", "w:wBefore",
    "w:wAfter", "w:cantSplit", "w:trHeight", "w:tblHeader", "w:tblCellSpacing",
    "w:jc", "w:hidden", "w:ins", "w:del", "w:trPrChange",
]

TBLPR_SEQ = [
    "w:tblStyle", "w:tblpPr", "w:tblOverlap", "w:bidiVisual",
    "w:tblStyleRowBandSize", "w:tblStyleColBandSize", "w:tblW", "w:jc",
    "w:tblCellSpacing", "w:tblInd", "w:tblBorders", "w:shd", "w:tblLayout",
    "w:tblCellMar", "w:tblLook", "w:tblCaption", "w:tblDescription",
    "w:tblPrChange",
]

PPR_SEQ = [
    "w:pStyle", "w:keepNext", "w:keepLines", "w:pageBreakBefore", "w:framePr",
    "w:widowControl", "w:numPr", "w:suppressLineNumbers", "w:pBdr", "w:shd",
    "w:tabs", "w:suppressAutoHyphens", "w:kinsoku", "w:wordWrap",
    "w:overflowPunct", "w:topLinePunct", "w:autoSpaceDE", "w:autoSpaceDN",
    "w:bidi", "w:adjustRightInd", "w:snapToGrid", "w:spacing", "w:ind",
    "w:contextualSpacing", "w:mirrorIndents", "w:suppressOverlap", "w:jc",
    "w:textDirection", "w:textAlignment", "w:textboxTightWrap",
    "w:outlineLvl", "w:divId", "w:cnfStyle", "w:rPr", "w:sectPr",
    "w:pPrChange",
]

RPR_SEQ = [
    "w:rStyle", "w:rFonts", "w:b", "w:bCs", "w:i", "w:iCs", "w:caps",
    "w:smallCaps", "w:strike", "w:dstrike", "w:outline", "w:shadow",
    "w:emboss", "w:imprint", "w:noProof", "w:snapToGrid", "w:vanish",
    "w:webHidden", "w:color", "w:spacing", "w:w", "w:kern", "w:position",
    "w:sz", "w:szCs", "w:highlight", "w:u", "w:effect", "w:bdr", "w:shd",
    "w:fitText", "w:vertAlign", "w:rtl", "w:cs", "w:em", "w:lang",
    "w:eastAsianLayout", "w:specVanish", "w:oMath",
]

TBLBORDERS_SEQ = [
    "w:top", "w:left", "w:bottom", "w:right", "w:insideH", "w:insideV",
    "w:tl2br", "w:tr2bl",
]

TCBORDERS_SEQ = TBLBORDERS_SEQ

SECTPR_SEQ = [
    "w:headerReference", "w:footerReference", "w:footnotePr", "w:endnotePr",
    "w:type", "w:pgSz", "w:pgMar", "w:paperSrc", "w:pgBorders", "w:lnNumType",
    "w:pgNumType", "w:cols", "w:formProt", "w:vAlign", "w:noEndnote",
    "w:titlePg", "w:textDirection", "w:bidi", "w:rtlGutter", "w:docGrid",
    "w:printerSettings", "w:sectPrChange",
]


def insert_ordered(parent, tag: str, seq) -> etree._Element:
    """在 ``parent`` 中按 ``seq`` 顺序获得（或插入）``tag`` 元素。"""
    existing = parent.find(qn(tag))
    if existing is not None:
        return existing
    el = OxmlElement(tag)
    try:
        idx = seq.index(tag)
    except ValueError:  # 未登记的元素：直接追加
        parent.append(el)
        return el
    for sib_tag in seq[idx + 1:]:
        sib = parent.find(qn(sib_tag))
        if sib is not None:
            sib.addprevious(el)
            return el
    parent.append(el)
    return el


def remove_children(parent, *tags: str) -> None:
    for tag in tags:
        for el in parent.findall(qn(tag)):
            parent.remove(el)


# ---------------------------------------------------------------------------
# 常用的高层写入函数
# ---------------------------------------------------------------------------
def set_shading(parent, fill: str, seq=TCPR_SEQ, tag: str = "w:shd"):
    """设置底纹。``fill`` 为 6 位十六进制（不含 #）。"""
    shd = insert_ordered(parent, tag, seq)
    shd.set(qn("w:val"), "clear")
    shd.set(qn("w:color"), "auto")
    shd.set(qn("w:fill"), fill)
    return shd


def clear_shading(parent, tag: str = "w:shd") -> None:
    for el in parent.findall(qn(tag)):
        parent.remove(el)


def set_run_fonts(rpr, east_asia: str | None, ascii_font: str | None) -> None:
    """同时写 ascii / hAnsi / eastAsia，避免中文回落成西文字体。

    Word 的字体主题属性（``w:asciiTheme`` / ``w:eastAsiaTheme`` 等）优先级高于
    显式字体名，必须删除，否则换了字体名也不生效。
    """
    rf = insert_ordered(rpr, "w:rFonts", RPR_SEQ)
    for attr in ("w:asciiTheme", "w:hAnsiTheme", "w:eastAsiaTheme", "w:cstheme",
                 "w:hint"):
        if rf.get(qn(attr)) is not None:
            del rf.attrib[qn(attr)]
    if ascii_font:
        rf.set(qn("w:ascii"), ascii_font)
        rf.set(qn("w:hAnsi"), ascii_font)
        rf.set(qn("w:cs"), ascii_font)
    if east_asia:
        rf.set(qn("w:eastAsia"), east_asia)


def set_run_size(rpr, size_pt: float) -> None:
    half = max(2, int(round(size_pt * 2)))
    for tag in ("w:sz", "w:szCs"):
        el = insert_ordered(rpr, tag, RPR_SEQ)
        el.set(qn("w:val"), str(half))


def set_run_bold(rpr, bold: bool) -> None:
    for tag in ("w:b", "w:bCs"):
        el = insert_ordered(rpr, tag, RPR_SEQ)
        el.set(qn("w:val"), "1" if bold else "0")


def set_run_color(rpr, hex_color: str) -> None:
    el = insert_ordered(rpr, "w:color", RPR_SEQ)
    el.set(qn("w:val"), hex_color.replace("#", "").upper())
    if el.get(qn("w:themeColor")):
        del el.attrib[qn("w:themeColor")]
    if el.get(qn("w:themeTint")):
        del el.attrib[qn("w:themeTint")]


def set_para_spacing(ppr, before_pt=None, after_pt=None, line=None,
                     line_rule: str | None = None) -> None:
    sp = insert_ordered(ppr, "w:spacing", PPR_SEQ)
    if before_pt is not None:
        sp.set(qn("w:before"), str(int(round(before_pt * 20))))
        if sp.get(qn("w:beforeLines")):
            del sp.attrib[qn("w:beforeLines")]
        sp.set(qn("w:beforeAutospacing"), "0")
    if after_pt is not None:
        sp.set(qn("w:after"), str(int(round(after_pt * 20))))
        if sp.get(qn("w:afterLines")):
            del sp.attrib[qn("w:afterLines")]
        sp.set(qn("w:afterAutospacing"), "0")
    if line is not None:
        sp.set(qn("w:line"), str(int(round(line * 240))))
        sp.set(qn("w:lineRule"), line_rule or "auto")


def set_para_indent(ppr, first_line_chars=None, hanging_cm=None,
                    left_cm=None, char_unit: float = 1.0) -> None:
    """首行缩进用"字符"单位（firstLineChars），Word 中显示为 2 字符。"""
    ind = insert_ordered(ppr, "w:ind", PPR_SEQ)
    if first_line_chars is not None:
        if first_line_chars > 0:
            chars = int(round(first_line_chars * 100))
            ind.set(qn("w:firstLineChars"), str(chars))
            ind.set(qn("w:firstLine"), str(int(round(first_line_chars * char_unit * 567))))
            ind.attrib.pop(qn("w:hanging"), None)
            ind.attrib.pop(qn("w:hangingChars"), None)
        else:
            for k in ("w:firstLine", "w:firstLineChars", "w:hanging",
                      "w:hangingChars"):
                ind.attrib.pop(qn(k), None)
    if hanging_cm is not None:
        ind.set(qn("w:hanging"), str(int(round(hanging_cm * 567))))
        ind.set(qn("w:hangingChars"), "0")
        ind.attrib.pop(qn("w:firstLine"), None)
        ind.attrib.pop(qn("w:firstLineChars"), None)
    if left_cm is not None:
        ind.set(qn("w:left"), str(int(round(left_cm * 567))))
        ind.set(qn("w:leftChars"), "0")


def set_align(ppr, align: str) -> None:
    mapping = {"left": "left", "center": "center", "right": "right",
               "justify": "both", "both": "both", "distribute": "distribute"}
    el = insert_ordered(ppr, "w:jc", PPR_SEQ)
    el.set(qn("w:val"), mapping.get(align, "both"))


def set_keep_next(ppr, on: bool = True) -> None:
    el = insert_ordered(ppr, "w:keepNext", PPR_SEQ)
    el.set(qn("w:val"), "1" if on else "0")


def set_keep_lines(ppr, on: bool = True) -> None:
    el = insert_ordered(ppr, "w:keepLines", PPR_SEQ)
    el.set(qn("w:val"), "1" if on else "0")


def set_page_break_before(ppr, on: bool = True) -> None:
    el = insert_ordered(ppr, "w:pageBreakBefore", PPR_SEQ)
    el.set(qn("w:val"), "1" if on else "0")


def set_outline_level(ppr, level: int) -> None:
    el = insert_ordered(ppr, "w:outlineLvl", PPR_SEQ)
    el.set(qn("w:val"), str(level))


def set_border(parent, edges, color: str, size_eighths: int = 4,
               seq=TBLBORDERS_SEQ, val: str = "single") -> None:
    """edges: (tag_name,) 如 ("w:top","w:insideH")。"""
    for edge in edges:
        el = insert_ordered(parent, edge, seq)
        el.set(qn("w:val"), val)
        el.set(qn("w:sz"), str(size_eighths))
        el.set(qn("w:space"), "0")
        el.set(qn("w:color"), color.replace("#", "").upper())


def strip_borders(parent, seq=TBLBORDERS_SEQ) -> None:
    for edge in TBLBORDERS_SEQ:
        for el in parent.findall(qn(edge)):
            parent.remove(el)


def ensure_tbl_layout_fixed(tblpr) -> None:
    lay = insert_ordered(tblpr, "w:tblLayout", TBLPR_SEQ)
    lay.set(qn("w:type"), "fixed")


def set_tbl_width_pct(tblpr, pct: int) -> None:
    w = insert_ordered(tblpr, "w:tblW", TBLPR_SEQ)
    w.set(qn("w:type"), "pct")
    w.set(qn("w:w"), str(int(pct * 50)))  # pct 单位是 1/50%


def localname(el) -> str:
    tag = el.tag
    return tag.split("}")[-1] if isinstance(tag, str) else ""
