"""图片尺寸自适应（可选开关）。

目标：让文档里的图片"放得下、看得清、不糊、不跑版"。

规则（只在**确实需要**时才改，绝不无故放大）：
1. 计算可用版面：Word 取当前节的「页宽 − 左右页边距」；Excel 取模板给的厘米数。
2. 图片宽度超过可用宽度 × ``image_max_width_ratio`` → 等比缩小；
3. 图片高度超过 ``image_max_height_cm`` → 等比缩小；
4. 宽度小于可用宽度 × ``image_min_width_ratio`` 的图片视为图标 / 小徽标，跳过；
5. 默认**不放大**（``image_allow_upscale=False``），避免小图拉伸变糊。

Word 同时处理"嵌入式"（inline）与"浮动式"（anchor）两种图片；
Excel 处理工作表中的图片对象（不改变锚点位置，只改显示尺寸）。
"""

from __future__ import annotations

import logging
from dataclasses import dataclass, field
from typing import Dict, List, Tuple

from docx.shared import Cm, Emu
from docx.oxml.ns import qn

log = logging.getLogger("wxfixer.images")

EMU_PER_CM = 360000
EMU_PER_PX = 9525          # 96 DPI
TWIP_PER_CM = 566.929


@dataclass
class ImageStats:
    total: int = 0
    resized: int = 0
    skipped_small: int = 0
    skipped_fit: int = 0
    details: List[str] = field(default_factory=list)

    def summary(self) -> str:
        if not self.total:
            return "文档内未发现图片"
        if not self.resized:
            return f"图片 {self.total} 张，尺寸均已合适，未调整"
        return f"调整 {self.resized}/{self.total} 张图片尺寸"


# ---------------------------------------------------------------------------
# Word
# ---------------------------------------------------------------------------
def _section_body_width_emu(section) -> int:
    """可用正文宽度（EMU）。取不到时按 A4 常规页边距估算。"""
    try:
        width = section.page_width
        left = section.left_margin
        right = section.right_margin
        if width and left is not None and right is not None:
            usable = int(width) - int(left) - int(right)
            if usable > EMU_PER_CM * 5:
                return usable
    except Exception:
        pass
    return int(Cm(16.0))


def _shape_owner_paragraph(doc, shape_element):
    """找到包含该 drawing 的段落对象（用于居中判断）。"""
    from docx.text.paragraph import Paragraph

    node = shape_element
    while node is not None and node.tag != qn("w:p"):
        node = node.getparent()
    if node is None:
        return None
    return Paragraph(node, doc)


def _scale_inline(shape, ratio: float) -> Tuple[int, int]:
    """嵌入式图形：直接改 InlineShape 的宽高（python-docx 会同步 XML）。"""
    new_w = max(1, int(int(shape.width) * ratio))
    new_h = max(1, int(int(shape.height) * ratio))
    shape.width = Emu(new_w)
    shape.height = Emu(new_h)
    return new_w, new_h


_NS_WP = "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"
_NS_A = "http://schemas.openxmlformats.org/drawingml/2006/main"


def _scale_anchor(anchor, ratio: float) -> Tuple[int, int] | None:
    """浮动式图形：wp:extent 与 a:ext 都要等比缩放，否则显示比例错乱。"""
    extent = anchor.find(qn("wp:extent"))
    if extent is None:
        return None
    cx = int(extent.get("cx") or 0)
    cy = int(extent.get("cy") or 0)
    if cx <= 0 or cy <= 0:
        return None
    new_cx = max(1, int(cx * ratio))
    new_cy = max(1, int(cy * ratio))
    extent.set("cx", str(new_cx))
    extent.set("cy", str(new_cy))
    for ext in anchor.iter():
        if not isinstance(ext.tag, str) or not ext.tag.endswith("}ext"):
            continue
        pcx, pcy = ext.get("cx"), ext.get("cy")
        if pcx is None or pcy is None:
            continue
        try:
            if abs(int(pcx) - cx) <= 2 and abs(int(pcy) - cy) <= 2:
                ext.set("cx", str(new_cx))
                ext.set("cy", str(new_cy))
        except (TypeError, ValueError):
            continue
    return new_cx, new_cy


def resize_word_images(doc, spec, enabled: bool) -> Tuple[ImageStats, int]:
    """按模板参数调整 Word 内图片尺寸，返回 (统计, 居中段落数)。"""
    stats = ImageStats()
    centered = 0
    if not enabled:
        return stats, centered

    try:
        sections = list(doc.sections)
    except Exception:
        sections = []
    default_section = sections[0] if sections else None

    max_ratio = max(0.2, min(1.0, float(getattr(spec, "image_max_width_ratio", 0.92))))
    min_ratio = max(0.0, min(0.9, float(getattr(spec, "image_min_width_ratio", 0.30))))
    max_h = int(Cm(float(getattr(spec, "image_max_height_cm", 17.5))))
    allow_up = bool(getattr(spec, "image_allow_upscale", False))
    do_center = bool(getattr(spec, "image_center", True))

    def _limit_ratio(cur_w: int, cur_h: int, usable: int) -> float:
        """返回应当乘上的比例（<=1 表示缩小）。"""
        caps = []
        if usable > 0 and cur_w > 0:
            caps.append((usable * max_ratio) / cur_w)
        if max_h > 0 and cur_h > 0:
            caps.append(max_h / cur_h)
        if not caps:
            return 1.0
        return min(caps)

    # ---- 嵌入式 ----
    for shape in list(doc.inline_shapes):
        stats.total += 1
        try:
            cur_w = int(shape.width or 0)
            cur_h = int(shape.height or 0)
        except Exception:
            continue
        if cur_w <= 0 or cur_h <= 0:
            continue
        para = _shape_owner_paragraph(doc, shape._inline)
        section = None
        if para is not None:
            # 图片可能位于不同的节（分节符），尽量取所在节
            idx = 0
            for i, s in enumerate(sections):
                try:
                    if s._sectPr is not None and s._sectPr.getparent() is not None:
                        pass
                except Exception:
                    continue
                idx = i
            section = sections[idx] if sections else None
        section = section or default_section
        usable = _section_body_width_emu(section) if section is not None else int(Cm(16.0))
        if usable > 0 and cur_w < usable * min_ratio and not allow_up:
            stats.skipped_small += 1
            continue
        ratio = _limit_ratio(cur_w, cur_h, usable)
        if allow_up and ratio < 1.0:
            ratio = 1.0
        if ratio >= 0.999:
            stats.skipped_fit += 1
            _maybe_center(doc, para, do_center)
            continue
        try:
            new_w, new_h = _scale_inline(shape, ratio)
        except Exception as exc:
            log.warning("图片缩放失败：%s", exc)
            continue
        stats.resized += 1
        stats.details.append(
            f"{cur_w / EMU_PER_CM:.1f}×{cur_h / EMU_PER_CM:.1f}cm → "
            f"{new_w / EMU_PER_CM:.1f}×{new_h / EMU_PER_CM:.1f}cm"
        )
        _maybe_center(doc, para, do_center)

    # ---- 浮动式（anchor）----
    for drawing in doc.element.body.iter(qn("w:drawing")):
        anchor = drawing.find(qn("wp:anchor"))
        if anchor is None:
            continue
        stats.total += 1
        extent = anchor.find(qn("wp:extent"))
        if extent is None:
            continue
        cur_w = int(extent.get("cx") or 0)
        cur_h = int(extent.get("cy") or 0)
        if cur_w <= 0 or cur_h <= 0:
            continue
        usable = _section_body_width_emu(default_section) if default_section is not None else int(Cm(16.0))
        if usable > 0 and cur_w < usable * min_ratio and not allow_up:
            stats.skipped_small += 1
            continue
        ratio = _limit_ratio(cur_w, cur_h, usable)
        if allow_up and ratio < 1.0:
            ratio = 1.0
        if ratio >= 0.999:
            stats.skipped_fit += 1
            continue
        result = _scale_anchor(anchor, ratio)
        if not result:
            continue
        stats.resized += 1
        stats.details.append(
            f"浮动图 {cur_w / EMU_PER_CM:.1f}×{cur_h / EMU_PER_CM:.1f}cm → "
            f"{result[0] / EMU_PER_CM:.1f}×{result[1] / EMU_PER_CM:.1f}cm"
        )
    return stats, centered


def _maybe_center(doc, para, do_center: bool) -> None:
    """图片单独成段时居中，避免图片贴左边缘。"""
    if not do_center or para is None:
        return
    try:
        text = "".join(t.text or "" for t in para._p.iter(qn("w:t")))
        if text.strip():
            return  # 图文混排，不动对齐
        from . import ooxml as X

        ppr = para._p.get_or_add_pPr()
        X.set_align(ppr, "center")
        # 居中后首行缩进会让图片整体右移，必须清零，否则"看起来没居中"
        X.set_para_indent(ppr, first_line_chars=0)
    except Exception:
        pass


# ---------------------------------------------------------------------------
# Excel
# ---------------------------------------------------------------------------
def resize_excel_images(ws, spec, enabled: bool) -> ImageStats:
    """调整工作表内图片尺寸（保留锚点，仅缩放显示大小）。

    注意单位：openpyxl 里 anchor.ext.cx/cy 是 **EMU**（1px = 9525 EMU），
    而 Image.width/height 是像素。之前混用导致 ext 被写成 680 EMU（≈0.07px），
    图片在 Excel/PDF 里几乎消失，这里统一用 EMU 计算。
    """
    stats = ImageStats()
    if not enabled:
        return stats
    images = getattr(ws, "_images", None) or []
    if not images:
        return stats
    max_w_cm = float(getattr(spec, "excel_image_max_width_cm", 18.0) or 18.0)
    max_h_cm = float(getattr(spec, "excel_image_max_height_cm", 12.0) or 12.0)
    max_w_emu = max(EMU_PER_PX * 16, int(max_w_cm * EMU_PER_CM))
    max_h_emu = max(EMU_PER_PX * 16, int(max_h_cm * EMU_PER_CM))
    for img in images:
        stats.total += 1
        cur_w, cur_h = _anchored_size(img)          # EMU
        if cur_w <= 0 or cur_h <= 0:
            continue
        ratio = min(1.0, max_w_emu / cur_w, max_h_emu / cur_h)
        if ratio >= 0.999:
            stats.skipped_fit += 1
            continue
        new_w = int(round(cur_w * ratio))
        new_h = int(round(cur_h * ratio))
        try:
            _apply_excel_size(img, new_w, new_h)
        except Exception as exc:  # noqa: BLE001
            log.warning("Excel 图片缩放失败：%s", exc)
            continue
        stats.resized += 1
        stats.details.append(
            f"{cur_w / EMU_PER_CM:.1f}×{cur_h / EMU_PER_CM:.1f}cm → "
            f"{new_w / EMU_PER_CM:.1f}×{new_h / EMU_PER_CM:.1f}cm")
    return stats


def _anchored_size(img) -> Tuple[int, int]:
    """取图片当前显示尺寸，统一返回 **EMU**。"""
    anchor = getattr(img, "anchor", None)
    ext = getattr(anchor, "ext", None)
    if ext is not None:
        try:
            w = int(getattr(ext, "cx", 0) or 0)
            h = int(getattr(ext, "cy", 0) or 0)
            if w > 0 and h > 0:
                return w, h
        except Exception:  # noqa: BLE001
            pass
    try:
        return (int(getattr(img, "width", 0) or 0) * EMU_PER_PX,
                int(getattr(img, "height", 0) or 0) * EMU_PER_PX)
    except Exception:  # noqa: BLE001
        return 0, 0


def _apply_excel_size(img, new_w_emu: int, new_h_emu: int) -> None:
    """写回尺寸：anchor.ext 用 EMU，Image.width/height 用像素，两者同步。"""
    new_w_emu = max(1, int(new_w_emu))
    new_h_emu = max(1, int(new_h_emu))
    anchor = getattr(img, "anchor", None)
    ext = getattr(anchor, "ext", None)
    if ext is not None:
        ext.cx = new_w_emu
        ext.cy = new_h_emu
    img.width = max(1, int(round(new_w_emu / EMU_PER_PX)))
    img.height = max(1, int(round(new_h_emu / EMU_PER_PX)))


def _sync_anchor_ext(img) -> None:
    """把 Image.width/height（像素）同步回 anchor.ext（EMU）。"""
    anchor = getattr(img, "anchor", None)
    ext = getattr(anchor, "ext", None)
    if ext is None:
        return
    try:
        ext.cx = int(img.width) * EMU_PER_PX
        ext.cy = int(img.height) * EMU_PER_PX
    except Exception:
        pass


# ---------------------------------------------------------------------------
def describe(stats: ImageStats) -> str:
    return stats.summary()
