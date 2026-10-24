"""Excel 排版优化器。

同样只动"格式"：字体、字号、颜色、填充、边框、行高列宽、对齐换行、
冻结窗格、筛选、页面设置与打印标题。**不修改任何单元格的值与公式。**

保留项（openpyxl 读写天然保留）：公式、批注、条件格式、数据验证、图片、
图表、合并单元格、名称定义。
"""

from __future__ import annotations

import hashlib
import logging
import re
from typing import Dict, List, Optional, Tuple

from openpyxl import load_workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.dimensions import ColumnDimension
from openpyxl.worksheet.page import PageMargins, PrintPageSetup

from ..templates.schema import LayoutSpec, ToggleOptions
from . import content_optimizer as C
from . import image_sizer as IMG

log = logging.getLogger("wxfixer.excel")

_MAX_SCAN_ROWS = 5000
_MAX_SCAN_COLS = 200
_NUMERIC_FORMATS = set("bfeEgG%0#")
# 首行只有一格有字、且文字是这些常见表头词时，不当作表格标题
_HEADER_WORDS = {
    "序号", "名称", "编号", "项目", "日期", "时间", "类型", "状态", "备注",
    "单位", "数量", "单价", "金额", "合计", "小计", "部门", "人员", "姓名",
}
_TITLE_FILLS = ("1F4E79", "2E5C8A", "16365C")


# ---------------------------------------------------------------------------
def extract_text_fingerprint(path: str) -> str:
    """所有单元格值的 sha256（值与公式字符串），用于完整性校验。"""
    h = hashlib.sha256()
    wb = load_workbook(path, data_only=False)
    try:
        for ws in wb.worksheets:
            h.update(ws.title.encode("utf-8"))
            h.update(b"\x1f")
            for row in ws.iter_rows():
                parts = []
                for cell in row:
                    v = cell.value
                    parts.append("" if v is None else str(v))
                h.update("\x1e".join(parts).encode("utf-8"))
                h.update(b"\x1d")
    finally:
        wb.close()
    return h.hexdigest()


# ---------------------------------------------------------------------------
def _hex(color: Optional[str], default: str = "000000") -> str:
    if not color:
        return default
    return str(color).replace("#", "").upper().lstrip("0") or "000000"


def _norm_color(color: Optional[str]) -> str:
    if color is None:
        return ""
    return str(color).replace("#", "").upper()


def _is_numeric(ws, row: int, col: int) -> bool:
    cell = ws.cell(row=row, column=col)
    v = cell.value
    if v is None:
        return False
    if isinstance(v, (int, float)) and not isinstance(v, bool):
        return True
    fmt = cell.number_format or "General"
    return fmt != "General" and any(ch in fmt for ch in _NUMERIC_FORMATS)


def _is_date(ws, row: int, col: int) -> bool:
    fmt = (ws.cell(row=row, column=col).number_format or "").lower()
    return any(tok in fmt for tok in ("yy", "dd", "mm/", "m/", "月", "日"))


class ExcelOptimizer:
    def __init__(self, path: str, spec: LayoutSpec, opts: ToggleOptions):
        self.path = path
        self.spec = spec
        self.opts = opts
        self.stats: Dict[str, int] = {
            "工作表": 0, "单元格": 0, "字体": 0, "表头": 0, "边框": 0,
            "列宽": 0, "行高": 0, "页面设置": 0, "冻结/筛选": 0, "背景色": 0,
            "标题行": 0, "写字优化": 0, "图片": 0,
        }
        self.warnings: List[str] = []
        self.analysis: Dict[str, int] = {}
        try:
            self.wb = load_workbook(path, rich_text=True)
        except Exception:
            self.wb = load_workbook(path)

    # ------------------------------------------------------------------
    def run(self) -> Tuple[Dict[str, int], Dict[str, int], List[str]]:
        total_cells = 0
        # ① 内容优化：只对文本单元格做"减法"，公式/数字一律不碰
        if self.opts.content:
            self._optimize_content()
        for ws in self.wb.worksheets:
            if ws.sheet_state != "visible":
                continue
            # ② 版式优化（受"自动排版"总开关控制）
            if self.opts.layout:
                try:
                    cells = self._format_sheet(ws)
                    total_cells += cells
                    if cells:
                        self.stats["工作表"] += 1
                    else:
                        self.warnings.append(f"工作表「{ws.title}」为空，已跳过。")
                except Exception as exc:  # 单表失败不影响其他表
                    log.exception("工作表 %s 排版失败", ws.title)
                    self.warnings.append(f"工作表「{ws.title}」排版失败：{exc}")
            # ③ 图片尺寸（独立开关）
            if self.opts.image_size:
                try:
                    img_stats = IMG.resize_excel_images(ws, self.spec, True)
                    self.stats["图片"] += img_stats.resized
                    if img_stats.resized:
                        log.info("工作表「%s」图片：%s", ws.title, img_stats.summary())
                except Exception as exc:
                    log.warning("工作表 %s 图片处理失败：%s", ws.title, exc)
        self.analysis = {
            "工作表数": len([w for w in self.wb.worksheets if w.sheet_state == "visible"]),
            "单元格数": total_cells if self.opts.layout else 0,
        }
        if self.opts.layout and self.stats["工作表"] == 0:
            self.warnings.append("未找到可见工作表，文件可能已损坏或被保护。")
        try:
            self.wb.save(self.path)
        finally:
            self.wb.close()
        return self.stats, self.analysis, self.warnings

    # ------------------------------------------------------------------
    def _optimize_content(self) -> None:
        """精简文本单元格：只删不改，关键信息（数字/公式）完全不动。"""
        profile = getattr(self.spec, "content_profile", "general") or "general"
        stats = C.ContentStats()
        changed = 0
        for ws in self.wb.worksheets:
            if ws.sheet_state != "visible":
                continue
            limit_row = min(ws.max_row or 1, _MAX_SCAN_ROWS)
            limit_col = min(ws.max_column or 1, _MAX_SCAN_COLS)
            for row in ws.iter_rows(min_row=1, max_row=limit_row,
                                    min_col=1, max_col=limit_col):
                for cell in row:
                    v = cell.value
                    if not isinstance(v, str) or not v.strip():
                        continue
                    if v.startswith("="):      # 公式，绝不改写
                        continue
                    stats.paragraphs += 1
                    new = C.optimize_text(v, profile, stats)
                    if new == v:
                        continue
                    if not C.is_char_subset(C.word_char_multiset(v),
                                            C.word_char_multiset(new)):
                        continue
                    cell.value = new
                    changed += 1
        self.stats["写字优化"] = changed
        if changed:
            log.info("内容优化：%s | %s", stats.summary(), C.profile_label(profile))

    # ------------------------------------------------------------------
    def _format_sheet(self, ws) -> int:
        bounds = self._find_region(ws)
        if bounds is None:
            log.debug("工作表 %s 是空白表，跳过", ws.title)
            return 0
        min_row, max_row, min_col, max_col = bounds
        title_row = self._detect_title_row(ws, min_row, max_row,
                                          min_col, max_col, False)
        header_row = min_row + 1 if title_row is not None else min_row
        has_header = self._detect_header(ws, header_row, max_row,
                                         min_col, max_col)
        self.stats["表头"] += 1 if has_header else 0
        if title_row is not None:
            self.stats["标题行"] += 1

        cell_count = 0
        for r in range(min_row, max_row + 1):
            is_title = title_row is not None and r == title_row
            is_header = (not is_title) and has_header and r == header_row
            zebra = (not is_header) and (not is_title) and self.spec.excel_zebra and \
                has_header and ((r - header_row) % 2 == 0)
            for c in range(min_col, max_col + 1):
                cell = ws.cell(row=r, column=c)
                if cell.value is None and r > min_row:
                    # 空单元格只补边框（保持表格完整）
                    if self.spec.excel_outline:
                        self._apply_border(cell, min_row, max_row, min_col, max_col, r, c)
                    continue
                if is_title:
                    self._format_title(cell, ws, r, c, min_row, max_row,
                                       min_col, max_col)
                else:
                    self._format_cell(cell, ws, r, c, is_header, zebra,
                                      min_row, max_row, min_col, max_col)
                cell_count += 1
                self.stats["单元格"] += 1
                if self.spec.excel_outline:
                    self._apply_border(cell, min_row, max_row, min_col,
                                       max_col, r, c)

        if cell_count == 0:
            return 0
        self._fit_columns(ws, min_row, max_row, min_col, max_col)
        self._setup_print(ws, min_row, max_row, min_col, max_col, has_header,
                          self.opts)
        self._freeze_and_filter(ws, min_row, max_row, min_col, max_col, has_header)
        return cell_count

    # ------------------------------------------------------------------
    def _find_region(self, ws):
        min_row = min_col = None
        max_row = max_col = 0
        limit_row = min(ws.max_row or 1, _MAX_SCAN_ROWS)
        limit_col = min(ws.max_column or 1, _MAX_SCAN_COLS)
        for row in ws.iter_rows(min_row=1, max_row=limit_row,
                                min_col=1, max_col=limit_col):
            for cell in row:
                if cell.value is None:
                    continue
                if isinstance(cell.value, str) and not cell.value.strip():
                    continue
                r, c = cell.row, cell.column
                min_row = r if min_row is None or r < min_row else min_row
                min_col = c if min_col is None or c < min_col else min_col
                max_row = max(max_row, r)
                max_col = max(max_col, c)
        if min_row is None:
            return None
        return min_row, max_row, min_col, max_col

    def _detect_title_row(self, ws, min_row: int, max_row: int,
                          min_col: int, max_col: int, has_header: bool):
        """识别表格上方单独占一行的表标题（如"××项目综合评分明细表"）。"""
        if max_row <= min_row:
            return None
        width = max_col - min_col + 1
        cells = [ws.cell(row=min_row, column=c) for c in range(min_col, max_col + 1)]
        values = [c.value for c in cells]
        filled = [v for v in values if v not in (None, "") and not (
            isinstance(v, str) and not v.strip())]
        if not filled or not all(isinstance(v, str) for v in filled):
            return None
        # 标题行特征：整行只有一段文字（其余为空），或被合并成一个大单元格
        merged_whole = any(
            (m.min_row <= min_row <= m.max_row and m.min_col <= min_col
             and m.max_col >= max_col)
            for m in ws.merged_cells.ranges
        )
        single_text = len(filled) == 1
        if not single_text and not merged_whole:
            return None
        text = filled[0].strip()
        if len(text) < 2:
            return None
        # 表头行里只有第一个单元格有字（如首列叫"序号"）——不像标题
        if single_text and not merged_whole and width >= 3:
            if text in _HEADER_WORDS:
                return None
        if single_text and not merged_whole:
            # 只有一格有字：宽表首行只写一格文字，视为标题；窄表可能是普通表头
            next_filled = sum(
                1 for c in range(min_col, max_col + 1)
                if ws.cell(row=min_row + 1, column=c).value not in (None, "")
            )
            if next_filled < 2 and width < 3:
                return None
        if merged_whole and not single_text:
            # 合并整行但里面是"序号/名称/…"多段文字 -> 是表头
            if has_header and next_filled >= max(2, int(width * 0.5)):
                return None
        return min_row

    def _format_title(self, cell, ws, r, c, min_row, max_row,
                      min_col, max_col) -> None:
        """标题行：整行居中、加粗、字号略大、深蓝底白字，不画表格边框。"""
        spec, opts = self.spec, self.opts
        base = spec.excel_header_font
        kw = {}
        if opts.font_family:
            kw["name"] = base.east_asia
        if opts.font_size:
            kw["sz"] = round(max(base.size_pt + 2.0, 12.0), 1)
        kw["b"] = True
        if opts.font_color:
            kw["color"] = "FFFFFF"
        try:
            cell.font = Font(**{k: v for k, v in kw.items() if v is not None})
            self.stats["字体"] += 1
        except Exception:
            pass
        if opts.background:
            cell.fill = PatternFill(fill_type="solid", start_color=_TITLE_FILLS[0],
                                    end_color=_TITLE_FILLS[0])
            self.stats["背景色"] += 1
        cell.alignment = Alignment(horizontal="center", vertical="center",
                                   wrap_text=False)
        try:
            cell.border = Border()
        except Exception:
            pass
        # 标题文字常常只写在首列（无合并），同行的其它列顺带铺上同样的底色，
        # 让标题在页面上是一条完整的色带
        if not opts.background:
            return
        span = [cc for cc in range(min_col, max_col + 1)]
        for cc in span:
            other = ws.cell(row=r, column=cc)
            if other is cell or (other.value not in (None, "")):
                continue
            other.fill = PatternFill(fill_type="solid", start_color=_TITLE_FILLS[0],
                                     end_color=_TITLE_FILLS[0])

    def _detect_header(self, ws, min_row: int, max_row: int,
                       min_col: int, max_col: int) -> bool:
        if max_row == min_row:
            return False
        row1_filled = sum(
            1 for c in range(min_col, max_col + 1)
            if ws.cell(row=min_row, column=c).value not in (None, "")
        )
        width = max_col - min_col + 1
        if row1_filled < max(2, int(width * 0.5)):
            return False
        # 表头行通常全是文本、且与下一行类型不同
        texts1 = sum(
            1 for c in range(min_col, max_col + 1)
            if isinstance(ws.cell(row=min_row, column=c).value, str)
        )
        if texts1 >= row1_filled * 0.8:
            return True
        nums2 = sum(
            1 for c in range(min_col, max_col + 1)
            if _is_numeric(ws, min_row + 1, c)
        )
        return nums2 >= max(1, width // 2)

    # ------------------------------------------------------------------
    def _format_cell(self, cell, ws, r, c, is_header, zebra,
                     min_row, max_row, min_col, max_col) -> None:
        spec, opts = self.spec, self.opts
        base = spec.excel_header_font if is_header else spec.excel_font

        old = cell.font
        changed = False
        kwargs = {}
        if old is not None:
            kwargs = dict(
                name=old.name, sz=old.sz, b=old.b, i=old.i, u=old.u,
                strike=old.strike,
                color=(old.color.rgb if old.color is not None
                       and isinstance(old.color.rgb, str) else None),
            )
        if opts.font_family:
            kwargs["name"] = base.east_asia
        if opts.font_size:
            kwargs["sz"] = base.size_pt
        if is_header:
            if base.bold:
                kwargs["b"] = True
        elif kwargs.get("b") is None and base.bold:
            kwargs["b"] = True
        if opts.font_color:
            if is_header and spec.excel_header_fg:
                kwargs["color"] = _norm_color(spec.excel_header_fg)
            elif (not is_header) and base.color:
                kwargs["color"] = _norm_color(base.color)
        kwargs = {k: v for k, v in kwargs.items() if v is not None}

        try:
            cell.font = Font(**kwargs)
            changed = True
        except Exception:
            cell.font = Font(name=base.east_asia, sz=base.size_pt,
                             bold=bool(is_header and base.bold))
            changed = True
        if changed:
            self.stats["字体"] += 1

        # 填充
        if opts.background:
            fill_hex = None
            if is_header and spec.excel_header_bg:
                fill_hex = spec.excel_header_bg
            elif zebra and spec.excel_zebra_bg:
                fill_hex = spec.excel_zebra_bg
            if fill_hex:
                cell.fill = PatternFill(
                    fill_type="solid", start_color=_norm_color(fill_hex),
                    end_color=_norm_color(fill_hex),
                )
                self.stats["背景色"] += 1
        elif is_header and spec.excel_header_font.color:
            # 不改背景时，表头文字用深色以保证可读
            pass

        # 对齐
        align_h = spec.excel_align_h
        if align_h == "auto":
            if is_header:
                horizontal = "center"
            elif _is_numeric(ws, r, c) or _is_date(ws, r, c):
                horizontal = "right"
            else:
                horizontal = "left"
        else:
            horizontal = align_h
        cell.alignment = Alignment(
            horizontal=horizontal,
            vertical="center",
            wrap_text=spec.excel_wrap_text,
            shrink_to_fit=False,
            indent=cell.alignment.indent if cell.alignment is not None else 0,
        )

        # 行高：1.35 倍基础行高，保证中文不挤、不裁切
        if spec.excel_row_height_auto:
            rd = ws.row_dimensions[r]
            if rd.height is None:
                rd.height = min(max(spec.excel_font.size_pt * 2.0, 18.0), 60.0)
                self.stats["行高"] += 1

    def _apply_border(self, cell, min_row, max_row, min_col, max_col, r, c) -> None:
        spec = self.spec
        thin = Side(style=spec.excel_border_style, color=_norm_color(spec.excel_border_color))
        medium = Side(style=spec.excel_border_outer_style,
                      color=_norm_color(spec.excel_border_color))
        cell.border = Border(
            left=medium if c == min_col else thin,
            right=medium if c == max_col else thin,
            top=medium if r == min_row else thin,
            bottom=medium if r == max_row else thin,
        )
        self.stats["边框"] += 1

    # ------------------------------------------------------------------
    def _fit_columns(self, ws, min_row, max_row, min_col, max_col) -> None:
        if not self.spec.excel_column_fit:
            return

        def disp_len(s: str) -> int:
            return sum(2 if ord(ch) > 127 else 1 for ch in s)

        merged_single_col = {
            rng.min_col for rng in ws.merged_cells.ranges
            if rng.max_col == rng.min_col and rng.max_row > rng.min_row
        }
        widths: Dict[int, float] = {}
        best_floor = 9.0  # 再窄就放不下主要表头文字
        for c in range(min_col, max_col + 1):
            best = 6.0
            total_len = 0.0
            count = 0
            for r in range(min_row, min(max_row, min_row + 400) + 1):
                cell = ws.cell(row=r, column=c)
                if cell.value is None:
                    continue
                text = str(cell.value)
                if cell.data_type == "f":
                    text = "000000"
                if text.startswith("="):
                    text = "000000"
                if len(text) > 60 and self.spec.excel_wrap_text:
                    # 长文本靠换行，列宽不无限拉长
                    tokens = re.split(r"[\s、，,；;。！？（）()：:]+", text)
                    longest_token = max((disp_len(t) for t in tokens), default=8)
                    length = min(max(longest_token, 12), 30)
                else:
                    length = disp_len(text)
                best = max(best, length + 3)
                total_len += disp_len(text)
                count += 1
            # 允许换行时，列宽取"够放下这一列大多数内容"的宽度，个别超长单元格
            # 换行显示；纵向合并单元格（如一列型号合并）不参与收窄，避免折行
            if self.spec.excel_wrap_text and c not in merged_single_col:
                need = sorted(
                    disp_len(str(ws.cell(row=r, column=c).value)) + 3
                    for r in range(min_row, min(max_row, min_row + 400) + 1)
                    if ws.cell(row=r, column=c).value is not None
                )
                if need:
                    k = max(0, int(len(need) * 0.8) - 1)
                    best = max(best_floor, min(best, need[k] + 2))
                    if count:
                        best = max(best, min(total_len / count + 4, 28))
            if self.spec.excel_column_fit >= 2:
                best = min(best, self.spec.excel_max_col_width)
            widths[c] = best
            if self.spec.excel_column_fit >= 2:
                best = min(best, self.spec.excel_max_col_width)
            widths[c] = best

        # 合并单元格：按合并区总宽分摊
        for rng in ws.merged_cells.ranges:
            try:
                span = rng.max_col - rng.min_col + 1
            except Exception:
                continue
            if span < 2:
                continue
            cell = ws.cell(row=rng.min_row, column=rng.min_col)
            if cell.value is None:
                continue
            if rng.min_row == min_row and self.spec.excel_wrap_text:
                continue  # 标题合并行交给换行
            need = disp_len(str(cell.value)) + 4
            covered = [c for c in range(rng.min_col, rng.max_col + 1) if c in widths]
            if not covered:
                continue
            current = sum(widths[c] for c in covered)
            if current < need:
                extra = (need - current) / len(covered)
                for c in covered:
                    widths[c] += extra

        for c, w in widths.items():
            # 设置 width 时 openpyxl 会自动把 customWidth 置位，直接赋值 customWidth 会报错
            ws.column_dimensions[get_column_letter(c)].width = round(w, 2)
            self.stats["列宽"] += 1

        # 标题行（整行合并）按当前总宽居中显示
        for rng in ws.merged_cells.ranges:
            if rng.min_row != min_row or rng.max_row != min_row:
                continue
            if rng.min_col > min_col or rng.max_col < max_col:
                continue
            cell = ws.cell(row=rng.min_row, column=rng.min_col)
            if cell.value is None:
                continue
            if not isinstance(cell.value, str):
                continue
            need = disp_len(cell.value.strip()) + 2
            current = sum(widths.get(c, 8.0)
                          for c in range(min_col, max_col + 1))
            if current < need:
                extra = (need - current) / max(1, max_col - min_col + 1)
                for c in range(min_col, max_col + 1):
                    if c in widths:
                        ws.column_dimensions[get_column_letter(c)].width = round(
                            widths[c] + extra, 2)
            break

    # ------------------------------------------------------------------
    def _setup_print(self, ws, min_row, max_row, min_col, max_col,
                     has_header, opts=None) -> None:
        spec = self.spec
        ws.page_margins = PageMargins(
            left=spec.excel_margins_cm["left"] / 2.54,
            right=spec.excel_margins_cm["right"] / 2.54,
            top=spec.excel_margins_cm["top"] / 2.54,
            bottom=spec.excel_margins_cm["bottom"] / 2.54,
            header=spec.excel_margins_cm.get("header", 0.76) / 2.54,
            footer=spec.excel_margins_cm.get("footer", 0.76) / 2.54,
        )
        ws.page_setup.paperSize = spec.excel_paper_size or 9  # 9 = A4

        orientation = spec.excel_orientation
        if orientation == "auto" or orientation is None:
            width = max_col - min_col + 1
            orientation = "landscape" if width > 8 else "portrait"
        ws.page_setup.orientation = orientation

        allow_paging = True if opts is None else bool(opts.page_break)
        if spec.excel_fit_to_width and allow_paging:
            ws.sheet_properties.pageSetUpPr.fitToPage = True
            ws.page_setup.fitToWidth = max(1, spec.excel_fit_width_pages)
            ws.page_setup.fitToHeight = max(0, spec.excel_fit_height_pages)
            # 兼容部分版本的属性名
            try:
                ws.print_options.horizontalCentered = True
                ws.print_options.verticalCentered = False
            except Exception:
                pass

        if has_header and spec.excel_header_repeat and allow_paging:
            try:
                ws.print_title_rows = f"{min_row}:{min_row}"
            except Exception as exc:
                log.debug("打印标题设置跳过: %s", exc)

        if spec.excel_footer_page_number:
            try:
                ws.oddFooter.center.text = "第 &P 页 / 共 &N 页"
                ws.oddFooter.center.size = 9
                ws.oddFooter.center.font = spec.excel_font.east_asia
            except Exception as exc:
                log.debug("页脚设置跳过: %s", exc)
        self.stats["页面设置"] += 1

    def _freeze_and_filter(self, ws, min_row, max_row, min_col, max_col,
                           has_header) -> None:
        spec = self.spec
        if has_header:
            if spec.excel_freeze_header:
                try:
                    ws.freeze_panes = ws.cell(row=min_row + 1, column=min_col)
                    self.stats["冻结/筛选"] += 1
                except Exception:
                    pass
            if spec.excel_autofilter:
                ref = f"{get_column_letter(min_col)}{min_row}:{get_column_letter(max_col)}{max_row}"
                try:
                    if ws.auto_filter.ref is None:
                        ws.auto_filter.ref = ref
                        self.stats["冻结/筛选"] += 1
                except Exception:
                    pass

    # ------------------------------------------------------------------
    def _set_row_heights(self) -> None:
        pass


def dump_text(path: str) -> str:
    """按工作表/行输出全部单元格文本，用于内容校验。"""
    wb = load_workbook(path, data_only=False)
    parts: List[str] = []
    try:
        for ws in wb.worksheets:
            parts.append(ws.title)
            for row in ws.iter_rows():
                for cell in row:
                    v = cell.value
                    if v is None:
                        continue
                    parts.append("" if isinstance(v, str) and not v.strip() else str(v))
    finally:
        wb.close()
    return "\n".join(parts)


def optimize_excel(path: str, spec: LayoutSpec, options) -> Tuple[Dict[str, int], Dict[str, int], List[str]]:
    optimizer = ExcelOptimizer(path, spec, options.toggles)
    return optimizer.run()
