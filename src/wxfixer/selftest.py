"""离线自检：不开界面，直接跑核心引擎，验证"只排版不改内容"。

用法:
    python src/main.py --selftest              # 使用内置样本（若存在）
    python src/main.py --selftest --samples D:\\samples --out D:\\out
退出码 0 表示全部通过；打包后的 exe 同样支持该参数。
"""

from __future__ import annotations

import glob
import os
import sys
import tempfile
import traceback
import zipfile

from .core import optimize_file
from .core.runner import OptimizeOptions, UnsupportedFileError, detect_file_type
from .templates import ALL_TEMPLATES, get_template
from .templates.schema import ToggleOptions
from .utils import paths

PASS: list[str] = []
FAIL: list[str] = []


def check(cond: bool, msg: str) -> None:
    (PASS if cond else FAIL).append(msg)


def _fingerprint(path: str, kind: str) -> str:
    if kind == "word":
        from .core.word_optimizer import extract_text_fingerprint

        return extract_text_fingerprint(path)
    from .core.excel_optimizer import extract_text_fingerprint

    return extract_text_fingerprint(path)


# ---------------------------------------------------------------------------
def build_builtin_samples(folder: str) -> list[str]:
    """用 python-docx / openpyxl 现场生成测试文档，exe 无外部样本时也能自检。"""
    from docx import Document
    from docx.enum.text import WD_ALIGN_PARAGRAPH
    from docx.shared import Pt
    from openpyxl import Workbook

    os.makedirs(folder, exist_ok=True)
    made: list[str] = []

    doc = Document()
    p = doc.add_paragraph("自检样本：技术方案书")
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p.runs[0].font.size = Pt(9)
    doc.add_paragraph("1 概述")
    doc.add_paragraph("本文用于自检排版引擎是否正常工作，正文内容必须保持不变。")
    doc.add_paragraph("1.1 目标")
    doc.add_paragraph("验证标题识别、字体字号、行距缩进、表格样式与页码。")
    doc.add_paragraph("一、注意事项")
    doc.add_paragraph("仅调整排版属性，不增删任何文字。")
    table = doc.add_table(rows=3, cols=3)
    for i, row in enumerate((("序号", "项目", "数值"), ("1", "示例", "100"),
                             ("2", "示例", "200"))):
        for j, val in enumerate(row):
            table.cell(i, j).text = val
    path = os.path.join(folder, "自检_技术方案.docx")
    doc.save(path)
    made.append(path)

    wb = Workbook()
    ws = wb.active
    ws.title = "自检表"
    ws.append(["序号", "名称", "数量", "单价", "金额"])
    for i in range(1, 6):
        ws.append([i, f"设备{i}", i * 2, 1000 + i * 50, f"=C{i + 1}*D{i + 1}"])
    path = os.path.join(folder, "自检_清单.xlsx")
    wb.save(path)
    made.append(path)
    return made


def run_documents(files: list[str], out: str) -> None:
    check(bool(files), f"找到测试样本 {len(files)} 个")
    for src in files:
        kind = detect_file_type(src)
        before = _fingerprint(src, kind)
        for tpl in ALL_TEMPLATES:
            try:
                res = optimize_file(src, tpl, OptimizeOptions(template_key=tpl.key), out)
            except Exception:
                FAIL.append(f"[异常] {os.path.basename(src)} / {tpl.name}")
                traceback.print_exc()
                continue
            after = _fingerprint(res.target, kind)
            tag = f"{os.path.basename(src)} + {tpl.name}"
            check(before == after, f"内容零改动: {tag}")
            check(os.path.exists(res.target), f"产出文件存在: {tag}")
            if kind == "word":
                from .core.ooxml_check import check_document

                issues = check_document(res.target)
                check(not issues, f"OOXML 结构自检通过: {tag}" +
                      (f" -> {issues[:3]}" if issues else ""))
            check(res.elapsed < 180, f"耗时合理: {tag} ({res.elapsed:.2f}s)")


def _png_bytes(w: int, h: int, rgb=(60, 110, 190)) -> bytes:
    """不依赖 Pillow 的最小 PNG 生成（冻结后的 exe 里也能用）。"""
    import struct
    import zlib

    raw = b""
    row = bytes(rgb) * w
    for _ in range(h):
        raw += b"\x00" + row

    def chunk(tag: bytes, data: bytes) -> bytes:
        return (struct.pack(">I", len(data)) + tag + data
                + struct.pack(">I", zlib.crc32(tag + data) & 0xFFFFFFFF))

    ihdr = struct.pack(">IIBBBBB", w, h, 8, 2, 0, 0, 0)
    return (b"\x89PNG\r\n\x1a\n" + chunk(b"IHDR", ihdr)
            + chunk(b"IDAT", zlib.compress(raw, 6)) + chunk(b"IEND", b""))


def build_image_sample(folder: str) -> str:
    """造一个含"超宽图 + 小图标"的 docx，用于验证图片尺寸开关。"""
    from docx import Document
    from docx.shared import Cm

    os.makedirs(folder, exist_ok=True)
    png = os.path.join(folder, "_selftest_img.png")
    with open(png, "wb") as fh:
        fh.write(_png_bytes(1200, 800))
    doc = Document()
    doc.add_paragraph("自检样本：图片尺寸")
    doc.add_paragraph("下面是一张远超版心宽度的图片。")
    doc.add_paragraph().add_run().add_picture(png, width=Cm(22))
    doc.add_paragraph("下面是一个小图标，不应被放大。")
    doc.add_paragraph().add_run().add_picture(png, width=Cm(1.2))
    path = os.path.join(folder, "自检_图片.docx")
    doc.save(path)
    return path


def _format_fingerprint(path: str) -> tuple:
    """Word 排版指纹：段落对齐/缩进/行距 + 每个 run 的字体与字号。"""
    from docx import Document

    doc = Document(path)
    rows = []
    for p in doc.paragraphs:
        pf = p.paragraph_format
        rows.append(("p", str(pf.alignment), str(pf.line_spacing),
                     str(pf.first_line_indent),
                     tuple((r.font.name, str(r.font.size), str(r.bold))
                           for r in p.runs)))
    return tuple(rows)


def _image_widths_cm(path: str) -> list:
    from docx import Document

    return [round(s.width.cm, 2) for s in Document(path).inline_shapes]


def run_toggles(files: list[str], out: str) -> None:
    words = [f for f in files if f.lower().endswith((".docx", ".docm"))]
    if not words:
        return
    src = words[0]
    before = _fingerprint(src, "word")
    combos = {
        "全关闭": ToggleOptions(font_family=False, font_size=False, font_color=False,
                             background=False, page_break=False),
        "仅字体": ToggleOptions(font_family=True, font_size=False, font_color=False,
                             background=False, page_break=False),
    }
    for name, toggles in combos.items():
        try:
            res = optimize_file(src, get_template("auto"),
                                OptimizeOptions(toggles=toggles), out)
            check(before == _fingerprint(res.target, "word"),
                  f"开关[{name}]内容零改动")
            from .core.ooxml_check import check_document

            check(not check_document(res.target), f"开关[{name}]结构自检通过")
            run_toggle_expectation(res.target, name, toggles)
        except Exception:
            FAIL.append(f"[异常] 开关 {name}")
            traceback.print_exc()


def run_toggle_expectation(path: str, name: str, toggles: ToggleOptions) -> None:
    """核对开关的"实际效果"：关掉分页时全文不应再有段前分页。"""
    from docx import Document

    doc = Document(path)
    pbb = [p for p in doc.paragraphs if p.paragraph_format.page_break_before]
    style_pbb = []
    for sname in ("Heading 1", "Heading 2", "Heading 3", "Title"):
        try:
            style = doc.styles[sname]
        except KeyError:
            continue
        if style.paragraph_format.page_break_before:
            style_pbb.append(sname)
    if not toggles.page_break:
        check(not pbb, f"开关[{name}]已关闭段前分页（残留 {len(pbb)} 处）")
        check(not style_pbb, f"开关[{name}]标题样式也无分页（残留 {style_pbb}）")
    else:
        check(True, f"开关[{name}]段前分页 {len(pbb)} 处")


def run_new_toggles(out: str) -> None:
    """v1.1 新增三个开关的"确有实效"验证。"""
    from docx import Document

    folder = os.path.join(out, "_toggles")
    os.makedirs(folder, exist_ok=True)

    # ---- A. 允许修改文件内容 -------------------------------------------
    src = os.path.join(folder, "内容样本.docx")
    doc = Document()
    doc.add_paragraph("项目 实施 通知")
    doc.add_paragraph("  众所周知， 该 设备 性能 很好。  ")
    doc.add_paragraph("基本上,运行稳定。")
    doc.add_paragraph("本项目由甲方负责验收，乙方负责交付。")
    doc.save(src)
    raw = [p.text for p in Document(src).paragraphs]

    try:
        res = optimize_file(src, get_template("auto"),
                            OptimizeOptions(toggles=ToggleOptions(content=True)), out)
        new = [p.text for p in Document(res.target).paragraphs]
        joined = "".join(new)
        check(len(joined) < len("".join(raw)), "内容开关：文字确有精简")
        check("性能" in joined and "验收" in joined and "交付" in joined,
              "内容开关：关键信息保留")
        check("众所周知" not in joined, "内容开关：赘语已删除")
        check(new[0].replace(" ", "") == raw[0].replace(" ", ""),
              "内容开关：标题文字未改动（仅清理空格）")
    except Exception:
        FAIL.append("[异常] 内容开关")
        traceback.print_exc()

    try:
        res = optimize_file(src, get_template("auto"),
                            OptimizeOptions(toggles=ToggleOptions(content=False)), out)
        same = [p.text for p in Document(res.target).paragraphs]
        check(same == raw, "内容开关关闭：文字完全未变")
    except Exception:
        FAIL.append("[异常] 内容开关关闭")
        traceback.print_exc()

    # ---- B. 允许修改图片尺寸 -------------------------------------------
    img_src = build_image_sample(folder)
    try:
        res = optimize_file(img_src, get_template("auto"),
                            OptimizeOptions(toggles=ToggleOptions(image_size=True)), out)
        widths = _image_widths_cm(res.target)
        check(len(widths) == 2, f"图片开关：识别到 {len(widths)} 张图片")
        if widths:
            check(widths[0] < 21.9, f"图片开关：超宽图已缩至 {widths[0]}cm")
            check(widths[0] > 5, f"图片开关：缩放幅度合理（{widths[0]}cm）")
        if len(widths) > 1:
            check(abs(widths[1] - 1.2) < 0.05,
                  f"图片开关：小图标未放大（{widths[1]}cm）")
    except Exception:
        FAIL.append("[异常] 图片开关")
        traceback.print_exc()

    try:
        res = optimize_file(img_src, get_template("auto"),
                            OptimizeOptions(toggles=ToggleOptions(image_size=False)), out)
        widths = _image_widths_cm(res.target)
        check(widths and widths[0] > 21.9, f"图片开关关闭：原尺寸保留（{widths[0]}cm）")
    except Exception:
        FAIL.append("[异常] 图片开关关闭")
        traceback.print_exc()

    # ---- C. 自动排版总开关 ---------------------------------------------
    body_src = os.path.join(folder, "排版样本.docx")
    doc = Document()
    doc.add_paragraph("一、背景")
    doc.add_paragraph("这是正文段落，用于核对排版开关是否生效。")
    doc.save(body_src)
    src_fp = _format_fingerprint(body_src)
    try:
        res = optimize_file(body_src, get_template("auto"),
                            OptimizeOptions(toggles=ToggleOptions(layout=False,
                                                                 content=False,
                                                                 image_size=False)), out)
        check(_format_fingerprint(res.target) == src_fp, "排版开关关闭：版式零改动")
    except Exception:
        FAIL.append("[异常] 排版开关关闭")
        traceback.print_exc()
    try:
        res = optimize_file(body_src, get_template("auto"),
                            OptimizeOptions(toggles=ToggleOptions(content=False,
                                                                 image_size=False)), out)
        check(_format_fingerprint(res.target) != src_fp, "排版开关开启：版式确有调整")
    except Exception:
        FAIL.append("[异常] 排版开关开启")
        traceback.print_exc()


def run_new_toggles_excel(out: str) -> None:
    """Excel 侧的新开关：内容精简 + 图片缩放。"""
    from openpyxl import Workbook, load_workbook
    from openpyxl.drawing.image import Image as XLImage

    folder = os.path.join(out, "_toggles_xl")
    os.makedirs(folder, exist_ok=True)
    png = os.path.join(folder, "_xl.png")
    with open(png, "wb") as fh:
        fh.write(_png_bytes(1600, 1000, (30, 90, 160)))

    src = os.path.join(folder, "自检_开关.xlsx")
    wb = Workbook()
    ws = wb.active
    ws.title = "清单"
    ws.append(["序号", "名称", "备注"])
    ws.append([1, "服务器", "  众所周知， 该 设备 性能 很好。  "])
    ws.append([2, "交换机", "基本上,运行稳定。"])
    img = XLImage(png)
    img.width, img.height = 1600, 1000
    ws.add_image(img, "E3")
    wb.save(src)

    before = _fingerprint(src, "excel")
    try:
        res = optimize_file(src, get_template("auto"),
                            OptimizeOptions(toggles=ToggleOptions(content=False,
                                                                  image_size=False)), out)
        check(_fingerprint(res.target, "excel") == before, "Excel 内容关闭：文字未变")
    except Exception:
        FAIL.append("[异常] Excel 内容关闭")
        traceback.print_exc()

    try:
        res = optimize_file(src, get_template("auto"),
                            OptimizeOptions(toggles=ToggleOptions(content=True)), out)
        wb2 = load_workbook(res.target)
        ws2 = wb2["清单"]
        check(ws2["B2"].value == "服务器", "Excel 内容开关：非文本列未变")
        c2 = str(ws2["C2"].value or "")
        check("众所周知" not in c2 and "性能" in c2,
              f"Excel 内容开关：赘语删除且保留要点（{c2!r}）")
        check(len(c2) < len("  众所周知， 该 设备 性能 很好。  "),
              "Excel 内容开关：文字确有精简")
        imgs = list(getattr(ws2, "_images", []))
        check(len(imgs) == 1, f"Excel 图片：仍存在 {len(imgs)} 张")
        if imgs:
            ext = getattr(imgs[0].anchor, "ext", None)
            cx = int(getattr(ext, "cx", 0) or 0)
            check(0 < cx < int(1600 * 9525), "Excel 图片开关：图片已缩小")
    except Exception:
        FAIL.append("[异常] Excel 内容开关")
        traceback.print_exc()


def run_negative() -> None:
    tmp = tempfile.mkdtemp()
    bad = os.path.join(tmp, "fake.docx")
    with open(bad, "wb") as fh:
        fh.write(b"this is not a zip")
    for path, label in ((bad, "非法文件"), (os.path.join(tmp, "x.txt"), "不支持的扩展名")):
        try:
            detect_file_type(path)
            FAIL.append(f"{label}应被拒绝")
        except UnsupportedFileError:
            PASS.append(f"{label}被正确拒绝")

    # 空 zip（伪装成 docx）
    empty = os.path.join(tmp, "empty.xlsx")
    with zipfile.ZipFile(empty, "w") as zf:
        zf.writestr("dummy.txt", "x")
    try:
        detect_file_type(empty)
        FAIL.append("空包应被拒绝")
    except UnsupportedFileError:
        PASS.append("空包被正确拒绝")


# ---------------------------------------------------------------------------
def main(argv: list[str] | None = None, writer=None) -> int:
    argv = list(sys.argv[1:] if argv is None else argv)

    def out(text: str = "") -> None:
        if writer is not None:
            writer(text)
        else:
            print(text)

    samples = None
    out_dir = None
    for i, arg in enumerate(argv):
        if arg == "--samples" and i + 1 < len(argv):
            samples = argv[i + 1]
        if arg == "--out" and i + 1 < len(argv):
            out_dir = argv[i + 1]

    PASS.clear()
    FAIL.clear()

    if samples is None:
        candidates = [
            os.path.join(paths.app_dir(), "samples"),
            os.path.join(paths.app_dir(), "..", "samples"),
        ]
        for cand in candidates:
            if os.path.isdir(cand):
                samples = cand
                break

    if out_dir is None:
        out_dir = os.path.join(tempfile.gettempdir(), "wxfixer_selftest")

    if samples and os.path.isdir(samples):
        files = sorted(
            f for f in glob.glob(os.path.join(samples, "*"))
            if os.path.splitext(f)[1].lower() in (".docx", ".xlsx")
        )
        if not files:
            files = build_builtin_samples(samples)
    else:
        samples = os.path.join(out_dir, "samples")
        files = build_builtin_samples(samples)

    out("样本目录: " + samples)
    out("输出目录: " + out_dir)
    os.makedirs(out_dir, exist_ok=True)

    run_negative()
    run_documents(files, out_dir)
    run_toggles(files, out_dir)
    run_new_toggles(out_dir)
    run_new_toggles_excel(out_dir)

    from .version import APP_NAME, APP_VERSION

    out("")
    out(f"{APP_NAME} v{APP_VERSION} 自检结果")
    out(f"  通过: {len(PASS)}")
    out(f"  失败: {len(FAIL)}")
    for f in FAIL:
        out("   × " + f)
    return 0 if not FAIL else 1
