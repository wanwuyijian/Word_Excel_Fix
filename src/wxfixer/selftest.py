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

    from .version import APP_NAME, APP_VERSION

    out("")
    out(f"{APP_NAME} v{APP_VERSION} 自检结果")
    out(f"  通过: {len(PASS)}")
    out(f"  失败: {len(FAIL)}")
    for f in FAIL:
        out("   × " + f)
    return 0 if not FAIL else 1
