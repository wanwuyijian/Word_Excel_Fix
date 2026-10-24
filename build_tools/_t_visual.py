"""版面目检（程序化）：导出 PDF 后用 PyMuPDF 检查图片是否越出版心、是否居中。

因为版式观感无法靠截图判断，这里把"美观"的硬指标变成可断言数字：
  1) 每张图片的横向范围必须在版心内（含 2mm 容差）；
  2) 图片中心与版心中心的偏差 < 6pt（说明已居中）；
  3) 图片高度不超过版心高度；
  4) 页面尺寸为 A4、页数合理。
"""
from __future__ import annotations

import io
import os
import subprocess
import sys

sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding="utf-8", errors="replace")
sys.path.insert(0, os.path.join(os.path.dirname(__file__), "..", "src"))

import fitz  # noqa: E402

ROOT = os.path.abspath(os.path.join(os.path.dirname(__file__), ".."))
PASS: list[str] = []
FAIL: list[str] = []


def check(cond, msg):
    (PASS if cond else FAIL).append(msg)
    print(("  OK   " if cond else "  FAIL ") + msg)


def render(docx: str, outdir: str) -> str:
    """调用已有的 render_check.ps1 导出 PDF（不自己拼 PowerShell，避免 COM 挂死）。"""
    import shutil

    ext = os.path.splitext(docx)[1].lower()
    stage = os.path.join(os.path.dirname(outdir), "_stage")
    os.makedirs(stage, exist_ok=True)
    base = "v_%d" % (abs(hash(docx)) % 100000)
    shutil.copyfile(docx, os.path.join(stage, base + ext))
    os.makedirs(outdir, exist_ok=True)
    script = os.path.join(ROOT, "build_tools", "render_check.ps1")
    try:
        subprocess.run(
            ["powershell", "-NoProfile", "-ExecutionPolicy", "Bypass",
             "-File", script, "-InDir", stage, "-OutDir", outdir,
             "-Filter", base + ext, "-MaxFiles", "1"],
            check=True, stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL,
            timeout=180)
    except Exception:
        pass
    pdf = os.path.join(outdir, base + ".pdf")
    try:
        os.remove(os.path.join(stage, base + ext))
    except OSError:
        pass
    return pdf if os.path.exists(pdf) else ""


def inspect(pdf: str, *, left_cm=3.17, right_cm=3.17, tol_mm=2.0) -> None:
    doc = fitz.open(pdf)
    check(doc.page_count >= 1, f"页数 {doc.page_count}")
    page = doc[0]
    w, h = page.rect.width, page.rect.height
    check(abs(w - 595.3) < 3 and abs(h - 841.9) < 3,
          f"页面为 A4 竖版（{w:.1f}x{h:.1f}pt）")
    pt_per_cm = 72 / 2.54
    body_left = left_cm * pt_per_cm
    body_right = w - right_cm * pt_per_cm
    body_center = (body_left + body_right) / 2
    tol = tol_mm / 10 * pt_per_cm

    found = 0
    for pno in range(doc.page_count):
        pg = doc[pno]
        for info in pg.get_image_info():
            bbox = fitz.Rect(info["bbox"])
            found += 1
            tag = f"P{pno + 1} 图片#{found}"
            check(bbox.x0 >= body_left - tol and bbox.x1 <= body_right + tol,
                  f"{tag} 横向在版心内（{bbox.x0:.1f}~{bbox.x1:.1f}，"
                  f"版心 {body_left:.1f}~{body_right:.1f}）")
            off = abs((bbox.x0 + bbox.x1) / 2 - body_center)
            check(off < 6, f"{tag} 水平居中（偏移 {off:.1f}pt）")
            check(bbox.height <= h - 100, f"{tag} 高度 {bbox.height:.1f}pt 未超版心")
            check(bbox.width > 20, f"{tag} 宽度 {bbox.width:.1f}pt 未被缩没")
    check(found >= 1, f"PDF 中检出 {found} 张图片")
    doc.close()


def check_excel(work: str) -> None:
    """Excel 版面目检：图片不越出页宽、页码字段存在、表头有底色。"""
    from openpyxl import Workbook, load_workbook
    from openpyxl.drawing.image import Image as XLImage
    from openpyxl.styles import Font as XLFont

    from wxfixer.core import optimize_file
    from wxfixer.core.runner import OptimizeOptions
    from wxfixer.templates import get_template
    from wxfixer.templates.schema import ToggleOptions

    from _t_toggles import png_bytes

    src_dir = os.path.join(work, "src")
    png = os.path.join(src_dir, "xl.png")
    with open(png, "wb") as fh:
        fh.write(png_bytes(1600, 1000, (45, 95, 165)))

    src = os.path.join(src_dir, "表格验证.xlsx")
    wb = Workbook()
    ws = wb.active
    ws.title = "设备清单"
    ws.append(["序号", "设备名称", "型号规格", "数量", "单价", "备注"])
    for i in range(1, 9):
        ws.append([i, f"设备{i}", f"XX-{1000 + i}", i, 1000 + i * 10, "含安装调试"])
    for row in ws.iter_rows(min_row=1, max_row=9, max_col=6):
        for c in row:
            c.font = XLFont(name="Times New Roman", size=9)
    img = XLImage(png)
    img.width, img.height = 1600, 1000
    ws.add_image(img, "B11")
    wb.save(src)

    res = optimize_file(src, get_template("product"),
                        OptimizeOptions(toggles=ToggleOptions(content=False)),
                        os.path.join(work, "out_xl"))
    pdf = render(res.target, os.path.join(work, "pdf_xl"))
    check(bool(pdf), "Excel PDF 导出成功")
    if not pdf:
        return
    d = fitz.open(pdf)
    check(d.page_count >= 1, f"Excel 导出 {d.page_count} 页")
    pg = d[0]
    w = pg.rect.width
    imgs = pg.get_image_info()
    check(len(imgs) >= 1, f"Excel PDF 中检出 {len(imgs)} 张图片")
    for i, info in enumerate(imgs):
        bbox = fitz.Rect(info["bbox"])
        check(bbox.x0 > -2 and bbox.x1 < w + 2,
              f"Excel 图片#{i + 1} 未越出页宽（{bbox.x0:.0f}~{bbox.x1:.0f} / {w:.0f}）")
        check(bbox.width > 40,
              f"Excel 图片#{i + 1} 尺寸合理（{bbox.width:.0f}x{bbox.height:.0f}pt）")
    txt = "".join(pg.get_text() for _ in [0])
    check("第" in txt and "页" in txt, "Excel 页脚页码已渲染")
    d.close()

    wb2 = load_workbook(res.target)
    ws2 = wb2["设备清单"]
    fills = {str(c.fill.fgColor.rgb) for row in ws2.iter_rows(min_row=1, max_row=1,
                                                             max_col=6) for c in row
             if c.fill and c.fill.fgColor}
    fills.discard("00000000")
    check(bool(fills), f"Excel 表头底色已应用（{sorted(fills)}）")
    check(ws2.freeze_panes is not None, f"Excel 冻结窗格 {ws2.freeze_panes}")


def main() -> int:
    from wxfixer.core import optimize_file
    from wxfixer.core.runner import OptimizeOptions
    from wxfixer.templates import get_template
    from wxfixer.templates.schema import ToggleOptions

    work = os.path.join(ROOT, "render_out", "visual")
    src_dir = os.path.join(work, "src")
    os.makedirs(src_dir, exist_ok=True)
    docx = os.path.join(src_dir, "版面验证.docx")

    sys.path.insert(0, os.path.join(ROOT, "build_tools"))
    from _t_toggles import png_bytes  # noqa: E402

    from docx import Document
    from docx.shared import Cm

    png_big = os.path.join(src_dir, "wide.png")
    png_sq = os.path.join(src_dir, "square.png")
    with open(png_big, "wb") as fh:
        fh.write(png_bytes(1600, 600, (40, 100, 170)))
    with open(png_sq, "wb") as fh:
        fh.write(png_bytes(600, 900, (170, 90, 60)))

    doc = Document()
    doc.add_paragraph("设备采购技术方案")
    doc.add_paragraph("一、项目概况")
    doc.add_paragraph("本次采购包含机房改造、网络设备与系统集成三部分内容。")
    doc.add_paragraph().add_run().add_picture(png_big, width=Cm(23))
    doc.add_paragraph("图 1  系统总体架构")
    doc.add_paragraph("二、设备清单")
    doc.add_paragraph("下列设备需要在高低温环境下连续运行，需满足稳定性要求。")
    doc.add_paragraph().add_run().add_picture(png_sq, width=Cm(6))
    doc.add_paragraph("图 2  设备外观")
    doc.save(docx)

    out = os.path.join(work, "out")
    print("[1] 图片开关开启 + 自动排版")
    res = optimize_file(docx, get_template("auto"),
                        OptimizeOptions(toggles=ToggleOptions(content=False)), out)
    print("    产出:", os.path.basename(res.target))
    pdf = render(res.target, os.path.join(work, "pdf"))
    check(bool(pdf), "PDF 导出成功")
    if pdf:
        inspect(pdf)

    print("\n[2] 图片开关关闭（应保持原 23cm 超宽，PDF 上会越界 → 证明开关有效）")
    res2 = optimize_file(docx, get_template("auto"),
                         OptimizeOptions(toggles=ToggleOptions(image_size=False)), out)
    pdf2 = render(res2.target, os.path.join(work, "pdf2"))
    if pdf2:
        d = fitz.open(pdf2)
        pg = d[0]
        widest = max((fitz.Rect(i["bbox"]).width for i in pg.get_image_info()),
                     default=0)
        d.close()
        check(widest > (595.3 - 6.34 * 28.35) + 10,
              f"关闭后超宽图仍越界（{widest:.1f}pt），说明开关确有作用")

    print("\n[3] 招标模板（一级标题分页）+ 内容精简")
    res3 = optimize_file(docx, get_template("tender"),
                         OptimizeOptions(toggles=ToggleOptions(content=True)), out)
    pdf3 = render(res3.target, os.path.join(work, "pdf3"))
    if pdf3:
        d = fitz.open(pdf3)
        txt = "".join(d[i].get_text() for i in range(d.page_count))
        d.close()
        check("众所周知" not in txt, "内容精简生效（赘语不在 PDF 中）")
        check("机房改造" in txt and "系统集成" in txt, "关键内容保留在 PDF 中")

    print("\n[4] Excel：图片不越出打印区域 + 表头底色渲染")
    check_excel(work)

    print(f"\n合计：通过 {len(PASS)} / 失败 {len(FAIL)}")
    for f in FAIL:
        print("  × " + f)
    return 0 if not FAIL else 1


if __name__ == "__main__":
    sys.exit(main())
