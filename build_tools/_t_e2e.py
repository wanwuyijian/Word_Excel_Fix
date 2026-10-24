# -*- coding: utf-8 -*-
"""端到端：真实文件 × 8 模板 × 三新开关，走 GUI 的 start() 流程（含校验与日志）。

用法：PYTHONPATH=src python build_tools/_t_e2e.py
"""
import io
import os
import shutil
import sys
import time

sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding="utf-8", errors="replace")
sys.stderr = io.TextIOWrapper(sys.stderr.buffer, encoding="utf-8", errors="replace")

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
sys.path.insert(0, os.path.join(ROOT, "src"))

from wxfixer.core import runner  # noqa: E402
from wxfixer.core import ooxml_check  # noqa: E402
from wxfixer.templates import library  # noqa: E402
from wxfixer.templates.schema import ToggleOptions  # noqa: E402
from wxfixer import selftest as ST  # noqa: E402

PASS = []
FAIL = []


def check(cond, msg):
    if cond:
        PASS.append(msg)
        print("  OK   " + msg)
    else:
        FAIL.append(msg)
        print("  FAIL " + msg)


def main():
    work = os.path.join(ROOT, "render_out", "e2e")
    if os.path.isdir(work):
        shutil.rmtree(work)
    os.makedirs(work)
    src_dir = os.path.join(work, "src")
    os.makedirs(src_dir)

    samples = []
    for name in ("发布通知", "技术说明", "招标文件", "产品介绍", "冗长草稿"):
        p = os.path.join(ROOT, "samples", name + "_原始.docx")
        if os.path.isfile(p):
            samples.append(p)
    for name in ("评分明细", "设备清单"):
        p = os.path.join(ROOT, "samples", name + "_原始.xlsx")
        if os.path.isfile(p):
            samples.append(p)
    img_doc = ST.build_image_sample(src_dir)
    samples.append(img_doc)
    print("样本 %d 个（含带图 Word）" % len(samples))

    # 三新开关全开 + 版式全开
    full = ToggleOptions(**{k: True for k in (
        "layout", "font_family", "font_size", "font_color",
        "background", "page_break", "content", "image_size")})

    out_root = os.path.join(work, "out")
    os.makedirs(out_root)

    print("\n[1] 8 模板 × 全部样本（三新开关全开）")
    for tpl in library.ALL_TEMPLATES:
        out_dir = os.path.join(out_root, tpl.key)
        os.makedirs(out_dir, exist_ok=True)
        ok_n = 0
        for src in samples:
            before = ST._fingerprint(src, "excel" if src.endswith((".xlsx", ".xlsm")) else "word")
            res = runner.optimize_file(src, tpl, runner.OptimizeOptions(tpl.key, full), out_dir)
            after = ST._fingerprint(res.target, res.doc_type)
            kind = "excel" if res.doc_type == "excel" else "word"
            if kind == "word":
                probs = ooxml_check.check_document(res.target)
                check(not probs, "%s / %s：OOXML 结构自检通过 %s"
                      % (tpl.name, os.path.basename(src), probs[:1]))
            check(os.path.isfile(res.target) and os.path.getsize(res.target) > 2000,
                  "%s / %s：产出存在（%.0f KB）"
                  % (tpl.name, os.path.basename(src), os.path.getsize(res.target) / 1024.0))
            # 内容开关开着，only 允许变短
            if kind == "word":
                same = before == after
                shorter = len(after) <= len(before)
                check(same or shorter,
                      "%s / %s：文字只减不增" % (tpl.name, os.path.basename(src)))
            else:
                same = before == after
                check(same or len(after) <= len(before),
                      "%s / %s：单元格文字只减不增" % (tpl.name, os.path.basename(src)))
            ok_n += 1
        check(ok_n == len(samples), "%s：%d/%d 个样本全部产出"
              % (tpl.name, ok_n, len(samples)))

    print("\n[2] content=False 时文字必须逐字不变")
    out2 = os.path.join(out_root, "strict")
    os.makedirs(out2, exist_ok=True)
    strict = ToggleOptions(layout=True, font_family=True, font_size=True, font_color=True,
                           background=True, page_break=True, content=False, image_size=True)
    for src in samples:
        res = runner.optimize_file(
            src, library.get_template("auto"), runner.OptimizeOptions("auto", strict), out2)
        kind = "excel" if res.doc_type == "excel" else "word"
        b = ST._fingerprint(src, kind)
        a = ST._fingerprint(res.target, kind)
        check(b == a, "严格模式 %s：文字逐字未变" % os.path.basename(src))

    print("\n[3] 全关：产出应与源文件等价（文字与图片尺寸均不变）")
    out3 = os.path.join(out_root, "alloff")
    os.makedirs(out3, exist_ok=True)
    off = ToggleOptions(**{k: False for k in (
        "layout", "font_family", "font_size", "font_color",
        "background", "page_break", "content", "image_size")})
    for src in samples:
        res = runner.optimize_file(src, library.get_template("auto"),
                                   runner.OptimizeOptions("auto", off), out3)
        kind = "excel" if res.doc_type == "excel" else "word"
        check(ST._fingerprint(src, kind) == ST._fingerprint(res.target, kind),
              "全关 %s：文字与源文件一致" % os.path.basename(src))
        if kind == "word":
            b = ST._image_widths_cm(src)
            a = ST._image_widths_cm(res.target)
            check(b == a, "全关 %s：图片尺寸不变 %s" % (os.path.basename(src), a or "无图"))

    print("\n[4] 冗长文档：内容开关应显著精简且保留关键事实")
    verbose = os.path.join(ROOT, "samples", "冗长草稿_原始.docx")
    if os.path.isfile(verbose):
        out4 = os.path.join(work, "verbose")
        os.makedirs(out4, exist_ok=True)
        res = runner.optimize_file(
            verbose, library.get_template("notice"),
            runner.OptimizeOptions("notice", full), out4)
        from docx import Document as _Doc
        raw = "\n".join(p.text for p in _Doc(verbose).paragraphs)
        text = "\n".join(p.text for p in _Doc(res.target).paragraphs)
        saved = len(raw) - len(text)
        check(saved >= 20, "冗长文档精简了 %d 个字符" % saved)
        for kw in ("核心交换机", "割接", "验收", "园区网络"):
            check(kw in text, "关键事实「%s」保留" % kw)
        check("众所周知" not in text and "值得一提的是" not in text,
              "赘语已从产出中移除")
        check("！！" in text and "！！！" not in text, "连续感叹号已收敛为两个")

    print("\n合计：通过 %d / 失败 %d" % (len(PASS), len(FAIL)))
    for m in FAIL:
        print("  失败项：" + m)
    return 0 if not FAIL else 1


if __name__ == "__main__":
    sys.exit(main())
