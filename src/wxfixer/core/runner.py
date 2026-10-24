"""优化任务统一入口。

* 自动识别文件类型（按扩展名 + 文件头魔数 double check）
* 保存到新文件，绝不覆盖源文件
* 内容完整性校验：优化前后正文文本指纹（sha256）必须完全一致
* 记录耗时、产生的修改统计，供界面/日志使用
"""

from __future__ import annotations

import hashlib
import logging
import os
import shutil
import time
from dataclasses import dataclass, field
from typing import Dict, List, Optional

from ..templates.schema import LayoutSpec, ToggleOptions

log = logging.getLogger("wxfixer.runner")

WORD_EXT = {".docx", ".docm"}
EXCEL_EXT = {".xlsx", ".xlsm"}
LEGACY_WORD_EXT = {".doc"}
LEGACY_EXCEL_EXT = {".xls"}

SUPPORTED_EXT = WORD_EXT | EXCEL_EXT | LEGACY_WORD_EXT | LEGACY_EXCEL_EXT


class UnsupportedFileError(Exception):
    """文件类型不受支持。"""


class OptimizeError(Exception):
    """优化失败。"""


@dataclass
class OptimizeOptions:
    template_key: str = "auto"
    toggles: ToggleOptions = field(default_factory=ToggleOptions)


@dataclass
class OptimizeResult:
    source: str
    target: str
    doc_type: str                 # "word" | "excel"
    template_name: str
    elapsed: float = 0.0
    stats: Dict[str, int] = field(default_factory=dict)
    analysis: Dict[str, int] = field(default_factory=dict)
    warnings: List[str] = field(default_factory=list)

    def summary_line(self) -> str:
        stat_txt = "，".join(f"{k} {v}" for k, v in self.stats.items() if v)
        return f"{os.path.basename(self.target)} | {stat_txt} | {self.elapsed:.2f}s"


# ---------------------------------------------------------------------------
def detect_file_type(path: str) -> str:
    """返回 'word' / 'excel' / 抛 UnsupportedFileError。"""
    ext = os.path.splitext(path)[1].lower()
    if ext in WORD_EXT:
        kind = "word"
    elif ext in EXCEL_EXT:
        kind = "excel"
    elif ext in LEGACY_WORD_EXT:
        raise UnsupportedFileError(
            "检测到旧版 .doc 文件。请先在 Word 中用「另存为」保存为 .docx 后再优化。"
        )
    elif ext in LEGACY_EXCEL_EXT:
        raise UnsupportedFileError(
            "检测到旧版 .xls 文件。请先在 Excel 中另存为 .xlsx 后再优化。"
        )
    else:
        raise UnsupportedFileError(
            f"不支持的文件类型：{ext or '（无扩展名）'}。仅支持 .docx / .docm / .xlsx / .xlsm"
        )
    return _verify_by_magic(path, kind)


def _verify_by_magic(path: str, claimed: str) -> str:
    """双重确认：ZIP 头 + [Content_Types].xml 里的类型声明。"""
    import zipfile

    if not zipfile.is_zipfile(path):
        raise UnsupportedFileError(
            "文件不是有效的 Office Open XML 文档（可能已损坏或扩展名被改过）。"
        )
    try:
        with zipfile.ZipFile(path) as zf:
            names = set(zf.namelist())
            try:
                ct = zf.read("[Content_Types].xml").decode("utf-8", "ignore")
            except KeyError:
                ct = ""
    except Exception as exc:
        raise UnsupportedFileError(f"无法读取文件内容：{exc}") from exc

    looks_word = "wordprocessingml" in ct or any(n.startswith("word/") for n in names)
    looks_excel = "spreadsheetml" in ct or any(n.startswith("xl/") for n in names)

    if claimed == "word" and looks_excel and not looks_word:
        return "excel"
    if claimed == "excel" and looks_word and not looks_excel:
        return "word"
    if not looks_word and not looks_excel:
        raise UnsupportedFileError("无法识别的 Office 文档结构。")
    return claimed


def build_output_path(src: str, out_dir: str, template_key: str) -> str:
    """生成不冲突的输出路径：原名_排版优化_<模板名>.docx。"""
    os.makedirs(out_dir, exist_ok=True)
    stem, ext = os.path.splitext(os.path.basename(src))
    tag = _safe_tag(template_key)
    candidate = os.path.join(out_dir, f"{stem}_排版优化_{tag}{ext}")
    n = 2
    while os.path.exists(candidate):
        candidate = os.path.join(out_dir, f"{stem}_排版优化_{tag}({n}){ext}")
        n += 1
    return candidate


def _safe_tag(key: str) -> str:
    from ..templates import get_template

    name = get_template(key).name
    return name.replace("/", "-")


def _temp_path(target: str) -> str:
    """同名临时文件（保留扩展名），避免 openpyxl 按扩展名拒绝 .part。"""
    stem, ext = os.path.splitext(target)
    return f"{stem}.__opt_tmp__{ext}"


def _text_fingerprint(path: str, doc_type: str) -> str:
    """正文文本指纹：只取文本，忽略所有格式属性。"""
    if doc_type == "word":
        from .word_optimizer import extract_text_fingerprint

        return extract_text_fingerprint(path)
    from .excel_optimizer import extract_text_fingerprint as xl_fp

    return xl_fp(path)


def _text_chars(path: str, doc_type: str):
    """正文的"文字字符"多重集（忽略标点与空白），用于内容精简后的校验。"""
    from .content_optimizer import word_char_multiset

    if doc_type == "word":
        from .word_optimizer import dump_text

        text = dump_text(path)
    else:
        from .excel_optimizer import dump_text as xl_dump

        text = xl_dump(path)
    return word_char_multiset(text)


def _is_content_subset(path_before: str, path_after: str, doc_type: str) -> bool:
    """勾选"优化文字内容"时的校验：优化后的文字字符必须是原文的子集。

    即：不允许出现任何一个原文没有（或数量更多）的文字字符，
    保证工具只做"精简"，绝不会凭空写出新内容。
    """
    try:
        before = _text_chars(path_before, doc_type)
        after = _text_chars(path_after, doc_type)
    except Exception as exc:  # 取不到文本时保守判为不通过
        log.warning("内容子集校验无法执行：%s", exc)
        return False
    from .content_optimizer import is_char_subset

    return is_char_subset(before, after)


def optimize_file(
    src: str,
    template: LayoutSpec,
    options: OptimizeOptions,
    out_dir: str,
    progress=None,
) -> OptimizeResult:
    """对单个文件执行排版优化，返回结果。"""
    src = os.path.abspath(src)
    if not os.path.isfile(src):
        raise OptimizeError(f"文件不存在：{src}")

    doc_type = detect_file_type(src)
    started = time.perf_counter()

    def report(pct: int, msg: str) -> None:
        if progress:
            try:
                progress(pct, msg)
            except Exception:
                pass

    report(5, f"识别为 {'Word' if doc_type == 'word' else 'Excel'} 文档")
    log.info("开始优化: %s | 类型=%s | 模板=%s | 开关=%s",
             src, doc_type, template.key, options.toggles.to_dict())

    before = _text_fingerprint(src, doc_type)
    target = build_output_path(src, out_dir, template.key)

    work_dir = os.path.dirname(target)
    os.makedirs(work_dir, exist_ok=True)
    # 先复制到临时文件，在同一副本上修改，最后原子替换 → 中途出错不会留下半成品
    # 注意：临时文件必须保留原扩展名，否则 openpyxl 会按扩展名校验格式而拒绝
    tmp = _temp_path(target)
    shutil.copy2(src, tmp)

    try:
        report(20, "正在应用排版方案")
        if doc_type == "word":
            from .word_optimizer import optimize_word

            stats, analysis, warnings = optimize_word(tmp, template, options)
        else:
            from .excel_optimizer import optimize_excel

            stats, analysis, warnings = optimize_excel(tmp, template, options)

        report(80, "正在校验内容完整性")
        after = _text_fingerprint(tmp, doc_type)
        if before != after:
            # 勾选了"优化文字内容"时允许正文变少（只删不增），此时用
            # 更宽松但仍严格的校验：文字字符必须是原文的子集。
            if not (options.toggles.content
                    and _is_content_subset(src, tmp, doc_type)):
                raise OptimizeError(
                    "内容完整性校验未通过（优化前后正文文本不一致），已放弃输出。"
                    "为防止误改文档，请把源文件反馈给开发者。"
                )
            report(85, "已按「优化文字内容」开关精简文字")
            log.info("内容校验：勾选了文字优化，文字仅做减法，校验通过")

        os.replace(tmp, target)
    except Exception:
        if os.path.exists(tmp):
            try:
                os.remove(tmp)
            except OSError:
                pass
        raise

    elapsed = time.perf_counter() - started
    report(100, "完成")

    result = OptimizeResult(
        source=src,
        target=target,
        doc_type=doc_type,
        template_name=template.name,
        elapsed=elapsed,
        stats=stats,
        analysis=analysis,
        warnings=warnings,
    )
    log.info("优化完成: %s", result.summary_line())
    return result
