"""OOXML 结构自检：不依赖 Word 即可发现"文件已损坏，需要修复"的隐患。

检查内容：
1. ``w:tcPr / w:trPr / w:tblPr / w:pPr / w:rPr / w:tblBorders / w:sectPr``
   的子元素顺序是否符合 schema（顺序错误是 Word 报损坏的头号原因）。
2. 包内关系（rels）与部件是否齐全、能否被 zip 正常读取。
"""

from __future__ import annotations

import zipfile
from typing import List

from lxml import etree
from docx.oxml.ns import qn

from .ooxml import PPR_SEQ, RPR_SEQ, SECTPR_SEQ, TBLBORDERS_SEQ, TBLPR_SEQ, TCPR_SEQ, TRPR_SEQ

W = "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}"

_SEQ_TABLE = {
    f"{W}tcPr": TCPR_SEQ,
    f"{W}trPr": TRPR_SEQ,
    f"{W}tblPr": TBLPR_SEQ,
    f"{W}pPr": PPR_SEQ,
    f"{W}rPr": RPR_SEQ,
    f"{W}sectPr": SECTPR_SEQ,
    f"{W}tblBorders": TBLBORDERS_SEQ,
    f"{W}tcBorders": TBLBORDERS_SEQ,
}


def _check_tree(root) -> List[str]:
    issues: List[str] = []
    for parent in root.iter():
        tag = parent.tag
        seq = _SEQ_TABLE.get(tag)
        if seq is None:
            continue
        idxs = []
        for child in parent:
            name = child.tag
            if not isinstance(name, str) or not name.startswith(W):
                continue
            wname = "w:" + name[len(W):]
            if wname in seq:
                idxs.append((seq.index(wname), wname))
        for (a, an), (b, bn) in zip(idxs, idxs[1:]):
            if b < a:
                issues.append(f"{tag.split('}')[-1]} 子元素顺序错误：{an} 在 {bn} 之前")
                break
    return issues


def check_document(path: str) -> List[str]:
    """检查 .docx 的结构顺序；返回问题列表（空表示通过）。"""
    issues: List[str] = []
    with zipfile.ZipFile(path) as zf:
        bad = zf.testzip()
        if bad is not None:
            issues.append(f"压缩包内 {bad} 校验失败")
        names = zf.namelist()
        for required in ("[Content_Types].xml", "word/document.xml"):
            if required not in names:
                issues.append(f"缺少必需部件 {required}")
        for name in names:
            if not name.endswith(".xml"):
                continue
            if not (name.startswith("word/") or name == "[Content_Types].xml"):
                continue
            try:
                root = etree.fromstring(zf.read(name))
            except etree.XMLSyntaxError as exc:
                issues.append(f"{name} XML 解析失败：{exc}")
                continue
            for issue in _check_tree(root):
                issues.append(f"{name}: {issue}")
            if len(issues) > 100:
                return issues
    return issues


def check_package_rels(path: str) -> List[str]:
    """检查所有 .rels 指向的部件是否存在。"""
    issues: List[str] = []
    with zipfile.ZipFile(path) as zf:
        names = set(zf.namelist())
        for name in names:
            if not name.endswith(".rels"):
                continue
            base = name.rsplit("_rels/", 1)[0]
            try:
                root = etree.fromstring(zf.read(name))
            except etree.XMLSyntaxError as exc:
                issues.append(f"{name} 解析失败：{exc}")
                continue
            for rel in root:
                target = rel.get("Target") or ""
                if rel.get("TargetMode") == "External" or target.startswith("http"):
                    continue
                resolved = _resolve(base, target)
                if resolved not in names:
                    issues.append(f"{name}: 关系目标缺失 {target}")
    return issues


def _resolve(base: str, target: str) -> str:
    if target.startswith("/"):
        return target.lstrip("/")
    parts = (base + target).split("/")
    out: List[str] = []
    for part in parts:
        if part in ("", "."):
            continue
        if part == "..":
            if out:
                out.pop()
            continue
        out.append(part)
    return "/".join(out)


__all__ = ["check_document", "check_package_rels", "qn"]
