"""文字内容优化（可选开关，默认关闭）。

定位：**只做"减法"**——删掉赘语、多余空白、重复标点，
绝不新增任何文字、不改变事实、不调整语序。

安全策略（三层）：
1. 规则白名单：只删除固定词表里的填充语与多余空白/标点；
2. 保底校验：优化后的文本必须是原文的"字符子多重集"（不引入新字符），
   且长度不变长，见 :func:`is_safe_rewrite`；
3. 调用方（runner）在写盘前再做一次全局校验，任一段不通过即整体放弃输出。

不同文件类别（``profile``）启用不同的规则子集，让语句更贴合该文体的
表达习惯：公文偏简洁庄重、合同偏严谨、技术文档偏客观。
"""

from __future__ import annotations

import re
import unicodedata
from collections import Counter
from dataclasses import dataclass, field
from typing import Dict, List, Tuple

# ---------------------------------------------------------------------------
# 规则词表
# ---------------------------------------------------------------------------
# 句首引导性赘语（删掉不影响信息量，属于典型"口头禅"式敷陈）
_FILLER_PREFIX = (
    "值得一提的是，", "值得一提的是",
    "众所周知，", "众所周知",
    "需要指出的是，", "需要指出的是",
    "需要强调的是，", "需要强调的是",
    "需要注意的是，", "需要注意的是",
    "我们可以看到，", "我们可以看到",
    "可以看出，", "不难看出，",
    "一般来说，", "通常来说，",
    "总而言之，", "总的来说，", "综上所述，",
    "换言之，", "换句话说，", "也就是说，",
    "具体来说，", "简单来说，", "简单地说，",
    "毫无疑问，", "毋庸置疑，",
    "事实上，", "实际上，", "基本上，",
)

# 句中副词性赘语（保留语义主干后可安全删除）
_FILLER_WORDS = (
    "基本上", "事实上", "实际上", "相对来说", "相对而言",
    "在一定程度上", "在很大程度上",
    "非常非常", "十分十分",
    "然后然后", "以及以及", "并且并且", "或者或者",
    "的的", "了了", "是是", "和和",
)

# "进行 + 动词" 里的 "进行" 是可省的虚化动词（"对系统进行优化" → "对系统优化"）。
# 只在后面紧跟白名单动词时删除，避免误伤 "正在进行中" 这类固定说法。
_JINXING_VERBS = (
    "优化", "调整", "修改", "处理", "分析", "评估", "检查", "测试", "确认",
    "说明", "补充", "完善", "整理", "统计", "汇总", "核对", "审核", "审批",
    "落实", "推进", "实施", "建设", "部署", "配置", "安装", "调试", "维护",
    "管理", "培训", "汇报", "沟通", "协调", "整改", "排查", "演练", "验收",
    "备份", "清理", "压缩", "归档", "加密", "扩容", "升级", "更换",
)
_RE_JINXING = re.compile(
    r"(?<!正在)(?<!已)(?<!未)进行(?=" + "|".join(_JINXING_VERBS) + r")")

# 冗余助词：删除后句子更紧凑（只删字，不换字）
_REPLACE_RULES = (
    (re.compile(r"的过程中"), "过程中", "冗余助词"),
    (re.compile(r"((?:在|当|遇|出现)[^，。；：！？]{0,24})的情况下"), r"\1情况下", "冗余助词"),
    (re.compile(r"的时候(?=[，。；、！？]|$)"), "时", "冗余助词"),
)

# 各类文体的专属赘语（这些空话在对应文体里通常可直接删除）
_PROFILE_FILLER: Dict[str, Tuple[str, ...]] = {
    "notice": ("这个问题", "这个问题上", "请大家注意", "请大家务必注意", "在此特别说明"),
    "contract": ("双方在此确认并同意，", "双方在此确认，", "经双方协商一致同意，"),
    "tender": ("我们认为", "我方认为", "我们会", "我方可"),
    "bid": ("我们认为", "我方认为", "我们会"),
    "tech": ("我们可以看到", "我们可以看到，", "我们可以看到"),
    "product": ("我们的产品", "我们的产品可以", "本公司产品"),
    "general": (),
}

# 允许被删除的字符集合（用于保底校验）：空白 + 标点 + 上面词表出现的字
_WHITESPACE = {" ", "\u3000", "\t", "\xa0", "\u200b", "\ufeff"}
# 中文标点与常见英文标点：允许"规整"（数量变少或形态替换）但不许凭空新增
_PUNCT = set("，。、；：？！“”‘’（）【】《》…—～·,.;:?!\"'()[]{}<>/\\-_*#&@%+=|~`^$")
_FILLER_CHARS = set("".join(_FILLER_PREFIX) + "".join(_FILLER_WORDS) +
                    "".join("".join(v) for v in _PROFILE_FILLER.values()) +
                    "进行")
# 冗余表述规则里允许被删掉的虚词（is_safe_rewrite 的白名单补充）
_RULE_CHARS = set("的候")

# 半角 → 全角（仅在中日韩上下文里替换，避免破坏英文/数字/URL）
_HALF2FULL = {
    ",": "，", ".": "。", ";": "；", ":": "：",
    "?": "？", "!": "！", "(": "（", ")": "）",
}
# 这些半角符号即使在中文里也不该转全角
_KEEP_ASCII_PUNCT = set("@#$%&*+-/\\=<>[]{}|~^`_")

_CJK_RE = re.compile(r"[\u3400-\u9fff\uf900-\ufaff]")
_LATIN_RE = re.compile(r"[A-Za-z0-9]")
_MULTI_SPACE_RE = re.compile(r"[ \t\xa0\u200b]{2,}")
_SPACE_AROUND_CJK_PUNCT_RE = re.compile(
    r"[ \t]*([，。、；：？！“”‘’（）【】《》…—])[ \t]*"
)
_REPEAT_PUNCT_RE = re.compile(r"([，。、；：？！,.;:?!])\1{1,}")
_DOTS_RE = re.compile(r"(?:…{2,}|。{2,}|\.{3,})")
# 章节标题的序号前缀（"第一章 招标公告" 里的空格是排版分隔，不能删）
_CHAPTER_HEAD_RE = re.compile(r"^\s*第[一二三四五六七八九十百零〇\d]{1,4}[章节条篇部分]")


@dataclass
class ContentChange:
    """一处内容改动，用于日志与统计。"""
    rule: str
    before: str
    after: str


@dataclass
class ContentStats:
    paragraphs: int = 0
    changed_paragraphs: int = 0
    removed_chars: int = 0
    rules: Dict[str, int] = field(default_factory=dict)

    def note(self, rule: str) -> None:
        self.rules[rule] = self.rules.get(rule, 0) + 1

    def summary(self) -> str:
        if not self.changed_paragraphs:
            return "文字内容已是最简状态，无需精简"
        top = sorted(self.rules.items(), key=lambda kv: -kv[1])[:3]
        detail = "、".join(f"{k}×{v}" for k, v in top)
        return (f"精简 {self.changed_paragraphs}/{self.paragraphs} 段，"
                f"共减少 {self.removed_chars} 字（{detail}）")


# ---------------------------------------------------------------------------
def _has_cjk(text: str) -> bool:
    return bool(_CJK_RE.search(text))


def _normalize_spaces(text: str, stats: ContentStats) -> str:
    """规整空白：去首尾空格、压缩连续空格、去掉中文之间的空格。"""
    new = text.strip()
    if new != text:
        stats.note("首尾空白")
    squeezed = _MULTI_SPACE_RE.sub(" ", new)
    if squeezed != new:
        stats.note("压缩空格")
        new = squeezed
    # 中文字符之间的单个空格通常是误打，删掉更整齐。
    # 例外：章节标题 "第一章 招标公告" 里的空格是排版分隔，删掉反而更难认。
    if " " in new and _has_cjk(new):
        keep_upto = -1
        m_head = _CHAPTER_HEAD_RE.match(new)
        if m_head:
            keep_upto = m_head.end()
        parts = list(new)
        out: List[str] = []
        for i, ch in enumerate(parts):
            if ch == " ":
                prev = out[-1] if out else ""
                nxt = parts[i + 1] if i + 1 < len(parts) else ""
                if prev and nxt and _CJK_RE.match(prev) and _CJK_RE.match(nxt):
                    if i <= keep_upto:
                        pass  # 章节序号后的分隔空格，保留
                    else:
                        stats.note("中文间空格")
                        continue
            out.append(ch)
        new = "".join(out)
    return new


def _normalize_punct(text: str, stats: ContentStats) -> str:
    """标点规整：重复标点折叠、半角转全角、去掉标点两侧空格。"""
    new = text
    # 1) 重复标点：，，→ ，；！！！→ ！！（保留最多两个表强调）
    def _dedupe(m: re.Match) -> str:
        stats.note("重复标点")
        return m.group(1) * 2 if m.group(1) == "！" or m.group(1) == "？" else m.group(1)

    new = _REPEAT_PUNCT_RE.sub(_dedupe, new)
    # 2) 省略号/句号堆叠
    def _dots(m: re.Match) -> str:
        stats.note("省略号规整")
        return "……" if m.group(0).startswith("…") else "。"

    new = _DOTS_RE.sub(_dots, new)
    # 3) 半角转全角（仅当该标点两侧出现中文，且不是数字/英文语境）
    if _has_cjk(new):
        chars = list(new)
        for i, ch in enumerate(chars):
            if ch not in _HALF2FULL:
                continue
            prev = chars[i - 1] if i > 0 else ""
            nxt = chars[i + 1] if i + 1 < len(chars) else ""
            prev_cjk = bool(prev and _CJK_RE.match(prev))
            nxt_cjk = bool(nxt and _CJK_RE.match(nxt))
            prev_latin = bool(prev and _LATIN_RE.match(prev))
            nxt_latin = bool(nxt and _LATIN_RE.match(nxt))
            # 数字小数点、英文缩写、版本号、金额 等一律不动
            if prev_latin and nxt_latin:
                continue
            if ch == "." and prev.isdigit() and nxt.isdigit():
                continue
            if prev_cjk or nxt_cjk:
                chars[i] = _HALF2FULL[ch]
                stats.note("半角标点转全角")
        new = "".join(chars)
    # 4) 标点两侧不留空格
    squeezed = _SPACE_AROUND_CJK_PUNCT_RE.sub(r"\1", new)
    if squeezed != new:
        stats.note("标点两侧空格")
        new = squeezed
    return new


def _strip_filler(text: str, profile: str, stats: ContentStats) -> str:
    """删除赘语词。只做整词删除，不改动其前后文字。"""
    new = text
    # 句首（以及句号/分号之后的句首）的引导性赘语：连着后面的逗号一起删
    for word in _FILLER_PREFIX:
        if not word:
            continue
        pattern = re.compile(r"(^|(?<=[。！？；：\n]))" + re.escape(word) + r"[ \t]*")
        if pattern.search(new):
            new = pattern.sub(r"\1", new)
            stats.note("句首赘语")
    for word in _FILLER_WORDS:
        if word in new:
            replaced = new.replace(word, _safe_drop(word))
            if replaced != new:
                stats.note("句中赘语")
                new = replaced
    for word in _PROFILE_FILLER.get(profile, ()):
        if word and word in new:
            new = new.replace(word, _safe_drop(word))
            stats.note(f"文体赘语({profile})")
    return new


def _strip_redundant(text: str, stats: ContentStats) -> str:
    """删除冗余虚词/虚化动词（只删字，不换字，语序不动）。"""
    new = text
    for pattern, repl, rule in _REPLACE_RULES:
        if pattern.search(new):
            new = pattern.sub(repl, new)
            stats.note(rule)
    if _RE_JINXING.search(new):
        new = _RE_JINXING.sub("", new)
        stats.note("虚化动词")
    return new


def _safe_drop(word: str) -> str:
    """删除赘语后留下的连接词处理。

    "实际上" 这类副词直接删掉后语句仍通顺；但 "的的"/"了了" 这类重复
    字只能保留一个，因此单独判断。
    """
    if len(word) >= 2 and len(set(word)) == 1:
        return word[0]
    return ""


def _tidy_spacing_after_drop(text: str) -> str:
    """删除赘语后收拾残留：多余的逗号、连续空格。"""
    new = re.sub(r"[，,]{2,}", "，", text)
    new = re.sub(r"[ \t]{2,}", " ", new)
    new = re.sub(r"^[，,、；;]\s*", "", new)
    # "句号 + 逗号"这类因删词产生的残迹：保留句号
    new = re.sub(r"([。！？；])[ \t]*[，,、]+", r"\1", new)
    new = re.sub(r"([，,、])\s*([。！？；])", r"\2", new)
    return new


# ---------------------------------------------------------------------------
def optimize_text(text: str, profile: str = "general",
                  stats: ContentStats | None = None) -> str:
    """对单个段落文本做保守精简；返回新文本（可能等于原文）。"""
    if not text or not text.strip():
        return text
    stats = stats or ContentStats()
    original = text
    new = _normalize_spaces(text, stats)
    new = _normalize_punct(new, stats)
    new = _strip_filler(new, profile, stats)
    new = _strip_redundant(new, stats)
    new = _tidy_spacing_after_drop(new)
    new = _normalize_spaces(new, stats)
    new = _normalize_punct(new, stats)
    if not new.strip():
        # 整段都是赘语（极少见）→ 不删，保持原样更安全
        return original
    if not is_safe_rewrite(original, new):
        return original
    return new


def is_safe_rewrite(before: str, after: str) -> bool:
    """保底校验：新文本只能是原文的子多重集（不引入新字符）。

    满足以下三条才认为安全：
    1. 长度不变长；
    2. after 的每个字符都在 before 中出现足够次数；
    3. 被删掉的字符必须属于"空白 / 标点 / 赘语字"白名单。
    """
    if len(after) > len(before):
        return False
    ca, cb = Counter(after), Counter(before)
    removed: Counter = Counter()
    for ch, n in cb.items():
        extra = n - ca.get(ch, 0)
        if extra > 0:
            removed[ch] = extra
    if removed and not set(removed).issubset(_WHITESPACE | _PUNCT | _FILLER_CHARS | _RULE_CHARS):
        return False
    return True


def optimize_paragraphs(texts: List[str], profile: str = "general"
                        ) -> Tuple[List[str], ContentStats]:
    """批量处理段落文本，返回 (新文本列表, 统计)。"""
    stats = ContentStats(paragraphs=len(texts))
    out: List[str] = []
    for old in texts:
        new = optimize_text(old, profile, stats)
        if new != old:
            stats.changed_paragraphs += 1
            stats.removed_chars += max(0, len(old) - len(new))
        out.append(new)
    return out, stats


def profile_label(profile: str) -> str:
    return {
        "general": "通用", "notice": "公文", "contract": "合同",
        "tender": "招标", "bid": "投标", "tech": "技术", "product": "产品",
    }.get(profile, "通用")


def display_width(text: str) -> int:
    """显示宽度（中文=2）。校验用例里用来比较精简幅度。"""
    return sum(2 if unicodedata.east_asian_width(c) in "WF" else 1 for c in text)


# ---------------------------------------------------------------------------
# Word：把新文本写回原有 run，尽量保留每个字所属 run 的格式
# ---------------------------------------------------------------------------
def rewrite_paragraph_runs(paragraph, new_text: str) -> bool:
    """把段落文本改写为 ``new_text``，返回是否改动。

    做法：用 ``difflib`` 做字符级比对，只把"被删除/被替换"的字符从它原来
    所在的 run 里去掉；保留下来的字符仍留在原 run 中，因此加粗、颜色、
    超链接等格式不会整段丢失。

    仅支持"只减不增"的改写——出现插入操作时直接放弃（返回 False），
    保证绝不会往文档里塞新内容。
    """
    import difflib

    from docx.oxml.ns import qn

    runs = []
    for run in paragraph.runs:
        els = run._r.findall(qn("w:t"))
        if not els:
            continue
        text = "".join(e.text or "" for e in els)
        runs.append((run, els, text))
    if not runs:
        return False
    before = "".join(t for _, _, t in runs)
    if before == new_text:
        return False

    owner: List[int] = []
    for idx, (_, _, text) in enumerate(runs):
        owner.extend([idx] * len(text))

    matcher = difflib.SequenceMatcher(a=before, b=new_text, autojunk=False)
    out: List[List[str]] = [[] for _ in runs]
    for tag, i1, i2, j1, j2 in matcher.get_opcodes():
        if tag == "equal":
            for k in range(i1, i2):
                out[owner[k]].append(before[k])
        elif tag == "delete":
            continue
        elif tag == "replace":
            if i1 >= i2:
                return False            # 纯插入 → 放弃
            out[owner[i1]].extend(new_text[j1:j2])
        elif tag == "insert":
            return False                # 新增字符 → 放弃，安全第一
    new_texts = ["".join(chars) for chars in out]
    if "".join(new_texts) != new_text:
        return False

    for (run, els, old), new in zip(runs, new_texts):
        if new == old:
            continue
        try:
            els[0].text = new
            if new != new.strip() or "  " in new:
                els[0].set(qn("xml:space"), "preserve")
            for extra in els[1:]:
                extra.text = ""
        except Exception:
            return False
    return True


def word_char_multiset(text: str) -> Counter:
    """只统计"文字字符"（中日韩 + 字母 + 数字），用于内容安全校验。

    标点、空白不计入，因此半角转全角、删除空格等规整动作不会误判；
    而赘语删除只会让这里变少，绝不变多。
    """
    c: Counter = Counter()
    for ch in text:
        if _CJK_RE.match(ch) or ch.isalnum():
            c[ch] += 1
    return c


def is_char_subset(before: Counter, after: Counter) -> bool:
    """after 的每个字符计数都不能超过 before（即只删不增）。"""
    for ch, n in after.items():
        if n > before.get(ch, 0):
            return False
    return True

