import sys, io
sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding="utf-8", errors="replace")
from wxfixer.core import content_optimizer as C

tests = [
    "  众所周知，本项目 的实施 需要 各部门 配合。  ",
    "我们需要，尽快完成这项工作，，并 提交 报告!!!",
    "基本上,该方案可行。实际上，它能满足需求。",
    "值得一提的是，我们 的 产品 支持 多平台。",
    "此项工作非常重要。",
    "合同金额为 1,234.56 元，详见附件 A.",
    "然后然后我们需要以及以及确认",
]
for t in tests:
    out, st = C.optimize_paragraphs([t], "general")
    print(repr(t))
    print("  ->", repr(out[0]), "| safe:", C.is_safe_rewrite(t, out[0]))
    print("   ", st.summary())
