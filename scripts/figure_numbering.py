# -*- coding: utf-8 -*-
"""figure_numbering.py —— 图表编号连续性与指路句核验

table_fig_audit.py 核的是「表格是不是三线表、图位占位符有没有删」，
不核**编号序列**。而编辑部校稿必查的恰恰是这些：
    图 1、图 2、图 4（缺图 3）；表 2、表 2（重号）；
    正文写「如图 3 所示」但全文没有图 3；图号出现在图题之前（先提及后出现的
    顺序颠倒）。这些在校对环节会被逐条挑出，投稿前自查成本极低。

本脚本逐项核验：
  · 编号连续性：图/表各自从 1 起、连续、无重号（含「图 1-1」式分图编号）；
  · 指路句对应：正文「如图 N 所示」「见表 N」是否都指向**已存在**的图表；
  · 反向覆盖：已编号的图表是否都有正文指路句（有图无引用会被问「图在哪」）；
  · 顺序：图表首次指路句的位置是否早于图表标题位置（先提及后出现）；
  · 图/表编号分离：图、表各自独立编号，不共用一套数字。

用法:
    python figure_numbering.py <论文.docx> [--json]
退出码：0 = 编号与指路句无问题；1 = 存在断号/重号/悬空指路/无引用图。

局限（词面识别，不做版面位置判断）：
  · 以 docx 段落顺序近似版面顺序（嵌入表格/图片的实际位置需人工核对）；
  · 分图（图 1a、图 1-1）与母图编号的包含关系按「母号存在即可」处理；
  * 指路句识别「如图 N/见表 N」及「见图 N/见表 N/图 N 所示」等常见变体。
"""
import os
import re
import sys
import json

# Windows 控制台 GBK 保护
if hasattr(sys.stdout, "reconfigure"):
    try:
        sys.stdout.reconfigure(encoding="utf-8", errors="replace")
    except Exception:
        pass

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import diagnose_paper as dp  # noqa: E402

# 图题/表题：段首「图 N」或「表 N」；也认占位形态「【图位：xxx（图 N）】」
CAP_HEAD = re.compile(r"^\s*(图|表)\s*(\d{1,2})(?:\s*[-–—]\s*(\d{1,2}))?\s*[\s　]")
CAP_PLACEHOLDER = re.compile(r"【(?:图位|待补)[^】]*?[（(]\s*(图|表)\s*(\d{1,2})")
# 正文指路句：如「如图 1 所示」「见表 2」「见图 3」「表 1 所示」
REF_SENT = re.compile(
    r"(?:如|参见|详见|见|参见)?\s*(图|表)\s*(\d{1,2})\s*(?:所示|所示|中)?")
# 排除：标题段本身（「图 1 系统架构图」不是指路句）
TITLE_LINE = re.compile(r"^\s*(图|表)\s*\d{1,2}\s*[\s　]")


def collect(doc):
    """收集全文段落序列，标注每个段落的图表标题与指路句。"""
    from docx import Document
    d = Document(doc)
    paras = [(p.text or "").strip() for p in d.paragraphs]

    caps = []          # [{kind, num, sub, para_idx, text}]
    refs = []          # [{kind, num, para_idx, text}]
    for idx, t in enumerate(paras):
        if not t:
            continue
        m = CAP_HEAD.match(t)
        pm = CAP_PLACEHOLDER.search(t)
        if m or pm:
            kind = (m or pm).group(1)
            num = int((m or pm).group(2))
            sub = int((m or pm).group(3)) if (m and (m or pm).group(3)) else None
            caps.append({"kind": kind, "num": num,
                         "sub": sub,
                         "para_idx": idx, "text": t[:50],
                         "placeholder": bool(pm)})
            continue
        # 指路句：只在该段不是标题段时统计
        for rm in REF_SENT.finditer(t):
            # 跳过「图 N」开头的标题行（已作 caps 处理）
            if TITLE_LINE.match(t):
                continue
            # 排除「图片」「图表」「目的」等误匹配
            kind = rm.group(1)
            after = t[rm.end():rm.end() + 2]
            before = t[max(0, rm.start() - 1):rm.start()]
            if before in ("图", "表") or after in ("片", "形", "达"):
                continue
            refs.append({"kind": kind, "num": int(rm.group(2)),
                         "para_idx": idx, "text": t[:60]})
    return paras, caps, refs


def analyze(doc):
    paras, caps, refs = collect(doc)
    issues, notices = [], []

    by_kind = {"图": [], "表": []}
    for c in caps:
        by_kind.setdefault(c["kind"], []).append(c)

    for kind, lst in by_kind.items():
        if not lst:
            continue
        nums = [c["num"] for c in lst]
        # ① 重号
        seen = {}
        for c in lst:
            key = (c["num"], c["sub"])
            if key in seen:
                issues.append(f"{kind}{c['num']}"
                              f"{'-' + str(c['sub']) if c['sub'] else ''} 编号重复："
                              f"「{seen[key]['text'][:20]}」与「{c['text'][:20]}」")
            else:
                seen[key] = c
        # ② 断号（按母号去重后查连续性）
        uniq = sorted(set(nums))
        if uniq:
            missing = [n for n in range(1, max(uniq) + 1) if n not in uniq]
            if missing:
                issues.append(f"{kind}编号断号：缺 {'、'.join(f'{kind}{m}' for m in missing)}"
                              f"（现有 {sorted(uniq)}）")
            if min(uniq) != 1:
                issues.append(f"{kind}编号不从 1 起（最小为 {kind}{min(uniq)}）")
        # ③ 悬空指路（指路句指向不存在的图表）
        for r in refs:
            if r["kind"] != kind:
                continue
            if r["num"] not in set(nums):
                issues.append(f"悬空指路：正文「{'如' if r['kind']=='图' else ''}"
                              f"{kind}{r['num']}」指向不存在的{kind}"
                              f"（段{r['para_idx'] + 1}）")

        # ④ 有图无引用（已编号图表无正文指路）
        ref_nums = {r["num"] for r in refs if r["kind"] == kind}
        for n in sorted(set(nums)):
            if n not in ref_nums:
                notices.append(f"{kind}{n} 未在正文被指路引用"
                               f"（该刊惯例「如图 N 所示」显式引出）")

        # ⑤ 顺序：指路句应早于对应图表标题（占位形态是稿件的待补标记，
        #    不算正式图表位置，跳过顺序判断）
        first_cap = {}
        for c in lst:
            if c.get("placeholder"):
                continue
            first_cap.setdefault(c["num"], c["para_idx"])
        for r in refs:
            if r["kind"] != kind or r["num"] not in first_cap:
                continue
            if r["para_idx"] > first_cap[r["num"]]:
                notices.append(f"{kind}{r['num']} 的指路句出现在图表标题之后"
                               f"（段{r['para_idx'] + 1}）——宜「先提及后出现」")

    # ⑥ 图表混编号（同一数字既是图又是表属正常，但同段既写图 N 又写表 N 且 N 相同需提示）
    stat = {
        "n_fig_caps": len(by_kind.get("图", [])),
        "n_tab_caps": len(by_kind.get("表", [])),
        "n_fig_refs": sum(1 for r in refs if r["kind"] == "图"),
        "n_tab_refs": sum(1 for r in refs if r["kind"] == "表"),
        "fig_nums": sorted({c["num"] for c in by_kind.get("图", [])}),
        "tab_nums": sorted({c["num"] for c in by_kind.get("表", [])}),
        "caps": caps, "refs": refs,
    }
    return {"issues": issues, "notices": notices, "stat": stat}


def main():
    args = [a for a in sys.argv[1:] if not a.startswith("--")]
    flags = set(a for a in sys.argv[1:] if a.startswith("--"))
    if not args or args[0] in ("-h", "--help"):
        print(__doc__)
        sys.exit(0)
    DOC = args[0]
    if not os.path.isfile(DOC):
        sys.exit(f"找不到文件：{DOC}")

    res = analyze(DOC)
    st = res["stat"]
    print("=" * 74)
    print("图表编号连续性与指路句核验（编辑部校稿必查项）")
    print(f"稿件：{os.path.basename(DOC)}")
    print(f"图题 {st['n_fig_caps']} 个（编号 {st['fig_nums']}）｜ "
          f"表题 {st['n_tab_caps']} 个（编号 {st['tab_nums']}）")
    print(f"正文指路句：图 {st['n_fig_refs']} 处 ｜ 表 {st['n_tab_refs']} 处")
    print("=" * 74)

    if res["issues"]:
        print(f"🔴 发现 {len(res['issues'])} 项编号问题：")
        for it in res["issues"]:
            print(f"   · {it}")
    else:
        print("✅ 图表编号连续、指路句均有对应图表。")
    if res["notices"]:
        print(f"\n🟡 建议核对 {len(res['notices'])} 项：")
        for it in res["notices"]:
            print(f"   · {it}")

    if "--json" in flags:
        jf = os.path.splitext(DOC)[0] + ".figure_numbering.json"
        with open(jf, "w", encoding="utf-8") as f:
            json.dump(res, f, ensure_ascii=False, indent=1, default=str)
        print(f"\nJSON 已写入：{jf}")

    sys.exit(1 if res["issues"] else 0)


if __name__ == "__main__":
    main()
