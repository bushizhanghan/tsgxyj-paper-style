# -*- coding: utf-8 -*-
"""review_triage.py —— 审稿意见自动分类与优先级助手（模式 F 前置环节）

返修第一步不是提笔回信，而是**把意见分好类**：审稿人常把「请核对样本量」
（必改）与「是否可以考虑补充案例？」（可选）写在同一段里，不分类就容易出现
两种失误——该改的没改（把必改当成可选意见），不该改的乱改（为迎合而改坏
研究逻辑）。人工分类依赖语感，条目一多就容易漏。
本脚本把审稿意见文件按条目切分后逐条打上四类标签：

  · 必改（事实性/规范性问题）：「必须」「应补充」「请补充」「请核对」
    「存在错误」「数据不一致」「格式不符」「建议补充」等确定性要求，
    外加**疑问句之外**的事实性/规范性问题陈述（不一致/矛盾/有误/缺失…）；
  · 疑问（需澄清）：「是否」「为何」「为什么」「请说明」「请解释」与问号；
  · 建议（可选优化）：「建议」「可考虑」「不妨」「如能」；
  · 表扬/正面（无需动作）：「观点鲜明」「有价值」「同意发表」类。

分类优先级：必改 > 疑问 > 建议 > 表扬——一条命中多类时取最高优先级
（「建议补充信效度检验」同时含「建议」与「建议补充」，判必改）。

输入支持 .docx（段落 + 表格）与 .md / .txt（含 markdown 表格）；条目切分、
表格列识别与多审稿人分组头判定复用 response_check.py 的解析函数
（read_paragraphs / split_items / parse_block / REVIEWER_HEAD_RE），
不重复实现一套解析。

用法:
    python review_triage.py --reviews <审稿意见.docx|md|txt> [--top 20] [--json]
退出码：0 = 分类完成（纯分类工具，不设硬伤，不因分类结果返回非零）；
        2 = 参数不全、文件不存在或无法解析。

局限（信号词启发式，不是语义理解，终审以人工为准）：
  · 分类只看**词面信号**：「建议补充」与「建议压缩」都判必改，轻重差异
    （审稿人私下说的「小修」「可不改」）机器读不出来；
  · 条目切分沿用 response_check.py 的编号/表格/段落启发式，自由书写的
    长段只能按段或逐行兜底，可能与审稿人原意不符；
  · 「疑问句外的确定性要求」通道只对不含问号的句子生效，
    「数据是否不一致？请补充说明」这类混合句按先命中的必改信号处理；
  · 未命中任何信号词的意见默认归入「建议」并单独计数，需人工确认；
  · 本工具只做分类与排序，**不代替人读意见**：回复撰写仍按
    references/response-template.md 逐条回应制执行。
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

HERE = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, HERE)
import response_check as rch  # noqa: E402  （复用其文件读取/条目切分/字段识别）

# ---------------- 分类定义 ----------------
CAT_MUST, CAT_Q, CAT_SUG, CAT_PRAISE = "必改", "疑问", "建议", "表扬"
# 优先级从高到低：一条命中多类时取最靠前者
CATS = [CAT_MUST, CAT_Q, CAT_SUG, CAT_PRAISE]
CAT_MARK = {CAT_MUST: "🔴", CAT_Q: "🟡", CAT_SUG: "🔵", CAT_PRAISE: "✅"}
CAT_DESC = {
    CAT_MUST: "事实性/规范性问题",
    CAT_Q: "需澄清",
    CAT_SUG: "可选优化",
    CAT_PRAISE: "正面评价，无需动作",
}

# 必改·显式信号：确定性要求（不论该句是否带问号）
MUST_EXPLICIT = [
    r"必须",
    r"应(?:当|予以|须)?(?:补充|增加|修改|更正|删除|统一|核实|核对|调整|压缩|精简|完善|规范)",
    r"需(?:要)?(?:补充|增加|修改|更正|删除|统一|核实|核对|完善|规范)",
    r"请(?:补充|增加|修改|更正|删除|统一|核实|核对|调整|压缩|精简|完善|规范|注意)",
    r"建议(?:补充|增加|修改|更正|删除|统一|核实|核对|调整|压缩|精简|完善|规范|明确)",
    r"有必要(?:补充|增加|修改|完善)",
    r"存在(?:错误|矛盾|问题|偏差|漏洞)",
    r"数据不一致|数据矛盾|前后不一致|相互矛盾|自相矛盾|前后矛盾|彼此矛盾",
    r"格式不符|体例不符|不符合(?:本刊|该刊|规范|要求|格式)|不规范|欠规范",
]
# 必改·确定性通道：仅对**不含问号**的句子生效（「疑问句外的确定性要求」）。
# 这些都是事实性/规范性问题陈述——审稿人不是在提问，是在指出错处。
MUST_DEFINITE = [
    r"不一致", r"矛盾", r"不符", r"有误", r"错误", r"偏差", r"偏误",
    r"缺失", r"缺少", r"遗漏", r"未标注", r"无单位", r"无法核对", r"不规范",
    r"漏洞", r"冗余", r"累赘", r"(?:超出|超过)(?:字数|篇幅|限制)|字数超标|篇幅过长",
]
QUESTION_SIGNALS = [
    r"是否", r"为何", r"为什么", r"何故", r"请说明", r"请解释", r"请澄清", r"请明确",
    r"能否", r"可否", r"是不是", r"有没有", r"指什么", r"什么意思", r"如何理解",
    r"[？?]",
]
SUGGEST_SIGNALS = [
    r"建议", r"可考虑", r"不妨", r"如能", r"似可", r"或可", r"可以进一步", r"可进一步",
    r"可适当", r"可酌情", r"可结合", r"可探讨", r"可补充", r"可加入", r"可完善", r"值得",
    r"希望",
]
PRAISE_SIGNALS = [
    r"同意发表", r"建议发表", r"观点鲜明", r"有价值", r"立意新颖", r"论述清晰",
    r"逻辑清晰", r"论证充分", r"数据翔实", r"内容丰富", r"结构完整", r"框架完整",
    r"富有启发", r"启发意义", r"获益良多", r"受益匪浅", r"写得很好", r"文笔流畅",
    r"可读性强", r"亮点", r"甚佳", r"优秀", r"赞成", r"赞同", r"认可",
]
SIGNAL_TABLE = {CAT_MUST: MUST_EXPLICIT, CAT_Q: QUESTION_SIGNALS,
                CAT_SUG: SUGGEST_SIGNALS, CAT_PRAISE: PRAISE_SIGNALS}

# 各类处理策略（联动 references/response-template.md 三类范式）
CAT_STRATEGY = {
    CAT_MUST: ("同意并说明改法 + 位置",
               "直接说明「已改为 X」并给出第 X 节第 X 段；不辩解、不赘述。"
               "确有不妥可部分同意，但须拿数据/文献说话。"),
    CAT_Q: ("先澄清再回应",
            "先把问题回答清楚（是/否 + 依据），再说明是否据此改动；"
            "不要只写「已核实」「详见正文」。"),
    CAT_SUG: ("说明采纳 / 不采纳理由",
              "采纳的写清改法与位置；不采纳的按「部分同意/不同意」范式给理由"
              "（标准条款、数据、文献），不做无原则让步。"),
    CAT_PRAISE: ("无需动作",
                 "稿件不改动；回复中一句致谢即可（感谢语克制，不堆「非常宝贵」）。"),
}


def trunc(s, n):
    return s if len(s) <= n else s[:n] + "…"


# 条目头前缀（切分后残留的「意见 1：」「问题2 」「（1）」「1. 」），
# 显示摘要前剥掉，避免与「意见 1」标签重复。
HEAD_PREFIX_RE = re.compile(
    r"^(?:审稿人?\s*\d+\s*[，,、]\s*)?(?:意见|问题|点评|疑问|质疑|建议|回复|comment|"
    r"question|point|item|q)\s*[0-9０-９一二三四五六七八九十]{1,3}\s*(?:[：:？?）)\]，,。.]|\s)"
    r"|^第\s*[0-9０-９一二三四五六七八九十]{1,3}\s*[条点个项]\s*[：:]\s*"
    r"|^[（(]\s*[0-9０-９一二三四五六七八九十]{1,3}\s*[)）]\s*"
    r"|^[0-9]{1,3}\s*[、.．)]\s*")


def strip_head(text):
    """剥掉条目头前缀（「意见 1：」/「（1）」/「1. 」），返回正文。"""
    t = (text or "").strip()
    m = HEAD_PREFIX_RE.match(t)
    if m:
        t = t[m.end():].strip()
    return t


def dedup_signals(words):
    """去掉被其他信号词包含的短词（「格式不符」「不符」只留前者）。"""
    return [w for w in words
            if not any(w != v and w in v for v in words)]


def sentences(text):
    """按中英文句末标点切句（保留标点；无标点则整段一句）。"""
    return [s for s in re.split(r"(?<=[。；;！!？?])", text or "") if s.strip()]


def matched_signals(text, patterns):
    """按表序返回 text 中命中的信号词（去重；问号统一显示为「问号」）。"""
    out = []
    for pat in patterns:
        m = re.search(pat, text or "")
        if not m:
            continue
        w = m.group(0)
        if w in ("？", "?"):
            w = "问号（？）"
        if w not in out:
            out.append(w)
    return out


def definite_problem(text):
    """疑问句外的确定性要求 → 命中的问题表述，否则 None。

    只扫描**不含问号**的句子：「数据是否不一致？」是疑问，
    「数据与正文不一致。」才是确定性要求。
    """
    for s in sentences(text):
        if re.search(r"[？?]", s):
            continue
        for pat in MUST_DEFINITE:
            m = re.search(pat, s)
            if m:
                return m.group(0)
    return None


def classify(text):
    """一条意见 → (类别, 依据信号词, 其他命中, 是否兜底)。

    按 CATS 顺序取首个有信号词的类别（必改 > 疑问 > 建议 > 表扬）；
    全不命中时兜底归入「建议」并置 fallback 标记。
    """
    t = text or ""
    hits = {CAT_MUST: matched_signals(t, MUST_EXPLICIT),
            CAT_Q: matched_signals(t, QUESTION_SIGNALS),
            CAT_SUG: matched_signals(t, SUGGEST_SIGNALS),
            CAT_PRAISE: matched_signals(t, PRAISE_SIGNALS)}
    d = definite_problem(t)
    if d and d not in hits[CAT_MUST]:
        hits[CAT_MUST].append(d)
    hits = {c: dedup_signals(w) for c, w in hits.items()}
    for c in CATS:
        if hits[c]:
            return c, hits[c], {k: v for k, v in hits.items() if v and k != c}, False
    return CAT_SUG, ["（未命中信号词，默认归入建议）"], {}, True


def reviewer_segments(lines):
    """审稿人分组头 → [(行号, 标签)]（分组头只作归属提示，不切成条目）。"""
    segs = []
    for i, t in enumerate(lines or []):
        if rch.REVIEWER_HEAD_RE.match(t):
            lab = re.sub(r"\s+", "", t).rstrip("：:")
            segs.append((i, lab or t.strip()))
    return segs


def reviewer_of(segs, idx):
    """idx 行所属的审稿人标签（取最近的前置分组头）。"""
    label = None
    for start, lab in segs:
        if start <= idx:
            label = lab
        else:
            break
    return label


def opinion_text(item):
    """条目 → 意见正文。

    复用 response_check.parse_block：意见文件若混有「回应/修改位置」列
    （编辑部返回的表格常见），只取「意见」列，避免把作者的回复当成意见分类。
    """
    block = item["lines"]
    f = rch.parse_block(block)
    text = (f.get("opinion") or "").strip()
    if not text:
        text = " ".join(block).strip()
    return re.sub(r"\s+", " ", text)


def triage(reviews_path):
    """读取审稿意见文件并逐条分类。

    返回 {"items": [...], "info": {...}}；条目结构与 response_check
    一致（{"no","par","lines"}），另加分类字段。
    """
    lines, heads = rch.read_paragraphs(reviews_path)
    items, prelude, mode = rch.split_items(lines, heads, "opinions")
    if not items:
        # 结构识别失败（自由书写的短行）：退化为逐行一条，避免空报告
        src = prelude if prelude else list(lines)
        items = [{"no": None, "par": i + 1, "lines": [t]}
                 for i, t in enumerate(src) if t and t.strip()]
        mode = "fallback_lines" if items else "none"
    segs = reviewer_segments(lines)
    # 分组头行自身不算意见：它被切进上一条的块尾（如「…同意发表。审稿人2」），
    # 取该行的行首标签部分从条目文本中剔除。
    seg_starts = {i: lab for i, lab in segs}

    out = []
    for k, it in enumerate(items, start=1):
        block = list(it["lines"])
        if block:
            last = block[-1]
            for start, lab in segs:
                if it["par"] - 1 < start <= it["par"] - 1 + len(block) - 1:
                    pos = start - (it["par"] - 1)
                    if pos == len(block) - 1:
                        block[-1] = last[:last.find(lab)].strip()
                        if not block[-1]:
                            block.pop()
                    break
        text = opinion_text(it) if block else opinion_text(it)
        if block and block != it["lines"]:
            text = re.sub(r"\s+", " ", " ".join(block).strip())
        if not text:
            continue
        cat, sigs, other, fb = classify(text)
        label = f"意见 {it['no']}" if it["no"] is not None else f"意见（段{it['par']}）"
        rev = reviewer_of(segs, it["par"] - 1)
        out.append({
            "seq": k, "label": label, "par": it["par"], "reviewer": rev,
            "text": strip_head(text), "summary": trunc(strip_head(text), 60),
            "cat": cat, "signals": sigs, "other": other, "fallback": fb,
        })
    info = {"n_items": len(out), "mode": mode,
            "reviewers": [lab for _, lab in segs],
            "fallback_n": sum(1 for r in out if r["fallback"])}
    return {"items": out, "info": info}


# ---------------- 输出 ----------------
def report(res, reviews_path, top):
    """按「必改 → 疑问 → 建议 → 表扬」打印分类统计、清单与处理策略。"""
    items, info = res["items"], res["info"]
    counts = {c: [r for r in items if r["cat"] == c] for c in CATS}
    stat = " ｜ ".join(f"{CAT_MARK[c]} {c} {len(counts[c])}" for c in CATS)
    print("=" * 70)
    print("审稿意见分类与优先级助手（模式 F 前置环节）")
    print(f"审稿意见：{os.path.basename(reviews_path)}"
          f"（{info['n_items']} 条意见；切分模式={info['mode']}）")
    if info["reviewers"]:
        print(f"检测到 {len(info['reviewers'])} 位审稿人分组头："
              f"{'、'.join(info['reviewers'])}")
    print("=" * 70)

    print(f"\n【分类统计】{stat}（共 {info['n_items']} 条）")
    print("处理优先级：必改 → 疑问 → 建议 → 表扬（一条命中多类时取最高优先级）")
    if info["fallback_n"]:
        print(f"其中 {info['fallback_n']} 条未命中信号词，已默认归入「建议」，请人工确认。")
    print("-" * 70)

    for c in CATS:
        group = counts[c]
        head, detail = CAT_STRATEGY[c]
        print(f"{CAT_MARK[c]} {c} {len(group)} 条（{CAT_DESC[c]}）")
        print(f"   处理策略：{head}——{detail}")
        if not group:
            print("   （无）")
            print()
            continue
        for i, r in enumerate(group[:top], start=1):
            who = f"{r['reviewer']}·" if r["reviewer"] else ""
            print(f"   · {i}. [{who}{r['label']}] {r['summary']}")
            basis = "、".join(r["signals"])
            print(f"         依据：{basis}（{c}）")
            if r["other"]:
                extra = "；".join(f"{k}「{'、'.join(v)}」" for k, v in r["other"].items())
                print(f"         另命中：{extra}（优先级低，已归入{c}）")
        if len(group) > top:
            print(f"   …另 {len(group) - top} 项，用 --top N 调整显示条数")
        print()

    print("-" * 70)
    print("提示：本工具只分类与排序；回复仍按 references/response-template.md "
          "逐条回应制撰写，")
    print("      提交前用 response_check.py 核「漏回 / 三件套 / 位置精确度」。")


def main():
    args = sys.argv[1:]
    flags = set(a for a in args if a.startswith("--"))
    if not args or "-h" in args or "--help" in args:
        print(__doc__)
        sys.exit(0)

    def _opt(name):
        if name in args:
            i = args.index(name)
            return args[i + 1] if i + 1 < len(args) else None
        return None

    reviews = _opt("--reviews")
    if not reviews:
        print("用法: python review_triage.py --reviews <审稿意见文件> [--top 20] [--json]")
        print("     支持 .docx（段落+表格）/ .md / .txt")
        sys.exit(2)
    if not os.path.isfile(reviews):
        print(f"[错误] 找不到文件：{reviews}")
        sys.exit(2)
    TOP = 20
    if "--top" in args:
        i = args.index("--top")
        try:
            TOP = int(args[i + 1])
        except (IndexError, ValueError):
            pass

    try:
        res = triage(reviews)
    except ValueError as e:
        print(f"[错误] {e}")
        sys.exit(2)

    if not res["items"]:
        print("[提示] 未从文件中识别到任何意见条目（空文件或全部为空行）。")
        sys.exit(0)

    report(res, reviews, TOP)

    if "--json" in flags:
        jf = os.path.splitext(reviews)[0] + ".review_triage.json"
        counts = {c: sum(1 for r in res["items"] if r["cat"] == c) for c in CATS}
        with open(jf, "w", encoding="utf-8") as f:
            json.dump({"reviews": os.path.basename(reviews),
                       "info": res["info"], "counts": counts,
                       "items": res["items"]},
                      f, ensure_ascii=False, indent=1)
        print(f"\nJSON 已写入：{jf}")

    sys.exit(0)


if __name__ == "__main__":
    main()
