# -*- coding: utf-8 -*-
"""analyze_corpus.py -- 基于 118 篇原文的写作模式挖掘（供写作/改写/新写直接取用）

从 paragraphs.json（118 篇 / 2658 段正文）中挖掘可复用的「该刊怎么写」证据：

    1. 段首句式     每段第一句的开头模式（政策切入/现象切入/文献对话/数据陈述/
                   过渡承接/结论先行的分布与高频开头）
    2. 转场句式     段间承接的固定表达（综上所述类除外，取该刊真实高频衔接）
    3. 引用句式     「叙述＋[n]」的周边形态（[n]前的高频引导词、句末挂尾统计）
    4. 数据陈述句   含数值/占比/排序的句式模板（带具体值，可套用）
    5. 对策段句式   「因此/应/需要」类对策句的骨架
    6. 标题句式     各级标题的命名范式（双动宾、「——基于……」型等）
    7. 摘要句式     四要素各自的实现句式（目的/方法/结果/价值各取高频模板）

用法:
    python analyze_corpus.py                      # 全部分析，人读报告
    python analyze_corpus.py --dim 段首句式        # 只看一个维度
    python analyze_corpus.py --top 15             # 每维度取前 N
    python analyze_corpus.py --json               # 输出 JSON（供其他工具消费）
    python analyze_corpus.py --corpus <dir>       # 指定语料库

产物：把结果写入 references/patterns.md（供模式 C 新写时直接取用）。
"""
import os
import re
import sys
import json
from collections import Counter, defaultdict

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import corpus  # noqa: E402

if hasattr(sys.stdout, "reconfigure"):
    try:
        sys.stdout.reconfigure(encoding="utf-8", errors="replace")
    except Exception:
        pass

DIMS = ["段首句式", "转场句式", "引用句式", "数据陈述句", "对策段句式", "标题句式", "摘要句式"]


# ---------------- 基础工具 ----------------
def norm(s):
    """全角→半角 + 去空白，用于模式匹配。"""
    return s.translate(str.maketrans("０１２３４５６７８９　", "0123456789 "))


def sentences(p):
    return [s.strip() for s in re.split(r"[。？！]", p) if len(s.strip()) >= 6]


def first_n_chars(s, n):
    return s[:n]


# ---------------- 1. 段首句式 ----------------
OPEN_PATS = [
    ("政策/文件切入", r"^(19|20)\d{2}年|印发|发布|颁布|出台|实施《|《[^》]{4,30}》"),
    ("现象/背景切入", r"^(随着|近年来|当前|目前|伴随|在.{2,20}(背景|趋势|过程)中)"),
    ("文献对话切入", r"^([\u4e00-\u9fff]{2,4}(等)?(\[|［)?\d|已有研究|既有研究|现有研究|相关研究)"),
    ("论断切入", r"^(.{2,30}?(是|为)[^。]{4,40}(之基|之要|核心|关键|基础|前提))"),
    ("结论先行", r"^(结果表明|研究发现|研究显示|结果显示|可见|由此)"),
    ("目的/行为切入", r"^(为|旨在|为了|本文|本研究)"),
    ("转折承接", r"^(然而|但是|但|不过|与此)"),
    ("顺承递进", r"^(同时|此外|另外|进而|从而|因此|由此)"),
]


def dim_openings(data):
    cnt = Counter()
    examples = defaultdict(list)
    n_paras = 0
    for item in data:
        for p in item["paras"]:
            if len(p) < 60:
                continue
            n_paras += 1
            ss = sentences(p)
            if not ss:
                continue
            head = norm(ss[0])
            for name, pat in OPEN_PATS:
                if re.search(pat, head):
                    cnt[name] += 1
                    if len(examples[name]) < 3:
                        examples[name].append(ss[0][:60])
                    break
            else:
                cnt["其他/直接陈述"] += 1
                if len(examples["其他/直接陈述"]) < 3:
                    examples["其他/直接陈述"].append(ss[0][:60])
    return {"n_paras": n_paras, "dist": cnt.most_common(),
            "examples": {k: v for k, v in examples.items()}}


# ---------------- 2. 转场句式（段首衔接词，排除引用/数据段） ----------------
TRANS_PATS = [
    ("总结收束", r"^(综上所述|总而言之|总体而言|总的来看|概括而言)"),
    ("递进补充", r"^(同时|此外|另外|还有|进一步|进而)"),
    ("因果推导", r"^(因此|因而|由此|从而|这[意表]明)"),
    ("转折对照", r"^(然而|但是|但|不过|与此|相比|不同于)"),
    ("回指上文", r"^(在这一|基于这一|以此|据此|在这一基础上|在此基础上)"),
    ("引出下文", r"^(下面|下文|接下来|以下|本章|本节)"),
]


def dim_transitions(data):
    cnt = Counter()
    ex = defaultdict(list)
    for item in data:
        for p in item["paras"]:
            if len(p) < 80:
                continue
            ss = sentences(p)
            if not ss:
                continue
            head = norm(ss[0])
            for name, pat in TRANS_PATS:
                if re.search(pat, head):
                    cnt[name] += 1
                    if len(ex[name]) < 2:
                        ex[name].append(ss[0][:50])
                    break
    return {"dist": cnt.most_common(), "examples": dict(ex)}


# ---------------- 3. 引用句式（[n] 前的引导词 + 挂尾率） ----------------
CITE_LEAD = [
    ("观点转述（认为/指出/提出）", r"(认为|指出|提出|强调|发现|表明)[^。]{0,30}\[\d"),
    ("事实陈述挂尾", r"[^。，]{4,40}\[\d"),
    ("多文献并列", r"\[\d{1,3}[-–—,，]\s*\d{1,3}\]"),
    ("定义/概念溯源", r"(概念|理论|框架|模型|方法)(源于|来自|源自|由)[^。]{0,40}\[\d"),
]


def dim_citations(data):
    lead_cnt = Counter()
    ex = defaultdict(list)
    # 句末挂尾率：引用后紧跟句末标点（。！？；）或段落结尾，视为「句末引用」
    tail = {"cite_total": 0, "sent_end": 0, "mid_sent": 0}
    for item in data:
        for p in item["paras"]:
            if len(p) < 60:
                continue
            t = norm(p)
            for name, pat in CITE_LEAD:
                ms = re.findall(pat, t)
                if ms:
                    lead_cnt[name] += len(ms)
            for m in re.finditer(r"\[\d{1,3}(?:[-–—,，]\s*\d{1,3})*\]", t):
                tail["cite_total"] += 1
                after = t[m.end():m.end() + 1]
                if after in "。！？；" or after == "":
                    tail["sent_end"] += 1
                else:
                    tail["mid_sent"] += 1
            if lead_cnt["观点转述（认为/指出/提出）"] and len(ex["lead"]) < 3:
                m = re.search(r"[^。]{4,50}(认为|指出|提出)[^。]{0,40}\[\d{1,3}[^\]]*\]", t)
                if m:
                    ex["lead"].append(m.group(0)[:70])
    tot = tail["cite_total"] or 1
    return {"lead_dist": lead_cnt.most_common(),
            "tail": tail,
            "sent_end_rate": round(tail["sent_end"] / tot * 100, 1),
            "examples": dict(ex)}


# ---------------- 4. 数据陈述句 ----------------
DATA_PATS = [
    ("占比", r"[^。]{0,25}占比[^。]{0,30}(%|％)"),
    ("频次/数量", r"[^。]{0,20}(共|共计|达|为|有)\s*\d{1,5}\s*(个|条|篇|家|种|次|人|项)"),
    ("排序/最高最低", r"[^。]{0,30}(最高|最低|居首|位列|排名)"),
    ("比较级", r"[^。]{0,25}(高于|低于|超过|不足|接近)\s*\d"),
    ("倍数/翻番", r"[^。]{0,25}\d+(\.\d+)?\s*倍"),
    ("均值为/平均", r"[^。]{0,20}(均值|平均|平均值为|中位数)[^。]{0,25}\d"),
]


def dim_data(data):
    cnt = Counter()
    ex = defaultdict(list)
    for item in data:
        for p in item["paras"]:
            if len(p) < 60:
                continue
            for s in sentences(p):
                t = norm(s)
                for name, pat in DATA_PATS:
                    m = re.search(pat, t)
                    if m:
                        cnt[name] += 1
                        if len(ex[name]) < 3:
                            ex[name].append(s.strip()[:70])
    return {"dist": cnt.most_common(), "examples": dict(ex)}


# ---------------- 5. 对策段句式 ----------------
POLICY_PATS = [
    ("应/需要+动词", r"^(因此|由此|为此|基于此)?[，,]?\s*(应|需要|应当|亟需|有必要)"),
    ("路径/策略总起", r"(路径|策略|对策|建议|措施)(如下|包括|主要有|可分为)"),
    ("分条展开", r"(其一|其二|其三|第一|第二|第三|首先|其次)[，,、]"),
    ("回扣发现", r"(这一?结论|上述结果|研究发现)(表明|指向|揭示|意味着)"),
    ("双动宾小标题式", r"^[\u4e00-\u9fff]{2,8}[，,][\u4e00-\u9fff]{2,12}$"),
]


def dim_policy(data):
    cnt = Counter()
    ex = defaultdict(list)
    for item in data:
        for p in item["paras"]:
            if len(p) < 60:
                continue
            for s in sentences(p):
                t = norm(s)
                for name, pat in POLICY_PATS:
                    m = re.search(pat, t)
                    if m:
                        cnt[name] += 1
                        if len(ex[name]) < 3:
                            ex[name].append(s.strip()[:70])
    return {"dist": cnt.most_common(), "examples": dict(ex)}


# ---------------- 6. 标题句式（从 corpus_stats.json 的 heads 字段） ----------------
TITLE_PATS = [
    ("双动宾结构", r"^\d+\s*[\u4e00-\u9fff]{2,10}[，,][\u4e00-\u9fff]{2,14}$"),
    ("「——基于」型", r"^[\u4e00-\u9fffA-Za-z]{4,}——基于"),
    ("冒号分层型", r"^\d+\s*[^:：]{2,20}[:：]"),
    ("「的」字短语", r"^\d+\s*[\u4e00-\u9fff]{2,20}的[\u4e00-\u9fff]{2,20}$"),
    ("问题/对策型", r"^\d+\s*.*(问题|困境|对策|策略|路径|启示)"),
    ("引言/结语标配", r"^\d*\s*(引言|结语|绪论)"),
]


def dim_titles():
    rows = corpus.load_corpus_stats()
    cnt = Counter()
    ex = defaultdict(list)
    all_heads = []
    for r in rows:
        for h in r.get("heads") or []:
            all_heads.append(h)
            t = norm(h)
            for name, pat in TITLE_PATS:
                if re.search(pat, t):
                    cnt[name] += 1
                    if len(ex[name]) < 4:
                        ex[name].append(h[:40])
                    break
            else:
                cnt["其他"] += 1
    return {"n_heads": len(all_heads), "dist": cnt.most_common(),
            "examples": dict(ex),
            "h3_ratio": round(sum(1 for h in all_heads
                                  if re.match(r"^\d+\.\d+\.\d+", norm(h)))
                              / max(len(all_heads), 1) * 100, 1)}


# ---------------- 7. 摘要句式（从 text/ 的【摘　要】块） ----------------
def load_abstracts(corpus_dir):
    """从语料库 text/*.md 提取摘要正文（论文文本在 text/ 子目录）。"""
    out = []
    if not corpus_dir:
        return out
    text_dir = os.path.join(corpus_dir, "text")
    if not os.path.isdir(text_dir):
        text_dir = corpus_dir  # 退化：md 就在根目录
    for fn in sorted(os.listdir(text_dir)):
        if not fn.endswith(".md"):
            continue
        try:
            t = open(os.path.join(text_dir, fn), encoding="utf-8").read()
        except Exception:
            continue
        # 【摘　要】…【关键词】（中间为全角空格；长度 80–1200 防跨段）
        m = re.search(r"【摘\s*要】(.{80,1200}?)【关\s*键\s*词】", t, re.S)
        if m:
            out.append(m.group(1))
    return out


ABS_PATS = [
    ("目的（针对/面向/旨在）", r"(针对|面向|旨在|以期为|为[了弄]?清|探究|分析)"),
    ("方法（采用/借助/基于/构建）", r"(采用|借助|基于|利用|通过|构建|运用)"),
    ("结果（发现/表明/形成/提出）", r"(发现|表明|显示|形成|提出|得出|揭示|划分为)"),
    ("价值（提供/为…提供）", r"(提供|具有.*意义|有助于|可[为能]|以期)"),
]


def dim_abstracts(corpus_dir):
    abs_list = load_abstracts(corpus_dir)
    cnt = Counter()
    ex = defaultdict(list)
    lens = []
    for a in abs_list:
        a_ = norm(re.sub(r"\s+", "", a))
        lens.append(len(a_))
        for name, pat in ABS_PATS:
            if re.search(pat, a_):
                cnt[name] += 1
        # 首句模板（目的句）
        first = a_.split("。")[0] if "。" in a_ else a_[:40]
        if len(ex["首句"]) < 5 and first:
            ex["首句"].append(first[:70])
    n = max(len(abs_list), 1)
    return {"n_abs": len(abs_list),
            "elem_rate": {k: round(v / n * 100, 1) for k, v in cnt.items()},
            "len_med": sorted(lens)[len(lens) // 2] if lens else 0,
            "examples": dict(ex)}


# ---------------- 输出 ----------------
def print_dim(name, r):
    print(f"\n### {name}")
    if "n_paras" in r:
        print(f"  （样本 {r['n_paras']} 段）")
    for k, v in r.get("dist", []):
        print(f"  {k}: {v}")
    if "n_heads" in r:
        print(f"  （标题 {r['n_heads']} 个；三级标题占比 {r['h3_ratio']}%）")
    if "n_abs" in r:
        print(f"  （摘要 {r['n_abs']} 篇；四要素出现率 {r['elem_rate']}；长度中位 {r['len_med']} 字）")
    if "tail" in r:
        print(f"  [n] 总数 {r['tail']['cite_total']}；句末引用占 {r['sent_end_rate']}%"
              f"（句中引用 {r['tail']['mid_sent']} 处，亦合法）")
    for k, lst in r.get("examples", {}).items():
        if lst:
            print(f"  例·{k}:")
            for e in lst[:3]:
                print(f"    {e}")


def main():
    args = sys.argv[1:]
    if not args or args[0] in ("-h", "--help"):
        print(__doc__)
        return 0
    flags = {a for a in args if a.startswith("--")}
    rest = [a for a in args if not a.startswith("--")]
    cli_corpus = None
    if "--corpus" in args:
        i = args.index("--corpus")
        cli_corpus = args[i + 1] if i + 1 < len(args) else None
    top_n = 15
    if "--top" in args:
        i = args.index("--top")
        try:
            top_n = int(args[i + 1])
        except (ValueError, IndexError):
            pass
    want = rest[0] if rest else None
    if want and want not in DIMS:
        print(f"未知维度「{want}」；可选：{' / '.join(DIMS)}")
        return 2

    corpus_dir = corpus.find_corpus_or_exit(cli_corpus, need="模式挖掘")
    data = corpus.load_paragraphs(corpus_dir)
    if not data:
        print("[错误] 语料库无 paragraphs.json")
        return 2

    funcs = {
        "段首句式": lambda: dim_openings(data),
        "转场句式": lambda: dim_transitions(data),
        "引用句式": lambda: dim_citations(data),
        "数据陈述句": lambda: dim_data(data),
        "对策段句式": lambda: dim_policy(data),
        "标题句式": dim_titles,
        "摘要句式": lambda: dim_abstracts(corpus_dir),
    }
    selected = [want] if want else DIMS

    print("=" * 66)
    print(f"《图书馆学研究》写作模式挖掘（{len(data)} 篇 / "
          f"{sum(len(i['paras']) for i in data)} 段）")
    print("=" * 66)
    results = {}
    for d in selected:
        r = funcs[d]()
        results[d] = r
        print_dim(d, r)

    if "--json" in flags:
        out = {"corpus": corpus_dir, "dims": results}
        jf = os.path.join(corpus_dir, "patterns.json")
        with open(jf, "w", encoding="utf-8") as f:
            json.dump(out, f, ensure_ascii=False, indent=1)
        print(f"\nJSON 已写入：{jf}")
    return 0


if __name__ == "__main__":
    sys.exit(main())
