# -*- coding: utf-8 -*-
"""polish_lexicon.py -- 润色语言层：基于 118 篇原文的词汇与句法证据

回答「润色时把词换成什么、句子怎么调」的三个原文依据：

    1. 口语→书面替换词表   在该刊原文中检索「口语词」与候选书面词的真实共现，
                          给出该刊实际使用的对应表达（不是词典，是语料证据）
    2. 程度副词与模糊限定   该刊怎么表达程度（有所/一定程度上/显著……），
                          以及「显著」等强断言词的真实分布
    3. 句法特征统计         「了/着/过」动态助词、被动句「被/由…」、
                           单句动词密度、「即/也就是」解释句、破折号用法

用法:
    python polish_lexicon.py                 # 全维度人读报告
    python polish_lexicon.py --dim 词汇替换  # 单维度
    python polish_lexicon.py --json          # 写语料库 polish_lexicon.json
    python polish_lexicon.py <词>            # 查单个词在该刊的真实用法（上下文示例）

产物：references/polish-lexicon.md（润色词典，改写时直接取用）
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

DIMS = ["词汇替换", "程度表达", "句法特征"]


def norm(s):
    return s.translate(str.maketrans("０１２３４５６７８９　", "0123456789 "))


def sents_of(p):
    return [s.strip() for s in re.split(r"[。？！]", norm(p)) if len(s.strip()) >= 8]


def all_sentences(data):
    for item in data:
        for p in item["paras"]:
            if len(p) < 30:
                continue
            for s in sents_of(p):
                yield s


# ---------------- 1. 口语→书面 ----------------
# 每项：(口语词/口语结构, 该刊候选书面表达们)。共现实证：在原文中统计
# 口语词出现次数 vs 候选书面词出现次数，并给出书面词的真实例句。
COLLOQUIAL_MAP = [
    ("很多", ["大量", "诸多", "多数", "不少", "为数众多"]),
    ("非常", ["极为", "尤为", "高度", "十分", "显著"]),
    ("特别", ["尤为", "格外", "特定", "专门"]),
    ("差不多", ["大致", "基本", "大体", "接近"]),
    ("大概", ["约", "大致", "估计", "近"]),
    ("其实", ["事实上", "实际上", "本质上", "究其实"]),
    ("当然", ["固然", "毋庸置疑", "诚然"]),
    ("看来", ["由此可见", "综上", "据此可见"]),
    ("越来越好", ["持续提升", "逐步完善", "日益", "不断优化"]),
    ("越来越好地", ["持续", "逐步", "日益"]),
    ("一下子", ["迅速", "短期内", "即刻"]),
    ("有时候", ["部分情况下", "某些情形下", "偶尔"]),
    ("搞得", ["致使", "导致", "造成"]),
    ("弄", ["厘清", "明晰", "廓清"]),
    ("东西", ["要素", "内容", "构件", "载体"]),
    ("事儿", ["事务", "事项", "工作"]),
    ("怎么", ["如何", "何以", "怎样"]),
    ("什么样", ["何种形态", "何种面貌", "特征"]),
    ("大家", ["公众", "用户", "读者群体", "各方"]),
    ("咱们", ["我国", "本领域", "学界"]),
    ("老百姓", ["民众", "公众", "城乡居民"]),
    ("活儿", ["业务", "任务", "职责"]),
    ("靠谱", ["可靠", "可信", "具有效度"]),
    ("给力", ["有效", "有力", "具备支撑作用"]),
    ("走心", ["细致", "深入", "精准"]),
    ("硬核", ["核心", "关键", "实质性"]),
]

# 该刊常见的弱化/审慎表达（润色时把绝对化改审慎）
HEDGE_WORDS = ["有所", "一定程度上", "相对", "趋于", "逐步", "日益", "大多", "多数",
               "往往", "通常", "一般", "可能", "或将", "有望", "初显", "渐趋"]
STRONG_WORDS = ["显著", "极大", "充分", "全面", "彻底", "完全", "绝对", "必然", "一定"]


def dim_vocab(data):
    """统计口语词与书面候选词在原文中的真实频次。"""
    all_text = norm("".join(p for it in data for p in it["paras"]))
    n_chars = max(len(all_text), 1)
    rows = []
    for col, formal_list in COLLOQUIAL_MAP:
        c_cnt = all_text.count(col)
        f_cnt = {f: all_text.count(f) for f in formal_list}
        f_best = max(f_cnt, key=f_cnt.get) if f_cnt else None
        rows.append({
            "colloquial": col,
            "colloquial_cnt": c_cnt,
            "formal": f_best,
            "formal_cnt": f_cnt.get(f_best, 0) if f_best else 0,
            "formal_all": f_cnt,
            "verdict": "该刊几乎不用，改写应替换" if c_cnt == 0 and f_best else
                       ("该刊亦用，视语境" if c_cnt > 0 and c_cnt < 20 else "该刊常用"),
        })
    # 审慎 vs 绝对化词频
    hedge = {w: all_text.count(w) for w in HEDGE_WORDS}
    strong = {w: all_text.count(w) for w in STRONG_WORDS}
    return {"rows": rows, "hedge": hedge, "strong": strong, "n_chars": n_chars}


# ---------------- 2. 程度表达 ----------------
DEGREE_PATS = [
    ("审慎弱化（有所/一定程度上）", r"有所(提升|改善|增加|体现|显现|缓解)|一定程度上"),
    ("趋势渐进（日益/逐步/渐趋）", r"日益|逐步|渐趋|渐次|持续"),
    ("相对比较（相对/较之）", r"相对(于|而言)|较之|相较于"),
    ("多数限定（大多/多数/多数情况）", r"大多|多数|大部分|绝大部分"),
    ("显著（统计口径）", r"显著(正|负)?(相关|影响|差异|提升|作用|效应)"),
    ("明显（描述口径）", r"明显(高于|低于|提升|改善|差异)"),
]


def dim_degree(data):
    cnt = Counter()
    ex = defaultdict(list)
    for s in all_sentences(data):
        for name, pat in DEGREE_PATS:
            ms = re.findall(pat, s)
            if ms:
                cnt[name] += len(ms) if isinstance(ms[0], str) else 1
                if len(ex[name]) < 3:
                    ex[name].append(s[:70])
    return {"dist": cnt.most_common(), "examples": dict(ex)}


# ---------------- 3. 句法特征 ----------------
SYN_PATS = [
    ("动态助词「了」", r"了"),
    ("动态助词「着」", r"着[。，；）]|着[^了]"),
    ("被动「被」", r"被[\u4e00-\u9fff]{1,6}(、|，|。|；)?"),
    ("被动「由…（来/所）」", r"由[\u4e00-\u9fff]{1,12}(来|所)?(完成|实现|构成|提供|确定|承担)"),
    ("「即/也就是说」解释句", r"即[^。]{4,60}|也就是说[^。]{4,60}"),
    ("破折号插入语", r"—[^—]{4,80}—"),
    ("「称为/称之为/视为」定义句", r"称为|称之为|视为|定义为|界定为"),
    ("「旨在/以期」目的句", r"旨在|以期|以便|为[了能][^。]{4,60}"),
    ("并列分句（；连接）", r"[^。；]{10,60}；[^。；]{10,60}"),
]


def dim_syntax(data):
    cnt = Counter()
    per_1k = {}
    ex = defaultdict(list)
    total_chars = 0
    for s in all_sentences(data):
        total_chars += len(s)
        for name, pat in SYN_PATS:
            ms = re.findall(pat, s)
            if ms:
                cnt[name] += len(ms)
                if len(ex[name]) < 3:
                    ex[name].append(s[:70])
    k = max(total_chars, 1) / 1000
    per_1k = {n: round(c / k, 2) for n, c in cnt.items()}
    return {"counts": cnt.most_common(), "per_1k": per_1k,
            "n_chars": total_chars, "examples": dict(ex)}


# ---------------- 单词查询 ----------------
def query_word(data, w):
    """查某个词在该刊原文中的真实用法。"""
    hits = []
    for it in data:
        for i, p in enumerate(it["paras"]):
            t = norm(p)
            if w in t:
                for m in re.finditer(re.escape(w), t):
                    s = max(0, m.start() - 40)
                    hits.append((f"[{it['no']}|{it.get('issue', '')}]",
                                 t[s:m.end() + 50]))
                if len(hits) > 60:
                    break
        if len(hits) > 60:
            break
    print(f"「{w}」在语料库中命中 {len(hits)} 处（仅显示前 20）\n")
    for loc, ctx in hits[:20]:
        print(f"{loc} …{ctx}…")
    if not hits:
        print("（该刊语料中无此词——若你的稿中用了它，考虑替换为该刊常用表达，"
              "用 `python polish_lexicon.py --dim 词汇替换` 查对应书面词）")


# ---------------- 输出 ----------------
def print_dim(name, r):
    print(f"\n### {name}")
    if name == "词汇替换":
        print(f"  （语料 {r['n_chars']} 字）")
        print(f"  {'口语词':<8} {'刊内次数':>6}  {'该刊对应书面词':<10} {'次数':>6}  判定")
        for row in r["rows"]:
            if row["colloquial_cnt"] == 0 and row["formal_cnt"] == 0:
                continue
            print(f"  {row['colloquial']:<8} {row['colloquial_cnt']:>6}  "
                  f"{row['formal'] or '-':<10} {row['formal_cnt']:>6}  {row['verdict']}")
        print("\n  [审慎词频次] " + " ".join(f"{k}:{v}" for k, v in
                                          sorted(r["hedge"].items(), key=lambda x: -x[1])[:8]))
        print("  [强断言词频次] " + " ".join(f"{k}:{v}" for k, v in
                                          sorted(r["strong"].items(), key=lambda x: -x[1])[:8]))
    elif name == "程度表达":
        for k, v in r["dist"]:
            print(f"  {k}: {v}")
        for k, lst in r["examples"].items():
            if lst:
                print(f"  例·{k}:")
                for e in lst[:2]:
                    print(f"    {e}")
    elif name == "句法特征":
        print(f"  （统计 {r['n_chars']} 字；频率单位=次/千字）")
        for name_, c in r["counts"]:
            print(f"  {name_}: {c} 处 = {r['per_1k'].get(name_, 0)}/千字")
        for k, lst in list(r["examples"].items())[:4]:
            if lst:
                print(f"  例·{k}: {lst[0][:66]}")


def main():
    args = [a for a in sys.argv[1:]]
    if not args or args[0] in ("-h", "--help"):
        print(__doc__)
        return 0
    flags = {a for a in args if a.startswith("--")}
    rest = [a for a in args if not a.startswith("--")]
    cli_corpus = None
    if "--corpus" in args:
        i = args.index("--corpus")
        cli_corpus = args[i + 1] if i + 1 < len(args) else None

    corpus_dir = corpus.find_corpus_or_exit(cli_corpus, need="润色词典")
    data = corpus.load_paragraphs(corpus_dir)
    if not data:
        print("[错误] 语料库无 paragraphs.json")
        return 2

    # 单词查询模式
    if rest and rest[0] not in DIMS:
        query_word(data, rest[0])
        return 0

    want = rest[0] if rest else None
    if want and want not in DIMS:
        print(f"未知维度「{want}」；可选：{' / '.join(DIMS)}，或直接传一个词查用法")
        return 2

    funcs = {"词汇替换": lambda: dim_vocab(data),
             "程度表达": lambda: dim_degree(data),
             "句法特征": lambda: dim_syntax(data)}
    selected = [want] if want else DIMS

    print("=" * 66)
    print(f"润色语言层：基于《图书馆学研究》{len(data)} 篇原文的词汇/句法证据")
    print("=" * 66)
    results = {}
    for d in selected:
        r = funcs[d]()
        results[d] = r
        print_dim(d, r)

    if "--json" in flags:
        jf = os.path.join(corpus_dir, "polish_lexicon.json")
        with open(jf, "w", encoding="utf-8") as f:
            json.dump({"corpus": corpus_dir, "dims": results}, f,
                      ensure_ascii=False, indent=1)
        print(f"\nJSON 已写入：{jf}")
    return 0


if __name__ == "__main__":
    sys.exit(main())
