# -*- coding: utf-8 -*-
"""term_consistency.py —— 术语漂移检测器（同义异形词混用专查）

审稿人最省力的退稿理由之一是「术语不统一」：摘要写「信息无障碍」、正文写
「信息可及性」、结语又写「无障碍服务」——三处指向同一概念却用三种写法，
既显草率，也让读者怀疑作者是否真懂这个概念。
本脚本把**该刊 118 篇原文**作为偏好证据，检测用户稿件里的同义异形词混用：

  · 同义词组（SYNONYM_GROUPS）：按领域常识归纳的同义/近义写法族，
    每组含 2–4 个成员（如 信息无障碍／信息可及性／数字无障碍）；
  · 该刊偏好词：在 118 篇原文中统计各成员的**出现篇数**（不是次数——
    篇数反映「多少篇论文把它当正式术语用」，比次数稳健），篇数最高者
     即偏好词；全组 0 篇时标注「该刊未见」；
  · 漂移判定：稿件中同时出现组内 ≥2 个**非白名单**成员即报漂移，给出
     「非偏好词 → 偏好词」的统一建议，附该刊真实篇数与总次数；
  · 白名单（WHITELISTED_GROUPS）：在 118 篇见刊论文中统计共现率，
    该刊自己也在篇内并用的写法对（如 大语言模型/大模型 14% 篇目并用、
    老年人/老年群体 8.5%、视障/盲人 7.6%）**不判漂移**——它们不是笔误，
    而是「首现全称后文简称」的学术惯例或各有分工的概念；
  · 单成员提示：组内只出现 1 个成员且不是偏好词时，报「非该刊偏好写法」
    （🟡 提示级，不阻断）。

用法:
    python term_consistency.py <论文.docx> [--corpus <dir>] [--top 20] [--json]
退出码：0 = 无漂移（可有提示级）；1 = 检出同义异形词混用；
       2 = 显式 --corpus 指定的目录不是语料库。

局限（词面匹配，不是语义理解）：
  · 同义词组是人工归纳的**领域常识**，不是该刊编辑部口径；组外仍有
    作者自创译名／新概念（如「提示素养」vs「指令素养」）未被覆盖；
  · 白名单依据「该刊自己也这么用」的经验证据，不是概念 equivalence 证明；
  · 成员互为子串时（「无障碍」⊂「信息无障碍」）按长词优先匹配，
    短词的计数已扣除长词命中，避免重复计数。
"""
import os
import re
import sys
import json
from collections import defaultdict, Counter

# Windows 控制台 GBK 保护
if hasattr(sys.stdout, "reconfigure"):
    try:
        sys.stdout.reconfigure(encoding="utf-8", errors="replace")
    except Exception:
        pass

HERE = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, HERE)
import corpus  # noqa: E402
import diagnose_paper as dp  # noqa: E402

# ---------------- 同义词组 ----------------
# 说明：每组为该刊语料覆盖范围内常见的同义/近义异形写法。
# 「组」的边界取「指向同一概念、可互相替换而不改变句义」的最小集合；
# 有争议的对（阅读障碍/读写障碍）也入组，但只作提示级，由作者定夺。
SYNONYM_GROUPS = [
    # 无障碍 / 可及性族（该刊信息无障碍服务研究的高频概念群）
    # 说明：「无障碍」与「可及性」在该刊语境高度重叠，单列一组；
    # 二者同时是限定形（信息无障碍/信息可及性/数字无障碍）的构词成分，
    # 稿件已用限定形时裸词组不再参与（见 main() 的限定形优先规则），
    # 避免同一处文本被两组重复报。
    ["信息无障碍", "信息可及性", "数字无障碍"],
    ["无障碍", "可及性"],
    # 适老化族
    ["适老化", "老龄友好", "老年友好"],
    # 视障/阅读障碍族（中文译名不统一是重灾区）
    ["视障", "视力障碍", "视觉障碍", "盲人"],
    ["阅读障碍", "读写障碍"],
    # 认知障碍族
    ["失智症", "痴呆", "认知障碍"],
    ["自闭症", "孤独症"],
    # 数字鸿沟/包容族
    ["数字鸿沟", "数字排拒"],
    ["数字包容", "数字接纳"],
    # 群体称谓族
    ["老年人", "老年群体", "老龄群体", "银发群体"],
    ["儿童", "少儿"],
    ["弱势群体", "特殊群体", "信息弱势群体"],
    ["乡村居民", "农村居民", "农民"],
    # 数据族（该刊 2026 年高频新概念，译名/缩略混用明显）
    ["开放数据", "数据开放"],
    ["公共数据", "政府数据"],
    # AI 族（中英混写是该刊真实痛点）
    ["生成式人工智能", "生成式AI", "生成式 Ai", "AIGC"],
    ["大语言模型", "大模型", "LLM"],
    ["fsQCA", "定性比较分析", "组态分析"],
    # 图书馆形态族
    ["智慧图书馆", "智能图书馆"],
    # 阅读族
    ["数字阅读", "在线阅读", "移动阅读"],
    ["全民阅读", "阅读推广"],
    ["碎片化阅读", "浅阅读"],
    # 空间族
    ["空间再造", "空间服务"],
    # 融合族
    ["文旅融合", "文化和旅游融合"],
    # 注：以下写法族经 118 篇见刊论文共现率验证，**该刊自己就在一篇内并用**，
    # 属各有分工的概念或「首现全称后文简称」的学术惯例，不列为检测组：
    #   馆员/图书馆员/图书馆工作人员（30% 篇目并用：叙述简称 vs 定义全称）
    #   知识服务/信息服务/智慧服务（15% 篇目并用：服务层级谱系）
    #   信息素养/数字素养（14% 篇目并用：素养谱系两支）
    #   大模型/LLM、生成式AI/AIGC 的简称对（12–14% 篇目并用）
    # 这些词的统一应由作者在术语表层面自查，本工具不报。
    # 公平族
    ["信息公平", "信息平等"],
    # 转型族
    ["数字化转型", "数智化"],
]

# 白名单：这些组合虽形似，但是**不同的概念、上下义或学术简称惯例**，不视为同义异形。
# 依据：在 118 篇已见刊论文中统计共现率，共现率 ≥3% 的组说明该刊自己也常
# 在一篇内并用这些写法——它们不是笔误，而是各有分工的写法，必须排除。
# （共现率验证口径：逐篇扫描组内成员，统计 ≥2 成员共现的篇数占比）
WHITELISTED_GROUPS = {
    # ── 共现率 ≥10%（该刊高频并用，绝非漂移）──
    # 大语言模型/大模型/LLM：14% 篇目并用（首现全称后文简称，学术惯例）
    frozenset(["大语言模型", "大模型"]),
    frozenset(["大语言模型", "LLM"]),
    frozenset(["大模型", "LLM"]),
    # 生成式人工智能/生成式AI/AIGC：12% 篇目并用（中英简称并存）
    frozenset(["生成式人工智能", "生成式AI"]),
    frozenset(["生成式人工智能", "生成式 Ai"]),
    frozenset(["生成式人工智能", "AIGC"]),
    frozenset(["生成式AI", "生成式 Ai"]),
    frozenset(["生成式AI", "AIGC"]),
    frozenset(["生成式 Ai", "AIGC"]),
    # ── 共现率 3–10%（该刊亦常用，各写法有分工）──
    # 视障族：7.6% 篇目并用（视障=状态、盲人=人群、视力障碍=医学口径）
    frozenset(["视障", "盲人"]),
    frozenset(["视力障碍", "盲人"]),
    # 老年人/老年群体：8.5% 篇目并用（主语位置不同）
    frozenset(["老年人", "老年群体"]),
    # 儿童/少儿：5.9% 篇目并用（正式程度不同）
    frozenset(["儿童", "少儿"]),
    # 无障碍/可及性：5.1% 篇目并用（物理维度 vs 信息维度）
    frozenset(["无障碍", "可及性"]),
    # 公共数据/政府数据：5.9% 篇目并用（口径不同）
    frozenset(["公共数据", "政府数据"]),
    # 开放数据/数据开放：4.2% 篇目并用（数据集 vs 行为）
    frozenset(["开放数据", "数据开放"]),
    # fsQCA/定性比较分析/组态分析：8.5% 篇目并用（方法名/中译/步骤名）
    frozenset(["fsQCA", "定性比较分析"]),
    frozenset(["fsQCA", "组态分析"]),
    frozenset(["定性比较分析", "组态分析"]),
    # 全民阅读/阅读推广：6.8% 篇目并用（活动 vs 业务）
    frozenset(["全民阅读", "阅读推广"]),
    # 弱势群体/信息弱势群体：6.8% 篇目并用（上下义）
    frozenset(["弱势群体", "信息弱势群体"]),
    # 空间再造/空间服务：3.4% 篇目并用（改造行为 vs 服务类型）
    frozenset(["空间再造", "空间服务"]),
    # 数字化转型/数智化：5.9% 篇目并用（过程 vs 状态）
    frozenset(["数字化转型", "数智化"]),
    # ── 概念确有区分，交由作者定夺 ──
    # 阅读障碍（泛称）/读写障碍（dyslexia 规范译名）
    frozenset(["阅读障碍", "读写障碍"]),
    # 数字阅读（总称）/在线阅读、移动阅读（细分渠道）
    frozenset(["数字阅读", "在线阅读"]),
    frozenset(["数字阅读", "移动阅读"]),
    frozenset(["在线阅读", "移动阅读"]),
}


def is_not_synonym(a, b):
    """判断两词是否在白名单（不同概念/上下义，不判漂移）。"""
    return frozenset((a, b)) in WHITELISTED_GROUPS


# ---------------- 语料统计 ----------------
def corpus_term_stats(corpus_dir=None):
    """统计同义词组成员在 118 篇原文中的出现篇数与总次数。

    返回 {term: {"papers": n, "hits": m}}；语料库缺失返回 {}。
    长词优先：统计短词时扣除长词命中（「信息无障碍」里的「无障碍」不算）。
    """
    paras = corpus.load_paragraphs(corpus_dir)
    if not paras:
        return {}
    texts = [dp.clean("".join(p.get("paras") or [])) for p in paras]
    all_terms = []
    for g in SYNONYM_GROUPS:
        for t in g:
            if t not in all_terms:
                all_terms.append(t)
    # 长词优先：按长度降序处理，先记录长词命中位置，短词扣除
    spans = {t: [] for t in all_terms}
    for t in sorted(all_terms, key=len, reverse=True):
        for pi, txt in enumerate(texts):
            start = 0
            while True:
                i = txt.find(t, start)
                if i < 0:
                    break
                spans[t].append((pi, i, i + len(t)))
                start = i + 1
    taken = []
    for t in sorted(all_terms, key=len, reverse=True):
        keep = []
        for pi, s0, s1 in spans[t]:
            if not any(pi == tp and (s0 < tb and ts < s1) for tp, ts, tb in taken):
                keep.append((pi, s0, s1))
        spans[t] = keep
        taken.extend(keep)
    out = {}
    for t in all_terms:
        ps = {sp[0] for sp in spans[t]}
        out[t] = {"papers": len(ps), "hits": len(spans[t])}
    return out

def preferred_term(group, stats):
    """组内偏好词：篇数最高者（并列时取总次数高者、再并列取组内靠前者）。"""
    best, best_key = None, (-1, -1, -1)
    for i, t in enumerate(group):
        s = stats.get(t, {"papers": 0, "hits": 0})
        key = (s["papers"], s["hits"], -i)
        if key > best_key:
            best, best_key = t, key
    return best


# ---------------- 稿件检测 ----------------
def scan_doc(path, stats):
    """扫描稿件：返回各组在稿件中的命中（含段落定位）。"""
    paras, refs, abs_text, kw_text, titles, body = dp.extract_paras(path)
    # 逐段文本：摘要 / 正文 / 结语 / 关键词（标题不入扫描，避免标题一词定调）
    regions = []
    if abs_text:
        regions.append(("摘要", [abs_text]))
    regions.append(("正文", list(body)))
    if kw_text:
        regions.append(("关键词", [kw_text]))
    # 结语：按收口章标题切分
    concl = []
    in_closing = False
    for style, t in paras:
        if not t:
            continue
        if t.startswith("参考文献"):
            break
        is_title = (re.match(r"^(0|1|2|3|4|5|6|7)\s*[^\d]", t) and len(t) < 40) or \
                   (style.startswith("Heading") and len(t) < 40)
        if is_title:
            in_closing = bool(re.search(r"结\s*语|结\s*论|结束语|小\s*结|研究结论", t))
            continue
        if in_closing:
            concl.append(t)
    if concl:
        regions.append(("结语", concl))

    hits = defaultdict(lambda: defaultdict(list))   # term -> region -> [(par, snip)]
    for region, plist in regions:
        for pi, raw in enumerate(plist):
            t = dp.clean(raw or "")
            if not t:
                continue
            for g in SYNONYM_GROUPS:
                for w in g:
                    start = 0
                    while True:
                        i = t.find(w, start)
                        if i < 0:
                            break
                        hits[w][region].append((pi + 1, snip(t, i)))
                        start = i + len(w)
    return dict(hits), regions


def snip(text, pos, back=20, fwd=14):
    """证据片段：词前后各取一段。"""
    a = max(0, pos - back)
    b = min(len(text), pos + fwd)
    s = text[a:b].replace("\n", " ")
    return ("…" if a > 0 else "") + s + ("…" if b < len(text) else "")


def trunc(s, n):
    return s if len(s) <= n else s[:n] + "…"


# ---------------- 主流程 ----------------
def main():
    args = [a for a in sys.argv[1:] if not a.startswith("--")]
    flags = set(a for a in sys.argv[1:] if a.startswith("--"))
    if not args or args[0] in ("-h", "--help"):
        print(__doc__)
        sys.exit(0)
    DOC = args[0]
    if not os.path.isfile(DOC):
        sys.exit(f"找不到文件：{DOC}")
    TOP = 20
    if "--top" in sys.argv:
        i = sys.argv.index("--top")
        try:
            TOP = int(sys.argv[i + 1])
        except (IndexError, ValueError):
            pass
    _corpus_arg = None
    if "--corpus" in sys.argv:
        i = sys.argv.index("--corpus")
        _corpus_arg = sys.argv[i + 1] if i + 1 < len(sys.argv) else None

    try:
        cdir = corpus.find_corpus_dir(_corpus_arg)
    except FileNotFoundError as e:
        # 显式指定的坏目录直接报错（与 corpus.py 其他工具一致）
        print(f"[错误] {e}")
        sys.exit(2)
    stats = corpus_term_stats(cdir)
    if not stats:
        # 语料库缺失：只做稿内混用检测，无法给出该刊偏好词
        print("[提示] 未找到语料库，仅做稿内同义异形词混用检测（无该刊偏好词证据）。")
        print("       可加 --corpus <目录> 或设置环境变量 TSGXYJ_CORPUS。")

    hits, regions = scan_doc(DOC, stats)

    print("=" * 70)
    print("术语漂移检测（同义异形词混用 · 该刊 118 篇原文偏好证据）")
    print(f"稿件：{os.path.basename(DOC)}")
    if stats:
        src = f"语料库 {len(corpus.load_paragraphs(cdir))} 篇" if cdir else "语料库"
        print(f"偏好证据：{src}（篇数 = 多少篇论文把该写法当正式术语用）")
    print(f"扫描区域：{' / '.join(r for r, _ in regions)}")
    print("=" * 70)

    drifts, hints = [], []
    for g in SYNONYM_GROUPS:
        present = [w for w in g if hits.get(w)]
        if not present:
            continue
        # 限定形优先：稿件已用限定形（如「信息可及性」）时，
        # 裸词组（「无障碍/可及性」）不再参与，避免同一处文本被两组重复报
        if g and g[0] in ("无障碍", "适老化", "视障", "老年人"):
            qualified = [w for grp in SYNONYM_GROUPS for w in grp
                         if len(w) > len(g[0]) and g[0] in w and hits.get(w)]
            if qualified:
                continue
        pref = preferred_term(g, stats) if stats else None
        # 稿内漂移：把「已见刊论文自己也常用」的白名单对从 present 中剔除后，
        # 若仍有 ≥2 个成员，说明存在真正的同义异形混用。
        # 白名单对逐个剔除（成对处理），保留「有真实混用」的成员。
        real = list(present)
        for a in present:
            for b in present:
                if a != b and is_not_synonym(a, b):
                    # a、b 是合法并存的不同概念，双双移出混用判定
                    if a in real:
                        real.remove(a)
                    if b in real:
                        real.remove(b)
        # 若剔除后为空，说明 present 内所有对都是白名单 → 无漂移
        if len(real) >= 2:
            # 无偏好词证据时（语料库缺失），取组内出现频次最高者作建议目标
            if not pref:
                pref = max(real, key=lambda w: sum(len(v) for v in hits[w].values()))
            non_pref = [w for w in real if w != pref]
            for w in non_pref:
                locs = []
                for region, lst in hits[w].items():
                    for par, sn in lst[:2]:
                        locs.append((region, par, sn))
                drifts.append({
                    "group": g, "word": w, "preferred": pref,
                    "papers": (stats.get(w) or {}).get("papers", 0),
                    "hits": (stats.get(w) or {}).get("hits", 0),
                    "pref_papers": (stats.get(pref) or {}).get("papers", 0) if pref else 0,
                    "pref_hits": (stats.get(pref) or {}).get("hits", 0) if pref else 0,
                    "n_doc": sum(len(v) for v in hits[w].values()),
                    "locs": locs, "type": "混用",
                })
        elif len(present) == 1 and pref and present[0] != pref:
            # 单成员且非偏好词 → 提示级
            w = present[0]
            s = stats.get(w) or {"papers": 0, "hits": 0}
            ps = stats.get(pref) or {"papers": 0, "hits": 0}
            if ps["papers"] >= 5 and s["papers"] <= 2:
                locs = []
                for region, lst in hits[w].items():
                    for par, sn in lst[:2]:
                        locs.append((region, par, sn))
                hints.append({
                    "group": g, "word": w, "preferred": pref,
                    "papers": s["papers"], "hits": s["hits"],
                    "pref_papers": ps["papers"], "pref_hits": ps["hits"],
                    "n_doc": sum(len(v) for v in hits[w].values()),
                    "locs": locs, "type": "非该刊偏好写法",
                })

    print(f"\n【检测结果】🔴 术语混用 {len(drifts)} 组 ｜ 🟡 非该刊偏好写法 {len(hints)} 处")
    print("-" * 70)
    if drifts:
        print(f"🔴 术语混用 {len(drifts)} 组（建议全文统一为该刊偏好词）：")
        for d in drifts[:TOP]:
            if stats:
                ev = (f"（本稿 {d['n_doc']} 处；该刊：{d['word']} {d['papers']} 篇 "
                      f"vs {d['preferred']} {d['pref_papers']} 篇）")
            else:
                ev = f"（本稿 {d['n_doc']} 处；语料库缺失，按本稿频次建议）"
            print(f"   · 「{d['word']}」→「{d['preferred']}」{ev}")
            for region, par, sn in d["locs"][:2]:
                print(f"       {region}·段{par}：{trunc(sn, 60)}")
            others = "、".join(w for w in d["group"] if w not in (d["word"], d["preferred"]))
            if others:
                print(f"       同组还有：{others}")
        if len(drifts) > TOP:
            print(f"   …另 {len(drifts) - TOP} 处，用 --top N 调整")
    else:
        print(f"🔴 术语混用 0 组")

    if hints:
        print(f"\n🟡 非该刊偏好写法 {len(hints)} 处（该刊另有更高频写法，非硬伤）：")
        for h in hints[:TOP]:
            print(f"   · 「{h['word']}」该刊 {h['papers']} 篇 ｜ 偏好词「{h['preferred']}」"
                  f" {h['pref_papers']} 篇（本稿 {h['n_doc']} 处）")
            for region, par, sn in h["locs"][:1]:
                print(f"       {region}·段{par}：{trunc(sn, 60)}")

    if not drifts and not hints:
        print("✅ 未检出同义异形词混用（词面匹配，仅覆盖已定义的同义词组）。")

    if "--json" in flags:
        jf = os.path.splitext(DOC)[0] + ".term_consistency.json"
        with open(jf, "w", encoding="utf-8") as f:
            json.dump({"doc": os.path.basename(DOC),
                       "corpus": bool(stats),
                       "drifts": drifts, "hints": hints,
                       "n_groups": len(SYNONYM_GROUPS)},
                      f, ensure_ascii=False, indent=1)
        print(f"\nJSON 已写入：{jf}")

    sys.exit(1 if drifts else 0)


if __name__ == "__main__":
    main()
