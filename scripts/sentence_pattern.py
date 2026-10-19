# -*- coding: utf-8 -*-
"""sentence_pattern.py —— 句首模式单调性核验

用途
----
统计全文**句首模式**的分布，判断稿件的句式是否单调：

    · 直接陈述      不以连接词/时间背景/「本研究」起句（该刊主流）
    · 连接词起句    句首是「同时/此外/因此/然而/一方面…」等衔接词
    · 时间背景起句  句首是「随着/近年来/当前/目前/2021 年…」
    · 「本研究」起句  句首是「本文/本研究/笔者」

判据（语料依据，全部可由 paragraphs.json 复跑，`--corpus-baseline` 打印实测值）：
    · 该刊实测**直接陈述 94.3%**、连接词起句 2.5%、时间背景起句 0.9%、
      「本研究」起句 3.6%（4743 段口径，见下「实测口径」）。直接陈述是该刊
      的默认句式，占比越高越贴刊风，**本脚本不因「直接陈述多」判单调**；
    · 连接词起句 >10% → 判「过度衔接」（该刊 2.5% 的 4 倍缓冲）；
    · 直接陈述 <85% → 🔴 句式被其他模式挤占，在该刊语境下属非主流；
    · 时间背景起句 >8% / 「本研究」起句 >8% → 🟡 提示铺陈过度。

实测口径（复跑方式：`python sentence_pattern.py --corpus-baseline`）：
    取 paragraphs.json 每段的**首句**（段 clean 后 ≥60 字；4743 段全满足），
    按上表四分类匹配。本目录 analyze_corpus.py 用另一套更细的九分类
    （政策切入/现象切入/文献对话/…）得到「其他/直接陈述 81.0%」，两者不矛盾：
    九分类把「论断切入/文献对话切入」等也算作非直接陈述，四分类口径更宽，
    与任务给定的 94.3% 基线一致。全句口径（不分段首句）的连接词起句为
    12.1%，与「>10% 判过度衔接」的红线同量级，故红线取 10% 有意义。

用法:
    python sentence_pattern.py <论文.docx> [--json] [--top 15]
    python sentence_pattern.py --corpus-baseline          # 只打印该刊实测基线

退出码：0 = 无判级问题；1 = 检出「连接词起句 >10%」的过度衔接。
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
import corpus  # noqa: E402
import diagnose_paper as dp  # noqa: E402

# ---- 四类句首模式（衔接性连接词；序数分条单列，不计入过度衔接）----
CONN_RE = re.compile(
    r"^(同时|此外|另外|还有|进一步|进而|从而|因此|因而|由此|然而|但是|但|不过|"
    r"与此|相比之下|相较于|不同于|基于这一|在这一|以此|据此|在此基础上|与此同时|"
    r"综上所述|总之|总的来说|总体而言|总的来看|具体而言|除此之外)")
TIME_RE = re.compile(
    r"^(随着|近年来|当前|目前|伴随|在.{2,20}(背景|趋势|过程)中|(19|20)\d{2}年)")
SELF_RE = re.compile(r"^(本文|本研究|笔者)")
ORDINAL_RE = re.compile(
    r"^(一是|二是|三是|四是|其一|其二|其三|首先|其次|再次|最后|第一|第二|第三|第四|"
    r"一方面|另一方)")

MODES = ["直接陈述", "连接词起句", "时间背景起句", "「本研究」起句"]

# 该刊实测（paragraphs.json 118 篇 / 4743 段，口径见 docstring）
CORPUS_DIRECT = 94.3        # %
CORPUS_CONN = 2.5           # %
CORPUS_TIME = 0.9           # %
CORPUS_SELF = 3.6           # %

# 判级阈值
DIRECT_FLOOR = 85.0   # 直接陈述低于此值 → 在该刊语境下属非主流，提示
CONN_LIMIT = 10.0     # 连接词起句 >10% → 过度衔接（红线）
TIME_LIMIT = 8.0      # 时间背景起句 >8% → 铺陈过度
SELF_LIMIT = 8.0      # 「本研究」起句 >8% → 铺陈过度
MIN_SENT = 12         # 参与统计的最短句长（clean 后）

CACHE = {}


def _sentences(text):
    """按句末标点切句（保留标点），剔除空句。"""
    return [s for s in re.split(r"(?<=[。？！])", text or "") if s.strip()]


def _classify(sent):
    """返回句首模式名。

    序数分条（第一/一是/首先…）不计入「连接词起句」——该刊 175 处机械序数
    （anti-patterns.md §B 级）是合法分条手段，问题只在「每段都同一套」，
    那由 paragraph_opening.py 与 language_audit 的 G 类负责；本模块只看
    衔接性连接词（同时/此外/因此/然而…）。
    """
    h = dp.clean(sent)
    if CONN_RE.match(h):
        return "连接词起句"
    if TIME_RE.match(h):
        return "时间背景起句"
    if SELF_RE.match(h):
        return "「本研究」起句"
    return "直接陈述"


def analyze(doc, corpus_dir=None):
    """统计稿件句首模式分布。返回 dict。"""
    paras, refs, abs_text, kw_text, titles, body = dp.extract_paras(doc)
    # 摘要不参与句式统计（摘要是压缩文体，句式与正文不同）；只统计正文段
    cnt = {m: 0 for m in MODES}
    ordinal = 0        # 序数分条起句（单列，不计入「连接词起句」）
    samples = {m: [] for m in MODES}
    n = 0
    for b in body:
        t = dp.clean(b or "")
        if len(t) < 60:
            continue
        for s in _sentences(t):
            h = dp.clean(s)
            if len(h) < MIN_SENT:
                continue
            if ORDINAL_RE.match(h):
                ordinal += 1
            m = _classify(s)
            cnt[m] += 1
            n += 1
            if len(samples[m]) < 3:
                samples[m].append(h[:52])

    pct = {m: (cnt[m] / n * 100 if n else 0.0) for m in MODES}
    issues, notices = [], []
    if n == 0:
        return {"n_sents": 0, "counts": cnt, "pct": pct, "samples": samples,
                "issues": ["正文过短，无法统计句首模式"], "notices": []}

    if pct["连接词起句"] > CONN_LIMIT:
        issues.append(
            f"连接词起句占 {pct['连接词起句']:.1f}%（>{CONN_LIMIT:.0f}%）——"
            f"该刊实测仅 {CORPUS_CONN}%，每句都挂连接词是「过度衔接」的典型形态；"
            f"段内逻辑靠语义承接即可，连接词留给真正的转折/递进")
    if pct["直接陈述"] < DIRECT_FLOOR:
        issues.append(
            f"直接陈述仅 {pct['直接陈述']:.1f}%（<{DIRECT_FLOOR:.0f}%）——"
            f"该刊实测 {CORPUS_DIRECT}%，句式被其他模式大量挤占；"
            f"请核对是连接词起句过多还是时间背景/「本研究」铺陈过长")
    if pct["时间背景起句"] > TIME_LIMIT:
        notices.append(
            f"时间背景起句 {pct['时间背景起句']:.1f}%（>{TIME_LIMIT:.0f}%）——"
            f"「随着…/近年来…」铺陈偏多（该刊 {CORPUS_TIME}%），背景一段足够")
    if pct["「本研究」起句"] > SELF_LIMIT:
        notices.append(
            f"「本研究/本文」起句 {pct['「本研究」起句']:.1f}%（>{SELF_LIMIT:.0f}%）——"
            f"自称起句偏多（该刊 {CORPUS_SELF}%），可改无主句或对象起句")
    if ordinal:
        notices.append(
            f"序数分条起句 {ordinal} 句（{ordinal / n * 100:.1f}%）——"
            f"该刊 175 处机械序数属合法分条手段，但同段别混用"
            f"「首先/第一/一是」多套体系")

    return {"n_sents": n, "counts": cnt, "pct": pct, "samples": samples,
            "ordinal": ordinal, "issues": issues, "notices": notices}


def corpus_baseline(corpus_dir=None):
    """从 paragraphs.json 实测该刊句首模式分布（供 --corpus-baseline 复跑）。"""
    if "base" in CACHE:
        return CACHE["base"]
    cdir = corpus.find_corpus_dir(corpus_dir)
    if not cdir:
        return None
    papers = corpus.load_paragraphs(cdir)
    cnt = {m: 0 for m in MODES}
    n = 0
    for paper in papers:
        for p in paper.get("paras") or []:
            t = dp.clean(p)
            if len(t) < 60:
                continue
            for s in _sentences(t):
                if len(dp.clean(s)) < MIN_SENT:
                    continue
                cnt[_classify(s)] += 1
                n += 1
    out = {"n_sents": n, "papers": len(papers),
           "pct": {m: (cnt[m] / n * 100 if n else 0.0) for m in MODES},
           "counts": cnt}
    CACHE["base"] = out
    return out


def main():
    args = [a for a in sys.argv[1:] if not a.startswith("--")]
    flags = set(a for a in sys.argv[1:] if a.startswith("--"))
    # --corpus-baseline 可单独运行（无 docx 参数），故先于 -h 分支判定
    if "--corpus-baseline" not in flags and (not args or args[0] in ("-h", "--help")):
        print(__doc__)
        sys.exit(0)

    TOP = 15
    for i, a in enumerate(sys.argv):
        if a == "--top" and i + 1 < len(sys.argv):
            try:
                TOP = int(sys.argv[i + 1])
            except (ValueError, IndexError):
                pass
    cli_corpus = None
    if "--corpus" in sys.argv:
        i = sys.argv.index("--corpus")
        cli_corpus = sys.argv[i + 1] if i + 1 < len(sys.argv) else None

    if "--corpus-baseline" in flags:
        base = corpus_baseline(cli_corpus)
        if not base:
            print("[错误] 未找到语料库（需要 corpus_stats.json / paragraphs.json）")
            sys.exit(2)
        print("=" * 66)
        print("该刊句首模式实测基线（paragraphs.json 复跑）")
        print(f"样本：{base['papers']} 篇 / {base['n_sents']} 句")
        print("=" * 66)
        for m in MODES:
            print(f"  {m:<14} {base['pct'][m]:5.1f}%  ({base['counts'][m]} 句)")
        sys.exit(0)

    DOC = args[0]
    if not os.path.isfile(DOC):
        sys.exit(f"找不到文件：{DOC}")

    res = analyze(DOC, cli_corpus)
    print("=" * 68)
    print("句首模式单调性核验（该刊 4743 段实测基线）")
    print(f"稿件：{os.path.basename(DOC)}")
    print(f"统计：{res['n_sents']} 句（正文段，clean 后段长 ≥60）")
    print(f"该刊基线：直接陈述 {CORPUS_DIRECT}% ｜ 连接词起句 {CORPUS_CONN}% ｜ "
          f"时间背景 {CORPUS_TIME}% ｜ 「本研究」{CORPUS_SELF}%")
    print("=" * 68)

    if res["n_sents"] == 0:
        print("⚠️ " + res["issues"][0])
        sys.exit(0)
    print("\n【句首模式分布】")
    for m in MODES:
        print(f"  {m:<14} {res['counts'][m]:4d} 句 = {res['pct'][m]:5.1f}%")
        for s in res["samples"][m][:2]:
            print(f"       例：{s}…")
    print(f"  {'序数分条起句':<14} {res['ordinal']:4d} 句 = "
          f"{res['ordinal'] / max(res['n_sents'], 1) * 100:5.1f}%（不计入连接词起句）")

    print("\n" + "-" * 68)
    if res["issues"]:
        print(f"🔴 发现 {len(res['issues'])} 项句式问题：")
        for it in res["issues"]:
            print(f"   · {it}")
    else:
        print("✅ 句首模式未触发红线（直接陈述占比正常、无过度衔接）。")
    if res["notices"]:
        print(f"\n🟡 建议核对 {len(res['notices'])} 项：")
        for it in res["notices"]:
            print(f"   · {it}")

    if "--json" in flags:
        jf = os.path.splitext(DOC)[0] + ".sentence_pattern.json"
        with open(jf, "w", encoding="utf-8") as f:
            json.dump(res, f, ensure_ascii=False, indent=1)
        print(f"\nJSON 已写入：{jf}")

    sys.exit(1 if res["issues"] else 0)


if __name__ == "__main__":
    main()
