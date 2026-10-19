# -*- coding: utf-8 -*-
"""paragraph_opening.py —— 段首词核验（对照该刊实测高频段首）

用途
----
核验稿件**每段开头**用词是否贴合《图书馆学研究》的段首习惯。该刊段首并
非任意：118 篇 4743 段实测，段首高频词高度集中在「自称起句（本研究/本文）、
回指起句（基于上述/基于此/针对上述）、序数分条（一是/第一/首先）、时间起句
（近年来/目前/当前）、转折递进（然而/此外/同时）」几类，且强过渡词起句
（此外/然而/同时/另一方面/与此同时）合计仅 45 段（≈0.95%）。

**词表构建方法（本脚本的语料依据，可由 --corpus-baseline 复跑）**：
    1. 读语料库 paragraphs.json（118 篇，`corpus.load_paragraphs()`）；
    2. 每段文本经 `dp.clean()` 去空白后，取**前 4 个字符**并去掉尾部标点
       （「一方面,」→「一方面」），作为该段的「段首词」；
    3. 只统计 clean 后段长 ≥60 字的段落（4743 段全部满足），用 Counter
       统计全部 3603 个不同段首词的频次；
    4. 取**频次最高的前 50 个**作为「该刊常见段首」词表 OPEN_TOP50
       （第 50 名频次 5，第 51 名「图书馆」频次 7——注意长尾极长，
        top50 只覆盖 15.8% 的段落，因此**不能**用「覆盖率」判偏离）。
    词表是**语料实测**的，不是人工拟定的；重建只需重跑 --corpus-baseline。

判定逻辑（为什么不用覆盖率）：
    该刊自身也只有 11.8% 的段落段首命中 top50（学科名词、机构名、引文
    作者名起句占绝大多数），所以「不在表内」本身不是问题。真正偏离该刊
    体例的是**强过渡词起句扎堆**——「此外/然而/同时/另一方面/与此同时/
    综上所述/总之/除此之外/相较之下」起句的段落，该刊合计 75/4743 段
    ≈ **1.58%**，而 AI 稿常把它们当段首默认词。故本脚本：
      · 🔴 强过渡词起句占比 >5%（约该刊 3 倍）→ 判「段首过度衔接」；
      · 🟡 占比 2%–5% → 提示核对；
      · 同时列出稿件段首词中不在词表的高频项与强过渡词逐段位置，便于定位。

用法:
    python paragraph_opening.py <论文.docx> [--corpus <dir>] [--json] [--top 15]
    python paragraph_opening.py --corpus-baseline    # 打印该刊 top50 词表与基线

退出码：0 = 段首形态正常；1 = 强过渡词起句占比 >5%（过度衔接）。
"""
import os
import re
import sys
import json
from collections import Counter

# Windows 控制台 GBK 保护
if hasattr(sys.stdout, "reconfigure"):
    try:
        sys.stdout.reconfigure(encoding="utf-8", errors="replace")
    except Exception:
        pass

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import corpus  # noqa: E402
import diagnose_paper as dp  # noqa: E402

HEAD_N = 4                 # 段首取前 N 字（词表构建口径）
MIN_PARA = 60              # 参与统计的最短段（clean 后）
TOP_K = 50                 # 词表规模

# 强过渡词起句：段首即衔接词，几乎不承载新信息（该刊极少这样起句）
STRONG_TRANS = ["此外", "然而", "同时", "另一方面", "与此同时", "总之",
                "综上所述", "由此可见", "换句话说", "除此之外", "相较之下",
                "另一方面", "但是", "不过"]

# 该刊实测基线（--corpus-baseline 可复跑）
CORPUS_N_PARAS = 4743
CORPUS_STRONG_TRANS_RATIO = 1.58   # %（75/4743）
CORPUS_TOP50_COVER = 11.8          # %（top50 词表对全部段落的覆盖率）

# 判级阈值
RED_RATIO = 5.0      # 强过渡词起句占比红线（%，约该刊 5 倍）
YELLOW_RATIO = 2.0   # 提示线

CACHE = {}


def head_word(text):
    """段首词：clean 后前 HEAD_N 字，去掉尾部标点。"""
    t = dp.clean(text or "")
    if len(t) < HEAD_N:
        return t
    return t[:HEAD_N].rstrip("，,、。；：）)」』")


def _strip_punct(s):
    return (s or "").rstrip("，,、。；：）)」』")


def build_open_table(corpus_dir=None):
    """从 paragraphs.json 构建「该刊常见段首」词表。

    返回 {"table": [(词, 频次)...top50], "n_paras": 段数, "cover": 覆盖率%,
          "strong": {词: 段数}, "strong_ratio": %}
    """
    if "table" in CACHE:
        return CACHE["table"]
    cdir = corpus.find_corpus_dir(corpus_dir)
    if not cdir:
        return None
    papers = corpus.load_paragraphs(cdir)
    cnt = Counter()
    strong = Counter()
    n = 0
    for paper in papers:
        for p in paper.get("paras") or []:
            t = dp.clean(p)
            if len(t) < MIN_PARA:
                continue
            n += 1
            cnt[head_word(p)] += 1
            for w in STRONG_TRANS:
                if t.startswith(w):
                    strong[w] += 1
                    break
    top = cnt.most_common(TOP_K)
    out = {
        "table": top,
        "n_paras": n,
        "papers": len(papers),
        "cover": sum(v for _, v in top) / max(n, 1) * 100,
        "strong": dict(strong),
        "strong_ratio": sum(strong.values()) / max(n, 1) * 100,
        "distinct": len(cnt),
    }
    CACHE["table"] = out
    return out


def analyze(doc, corpus_dir=None):
    """核验稿件段首词。返回 dict。"""
    tbl = build_open_table(corpus_dir)
    known = {w for w, _ in (tbl["table"] if tbl else [])}

    paras, refs, abs_text, kw_text, titles, body = dp.extract_paras(doc)
    rows = []          # (段号, 段首词, 段首片段, 命中/强过渡/表外)
    cnt_in = cnt_out = 0
    strong_rows = []
    for pi, b in enumerate(body):
        t = dp.clean(b or "")
        if len(t) < MIN_PARA:
            continue
        hw = head_word(t)
        if hw in known:
            cnt_in += 1
            rows.append((pi + 1, hw, t[:30], "命中"))
            continue
        sw = next((w for w in STRONG_TRANS if t.startswith(w)), None)
        if sw:
            cnt_out += 1
            strong_rows.append((pi + 1, sw, t[:46]))
            rows.append((pi + 1, hw, t[:30], f"强过渡「{sw}」"))
        else:
            cnt_out += 1
            rows.append((pi + 1, hw, t[:30], "表外"))

    n = cnt_in + cnt_out
    strong_ratio = len(strong_rows) / max(n, 1) * 100
    # 表外高频段首词（供人工核对是否成习惯）
    out_words = Counter(hw for _, hw, _, tag in rows if tag != "命中")

    issues, notices = [], []
    if n == 0:
        return {"n_paras": 0, "issues": ["正文过短，无法统计段首词"],
                "notices": [], "rows": rows}
    if strong_ratio > RED_RATIO:
        issues.append(
            f"强过渡词起句 {len(strong_rows)}/{n} 段（{strong_ratio:.1f}%）——"
            f"该刊实测仅 {CORPUS_STRONG_TRANS_RATIO}%（75/4743 段）；"
            f"段首默认挂「此外/然而/同时」是过度衔接形态，"
            f"改为直接陈述或用「基于上述/针对上述」式回指")
    elif strong_ratio >= YELLOW_RATIO:
        notices.append(
            f"强过渡词起句 {len(strong_rows)}/{n} 段（{strong_ratio:.1f}%）——"
            f"已高于该刊 {CORPUS_STRONG_TRANS_RATIO}%，建议压缩")
    if tbl and cnt_in / max(n, 1) * 100 < CORPUS_TOP50_COVER * 0.5:
        notices.append(
            f"段首命中该刊 top50 词表 {cnt_in}/{n} 段"
            f"（{cnt_in / max(n, 1) * 100:.1f}%）——低于该刊覆盖率"
            f"（{CORPUS_TOP50_COVER}%）的一半；不代表一定有问题"
            f"（该刊段首以学科名词/引文作者名为主），但可对照下表核对")

    return {"n_paras": n, "in_table": cnt_in, "out_table": cnt_out,
            "strong_rows": strong_rows, "strong_ratio": round(strong_ratio, 2),
            "out_words": out_words.most_common(12), "rows": rows,
            "table": tbl, "issues": issues, "notices": notices}


def main():
    args = [a for a in sys.argv[1:] if not a.startswith("--")]
    flags = set(a for a in sys.argv[1:] if a.startswith("--"))
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
        tbl = build_open_table(cli_corpus)
        if not tbl:
            print("[错误] 未找到语料库（需要 corpus_stats.json / paragraphs.json）")
            sys.exit(2)
        print("=" * 68)
        print("该刊常见段首词表（paragraphs.json 实测，取每段前 4 字 top 50）")
        print(f"样本：{tbl['papers']} 篇 / {tbl['n_paras']} 段"
              f"（不同段首词 {tbl['distinct']} 个）")
        print(f"top50 覆盖率 {tbl['cover']:.1f}% ｜ 强过渡词起句 "
              f"{tbl['strong_ratio']:.2f}%")
        print("=" * 68)
        for i, (w, v) in enumerate(tbl["table"], 1):
            print(f"  {i:2d}. {w}({v})", end="\n" if i % 5 == 0 else "  ")
        print("\n\n强过渡词起句（该刊逐词段数）：")
        for w, v in sorted(tbl["strong"].items(), key=lambda x: -x[1]):
            print(f"  {w}: {v} 段")
        sys.exit(0)

    DOC = args[0]
    if not os.path.isfile(DOC):
        sys.exit(f"找不到文件：{DOC}")

    res = analyze(DOC, cli_corpus)
    tbl = res.get("table")
    print("=" * 68)
    print("段首词核验（对照该刊 4743 段实测高频段首）")
    print(f"稿件：{os.path.basename(DOC)}")
    print(f"统计：{res['n_paras']} 段（clean 后段长 ≥60）")
    print(f"该刊基线：top50 词表覆盖率 {CORPUS_TOP50_COVER}% ｜ "
          f"强过渡词起句 {CORPUS_STRONG_TRANS_RATIO}%")
    if tbl:
        print(f"词表来源：paragraphs.json 实测 top50（第 50 名频次 "
              f"{tbl['table'][-1][1]}）")
    print("=" * 68)

    if res["n_paras"] == 0:
        print("⚠️ " + res["issues"][0])
        sys.exit(0)

    print(f"\n【段首命中】命中该刊 top50 词表 {res['in_table']} 段 ｜ "
          f"表外 {res['out_table']} 段")
    print(f"【强过渡词起句】{len(res['strong_rows'])} 段 = "
          f"{res['strong_ratio']:.1f}%（该刊 {CORPUS_STRONG_TRANS_RATIO}%）")
    for pi, sw, frag in res["strong_rows"][:TOP]:
        print(f"   段{pi}「{sw}」：{frag}…")

    print("\n【稿件高频表外段首词】（人工核对是否成习惯）")
    print("  " + "  ".join(f"{w}({v})" for w, v in res["out_words"]))

    print("\n" + "-" * 68)
    if res["issues"]:
        print(f"🔴 发现 {len(res['issues'])} 项段首问题：")
        for it in res["issues"]:
            print(f"   · {it}")
    else:
        print("✅ 段首形态正常（无强过渡词起句扎堆）。")
    if res["notices"]:
        print(f"\n🟡 建议核对 {len(res['notices'])} 项：")
        for it in res["notices"]:
            print(f"   · {it}")

    if "--json" in flags:
        jf = os.path.splitext(DOC)[0] + ".paragraph_opening.json"
        with open(jf, "w", encoding="utf-8") as f:
            json.dump({k: v for k, v in res.items() if k not in ("table", "rows")},
                      f, ensure_ascii=False, indent=1)
        print(f"\nJSON 已写入：{jf}")

    sys.exit(1 if res["issues"] else 0)


if __name__ == "__main__":
    main()
