# -*- coding: utf-8 -*-
"""conn_density.py —— 连接词密度核验（逐词对照 + 空转连接专项）

用途
----
 diagnose_paper.py ③b 已给出单篇连接词频率汇总（次/千字，对照 style_stats
 .md 静态基线），本脚本做两件它没做的事：

  1. **逐词对照**：对该刊实测的 8 个高频连接词，逐个给出
     「该刊次/千字 → 本文次/千字 → 倍数」，按 >2 倍判偏高、<0.4 倍判偏低；
     并对每个偏高的词列出**出现位置（段号 + 原句片段）**，可直接定位改写。
  2. **空转连接专项**：「综上 / 在此基础上 / 与此同时」这类不承载新信息、
     只起「段间润滑」作用的连接，该刊 121 万字合计仅 66+36+23 = 125 处
     （anti-patterns.md §B 级：一篇论文用 2–3 处正常）。AI 稿常在每段
     开头都挂一个，本脚本按「每千字出现次数」与「绝对处数」双口径判定。

语料依据（corpus_stats.json 的 conn 字段，108 篇 body_chars 合计
1,096,058 字实测换算；--corpus-baseline 可复跑）：
    同时 0.55 ｜ 因此 0.35 ｜ 一方面 0.28 ｜ 从而 0.22
    此外 0.19 ｜ 另一方面 0.14 ｜ 进而 0.11 ｜ 然而 0.08    （单位：次/千字）

用法:
    python conn_density.py <论文.docx> [--corpus <dir>] [--json] [--top 15]
    python conn_density.py --corpus-baseline        # 打印该刊连接词实测基线

退出码：0 = 无偏高/偏低词；1 = 检出 >2 倍偏高的连接词或空转连接超标。
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

# 该刊实测基线（次/千字）：corpus_stats.json conn 字段 108 篇汇总换算
CORPUS_CONN = {
    "同时": 0.55, "因此": 0.35, "一方面": 0.28, "从而": 0.22,
    "此外": 0.19, "另一方面": 0.14, "进而": 0.11, "然而": 0.08,
}
CONN_ORDER = ["同时", "因此", "一方面", "从而", "此外", "另一方面", "进而", "然而"]

# 空转连接：anti-patterns.md §B 级（121 万字合计 66+36+23 处）
# 每篇论文 2–3 处正常；按千字基准 125 处/121 万字 = 0.103/千字。
IDLE_CONN = ["综上", "在此基础上", "与此同时"]
IDLE_CORPUS_PER_1K = 0.103     # 次/千字（125 处 / 121 万字）
IDLE_ABS_HINT = 3              # 单篇超过此数即提示（anti-patterns.md 口径）

# 判级阈值（任务给定）
HIGH_MULT = 2.0    # >2 倍该刊基线 → 偏高
LOW_MULT = 0.4     # <0.4 倍该刊基线 → 偏低
MIN_PARA = 60      # 参与统计的最短段（clean 后）
MIN_SAMPLE_CHARS = 800   # 全文不足此字数时不判密度（短稿偶然性大）

CACHE = {}


def corpus_baseline(corpus_dir=None):
    """从 corpus_stats.json 实测该刊连接词密度（次/千字）。"""
    if "base" in CACHE:
        return CACHE["base"]
    cdir = corpus.find_corpus_dir(corpus_dir)
    if not cdir:
        return None
    rows = corpus.load_corpus_stats(cdir)
    total = sum(r.get("body_chars", 0) for r in rows)
    cnt = Counter()
    for r in rows:
        for k, v in (r.get("conn") or {}).items():
            cnt[k] += v
    out = {"n_papers": len(rows), "chars": total,
           "conn": {k: cnt[k] / total * 1000 for k, v in sorted(cnt.items(), key=lambda x: -x[1]) if v}}
    CACHE["base"] = out
    return out


def _find_locations(body, word, limit=3):
    """返回该词出现的 (段号, 片段) 列表。"""
    hits = []
    for pi, b in enumerate(body):
        t = dp.clean(b or "")
        if len(t) < MIN_PARA:
            continue
        for m in re.finditer(re.escape(word), t):
            s = max(0, m.start() - 14)
            hits.append((pi + 1, t[s:m.end() + 18]))
            break                      # 每段最多记一条，避免刷屏
        if len(hits) >= limit:
            break
    return hits


def analyze(doc, corpus_dir=None):
    """逐词对照 + 空转连接专项。返回 dict。"""
    paras, refs, abs_text, kw_text, titles, body = dp.extract_paras(doc)
    text = dp.clean(abs_text or "") + "".join(dp.clean(b or "") for b in body)
    n = len(text)

    rows, issues, notices = [], [], []
    for w in CONN_ORDER:
        c = text.count(w)
        rate = c / max(n, 1) * 1000
        base = CORPUS_CONN[w]
        mult = rate / base if base else 0.0
        tag = "≈"
        if mult > HIGH_MULT:
            tag = "▲ 偏高"
        elif mult < LOW_MULT:
            tag = "▼ 偏低"
        rows.append({"word": w, "n": c, "rate": round(rate, 3), "base": base,
                     "mult": round(mult, 2), "tag": tag,
                     "locs": _find_locations(body, w) if c else []})
        if tag == "▲ 偏高":
            issues.append(
                f"「{w}」{c} 次 = {rate:.2f}/千字，为该刊基线（{base}）的 "
                f"{mult:.1f} 倍——每处都核一下是否必要，能删的删")
        elif tag == "▼ 偏低" and n >= MIN_SAMPLE_CHARS:
            notices.append(
                f"「{w}」{c} 次 = {rate:.2f}/千字，仅为该刊基线的 {mult:.2f} 倍"
                f"——若稿中确有该类逻辑关系，考虑补上")

    # 空转连接专项
    idle = []
    for w in IDLE_CONN:
        c = text.count(w)
        idle.append({"word": w, "n": c, "rate": round(c / max(n, 1) * 1000, 3),
                     "base": IDLE_CORPUS_PER_1K,
                     "locs": _find_locations(body, w, limit=5) if c else []})
    idle_total = sum(x["n"] for x in idle)
    idle_rate = idle_total / max(n, 1) * 1000
    if idle_rate > IDLE_CORPUS_PER_1K * 2 or idle_total > IDLE_ABS_HINT:
        detail = "、".join(f"{x['word']}×{x['n']}" for x in idle if x["n"])
        issues.append(
            f"空转连接「综上/在此基础上/与此同时」合计 {idle_total} 处"
            f"（{idle_rate:.2f}/千字，该刊 {IDLE_CORPUS_PER_1K}/千字）——"
            f"一篇论文用 2–3 处正常；{detail}。逐处核对是否只是「段间润滑」，"
            f"是则删（该刊 63% 段落不用任何段首连接）")

    return {"chars": n, "conn_rows": rows, "idle": idle,
            "idle_total": idle_total, "idle_rate": round(idle_rate, 3),
            "issues": issues, "notices": notices}


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
        base = corpus_baseline(cli_corpus)
        if not base:
            print("[错误] 未找到语料库（需要 corpus_stats.json / paragraphs.json）")
            sys.exit(2)
        print("=" * 66)
        print("该刊连接词密度实测基线（corpus_stats.json conn 字段换算）")
        print(f"样本：{base['n_papers']} 篇 / 正文 {base['chars']} 字")
        print("=" * 66)
        for k, v in base["conn"].items():
            print(f"  {k}: {v:.3f} 次/千字")
        sys.exit(0)

    DOC = args[0]
    if not os.path.isfile(DOC):
        sys.exit(f"找不到文件：{DOC}")

    res = analyze(DOC, cli_corpus)
    print("=" * 74)
    print("连接词密度核验（逐词对照 + 空转连接专项）")
    print(f"稿件：{os.path.basename(DOC)}")
    print(f"统计口径：摘要+正文共 {res['chars']} 字（clean 后）")
    print("=" * 74)

    print("\n【① 逐词对照】单位：次/千字")
    print(f"  {'连接词':<6}{'本文':>6}{'该刊':>7}{'倍数':>7}  判定")
    for r in res["conn_rows"]:
        print(f"  {r['word']:<6}{r['rate']:>6.2f}{r['base']:>7.2f}"
              f"{r['mult']:>7.2f}  {r['tag']}")
        if r["locs"] and r["tag"] == "▲ 偏高":
            for pi, frag in r["locs"][:2]:
                print(f"        段{pi}：…{frag}…")

    print("\n【② 空转连接专项】该刊 121 万字合计 66+36+23 = 125 处"
          f"（{IDLE_CORPUS_PER_1K}/千字），一篇论文 2–3 处正常")
    for x in res["idle"]:
        print(f"  {x['word']}: {x['n']} 处 = {x['rate']:.3f}/千字"
              f"（该刊 {x['base']}）")
        for pi, frag in x["locs"][:3]:
            print(f"      段{pi}：…{frag}…")
    print(f"  合计 {res['idle_total']} 处 = {res['idle_rate']:.3f}/千字")

    print("\n" + "-" * 74)
    if res["issues"]:
        print(f"🔴 发现 {len(res['issues'])} 项连接词问题：")
        for it in res["issues"]:
            print(f"   · {it}")
    else:
        print("✅ 连接词密度正常（无 >2 倍偏高词，空转连接未超标）。")
    if res["notices"]:
        print(f"\n🟡 建议核对 {len(res['notices'])} 项：")
        for it in res["notices"]:
            print(f"   · {it}")

    if "--json" in flags:
        jf = os.path.splitext(DOC)[0] + ".conn_density.json"
        with open(jf, "w", encoding="utf-8") as f:
            json.dump(res, f, ensure_ascii=False, indent=1)
        print(f"\nJSON 已写入：{jf}")

    sys.exit(1 if res["issues"] else 0)


if __name__ == "__main__":
    main()
