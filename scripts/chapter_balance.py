# -*- coding: utf-8 -*-
"""chapter_balance.py —— 单章字数失衡核验（结构极值专查）

diagnose_paper 的章节占比对照 SEC_BASE 给「偏高/偏低」标签，但只对
**已知类别的标配章**（引言/对策/结语）作比较，不判**单章极值**——而
「一章占掉半篇文章」或「某章只写了三行」是结构失衡最直观的信号：

  · 单章占正文 >40% → 🔴 严重失衡（该刊单章占比实测 P90 约 30%）；
  · 单章占正文 <2%（且 >0）→ 🟡 疑似残章（该刊单章 P10 约 1.5%，
    低于此多为「挂名章」——有标题无实质内容）；
  · 同篇内最大/最小章字数比 >25 → 🟡 章间严重不均。

本脚本输出逐章占比表 + 极值定位，并附该刊单章占比分位作参照。

用法:
    python chapter_balance.py <论文.docx> [--json]
退出码：1 = 检出单章 >40%（严重失衡）。
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

# 该刊单章占比分位（108 篇实测，按章展开）
CH_P50, CH_P90 = 0.13, 0.30
HUGE = 0.40             # >40% 判严重失衡
TINY = 0.02             # <2% 判疑似残章
RATIO_LIMIT = 25        # 最大/最小章字数比


def analyze(doc):
    paras, refs, abs_text, kw_text, titles, body = dp.extract_paras(doc)
    secs = dp.scan_sections(paras, titles)
    if not secs or not secs.get("secs"):
        return {"issues": ["未检出章节结构"], "notices": [], "rows": []}

    rows = []
    for s in secs["secs"]:
        rows.append({"name": s["name"], "chars": s["chars"],
                     "pct": round(s["pct"], 2), "cls": s.get("cls")})
    total = secs.get("total") or sum(r["chars"] for r in rows)

    issues, notices = [], []
    for r in rows:
        p = r["pct"] / 100.0
        if p > HUGE:
            issues.append(f"「{r['name'][:20]}」占正文 {r['pct']:.1f}%（>{HUGE:.0%}）——"
                          f"单章严重失衡，该刊单章 P90 约 {CH_P90:.0%}；"
                          f"宜拆分或把论证分摊到相关章")
        elif p < TINY and r["chars"] > 0:
            notices.append(f"「{r['name'][:20]}」仅 {r['chars']} 字"
                           f"（{r['pct']:.1f}%）——疑似挂名章（有标题无实质内容），"
                           f"宜补充或合并")

    # 章间极值比
    real = [r for r in rows if r["chars"] > 0]
    if len(real) >= 3:
        mx, mn = max(r["chars"] for r in real), min(r["chars"] for r in real)
        ratio = mx / mn if mn else 0
        if ratio > RATIO_LIMIT:
            notices.append(f"章间字数比 {ratio:.0f}:1（最大 {mx} 字 vs 最小 {mn} 字），"
                           f"超过 {RATIO_LIMIT}:1——结构严重不均，考虑重新分配论证篇幅")

    return {"rows": rows, "total": total,
            "issues": issues, "notices": notices}


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
    print("=" * 74)
    print("单章字数失衡核验（结构极值专查）")
    print(f"稿件：{os.path.basename(DOC)}")
    print(f"正文合计 {res.get('total', 0)} 字 ｜ 该刊单章占比 P50≈{CH_P50:.0%}、"
          f"P90≈{CH_P90:.0%}")
    print("=" * 74)

    if res["rows"]:
        print("【逐章占比】")
        for r in sorted(res["rows"], key=lambda x: -x["pct"]):
            bar = "█" * max(1, int(r["pct"] / 3))
            print(f"  {r['name'][:22]:<24}{r['chars']:>6}字 {r['pct']:>6.1f}%  {bar}")

    print("\n" + "-" * 74)
    if res["issues"]:
        print(f"🔴 发现 {len(res['issues'])} 项严重失衡：")
        for it in res["issues"]:
            print(f"   · {it}")
    else:
        print("✅ 无单章严重失衡。")
    if res["notices"]:
        print(f"\n🟡 建议核对 {len(res['notices'])} 项：")
        for it in res["notices"]:
            print(f"   · {it}")

    if "--json" in flags:
        jf = os.path.splitext(DOC)[0] + ".chapter_balance.json"
        with open(jf, "w", encoding="utf-8") as f:
            json.dump(res, f, ensure_ascii=False, indent=1)
        print(f"\nJSON 已写入：{jf}")

    sys.exit(1 if res["issues"] else 0)


if __name__ == "__main__":
    main()
