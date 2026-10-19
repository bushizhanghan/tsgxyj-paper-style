# -*- coding: utf-8 -*-
"""le_density.py —— 「了」密度与属性陈述核验

language_audit 的 D 类查「了」**属性陈述堆积**（句级模式匹配），本脚本从
**密度分布**角度补一刀——该刊 118 篇 4743 段实测：

  · 全文「了」密度 26.60/万字；
  · 单段 ≥3 个「了」的段落占 7.8%（即该刊 92% 的段落「了」不超过 2 个）。

「了」在中文里有两种用法：**叙述过程**（「提取了 343 个主题词」——该刊正常）
与**属性陈述**（「系统提升了效率」——该刊倾向改「能够提升」）。本工具不区分
语义（那是 language_audit D 类的职责），只按**密度**判形态：

  · 单段「了」≥4 → 🔴 该刊 92% 段落 ≤2 个，4 个已属明显堆积；
  · 单段「了」=3 → 🟡 提示核对是否属性陈述；
  · 全文密度 >50/万字 → 🟡 整体偏高（该刊 26.6，P90 约 55）。

用法:
    python le_density.py <论文.docx> [--top 15] [--json]
退出码：1 = 检出单段 ≥4 个「了」。
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

# 该刊 118 篇 4743 段实测
CORPUS_LE_RATE = 26.60       # /万字
CORPUS_PARA_GE3 = 0.078      # 单段 ≥3 个「了」的段落占比
MIN_PARA = 40                # 参与统计的最短段
SEG_RED, SEG_YELLOW = 4, 3   # 单段「了」数红/黄线
DOC_LIMIT = 50.0             # 全文密度上限（/万字）


def analyze(doc):
    paras, refs, abs_text, kw_text, titles, body = dp.extract_paras(doc)
    total_le, total_chars, n_para = 0, 0, 0
    seg_rows = []
    for pi, raw in enumerate(body):
        t = dp.clean(raw or "")
        if len(t) < MIN_PARA:
            continue
        n_para += 1
        total_chars += len(t)
        c = t.count("了")
        total_le += c
        if c >= SEG_YELLOW:
            seg_rows.append({"para": pi + 1, "n": c, "chars": len(t),
                             "head": t[:44]})
    rate = total_le / max(total_chars, 1) * 10000

    issues, notices = [], []
    red = [r for r in seg_rows if r["n"] >= SEG_RED]
    yel = [r for r in seg_rows if r["n"] == SEG_YELLOW]
    for r in red:
        issues.append(f"段{r['para']} 含「了」×{r['n']}（≥{SEG_RED}）——"
                      f"该刊 92% 段落「了」≤2 个，此处明显堆积：{r['head']}…")
    if yel:
        notices.append(f"单段「了」=3 的段落 {len(yel)} 处（段"
                       + "、".join(str(r["para"]) for r in yel[:8])
                       + "）——核对是否属性陈述（「提升了效率」式），"
                         "该刊倾向改「能够提升」")
    if rate > DOC_LIMIT:
        notices.append(f"全文「了」密度 {rate:.1f}/万字，超过 {DOC_LIMIT:.0f}"
                       f"（该刊实测 {CORPUS_LE_RATE}）——整体偏高")

    return {"rate": round(rate, 2), "total_le": total_le,
            "n_para": n_para, "seg_rows": seg_rows, "red": red,
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
    TOP = 15
    for i, a in enumerate(sys.argv):
        if a == "--top" and i + 1 < len(sys.argv):
            try:
                TOP = int(sys.argv[i + 1])
            except ValueError:
                pass

    res = analyze(DOC)
    print("=" * 74)
    print("「了」密度与属性陈述核验（该刊 4743 段实测基线）")
    print(f"稿件：{os.path.basename(DOC)}")
    print(f"全文「了」{res['total_le']} 个 = {res['rate']}/万字"
          f"（该刊 {CORPUS_LE_RATE}/万字）｜ 统计 {res['n_para']} 段")
    print(f"单段≥3 个「了」的段落 {len(res['seg_rows'])} 处"
          f"（该刊同类占比 {CORPUS_PARA_GE3:.1%}）")
    print("=" * 74)

    if res["issues"]:
        print(f"🔴 发现 {len(res['issues'])} 项「了」堆积：")
        for it in res["issues"]:
            print(f"   · {it}")
    else:
        print("✅ 无单段「了」≥4 的堆积段。")
    if res["notices"]:
        print(f"\n🟡 建议核对 {len(res['notices'])} 项：")
        for it in res["notices"]:
            print(f"   · {it}")
    if res["seg_rows"]:
        print("\n【「了」集中段明细】")
        for r in res["seg_rows"][:TOP]:
            print(f"   段{r['para']}（{r['n']} 个/{r['chars']} 字）：{r['head']}…")

    if "--json" in flags:
        jf = os.path.splitext(DOC)[0] + ".le_density.json"
        with open(jf, "w", encoding="utf-8") as f:
            json.dump(res, f, ensure_ascii=False, indent=1)
        print(f"\nJSON 已写入：{jf}")

    sys.exit(1 if res["issues"] else 0)


if __name__ == "__main__":
    main()
