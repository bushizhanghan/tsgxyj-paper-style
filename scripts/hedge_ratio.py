# -*- coding: utf-8 -*-
"""hedge_ratio.py —— 审慎词与强断言词比率核验

polish-lexicon.md §2 记录了该刊的程度表达分布（118 篇 121 万字实测）：
**审慎/限定词 1427 次 vs 强断言词 1303 次——审慎占 52.3%**。
该刊主流是「可能/相对/逐步/通常/多数/一般/往往」式审慎表达；
「显著/全面/充分/一定/完全/极大/必然/绝对」式强断言虽也在用
（显著 461 次是该刊最高频强断言），但**无数据支撑时**用就会被评为
「形容词通胀」（anti-patterns.md B 级）。

language_audit 的 C 类查「无支撑强断言」的**句子模式**，本脚本从**词频比率**
角度补一刀：

  · 强断言占比 >60% → 🟡 明显偏离该刊（52.3%），检查是否有无数据支撑的强断言；
  · 强断言占比 >75% → 🔴 严重失衡，学术写作的谨慎性不足；
  · 逐词列出双方频次，供对照 polish-lexicon 调整。

用法:
    python hedge_ratio.py <论文.docx> [--json]
退出码：1 = 强断言占比 >75%。
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
import diagnose_paper as dp  # noqa: E402

# 该刊 121 万字实测（polish-lexicon.md §2）
CAUTIOUS = ["可能", "相对", "逐步", "通常", "多数", "一般", "往往", "基本",
            "大致", "一定程度上", "某种程度上", "约", "近", "趋于"]
STRONG = ["显著", "全面", "充分", "一定", "完全", "极大", "必然", "绝对",
          "彻底", "完全", "极大", "完全"]
STRONG = list(dict.fromkeys(STRONG))     # 去重保序
CORPUS_CAUTIOUS_TOTAL = 1427
CORPUS_STRONG_TOTAL = 1303
CORPUS_RATIO = 0.523                     # 审慎占比
YELLOW_LIMIT, RED_LIMIT = 0.60, 0.75     # 强断言占比黄/红线


def analyze(doc):
    paras, refs, abs_text, kw_text, titles, body = dp.extract_paras(doc)
    text = dp.clean(abs_text or "") + dp.clean("\n".join(body))
    if len(text) < 100:
        return {"issues": [], "notices": ["正文过短，无法统计"]}

    c_cnt = Counter()
    s_cnt = Counter()
    for w in CAUTIOUS:
        c = text.count(w)
        if c:
            c_cnt[w] = c
    for w in STRONG:
        c = text.count(w)
        if c:
            s_cnt[w] = c

    tc, ts = sum(c_cnt.values()), sum(s_cnt.values())
    total = tc + ts
    strong_ratio = ts / total if total else 0.0

    issues, notices = [], []
    MIN_SAMPLE = 5           # 最少样本量：低于此不判（防短稿偶然命中）
    if total >= MIN_SAMPLE and strong_ratio > RED_LIMIT:
        issues.append(f"强断言词占比 {strong_ratio:.0%}（>{RED_LIMIT:.0%}）——"
                      f"审慎 {tc} 次 vs 强断言 {ts} 次；该刊实测审慎占 "
                      f"{CORPUS_RATIO:.1%}，学术写作的谨慎性严重不足")
    elif total >= MIN_SAMPLE and strong_ratio > YELLOW_LIMIT:
        notices.append(f"强断言词占比 {strong_ratio:.0%}（>{YELLOW_LIMIT:.0%}）——"
                       f"审慎 {tc} 次 vs 强断言 {ts} 次，偏离该刊（审慎占 "
                       f"{CORPUS_RATIO:.1%}）；请核对强断言是否都有数据/文献支撑"
                       f"（无支撑的按 anti-patterns.md 改「有所/一定程度上/相对」）")

    return {"cautious": dict(c_cnt), "strong": dict(s_cnt),
            "tc": tc, "ts": ts, "ratio": round(strong_ratio, 3),
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
    print("审慎词与强断言词比率核验（该刊 121 万字实测基线）")
    print(f"稿件：{os.path.basename(DOC)}")
    print(f"审慎词 {res.get('tc', 0)} 次 ｜ 强断言词 {res.get('ts', 0)} 次 ｜ "
          f"强断言占比 {res.get('ratio', 0):.1%}")
    print(f"该刊基线：审慎 {CORPUS_CAUTIOUS_TOTAL} 次 vs 强断言 "
          f"{CORPUS_STRONG_TOTAL} 次（审慎占 {CORPUS_RATIO:.1%}）")
    print("=" * 74)

    print("【审慎词】" + ("、".join(f"{k}×{v}" for k, v in
                                  sorted(res["cautious"].items(),
                                         key=lambda x: -x[1])) if res.get("cautious") else "（无）"))
    print("【强断言词】" + ("、".join(f"{k}×{v}" for k, v in
                                    sorted(res["strong"].items(),
                                           key=lambda x: -x[1])) if res.get("strong") else "（无）"))

    print("\n" + "-" * 74)
    if res["issues"]:
        print(f"🔴 发现 {len(res['issues'])} 项问题：")
        for it in res["issues"]:
            print(f"   · {it}")
    else:
        print("✅ 程度词比率未严重失衡。")
    if res["notices"]:
        print(f"\n🟡 建议核对 {len(res['notices'])} 项：")
        for it in res["notices"]:
            print(f"   · {it}")

    if "--json" in flags:
        jf = os.path.splitext(DOC)[0] + ".hedge_ratio.json"
        with open(jf, "w", encoding="utf-8") as f:
            json.dump(res, f, ensure_ascii=False, indent=1)
        print(f"\nJSON 已写入：{jf}")

    sys.exit(1 if res["issues"] else 0)


if __name__ == "__main__":
    main()
