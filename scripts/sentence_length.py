# -*- coding: utf-8 -*-
"""sentence_length.py —— 句长分布核验（书面长句主体专查）

该刊句子实测（118 篇 18601 句）：P25=33、P50=55、P75=81、P90=118，
**40–90 字书面长句占 48.1%**（近半）。而 AI 生成稿与学生稿的典型问题是
「碎句化」——满篇 15–25 字短句，读起来像提纲，不像期刊论文。

diagnose_paper 只报「句均字数」一个平均值；平均值会掩盖分布问题：
句均 57 字可能由「半数 20 字碎句 + 半数 100 字长句」凑成，后者其实尚可，
前者必须改。本工具看**分布**：

  · 短句率：<30 字的句子占比（该刊 P25=33，即 25% 分位是 33 字；
    超过 45% 即为明显碎句化）；
  · 长句率：>120 字的句子占比（该刊 P90=118；超过 12% 说明臃肿）；
  · 40–90 字主体句占比（该刊 48.1%；低于 30% 说明句长结构偏离刊风）；
  · 定位：输出超限段落的段号与该段短句/长句数，便于定向合并或拆分。

用法:
    python sentence_length.py <论文.docx> [--top 20] [--json]
退出码：0 = 分布在容差内；1 = 短句率或长句率超限（提示级，不计红线）。

口径：阈值取该刊实测分位（P25/P90）上浮一档，见下方常量注释。
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

# 阈值（依据 118 篇 18601 句实测：P25=33 / P90=118 / 40–90 字占 48.1%）
SHORT_MAX = 30          # <30 字算短句
SHORT_RATE_LIMIT = 0.45  # 短句率上限（该刊约 25%，留 1.8 倍容差）
LONG_MIN = 120          # >120 字算长句
LONG_RATE_LIMIT = 0.12   # 长句率上限（该刊 P90=118，即约 10%）
MAIN_RANGE = (40, 90)   # 该刊主体句长区间
MAIN_RATE_LIMIT = 0.30  # 主体句占比下限（该刊 48.1%，容差后取 30%）
MIN_SENT = 6            # 参与统计的最短句


def sent_lens(paras):
    """返回 [(段号, [句长...])]，只统计 ≥40 字的正文段。"""
    out = []
    for pi, raw in enumerate(paras):
        t = dp.clean(raw or "")
        if len(t) < 40:
            continue
        sents = [s for s in re.split(r"(?<=[。？！])", t) if s.strip()]
        lens = [len(dp.clean(s)) for s in sents if len(dp.clean(s)) >= MIN_SENT]
        if lens:
            out.append((pi + 1, lens))
    return out


def analyze(doc):
    paras, refs, abs_text, kw_text, titles, body = dp.extract_paras(doc)
    per_para = sent_lens(body)
    all_lens = [x for _, ls in per_para for x in ls]
    n = len(all_lens)
    if not n:
        return {"n_sents": 0, "issues": [], "notices": [], "paras": []}

    short = [x for x in all_lens if x < SHORT_MAX]
    long_ = [x for x in all_lens if x > LONG_MIN]
    main = [x for x in all_lens if MAIN_RANGE[0] <= x <= MAIN_RANGE[1]]
    short_rate = len(short) / n
    long_rate = len(long_) / n
    main_rate = len(main) / n

    issues, notices = [], []
    if short_rate > SHORT_RATE_LIMIT:
        issues.append(f"短句率 {short_rate:.1%}（>{SHORT_RATE_LIMIT:.0%}）——"
                      f"<{SHORT_MAX} 字的句子占 {len(short)}/{n}，"
                      f"该刊 P25=33 字、40–90 字长句占 48%，"
                      f"碎句化会显得像提纲而非期刊论文，宜按段落合并分句")
    if long_rate > LONG_RATE_LIMIT:
        issues.append(f"长句率 {long_rate:.1%}（>{LONG_RATE_LIMIT:.0%}）——"
                      f">{LONG_MIN} 字的句子占 {len(long_)}/{n}，该刊 P90=118 字，"
                      f"臃肿句宜拆分为两句")
    if main_rate < MAIN_RATE_LIMIT and not issues:
        notices.append(f"主体句（{MAIN_RANGE[0]}–{MAIN_RANGE[1]} 字）占比 "
                       f"{main_rate:.1%}，低于该刊 48.1% 较多——句长结构偏离刊风")

    # 定位：短句最集中的段落（短句率 >60% 且短句数 ≥3）
    hotspots = []
    for pno, ls in per_para:
        ns = sum(1 for x in ls if x < SHORT_MAX)
        if len(ls) >= 3 and ns / len(ls) > 0.6:
            hotspots.append({"para": pno, "short": ns, "total": len(ls),
                             "lens": sorted(ls)[:8]})
    hotspots.sort(key=lambda h: -h["short"])
    if hotspots and short_rate > SHORT_RATE_LIMIT:
        notices.append(f"短句集中段 {len(hotspots)} 个（段号见下方明细）")

    return {"n_sents": n, "short_rate": round(short_rate, 3),
            "long_rate": round(long_rate, 3), "main_rate": round(main_rate, 3),
            "issues": issues, "notices": notices, "hotspots": hotspots,
            "paras": [{"para": p, "n": len(ls),
                       "mean": round(sum(ls) / len(ls), 1)} for p, ls in per_para]}


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
    for i, a in enumerate(sys.argv):
        if a == "--top" and i + 1 < len(sys.argv):
            try:
                TOP = int(sys.argv[i + 1])
            except ValueError:
                pass

    res = analyze(DOC)
    print("=" * 74)
    print("句长分布核验（书面长句主体专查）")
    print(f"稿件：{os.path.basename(DOC)}")
    if not res["n_sents"]:
        print("未检出可统计句子。")
        sys.exit(0)
    print(f"统计 {res['n_sents']} 句")
    print(f"短句率（<{SHORT_MAX}字）{res['short_rate']:.1%} ｜ "
          f"长句率（>{LONG_MIN}字）{res['long_rate']:.1%} ｜ "
          f"主体句（{MAIN_RANGE[0]}–{MAIN_RANGE[1]}字）{res['main_rate']:.1%}")
    print(f"该刊基线：短句约 25% ｜ 长句约 10% ｜ 主体句 48.1%")
    print("=" * 74)

    if res["issues"]:
        print(f"🔴 发现 {len(res['issues'])} 项分布问题：")
        for it in res["issues"]:
            print(f"   · {it}")
    else:
        print("✅ 句长分布在该刊容差范围内。")
    if res["notices"]:
        print(f"\n🟡 建议核对 {len(res['notices'])} 项：")
        for it in res["notices"]:
            print(f"   · {it}")
    if res["hotspots"]:
        print("\n【短句集中段】（该段短句占比 >60%，宜合并分句）")
        for h in res["hotspots"][:TOP]:
            print(f"   段{h['para']}：短句 {h['short']}/{h['total']}，"
                  f"各句长 {h['lens']}")

    if "--json" in flags:
        jf = os.path.splitext(DOC)[0] + ".sentence_length.json"
        with open(jf, "w", encoding="utf-8") as f:
            json.dump(res, f, ensure_ascii=False, indent=1)
        print(f"\nJSON 已写入：{jf}")

    sys.exit(1 if res["issues"] else 0)


if __name__ == "__main__":
    main()
