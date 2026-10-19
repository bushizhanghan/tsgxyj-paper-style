# -*- coding: utf-8 -*-
"""structure_scale.py —— 篇章规模核验（章节数/段落数/正文字数）

diagnose_paper 报「正文 11711 字｜段 62｜标题 21 个」的原始值，但不与
该刊分布对照判定。而**规模失配**是初审最直观的筛选项之一：

  · 一级章节数（该刊 108 篇：P25=1、P50=4、P75=5、P90=6）——
    只有 1–2 章的多半是「论述单薄」或「结构未展开」；超过 7 章则偏碎；
  · 正文段落数（P25=33、P50=39、P75=47、P90=53）——段落数明显偏少
    常意味着长段堆叠（该刊 2–6 句/段），偏多则意味着碎段化；
  · 正文字数（投稿者实测集中 10000–15000 字，见 submission-reality.md）——
    低于 8000 字显单薄，超过 18000 字超出该刊单篇容纳惯例。

本工具把这三项与该刊实测分位对照，给出「规模定位」与偏离提示。

用法:
    python structure_scale.py <论文.docx> [--corpus <dir>] [--json]
退出码：0 = 规模在常规区间；1 = 明显偏离（章节过少/字数越界）。
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

# 该刊 108 篇实测分位
SEC_P25, SEC_P50, SEC_P75, SEC_P90 = 1, 4, 5, 6
PARA_P25, PARA_P50, PARA_P75, PARA_P90 = 33, 39, 47, 53
# 正文字数：投稿者实测集中区间（submission-reality.md）
CHARS_MIN, CHARS_MAX = 8000, 18000


def count_h1(titles):
    n = 0
    for t in titles:
        s = re.sub(r"\s+", " ", (t or "").strip())
        if re.match(r"^\d+\s*[^\d.]", s) or re.match(r"^\s*(引\s*言|结\s*语)\s*$", s):
            n += 1
    return n


def analyze(doc):
    paras, refs, abs_text, kw_text, titles, body = dp.extract_paras(doc)
    m = dp.compute_metrics(body, refs, abs_text, kw_text, titles)
    n_h1 = count_h1(titles)
    n_para = m["para_n"]
    n_chars = m["body_chars"]

    issues, notices = [], []
    if n_h1 < SEC_P25:
        issues.append(f"一级章节仅 {n_h1} 个（该刊 P25={SEC_P25}、P50={SEC_P50}）——"
                      f"结构未展开，初审易判「论述单薄」；可将并列内容拆为独立章")
    elif n_h1 > SEC_P90:
        notices.append(f"一级章节 {n_h1} 个，超过该刊 P90（{SEC_P90}）——"
                       f"结构偏碎，考虑合并相关章")
    if n_para < PARA_P25:
        notices.append(f"正文 {n_para} 段，低于该刊 P25（{PARA_P25}）——"
                       f"长段堆叠倾向，该刊主体 2–6 句/段")
    elif n_para > PARA_P90:
        notices.append(f"正文 {n_para} 段，超过该刊 P90（{PARA_P90}）——"
                       f"碎段化倾向，考虑合并相邻短段")
    if n_chars < CHARS_MIN:
        issues.append(f"正文 {n_chars} 字，低于该刊单篇常规下限（{CHARS_MIN}，"
                      f"投稿者实测集中 10000–15000 字）——篇幅单薄")
    elif n_chars > CHARS_MAX:
        notices.append(f"正文 {n_chars} 字，超过该刊单篇常规上限（{CHARS_MAX}）——"
                       f"考虑压缩或与编辑沟通")

    return {"n_h1": n_h1, "n_para": n_para, "n_chars": n_chars,
            "n_titles": len(titles), "issues": issues, "notices": notices}


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
    print("篇章规模核验（章节数/段落数/字数 vs 该刊实测分位）")
    print(f"稿件：{os.path.basename(DOC)}")
    print(f"一级章节 {res['n_h1']} 个（刊 P25–P90：{SEC_P25}–{SEC_P90}，中位 {SEC_P50}）")
    print(f"正文段数 {res['n_para']}（刊 P25–P90：{PARA_P25}–{PARA_P90}，中位 {PARA_P50}）")
    print(f"正文字数 {res['n_chars']}（该刊常规 {CHARS_MIN}–{CHARS_MAX}）")
    print(f"标题总数 {res['n_titles']} 个")
    print("=" * 74)

    if res["issues"]:
        print(f"🔴 发现 {len(res['issues'])} 项规模问题：")
        for it in res["issues"]:
            print(f"   · {it}")
    else:
        print("✅ 篇章规模在该刊常规区间内。")
    if res["notices"]:
        print(f"\n🟡 建议核对 {len(res['notices'])} 项：")
        for it in res["notices"]:
            print(f"   · {it}")

    if "--json" in flags:
        jf = os.path.splitext(DOC)[0] + ".structure_scale.json"
        with open(jf, "w", encoding="utf-8") as f:
            json.dump(res, f, ensure_ascii=False, indent=1)
        print(f"\nJSON 已写入：{jf}")

    sys.exit(1 if res["issues"] else 0)


if __name__ == "__main__":
    main()
