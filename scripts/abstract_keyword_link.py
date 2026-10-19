# -*- coding: utf-8 -*-
"""abstract_keyword_link.py —— 摘要与关键词的关联核验

关键词的职责是「让读者检索到论文」。除了 keyword_echo.py 查的「关键词在
全文是否出现」（覆盖度），还有一层：**关键词应从摘要的核心概念派生**——
审稿人与编辑部常通过摘要判断关键词选得准不准。

本脚本只做**确定性词面核验**（不用「覆盖率」这类无判别力指标——曾尝试用
「关键词对摘要高频实词的覆盖率」衡量，实测在无分词器条件下摘要实词表
被「碍群体利」「体利用数」这类跨词碎片占满，覆盖率仅 15%，完全无意义，
故已移除）：

  · 每个关键词是否在摘要中出现（整词包含，或作为摘要中更长中文串的
    子串——如「智慧图书馆」在摘要以「智慧图书馆服务」出现也算）；
  · 摘要未出现的关键词占比 >50% → 🔴（关键词与摘要脱节，若非刻意补充
    的检索热词，应换成摘要实际论述的概念）；
  · 摘要未出现的关键词 ≤50% → 🟡 逐个列出，供作者确认。

用法:
    python abstract_keyword_link.py <论文.docx> [--json]
退出码：1 = 过半关键词未在摘要出现。
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

MISS_LIMIT = 0.5      # 摘要未出现的关键词占比上限（超过判 🔴）


def split_kws(kw_text):
    t = (kw_text or "").strip()
    t = re.sub(r"^【?\s*(?:关\s*键\s*词|Key\s*words?)\s*】?\s*[:：]?", "", t, flags=re.I)
    return [p.strip() for p in re.split(r"[;；,，、]\s*|\s{2,}", t) if p.strip()]


def kw_in_abstract(kw, abs_c):
    """关键词是否在摘要出现：整词包含，或作为摘要更长中文串的子串。"""
    if kw in abs_c:
        return True
    # 摘要中提取连续中文串，看关键词是否为某个串的子串
    for blk in re.findall(r"[\u4e00-\u9fff]{4,}", abs_c):
        if kw in blk:
            return True
    return False


def analyze(doc):
    paras, refs, abs_text, kw_text, titles, body = dp.extract_paras(doc)
    kws = split_kws(kw_text)
    abs_c = dp.clean(abs_text or "")
    if not kws or not abs_c:
        return {"issues": [], "notices": ["摘要或关键词缺失，无法核验"],
                "rows": [], "n_kw": 0, "abs_chars": 0}

    rows, miss = [], []
    for i, kw in enumerate(kws, 1):
        hit = kw_in_abstract(kw, abs_c)
        rows.append({"i": i, "kw": kw, "in_abs": hit})
        if not hit:
            miss.append(kw)

    miss_ratio = len(miss) / len(kws) if kws else 0.0
    issues, notices = [], []
    if miss_ratio > MISS_LIMIT:
        issues.append(f"{len(miss)}/{len(kws)} 个关键词未在摘要出现"
                      f"（{'、'.join(miss)}）——关键词与摘要脱节；"
                      f"若非刻意补充的检索热词，建议换成摘要实际论述的核心概念")
    elif miss:
        notices.append(f"未在摘要出现的关键词：{'、'.join(miss)}——"
                       f"若为检索热词可保留，否则建议替换")

    return {"rows": rows, "miss": miss, "miss_ratio": round(miss_ratio, 3),
            "n_kw": len(kws), "abs_chars": len(abs_c),
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
    print("摘要与关键词关联核验（关键词是否从摘要核心概念派生）")
    print(f"稿件：{os.path.basename(DOC)}")
    print(f"关键词 {res.get('n_kw', 0)} 个 ｜ 摘要 {res.get('abs_chars', 0)} 字 ｜ "
          f"未在摘要出现 {len(res.get('miss', []))} 个")
    print("=" * 74)

    if res["rows"]:
        print("【关键词在摘要的出现情况】")
        for r in res["rows"]:
            mark = "✓" if r["in_abs"] else "✗"
            print(f"  {mark} {r['i']}. {r['kw']}")

    print("\n" + "-" * 74)
    if res["issues"]:
        print(f"🔴 发现 {len(res['issues'])} 项关联问题：")
        for it in res["issues"]:
            print(f"   · {it}")
    else:
        print("✅ 关键词与摘要关联良好。")
    if res["notices"]:
        print(f"\n🟡 建议核对 {len(res['notices'])} 项：")
        for it in res["notices"]:
            print(f"   · {it}")

    if "--json" in flags:
        jf = os.path.splitext(DOC)[0] + ".abstract_keyword_link.json"
        with open(jf, "w", encoding="utf-8") as f:
            json.dump(res, f, ensure_ascii=False, indent=1)
        print(f"\nJSON 已写入：{jf}")

    sys.exit(1 if res["issues"] else 0)


if __name__ == "__main__":
    main()
