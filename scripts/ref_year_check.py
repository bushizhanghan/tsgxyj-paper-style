# -*- coding: utf-8 -*-
"""ref_year_check.py —— 引文出版年越界与陈旧核验

引文表里出现「出版年晚于当前年」的条目（如 2027 年的文献），是**硬伤**——
要么是笔误，要么是引了不存在的文献，审稿人一抓一个准。此外：

  · 越界年（>当前年 2026）→ 🔴 硬伤，必须核对；
  · 早于 1990 年的条目 → 🟡 提示（经典文献可追溯，但该刊引文以近 10 年
    为主干，早年文献需确认确有必要）；
  · 无年份条目 → 🟡 提示（GB/T 7714 要求出版年必填）。

该刊实测（118 篇 5115 条引文）：篇内引文年份中位 2022、近 5 年占比中位
67.4%——引文以近 10 年为主干，少量经典追溯。

用法:
    python ref_year_check.py <论文.docx> [--json]
退出码：1 = 检出越界年（未来文献）。
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

CUR_YEAR = 2026          # 与该刊语料库同期（ref_format.CUR_YEAR 同值）
OLD_YEAR = 1990          # 早于此年提示（经典追溯需确认必要性）
# 年份模式：4 位年，排除 19xx-20xx 的日期段与标准号
YEAR_RE = re.compile(r"(?<!\d)(19\d{2}|20\d{2})(?!\d)")
DATE_RE = re.compile(r"\(\s*\d{4}-\d{2}-\d{2}\s*\)|\d{4}-\d{2}-\d{2}")
# 标准号里的年份（GB/T 37668—2019）不算出版年
STD_RE = re.compile(r"(?:GB/T|GB|ISO|IEC)\s*\d{5,6}[—\-–](\d{4})")


def ref_num(text):
    m = re.match(r"^[\[［]?\s*(\d{1,3})", text or "")
    return int(m.group(1)) if m else None


def pub_years(text):
    """抽取条目的出版年。

    注意：电子文献的「(2018-06-05)[2026-09-01]」里，前一段是**出版/发布日**，
    后一段才是**访问日期**。故只剥离方括号内的访问日期，保留圆括号内的
    出版日期段；标准号年份（GB/T 37668—2019）不算出版年，仍剥离。
    """
    t = text or ""
    # 只剥离方括号内的访问日期（[2026-09-01]），保留圆括号出版日期
    t = re.sub(r"\[\s*\d{4}-\d{2}-\d{2}\s*\]", " ", t)
    # 剥离标准号年份（GB/T 37668—2019）
    t = STD_RE.sub(" ", t)
    return [int(y) for y in YEAR_RE.findall(t)]


def analyze(doc):
    paras, refs, abs_text, kw_text, titles, body = dp.extract_paras(doc)
    future, old, no_year = [], [], []
    all_years = []
    for r in refs:
        n = ref_num(r)
        ys = pub_years(r)
        if not ys:
            no_year.append({"no": n, "text": r[:60]})
            continue
        # 取最大年作为出版年（卷期年在前、访问年在后已被剥离）
        y = max(ys)
        all_years.append(y)
        if y > CUR_YEAR:
            future.append({"no": n, "year": y, "text": r[:70]})
        elif y < OLD_YEAR:
            old.append({"no": n, "year": y, "text": r[:70]})

    issues, notices = [], []
    for f in future:
        issues.append(f"文献[{f['no']}] 出版年 {f['year']} 晚于当前年（{CUR_YEAR}）——"
                      f"「未来文献」是硬伤，核对是否笔误或文献不存在：{f['text'][:40]}")
    if old:
        notices.append(f"早于 {OLD_YEAR} 年的文献 {len(old)} 条："
                       + "、".join(f"[{o['no']}]({o['year']})" for o in old)
                       + "——该刊引文以近 10 年为主干（篇内年份中位 2022），"
                         "经典追溯请确认必要性")
    if no_year:
        notices.append(f"无出版年的条目 {len(no_year)} 条："
                       + "、".join(f"[{x['no']}]" for x in no_year)
                       + "——GB/T 7714 要求出版年必填")

    return {"future": future, "old": old, "no_year": no_year,
            "n_refs": len(refs), "n_years": len(all_years),
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
    print("引文出版年越界与陈旧核验")
    print(f"稿件：{os.path.basename(DOC)}")
    print(f"文献 {res['n_refs']} 条（可解析年份 {res['n_years']} 条）｜ "
          f"当前年基准 {CUR_YEAR}")
    print("=" * 74)

    if res["issues"]:
        print(f"🔴 发现 {len(res['issues'])} 项越界年硬伤：")
        for it in res["issues"]:
            print(f"   · {it}")
    else:
        print("✅ 无越界年（未来文献）。")
    if res["notices"]:
        print(f"\n🟡 建议核对 {len(res['notices'])} 项：")
        for it in res["notices"]:
            print(f"   · {it}")

    if "--json" in flags:
        jf = os.path.splitext(DOC)[0] + ".ref_year.json"
        with open(jf, "w", encoding="utf-8") as f:
            json.dump(res, f, ensure_ascii=False, indent=1)
        print(f"\nJSON 已写入：{jf}")

    sys.exit(1 if res["issues"] else 0)


if __name__ == "__main__":
    main()
