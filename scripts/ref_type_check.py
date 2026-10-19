# -*- coding: utf-8 -*-
"""ref_type_check.py —— 文献类型标识缺失与形态核验

GB/T 7714-2015 要求每条文献带类型标识（[J] 期刊／[M] 专著／[D] 学位论文／
[C] 会议／[S] 标准／[EB/OL] 电子文献…）。该刊官方 §7 更明文给出各类格式。

ref_format.py 已核「要素完整性」，但**类型标识这一项**值得单独深挖：
该刊 118 篇 5115 条引文实测，**14.8% 的条目无类型标识**——那多是早期
著录不规范的刊文本身；而**投稿稿应 100% 补齐**（无标识是初审最易挑的
格式问题之一）。

本脚本核验：
  · 无类型标识条目 → 🔴（投稿稿必须补全，附条目与建议标识推断）；
  · 标识大小写/形态异常（如 [j] 小写、[J/OL] 缺 /OL）→ 🟡；
  · 标识与内容不符的强信号（如带「http」或「[访问日期]」却标 [J] 而非
    [EB/OL]）→ 🟡 提示核对；
  · 类型分布对照该刊基准（J 68%、EB/OL 8.6%、M 3.8%、C 1.6%、D 1.4%）。

用法:
    python ref_type_check.py <论文.docx> [--json]
退出码：1 = 存在无类型标识条目。
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

# 合法类型标识（GB/T 7714-2015 表 A.1 + 该刊官方 §7 用到的）
TYPE_RE = re.compile(
    r"\[(J/OL|M/OL|D/OL|C/OL|S/OL|N/OL|EB/OL|DB/OL|DB/MT|CP/DK|Z/OL"
    r"|J|M|D|C|S|P|R|A|N|Z)\]")
# 电子文献强信号
ONLINE_SIG = re.compile(r"https?://|\[访问日期\]|\[\d{4}-\d{2}-\d{2}\]")
# 该刊类型分布基准（citation-benchmarks.md）
BASE_DIST = {"J": 68.0, "无标识": 14.8, "EB/OL": 8.6, "M": 3.8, "C": 1.6,
             "D": 1.4, "J/OL": 1.0, "N": 0.5}


def ref_num(text):
    m = re.match(r"^[\[［]?\s*(\d{1,3})", text or "")
    return int(m.group(1)) if m else None


def infer_type(text):
    """无标识时按内容推断建议标识。"""
    if ONLINE_SIG.search(text):
        return "EB/OL"
    if re.search(r"\[S\]|标准|GB/|ISO/", text):
        return "S"
    if re.search(r"学报|期刊|杂志|\.\s*19|20\d{2}\s*\(", text) and \
       re.search(r"\d{4}\s*[,，(]\s*\d{1,2}\s*[)）]?\s*[:：]\s*\d+", text):
        return "J"
    if re.search(r"出版社|出版地|：\s*[\u4e00-\u9fff]{2,6}\s*，|北京：|上海：", text):
        return "M"
    if re.search(r"会议|论文集|Proceedings|Conference", text, re.I):
        return "C"
    if re.search(r"学位论文|博士|硕士| dissertation|thesis", text, re.I):
        return "D"
    return None


def analyze(doc):
    paras, refs, abs_text, kw_text, titles, body = dp.extract_paras(doc)
    missing, bad_form, mismatch = [], [], []
    dist = Counter()
    for r in refs:
        n = ref_num(r)
        m = TYPE_RE.search(r)
        if not m:
            missing.append({"no": n, "text": r[:60], "guess": infer_type(r)})
            dist["无标识"] += 1
            continue
        t = m.group(1)
        dist[t] += 1
        # 形态异常：小写
        if t != t.upper():
            bad_form.append({"no": n, "type": t, "text": r[:50]})
        # 电子文献标了非 /OL 类型
        if ONLINE_SIG.search(r) and "/OL" not in t and t not in ("EB/OL",):
            mismatch.append({"no": n, "type": t, "text": r[:55]})

    issues, notices = [], []
    if missing:
        for x in missing:
            g = f"，建议补 [{x['guess']}]" if x["guess"] else ""
            issues.append(f"文献[{x['no']}] 无类型标识{g}——GB/T 7714 必填，"
                          f"该刊投稿稿应 100% 补齐：{x['text'][:36]}")
    if bad_form:
        notices.append(f"类型标识大小写异常 {len(bad_form)} 条："
                       + "、".join(f"[{x['no']}]({x['type']})" for x in bad_form)
                       + "——应为大写")
    if mismatch:
        notices.append(f"含网址/访问日期但类型标识非电子文献 {len(mismatch)} 条："
                       + "、".join(f"[{x['no']}]({x['type']})" for x in mismatch)
                       + "——电子文献应标 [EB/OL] 或 [J/OL] 等")

    # 类型分布（仅提示，该刊本身 14.8% 无标识）
    n = len(refs) or 1
    dist_txt = "、".join(f"{k} {v}条({v / n:.0%})" for k, v in dist.most_common())
    notices.append(f"本文类型分布：{dist_txt}（该刊基准 J 68%／无标识 14.8%／"
                   f"EB/OL 8.6%／M 3.8%）")

    return {"missing": missing, "bad_form": bad_form, "mismatch": mismatch,
            "dist": dict(dist), "n_refs": len(refs),
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
    print("文献类型标识缺失与形态核验（GB/T 7714 + 该刊官方 §7）")
    print(f"稿件：{os.path.basename(DOC)}")
    print(f"文献 {res['n_refs']} 条")
    print("=" * 74)

    if res["issues"]:
        print(f"🔴 发现 {len(res['issues'])} 项类型标识缺失：")
        for it in res["issues"]:
            print(f"   · {it}")
    else:
        print("✅ 全部条目均带类型标识。")
    if res["notices"]:
        print(f"\n🟡 建议核对 {len(res['notices'])} 项：")
        for it in res["notices"]:
            print(f"   · {it}")

    if "--json" in flags:
        jf = os.path.splitext(DOC)[0] + ".ref_type.json"
        with open(jf, "w", encoding="utf-8") as f:
            json.dump(res, f, ensure_ascii=False, indent=1)
        print(f"\nJSON 已写入：{jf}")

    sys.exit(1 if res["issues"] else 0)


if __name__ == "__main__":
    main()
