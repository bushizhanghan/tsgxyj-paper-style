# -*- coding: utf-8 -*-
"""ref_dup_check.py —— 文献表重复与高相似条目核验

文献表里同一条文献出现两次（或仅页码/卷期不同的两条），是格式硬伤：
GB/T 7714 与官方 §7 都要求「同一文献多次引用使用同一注释序号，不再重复
给号」——重复列条目既暴露校对不严，也让审稿人怀疑引用态度。

ref_format.py 已核「完全相同的重复条目」；本脚本补两档更细的：

  · 🔴 完全重复：归一化（去空白/标点）后两条完全一致；
  · 🔴 同题名重复：作者与题名相同、仅出版信息（页码/卷期/网址）不同——
    这通常是同一文献的两次著录，必须合并；
  · 🟡 同作者同主题高相似：同第一作者 + 题名主题词高度重叠（≥0.6），
    可能是同系列研究的误列，提示人工确认。

用法:
    python ref_dup_check.py <论文.docx> [--json]
退出码：1 = 检出完全重复或同题名重复。
"""
import os
import re
import sys
import json
import itertools

# Windows 控制台 GBK 保护
if hasattr(sys.stdout, "reconfigure"):
    try:
        sys.stdout.reconfigure(encoding="utf-8", errors="replace")
    except Exception:
        pass

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import diagnose_paper as dp  # noqa: E402
import argument_check as ac  # noqa: E402


def ref_num(text):
    m = re.match(r"^[\[［]?\s*(\d{1,3})", text or "")
    return int(m.group(1)) if m else None


def norm_entry(text):
    """归一化条目：去序号、空白、标点。"""
    t = re.sub(r"^[\[［]?\s*\d{1,3}\s*[\]］]?\s*", "", text or "")
    t = re.sub(r"[\[\]【】（）()．.，,。；;：:\s]", "", t)
    return t.lower()


def author_of(text):
    """第一作者：序号后、第一个标点前的部分。"""
    t = re.sub(r"^[\[［]?\s*\d{1,3}\s*[\]］]?\s*", "", text or "")
    m = re.match(r"^([^,.．,，\[［]{2,20}?)[,.．,，\[［]", t)
    if m:
        return re.sub(r"\s", "", m.group(1))
    # 无标点时取前 8 字
    return re.sub(r"\s", "", t[:8])


def title_of(text):
    """题名：类型标识前、第一个句点后的部分（粗粒度）。"""
    t = re.sub(r"^[\[［]?\s*\d{1,3}\s*[\]］]?\s*", "", text or "")
    # 去掉第一作者段
    t = re.sub(r"^[^,.．,，]{2,20}?[,.．,，]\s*", "", t)
    # 取到类型标识或第二个标点为止
    m = re.split(r"\[(?:[A-Z/]+)\]|[,.．,，]", t, maxsplit=1)
    return re.sub(r"\s", "", m[0]) if m else t[:40]


def analyze(doc):
    paras, refs, abs_text, kw_text, titles, body = dp.extract_paras(doc)
    exact, same_title, high_sim = [], [], []

    entries = [(ref_num(r), r) for r in refs]
    # ① 完全重复
    seen = {}
    for n, r in entries:
        k = norm_entry(r)
        if k in seen:
            exact.append({"a": seen[k], "b": n, "text": r[:50]})
        else:
            seen[k] = n

    # ② 同题名重复（作者相同 + 题名相同）
    by_ta = {}
    for n, r in entries:
        a, t = author_of(r), title_of(r)
        if len(t) < 6:
            continue
        k = (a, t[:24])
        if k in by_ta:
            same_title.append({"a": by_ta[k], "b": n, "author": a, "title": t[:30]})
        else:
            by_ta[k] = n

    # ③ 同第一作者 + 题名主题词高相似
    by_author = {}
    for n, r in entries:
        a = author_of(r)
        if a:
            by_author.setdefault(a, []).append((n, r))
    for a, lst in by_author.items():
        if len(lst) < 2:
            continue
        for (n1, r1), (n2, r2) in itertools.combinations(lst, 2):
            t1, t2 = title_of(r1), title_of(r2)
            if len(t1) < 8 or len(t2) < 8:
                continue
            v1, v2 = ac.topic_vec(t1), ac.topic_vec(t2)
            sc = ac.overlap_score(v1, v2)
            # 归一化到 0-1
            max_possible = max(len(v1), len(v2), 1)
            ratio = sc / max_possible
            if ratio >= 0.6:
                high_sim.append({"a": n1, "b": n2, "author": a,
                                 "score": round(ratio, 2),
                                 "t1": t1[:26], "t2": t2[:26]})

    issues, notices = [], []
    for x in exact:
        issues.append(f"文献[{x['a']}] 与 [{x['b']}] 完全重复——须合并为一条、"
                      f"共用同一序号：{x['text'][:34]}")
    for x in same_title:
        issues.append(f"文献[{x['a']}] 与 [{x['b']}] 同作者同题名（{x['author']}："
                      f"{x['title']}）——通常是同一文献两次著录，须合并")
    for x in high_sim:
        notices.append(f"文献[{x['a']}] 与 [{x['b']}] 同作者且题名高度相似"
                       f"（{x['score']}）：「{x['t1']}」vs「{x['t2']}」——"
                       f"请确认是否为同系列研究的误列")

    return {"exact": exact, "same_title": same_title, "high_sim": high_sim,
            "n_refs": len(refs), "issues": issues, "notices": notices}


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
    print("文献表重复与高相似条目核验（GB/T 7714「不重复给号」）")
    print(f"稿件：{os.path.basename(DOC)}")
    print(f"文献 {res['n_refs']} 条")
    print("=" * 74)

    if res["issues"]:
        print(f"🔴 发现 {len(res['issues'])} 项重复：")
        for it in res["issues"]:
            print(f"   · {it}")
    else:
        print("✅ 无完全重复或同题名重复条目。")
    if res["notices"]:
        print(f"\n🟡 建议核对 {len(res['notices'])} 项：")
        for it in res["notices"]:
            print(f"   · {it}")

    if "--json" in flags:
        jf = os.path.splitext(DOC)[0] + ".ref_dup.json"
        with open(jf, "w", encoding="utf-8") as f:
            json.dump(res, f, ensure_ascii=False, indent=1)
        print(f"\nJSON 已写入：{jf}")

    sys.exit(1 if res["issues"] else 0)


if __name__ == "__main__":
    main()
