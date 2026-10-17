# -*- coding: utf-8 -*-
"""citation_analysis.py -- 引文结构与时效性分析（对照 118 篇刊文实测基准）

回答审稿人最常问的两个问题：
    1. 「你的参考文献够新吗？」  —— 引文年份分布 vs 该刊基准（首年份中位 2022、2021+ 占 65%）
    2. 「来源结构合理吗？」      —— J/M/C/D/EB-OL 类型占比 vs 该刊基准

两种用法：
    A. 分析语料库本身（挖掘基准，产出 references/citation-benchmarks.md）
       python citation_analysis.py --corpus-stats
    B. 分析你的论文（对照基准给出诊断）
       python citation_analysis.py <论文.docx>

输出的诊断项：
    · 引文时效：近 5 年占比、中位年份、最新/最旧年份
    · 类型结构：各类型条数与占比，与该刊对比
    · 高频来源：你的文献集中在哪些刊/出版社（与该刊高频刊源对照）
    · 风险提示：陈旧文献过多 / 来源单一 / 类型失衡

依据：5118 条该刊参考文献（118 篇，2026 年 1—9 期）
"""
import os
import re
import sys
import json
import collections
import statistics as st

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import corpus  # noqa: E402

if hasattr(sys.stdout, "reconfigure"):
    try:
        sys.stdout.reconfigure(encoding="utf-8", errors="replace")
    except Exception:
        pass

CUR_YEAR = 2026
TYPE_RE = re.compile(r"\[(J/OL|M/OL|J|M|D|C|N|Z|S|R|P|EB/OL|DB/OL|CP/DK)\]")


def norm(s):
    """全角→半角（含字母、数字、标点），去全角空格。"""
    return "".join(chr(ord(c) - 0xFEE0) if 0xFF01 <= ord(c) <= 0xFF5E else c
                   for c in s).replace("\u3000", " ")


def parse_refs(text_block):
    """从「参考文献」起始的文本块切分条目。返回 [(num, body)]。"""
    parts = re.split(r"\[(\d{1,3})\]", text_block)
    out = []
    for k in range(1, len(parts) - 1, 2):
        body = parts[k + 1].strip()
        if len(body) > 12 and not body.startswith("["):
            out.append((int(parts[k]), body))
    return out


def ref_year(body):
    """提取出版年：优先「年(期)」「年,」「年:」，排除 URL 与访问日期。"""
    b = re.sub(r"https?://\S+", "", body)
    b = re.sub(r"\[(19|20)\d{2}-\d{2}-\d{2}\]", "", b)   # 访问日期 [2026-09-01]
    m = re.search(r"((?:19|20)\d{2})\s*[\(（,，:：]", b)
    if m:
        return int(m.group(1))
    m = re.search(r"((?:19|20)\d{2})", b)
    return int(m.group(1)) if m else None


def ref_type(body):
    m = TYPE_RE.search(body)
    return m.group(1).upper() if m else "无标识"


def ref_source(body, rtype):
    """提取来源：J 类取刊名；M 类取出版者；EB/OL 取完整域名。

    J 类形态（GB/T 7714）：作者. 题名[J]. 刊名, 年(期): 页码.
    策略：定位 `[J]` 后的分隔点，取到紧随其后的中文/英文逗号前的部分。
    """
    if rtype and rtype.startswith("J"):
        m = re.search(r"\[J(?:/OL)?\]", body)
        if m:
            tail = body[m.end():]
            # 跳过 [J] 后可能的「.」「. 」等
            tail = re.sub(r"^[\s.。]*", "", tail)
            # 刊名到下一个英文/中文逗号或句点为止
            mm = re.match(r"([^,，.。]{2,32})", tail)
            if mm:
                name = mm.group(1).strip()
                # 排除误匹配：作者-机构串、纯数字、过短
                if (len(name) >= 2 and not re.search(r"\d{4}", name)
                        and not re.search(r"(大学|学院|研究所)\s*(副|讲师|教授|博士|硕士)", name)):
                    return name
    if rtype == "M":
        m = re.search(r"\[M\]", body)
        if m:
            tail = re.sub(r"^[\s.。]*", "", body[m.end():])
            mm = re.search(r"[:：]\s*([^,，.。]{2,30})", tail)
            if mm:
                return mm.group(1).strip()
    if "EB/OL" in body or "http" in body:
        m = re.search(r"https?://([^/\s，,。]+)", body)
        if m:
            host = m.group(1).strip().lower()
            host = re.sub(r"^www\.", "", host)
            parts = host.split(".")
            if len(parts) >= 3 and parts[-2] in ("com", "edu", "gov", "org", "net", "ac"):
                return ".".join(parts[-3:])
            return ".".join(parts[-2:]) if len(parts) >= 2 else host
    return None


# ---------- A. 语料库基准统计 ----------
def corpus_stats(corpus_dir):
    tdir = os.path.join(corpus_dir, "text")
    if not os.path.isdir(tdir):
        tdir = corpus_dir
    per_paper = []
    all_refs = []
    for fn in sorted(os.listdir(tdir)):
        if not fn.endswith(".md"):
            continue
        try:
            t = norm(open(os.path.join(tdir, fn), encoding="utf-8").read())
        except Exception:
            continue
        i = t.find("参考文献")
        if i < 0:
            continue
        items = parse_refs(t[i:])
        if len(items) < 5:
            continue
        yrs = [ref_year(b) for _, b in items]
        yrs = [y for y in yrs if y]
        types = collections.Counter(ref_type(b) for _, b in items)
        srcs = [ref_source(b, ref_type(b)) for _, b in items]
        srcs = [s for s in srcs if s]
        per_paper.append({
            "file": fn, "n": len(items),
            "year_med": st.median(yrs) if yrs else None,
            "recent5": sum(1 for y in yrs if y >= CUR_YEAR - 5) / len(yrs) * 100 if yrs else 0,
            "newest": max(yrs) if yrs else None,
            "oldest": min(yrs) if yrs else None,
            "types": dict(types),
            "sources": srcs,
        })
        all_refs.extend(items)

    if not per_paper:
        return None
    # 汇总基准
    meds = [p["year_med"] for p in per_paper if p["year_med"]]
    recents = [p["recent5"] for p in per_paper]
    typecnt = collections.Counter()
    srccnt = collections.Counter()
    for p in per_paper:
        typecnt.update(p["types"])
        srccnt.update(p["sources"])
    total_types = sum(typecnt.values())
    yearcnt = collections.Counter()
    for _, b in all_refs:
        y = ref_year(b)
        if y:
            yearcnt[y] += 1
    return {
        "n_papers": len(per_paper), "n_refs": len(all_refs),
        "year_med_med": st.median(meds) if meds else None,
        "year_med_mean": st.mean(meds) if meds else None,
        "recent5_med": st.median(recents) if recents else None,
        "recent5_mean": st.mean(recents) if recents else None,
        "recent5_p25": sorted(recents)[len(recents) // 4] if recents else None,
        "types": {k: round(v / total_types * 100, 1) for k, v in typecnt.most_common()},
        "type_counts": dict(typecnt),
        "top_sources": srccnt.most_common(25),
        "year_dist": sorted(yearcnt.items()),
        "per_paper": per_paper,
    }


# ---------- B. 论文诊断 ----------
def analyze_doc(docx_path, base):
    import diagnose_paper as dp  # 复用提取
    paras, refs, abs_text, kw_text, titles, body = dp.extract_paras(docx_path)
    if not refs:
        return None
    items = []
    for r in refs:
        m = re.match(r"^[\[［]?\s*(\d{1,3})\s*[\]］]?\s*[:：.]?\s*(.+)$", r)
        items.append((int(m.group(1)) if m else 0, m.group(2) if m else r))
    normed = [(n, norm(b)) for n, b in items]
    yrs = [ref_year(b) for _, b in normed]
    yrs = [y for y in yrs if y]
    types = collections.Counter(ref_type(b) for _, b in normed)
    srcs = [ref_source(b, ref_type(b)) for _, b in normed]
    srcs = [s for s in srcs if s]
    return {
        "n": len(normed),
        "years": yrs,
        "year_med": st.median(yrs) if yrs else None,
        "recent5": sum(1 for y in yrs if y >= CUR_YEAR - 5) / len(yrs) * 100 if yrs else 0,
        "newest": max(yrs) if yrs else None,
        "oldest": min(yrs) if yrs else None,
        "types": dict(types),
        "sources": srcs,
        "src_counter": collections.Counter(srcs),
        "old_items": [(n, b) for (n, b), y in zip(normed, yrs) if y and y < CUR_YEAR - 10][:8],
    }


def main():
    args = sys.argv[1:]
    if not args or args[0] in ("-h", "--help"):
        print(__doc__)
        return 0
    flags = {a for a in args if a.startswith("--")}
    rest = [a for a in args if not a.startswith("--")]
    cli_corpus = None
    if "--corpus" in args:
        i = args.index("--corpus")
        cli_corpus = args[i + 1] if i + 1 < len(args) else None

    corpus_dir = corpus.find_corpus_or_exit(cli_corpus, need="引文分析")

    if "--corpus-stats" in flags:
        s = corpus_stats(corpus_dir)
        if not s:
            print("[错误] 无法从语料库提取参考文献")
            return 2
        print("=" * 66)
        print("该刊引文基准（118 篇参考文献实测）")
        print("=" * 66)
        print(f"篇数 {s['n_papers']} | 文献 {s['n_refs']} 条 | 篇均 {s['n_refs']/s['n_papers']:.1f} 条")
        print(f"\n【时效性】")
        print(f"  篇内年份中位的中位：{s['year_med_med']:.0f}（均值 {s['year_med_mean']:.1f}）")
        print(f"  近 5 年文献占比：中位 {s['recent5_med']:.1f}%（P25 {s['recent5_p25']:.1f}%，均值 {s['recent5_mean']:.1f}%）")
        print(f"\n【类型结构】")
        for k, v in s["types"].items():
            print(f"  {k}: {v}% ({s['type_counts'][k]} 条)")
        print(f"\n【高频来源 top15】")
        for k, v in s["top_sources"][:15]:
            print(f"  {k}: {v}")
        print(f"\n【年份分布】")
        yd = s["year_dist"]
        tot = sum(c for _, c in yd)
        for y, c in yd:
            if c >= tot * 0.01:
                bar = "█" * int(c / max(x[1] for x in yd) * 40)
                print(f"  {y}: {c:>4} {bar}")
        if "--json" in flags:
            jf = os.path.join(corpus_dir, "citation_benchmarks.json")
            with open(jf, "w", encoding="utf-8") as f:
                json.dump({k: v for k, v in s.items() if k != "per_paper"},
                          f, ensure_ascii=False, indent=1)
            print(f"\nJSON 已写入：{jf}")
        return 0

    # B. 分析论文
    if not rest:
        print(__doc__)
        return 2
    DOC = rest[0]
    if not os.path.isfile(DOC):
        print(f"找不到文件：{DOC}")
        return 2
    base = corpus_stats(corpus_dir)
    if not base:
        print("[错误] 无法建立语料库基准")
        return 2
    r = analyze_doc(DOC, base)
    if not r:
        print("未解析到参考文献——请确认有独立「参考文献」段且条目带编号")
        return 1

    print("=" * 66)
    print(f"引文诊断：{os.path.basename(DOC)}")
    print(f"对照基准：{base['n_papers']} 篇刊文 / {base['n_refs']} 条文献")
    print("=" * 66)

    print(f"\n【① 时效性】文献 {r['n']} 条")
    print(f"  年份：{r['oldest']}–{r['newest']}｜中位 {r['year_med']:.0f}"
          f"（刊中位 {base['year_med_med']:.0f}）")
    print(f"  近 5 年（{CUR_YEAR-5}+）占比：{r['recent5']:.1f}%"
          f"（刊中位 {base['recent5_med']:.1f}%，P25 {base['recent5_p25']:.1f}%）")
    risk = []
    if r["recent5"] < base["recent5_p25"]:
        risk.append(f"⚠️ 近 5 年文献仅 {r['recent5']:.1f}%，低于该刊 P25（{base['recent5_p25']:.1f}%）"
                    f"——审稿意见常见「文献陈旧」")
    elif r["recent5"] < base["recent5_med"]:
        print(f"  · 近 5 年占比低于刊中位但仍在区间内，可再补 2–3 篇新文献")
    else:
        print("  ✓ 时效性达标")
    if r["old_items"]:
        print(f"  · 超过 10 年的陈旧文献 {len(r['old_items'])} 条（示例）：")
        for n, b in r["old_items"][:3]:
            print(f"      [{n}] {b[:60]}…")

    print(f"\n【② 类型结构】")
    total = sum(r["types"].values()) or 1
    for k, v in sorted(r["types"].items(), key=lambda x: -x[1]):
        mine = v / total * 100
        theirs = base["types"].get(k, 0)
        gap = ""
        if theirs and mine > max(theirs * 2.5, theirs + 4):
            gap = f"  ▲ 偏高（刊 {theirs}%）"
            # S（标准）在无障碍/技术规范类选题中不可替代，只提示不警告
            if k == "S":
                gap += "  · 标准类文献在该刊罕见但对本选题必要，属合理使用"
            elif k in ("C", "D", "N", "Z", "R", "P"):
                risk.append(f"⚠️ {k} 类占比 {mine:.1f}%（{v} 条），该刊仅 {theirs}%"
                            f"——此类文献该刊少引，建议压缩至 1–2 条，优先换成期刊论文")
        elif theirs and mine < theirs * 0.4:
            gap = f"  ▼ 偏低（刊 {theirs}%）"
            if k == "EB/OL" and theirs >= 3:
                risk.append(f"⚠️ 未引网络资源（该刊 EB/OL 占 {theirs}%）"
                            f"——政策文件/标准/机构报告可用 [EB/OL] 补充时效性")
        print(f"  {k}: {v} 条 = {mine:.1f}%（刊 {theirs}%）{gap}")

    print(f"\n【③ 来源分布】识别到 {len(r['src_counter'])} 个来源")
    for k, v in r["src_counter"].most_common(8):
        in_corpus = dict(base["top_sources"]).get(k, 0)
        tag = f"  （该刊高频源，{in_corpus} 次）" if in_corpus >= 5 else ""
        print(f"  {k}: {v}{tag}")
    if len(r["src_counter"]) <= 5 and r["n"] >= 15:
        risk.append(f"⚠️ 来源仅 {len(r['src_counter'])} 种，集中度过高——审稿意见常见「文献来源单一」")

    if risk:
        print("\n【风险提示】")
        for x in risk:
            print("  " + x)
    else:
        print("\n【结论】引文结构与时效性未见明显问题。")

    print("\n【依据】该刊 118 篇刊文实测基准；类型占比为该刊各类型条数占比。")

    if "--json" in flags:
        out = {"doc": os.path.basename(DOC), "n": r["n"],
               "year_med": r["year_med"], "recent5": round(r["recent5"], 1),
               "types": r["types"], "sources": dict(r["src_counter"]),
               "baseline": {"year_med": base["year_med_med"],
                            "recent5_med": base["recent5_med"],
                            "recent5_p25": base["recent5_p25"],
                            "types": base["types"]},
               "risks": risk}
        jf = os.path.splitext(DOC)[0] + ".citation.json"
        with open(jf, "w", encoding="utf-8") as f:
            json.dump(out, f, ensure_ascii=False, indent=1)
        print(f"\nJSON 已写入：{jf}")
    return 0


if __name__ == "__main__":
    sys.exit(main())
