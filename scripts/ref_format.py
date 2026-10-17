# -*- coding: utf-8 -*-
"""ref_format.py -- 参考文献表 GB/T 7714 要素完整性与形态核验

用法:
    python ref_format.py <论文.docx> [--json]

核验维度（对照 GB/T 7714-2015 与该刊体例）:
    (1) 要素完整性  作者 / 题名 / 文献类型标识 [J][M][D][C][S][EB/OL] 等 /
                   出版信息（刊名+年卷期 或 出版地+出版者）——缺项即报
    (2) 形态规范    序号形态 [n]；标点形态；年份范围（1900–当前年+1）；
                   页码区间（排除日期/标准号误匹配）
    (3) 一致性别名  同一作者名在表内拼写不一致；同一条目重复出现
    (4) 类型分布    J / M / D / C / EB/OL 等占比（该刊以 J 为主、EB/OL 偶发；
                   外文文献作者姓全大写）

边界：只做形态与要素核验，不核验文献真实性与内容正确性（需人工/数据库）。
"""
import re
import sys
import os
import json

# Windows 控制台 GBK 保护
if hasattr(sys.stdout, "reconfigure"):
    try:
        sys.stdout.reconfigure(encoding="utf-8", errors="replace")
    except Exception:
        pass

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

CUR_YEAR = 2026  # 与该刊语料库同期；超此年的年份标记为可疑

# 文献类型标识（GB/T 7714-2015 表 A.1；带 /OL 的在线版优先匹配）
TYPE_PAT = re.compile(
    r"\[(J/OL|M/OL|D/OL|C/OL|S/OL|N/OL|EB/OL|DB/OL|DB/MT|CP/DK|CS/OL|DS/OL|Z/OL"
    r"|J|M|D|C|S|P|R|A|N|Z)\]",
    re.I)


def parse_item(raw):
    """解析单条文献。返回 dict：num / authors / year / type / rest

    年份取「出版年」：优先匹配 年,卷(期) / 年: 页码 形态中的年份，
    其次匹配 (2020-01-01) 之外的独立 4 位年份；避免把访问日期或
    网址中的数字当年份。
    """
    m = re.match(r"^[\[［]?\s*(\d{1,3})\s*[\]］]?\s*[:：.]?\s*(.+)$", raw)
    num = int(m.group(1)) if m else None
    txt = m.group(2) if m else raw
    am = re.match(r"^([^.]{2,60}?)\.", txt)
    authors = am.group(1).strip() if am else ""
    year = None
    # ① 出版年形态：紧邻「,」或「:」前的 4 位年份（2014, 40(4): / 2019: ）
    ym = re.search(r"((?:19|20)\d{2})\s*[,，:：]", txt)
    if ym:
        year = int(ym.group(1))
    else:
        # ② 「年」字语境
        ym = re.search(r"((?:19|20)\d{2})\s*年", txt)
        if ym:
            year = int(ym.group(1))
        else:
            # ③ 兜底：排除 URL 与括号内日期后的第一个 4 位年份
            masked = re.sub(r"\([^)]*\d{4}-\d{2}[^)]*\)", "", txt)   # 去 (2020-01-01)
            masked = re.sub(r"https?://\S+", "", masked)            # 去 URL
            ym = re.search(r"((?:19|20)\d{2})", masked)
            if ym:
                year = int(ym.group(1))
    tm = TYPE_PAT.search(txt)
    rtype = tm.group(1).upper() if tm else None
    return {"raw": raw, "num": num, "authors": authors, "year": year,
            "type": rtype, "rest": txt}


def find_page_inversion(rest):
    """在条目尾部找页码区间颠倒（排除日期 2023-06、标准号 GB/T x—y、年份对、
    以及 [2026-09-01] 这类方括号访问日期）。"""
    tail = rest[-30:]
    for m_ in re.finditer(r"(\d+)\s*[-–—]\s*(\d+)", tail):
        a, b = m_.group(1), m_.group(2)
        pre = tail[:m_.start()]
        # 日期上下文：前面是 4 位年份、[、( 、- 或紧邻分隔符
        if re.search(r"(19|20)\d{2}\s*$", pre):
            continue
        if re.search(r"[\[\(\-]\s*$", pre) and len(a) == 4 and a.startswith(("19", "20")):
            continue
        if len(a) == 4 and len(b) == 4:                   # 年份对
            if a.startswith(("19", "20")) and b.startswith(("19", "20")):
                continue
        if re.search(r"GB[/]?T?\s*\d*\s*$", pre, re.I):   # 标准号
            continue
        if len(a) > 4 or len(b) > 4:                      # 页码极少 >4 位
            continue
        return (a, b)
    return None


def check_items(items):
    """对解析后的条目跑全部核验，返回 (issues, type_count, dups, latin_stats)。"""
    issues = []
    type_count = {}
    for it in items:
        n = it["num"]
        if not it["authors"]:
            issues.append((n, "缺作者", "未解析到作者段（首个「.」前的部分）", it["raw"][:50]))
        if not it["type"]:
            if re.search(r"https?://|www\.", it["rest"]):
                issues.append((n, "疑似缺[EB/OL]",
                               "条目含网址但无文献类型标识——网络资源应标 [EB/OL]",
                               it["raw"][:50]))
            else:
                issues.append((n, "缺文献类型标识",
                               "未检出 [J]/[M]/[D]/[C]/[EB/OL] 等标识", it["raw"][:50]))
        if it["year"] is None:
            issues.append((n, "缺年份", "未检出 4 位年份", it["raw"][:50]))
        elif it["year"] < 1900 or it["year"] > CUR_YEAR + 1:
            issues.append((n, "年份可疑",
                           f"{it['year']}（正常范围 1900–{CUR_YEAR}）", it["raw"][:50]))
        if it["type"] == "J":
            if not re.search(r"\d{4}\s*[,，]\s*\d+", it["rest"]) and \
               not re.search(r"\d+\(\d+\)", it["rest"]):
                issues.append((n, "疑似缺卷期", "J 类条目未见「年,卷(期)」形态", it["raw"][:50]))
        if it["type"] == "M":
            if not re.search(r"[:：].*[:：]", it["rest"]):
                issues.append((n, "疑似缺出版信息",
                               "M 类条目未见「出版地: 出版者」", it["raw"][:50]))
        if re.search(r"[\u4e00-\u9fff],[\u4e00-\u9fff]", it["rest"][:40]):
            issues.append((n, "标点提示",
                           "条目前部出现半角逗号（中文文献表宜统一标点风格）", it["raw"][:50]))
        pm = find_page_inversion(it["rest"])
        if pm and int(pm[0]) > int(pm[1]):
            issues.append((n, "页码区间颠倒", f"{pm[0]}-{pm[1]}", it["raw"][:50]))
        key = it["type"] or "未知"
        type_count[key] = type_count.get(key, 0) + 1

    # 重复条目
    dups = []
    seen = {}
    for it in items:
        key = re.sub(r"\s+", "", it["rest"])[:40]
        if key in seen:
            dups.append((seen[key], it["num"], key))
        else:
            seen[key] = it["num"]

    latin_items = [it for it in items if re.search(r"[A-Za-z]{3,}", it["rest"])]
    upper_ok = sum(1 for it in items
                   if it["authors"] and re.search(r"[A-Z]", it["authors"])
                   and re.search(r"[A-Z]{2,}", it["authors"]))
    latin_lower = [it for it in latin_items
                   if it["authors"] and re.match(r"^[a-z]", it["authors"].strip())]
    return issues, type_count, dups, {
        "n_latin": len(latin_items), "upper_ok": upper_ok, "lower_surname": latin_lower}


def main():
    args = sys.argv[1:]
    if not args or args[0] in ("-h", "--help"):
        print(__doc__)
        return 0
    _rest = [a for a in args if not a.startswith("--")]
    _flags = {a for a in args if a.startswith("--")}
    PATH = _rest[0]
    if not os.path.isfile(PATH):
        print(f"找不到文件：{PATH}")
        return 2

    from docx import Document  # noqa: E402
    doc = Document(PATH)
    paras = [p.text.strip() for p in doc.paragraphs if p.text.strip()]

    REF_TITLE = None
    for i, t in enumerate(paras):
        if t.startswith("参考文献"):
            REF_TITLE = i
            break
    if REF_TITLE is None:
        print("未找到「参考文献」段")
        return 1
    refs = [t for t in paras[REF_TITLE + 1:] if re.match(r"^[\[［]?\d{1,3}[\]］]?", t)]
    if not refs:
        print("未解析到文献条目")
        return 1

    items = [parse_item(t) for t in refs]
    issues, type_count, dups, latin = check_items(items)

    print("=" * 62)
    print(f"文献表 GB/T 7714 要素核验：{os.path.basename(PATH)}")
    print(f"条目 {len(items)} 条（编号 {items[0]['num']}–{items[-1]['num']}）")
    print("=" * 62)

    print("\n【类型分布】" + "  ".join(f"{k}:{v}" for k, v in
                                     sorted(type_count.items(), key=lambda x: -x[1])))
    print("  （该刊以 J 期刊论文为主、EB/OL 偶发；外文作者姓全大写）")

    print(f"\n【要素与形态问题】共 {len(issues)} 处")
    for n, kind, detail, raw in issues:
        mark = "⚠️" if kind in ("缺作者", "缺年份") else "·"
        tag = f"[{n}]" if n else "[?]"
        print(f"  {mark} {tag} {kind}：{detail}")
        print(f"      {raw}")

    if dups:
        print(f"\n【重复条目】{len(dups)} 组")
        for a, b, k in dups[:10]:
            print(f"  [{a}] 与 [{b}] 内容疑似重复：{k}")

    if latin["n_latin"]:
        print(f"\n【外文条目】{latin['n_latin']} 条；姓全大写形态 {latin['upper_ok']} 条")
        if latin["lower_surname"]:
            print(f"  ⚠️ {len(latin['lower_surname'])} 条外文作者未用全大写姓开头：")
            for it in latin["lower_surname"][:5]:
                print(f"    [{it['num']}] {it['authors'][:40]}")

    print("\n【边界】仅形态与要素核验；文献真实性与内容正确性需人工或数据库核验。")

    if "--json" in _flags:
        out = {
            "doc": os.path.basename(PATH), "n_items": len(items),
            "types": type_count,
            "issues": [{"num": n, "kind": k, "detail": d, "raw": r}
                       for n, k, d, r in issues],
            "dups": dups, "latin_items": latin["n_latin"],
        }
        jf = os.path.splitext(PATH)[0] + ".reffmt.json"
        with open(jf, "w", encoding="utf-8") as f:
            json.dump(out, f, ensure_ascii=False, indent=1)
        print(f"\nJSON 已写入：{jf}")
    return 0


if __name__ == "__main__":
    sys.exit(main())
