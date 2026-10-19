# -*- coding: utf-8 -*-
"""heading_audit.py —— 标题层级与编号体例核验

官方格式序列（见 references/official-guidelines.md §4）：

    0 引言
    1（一级标题）
    1.1（二级标题）
    1.2（二级标题）
    1.2.1（三级标题）
    （1）①②③
    （2）
    2（一级标题）
    2.1（二级标题）
    ……
    结语

初审退修的常见技术性问题正是**编号体例**：跳级（1 → 1.2.1 缺 1.2）、重号
（出现两个 2.3）、断号（1.1、1.3 却无 1.2）、空格不符（「1.1研究」缺空格）、
末章不叫「结语」、层级过深（四级以下）。

本脚本逐项核验：
  · 编号序列：一级/二级/三级各自的数字是否从 1 起、连续、无重号；
  · 层级跳越：出现 1.2.1 但整篇无 1.2 → 跳级；
  · 编号后空格：官方体例「1.1 标题」编号后须有空格（全角或半角均可）；
  · 层级深度：≥四级编号（1.2.3.4）提示该刊罕见；
  · 末章名称：官方序列写「结语」，用「结论/结束语/余论」则提示对齐；
  · 三级标题过量：超过实测 P90（约 22 个）提示是否过碎。

用法:
    python heading_audit.py <论文.docx> [--json]
退出码：0 = 编号体例未发现问题；1 = 存在跳级/重号/断号等问题。
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

# 标题形态：数字编号（1 / 1.1 / 1.2.3）或纯章名（引言/结语）
HEAD_RE = re.compile(r"^\s*(\d+(?:\.\d+){0,3})\s*(.*)$")
# 章名类标题（无编号）
NAME_HEADS = re.compile(r"^\s*(引\s*言|结\s*语|结\s*论|结束语|余\s*论|参考文献|注\s*释)\s*$")

H3_MAX = 22   # 该刊实测 P90，超过提示过碎


def extract_headings(titles, paras):
    """返回标题清单：[{raw, num, level, name, has_space}]。"""
    out = []
    for t in titles:
        s = (t or "").strip()
        if not s:
            continue
        m = HEAD_RE.match(s)
        if m:
            num, name = m.group(1), m.group(2).strip()
            # 编号后是否有空格（原始串里 num 与 name 之间）
            raw_after = s[len(m.group(1)):]
            has_space = bool(re.match(r"^\s+\S", raw_after)) or bool(re.match(r"^\u3000", raw_after))
            out.append({"raw": s, "num": num, "level": num.count(".") + 1,
                        "name": name, "has_space": has_space})
        elif NAME_HEADS.match(s):
            out.append({"raw": s, "num": None, "level": 1,
                        "name": re.sub(r"\s+", "", s), "has_space": True})
    return out


def audit(doc):
    paras, refs, abs_text, kw_text, titles, body = dp.extract_paras(doc)
    heads = extract_headings(titles, paras)
    issues, notices = [], []

    numbered = [h for h in heads if h["num"]]
    renamed = [h for h in heads if not h["num"]]

    # ---------- ① 各级编号连续性与重号 ----------
    by_level = {}
    for h in numbered:
        parts = [int(x) for x in h["num"].split(".")]
        by_level.setdefault(len(parts), []).append((h, parts))

    for lvl in sorted(by_level):
        items = by_level[lvl]
        # 重号
        seen = {}
        for h, parts in items:
            key = tuple(parts)
            if key in seen:
                issues.append(f"编号重号：出现两个「{h['num']}」"
                              f"（「{seen[key]['raw'][:24]}」与「{h['raw'][:24]}」）")
            else:
                seen[key] = h

    # 一级编号：从 1 起、连续（0 引言另计）
    lvl1 = [p[0] for _, p in by_level.get(1, [])]
    if lvl1:
        has_zero = 0 in lvl1
        seq = [x for x in lvl1 if x > 0]
        if seq:
            missing = [n for n in range(1, max(seq) + 1) if n not in seq]
            if missing:
                issues.append(f"一级标题断号：缺「{'、'.join(map(str, missing))}」"
                              f"（现有 {sorted(set(seq))}）")
        if has_zero != bool(re.search(r"^\s*0\s*引\s*言|^\s*引\s*言", "\n".join(titles), re.M)):
            notices.append("一级编号含 0（引言），官方序列确实以「0 引言」起——保持即可")

    # ---------- ② 二级/三级：按父号分组查连续性 ----------
    for lvl in (2, 3, 4):
        groups = {}
        for h, parts in by_level.get(lvl, []):
            parent = tuple(parts[:-1])
            groups.setdefault(parent, []).append(parts[-1])
        for parent, subs in groups.items():
            subs_sorted = sorted(set(subs))
            if subs_sorted:
                miss = [n for n in range(1, max(subs_sorted) + 1) if n not in subs_sorted]
                if miss:
                    pstr = ".".join(map(str, parent))
                    issues.append(f"{lvl} 级编号断号：{pstr}.x 缺 "
                                  f"「{'、'.join(f'{pstr}.{m}' for m in miss)}」")

    # ---------- ③ 层级跳越（出现 1.2.1 但无 1.2） ----------
    have = {tuple(int(x) for x in h["num"].split(".")) for h in numbered}
    for h in numbered:
        parts = tuple(int(x) for x in h["num"].split("."))
        if len(parts) >= 3:
            parent = parts[:-1]
            if parent not in have and any(len(p) == len(parent) for p in have):
                issues.append(f"层级跳越：「{h['num']}」的上级 "
                              f"「{'.'.join(map(str, parent))}」不存在")

    # ---------- ④ 编号后空格 ----------
    nospace = [h for h in numbered if not h["has_space"]]
    if nospace:
        issues.append(f"编号后缺空格 {len(nospace)} 处（官方体例「1.1 标题」）——"
                      f"例：{'、'.join(h['raw'][:18] for h in nospace[:3])}")

    # ---------- ⑤ 层级深度 ----------
    deeep = [h for h in numbered if h["level"] >= 4]
    if deeep:
        notices.append(f"四级及以下标题 {len(deeep)} 个（该刊罕见，建议合并到三级）——"
                       f"例：{deeep[0]['raw'][:20]}")

    # ---------- ⑥ 末章名称 ----------
    has_jieyu = any(re.search(r"结\s*语", t) for t in titles)
    has_other = [t for t in titles if re.search(r"结\s*论|结束语|余\s*论", t)]
    if has_other and not has_jieyu:
        issues.append(f"末章名为「{has_other[0][:20]}」——官方格式序列写「结语」，建议对齐")

    # ---------- ⑦ 三级标题过量 ----------
    h3 = by_level.get(3, [])
    if len(h3) > H3_MAX:
        notices.append(f"三级标题 {len(h3)} 个，超过该刊实测 P90（约 {H3_MAX} 个），考虑是否过碎")

    # ---------- ⑧ 编号起点 ----------
    for lvl in (2, 3):
        groups = {}
        for h, parts in by_level.get(lvl, []):
            groups.setdefault(tuple(parts[:-1]), []).append(parts[-1])
        for parent, subs in groups.items():
            if subs and min(subs) != 1:
                pstr = ".".join(map(str, parent))
                issues.append(f"{lvl} 级编号不从 1 起：{pstr} 下最小为 {pstr}.{min(subs)}")

    stat = {
        "total_heads": len(heads),
        "lvl1": len(by_level.get(1, [])),
        "lvl2": len(by_level.get(2, [])),
        "lvl3": len(by_level.get(3, [])),
        "unnumbered": len(renamed),
        "max_level": max([h["level"] for h in heads], default=0),
    }
    return {"issues": issues, "notices": notices, "stat": stat,
            "heads": [h["raw"] for h in heads]}


def main():
    args = [a for a in sys.argv[1:] if not a.startswith("--")]
    flags = set(a for a in sys.argv[1:] if a.startswith("--"))
    if not args or args[0] in ("-h", "--help"):
        print(__doc__)
        sys.exit(0)
    DOC = args[0]
    if not os.path.isfile(DOC):
        sys.exit(f"找不到文件：{DOC}")

    res = audit(DOC)
    st = res["stat"]
    print("=" * 72)
    print("标题层级与编号体例核验（官方格式序列）")
    print(f"稿件：{os.path.basename(DOC)}")
    print(f"标题 {st['total_heads']} 个：一级 {st['lvl1']} ／ 二级 {st['lvl2']} ／ "
          f"三级 {st['lvl3']} ／ 无编号 {st['unnumbered']}｜最深 {st['max_level']} 级")
    print("=" * 72)

    if res["issues"]:
        print(f"🔴 发现 {len(res['issues'])} 项编号体例问题：")
        for it in res["issues"]:
            print(f"   · {it}")
    else:
        print("✅ 编号体例：未发现跳级/重号/断号/空格问题")
    if res["notices"]:
        print(f"\n🟡 建议核对 {len(res['notices'])} 项：")
        for it in res["notices"]:
            print(f"   · {it}")

    if "--json" in flags:
        jf = os.path.splitext(DOC)[0] + ".heading_audit.json"
        with open(jf, "w", encoding="utf-8") as f:
            json.dump(res, f, ensure_ascii=False, indent=1)
        print(f"\nJSON 已写入：{jf}")

    sys.exit(1 if res["issues"] else 0)


if __name__ == "__main__":
    main()
