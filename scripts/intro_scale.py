# -*- coding: utf-8 -*-
"""intro_scale.py —— 引言篇幅形态核验（过度铺陈/铺垫不足专查）

该刊引言章的实测画像（108 篇中 56 篇单列引言章）：占正文**中位 7.1%**，
是五段式里最短的一章（除结语 4.1%）。引言的职能是「背景 + 缺口」，
铺陈过多会挤占论证空间，过少则问题意识不足：

  · 引言占比 >15% → 🟡 过度铺陈（该刊单列章中位 7.1%，P90 约 14%），
    提示把背景/政策内容压缩，缺口句前置；
  · 引言占比 <3% → 🟡 铺垫不足，问题意识可能没立住；
  · 引言内段落数 <2 或 >5 → 🟡 形态偏离（该刊引言 2–3 句/段、
    通常 2–4 段：背景段 + 缺口段 + 本文工作段）；
  · 引言是否含缺口句（「尚缺乏/不足/有待」式）→ 该刊引言的必备件，
    缺了报 🔴（模式 E 第 1 维也核此项，此处按引言章定位）。

用法:
    python intro_scale.py <论文.docx> [--json]
退出码：1 = 引言缺缺口句（问题意识未立）。
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

# 该刊引言章实测（108 篇，56 篇单列）
INTRO_MED, INTRO_P90 = 7.1, 14.0
OVER_LIMIT, UNDER_LIMIT = 15.0, 3.0
PARA_MIN, PARA_MAX = 2, 5
# 缺口句信号（与 mock_review 第 1 维同口径）
GAP_RE = re.compile(
    r"(尚未|尚缺乏|缺乏|不足|有待|仍未|未得到充分|鲜有|较少|局限|未能).{0,30}"
    r"(研究|探讨|讨论|关注|支撑|证据|覆盖|涉及)")
GAP_RE2 = re.compile(r"(既有|现有|已有)研究.{0,40}(多聚焦|主要关注|较少|尚未)")


def analyze(doc):
    paras, refs, abs_text, kw_text, titles, body = dp.extract_paras(doc)
    secs = dp.scan_sections(paras, titles)
    if not secs or not secs.get("secs"):
        return {"issues": ["未检出章节结构"], "notices": [], "pct": None}

    # 找引言章（cls=引言，或名称含「引言」）
    intro = None
    for s in secs["secs"]:
        if s.get("cls") == "引言" or re.search(r"引\s*言", s["name"]):
            intro = s
            break
    if intro is None:
        return {"issues": [], "notices": ["未单列引言章（该刊 52% 篇目单列，"
                                        "其余并入正文首段）——若确已并入，请确认"
                                        "问题意识在正文前两段立住"],
                "pct": None}

    pct = intro["pct"]
    # 引言章文本（从 paras 切）
    intro_txt, in_intro = [], False
    for style, t in paras:
        if not t:
            continue
        s = re.sub(r"\s+", " ", t.strip())
        is_title = (len(s) < 40 and ((re.match(r"^(0|1|2|3|4|5|6|7)\s*[^\d]", s))
                                     or style.startswith("Heading")))
        if is_title:
            in_intro = bool(re.search(r"引\s*言", s))
            continue
        if in_intro:
            intro_txt.append(t)
    txt = "".join(intro_txt)
    n_paras = len(intro_txt)

    issues, notices = [], []
    if pct > OVER_LIMIT:
        notices.append(f"引言占正文 {pct:.1f}%，超过过度铺陈线（{OVER_LIMIT:.0f}%；"
                       f"该刊单列章中位 {INTRO_MED}%、P90≈{INTRO_P90:.0f}%）——"
                       f"压缩背景/政策铺陈，让缺口句前置")
    elif pct < UNDER_LIMIT:
        notices.append(f"引言仅占正文 {pct:.1f}%（<{UNDER_LIMIT:.0f}%）——"
                       f"铺垫不足，问题意识可能没立住")
    if n_paras < PARA_MIN:
        notices.append(f"引言仅 {n_paras} 段——该刊引言通常 2–4 段"
                       f"（背景段/缺口段/本文工作段）")
    elif n_paras > PARA_MAX:
        notices.append(f"引言 {n_paras} 段，偏多——该刊引言通常 2–4 段")
    has_gap = bool(GAP_RE.search(txt) or GAP_RE2.search(txt))
    if not has_gap:
        issues.append(f"引言章未检出缺口句（「尚缺乏/不足/有待…研究」式）——"
                      f"该刊引言必备件，模式 E 第 1 维同样核此项；"
                      f"请补一句明确的问题意识陈述")

    return {"pct": round(pct, 2), "chars": intro["chars"], "n_paras": n_paras,
            "has_gap": has_gap, "issues": issues, "notices": notices}


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
    print("引言篇幅形态核验（过度铺陈/铺垫不足/缺口句）")
    print(f"稿件：{os.path.basename(DOC)}")
    if res.get("pct") is not None:
        print(f"引言 {res['chars']} 字，占正文 {res['pct']}%，共 {res['n_paras']} 段")
        print(f"该刊基线：单列章中位 {INTRO_MED}%、P90≈{INTRO_P90:.0f}%；"
              f"缺口句 {'✓ 有' if res.get('has_gap') else '✗ 缺'}")
    print("=" * 74)

    if res["issues"]:
        print(f"🔴 发现 {len(res['issues'])} 项问题：")
        for it in res["issues"]:
            print(f"   · {it}")
    elif res.get("pct") is not None:
        print("✅ 引言含缺口句，篇幅形态正常。")
    if res["notices"]:
        print(f"\n🟡 建议核对 {len(res['notices'])} 项：")
        for it in res["notices"]:
            print(f"   · {it}")

    if "--json" in flags:
        jf = os.path.splitext(DOC)[0] + ".intro_scale.json"
        with open(jf, "w", encoding="utf-8") as f:
            json.dump(res, f, ensure_ascii=False, indent=1)
        print(f"\nJSON 已写入：{jf}")

    sys.exit(1 if res["issues"] else 0)


if __name__ == "__main__":
    main()
