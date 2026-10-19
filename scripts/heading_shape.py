# -*- coding: utf-8 -*-
"""heading_shape.py —— 标题形态核验（字数/层级命名/并列结构）

heading_audit.py 核的是**编号序列**（跳级/重号/断号），不核**标题本身**的形态。
而该刊标题有明确的实测画像（118 篇 2779 个标题）：

  · 字数：P10=6、P25=9、P50=13、P75=19、P90=24——中位 13 字，
    ≥20 字占 23.1%（副标题式长题），≤8 字占 23.0%（短平快）；
  · 末章命名：官方序列写「结语」，且该刊一级标题命名有固定偏好
    （「对策与启示」「研究设计」「结果与分析」等）；
  · 标题不带对象：「有关研究」「若干思考」「初探」这类虚标题——
    该刊标题必含对象或方法词；
  · 并列结构：标题中的并列项（顿号列举）通常 ≤3 个，超过则显堆砌。

本脚本逐项核验，输出超限标题的原文与所在层级。

用法:
    python heading_shape.py <论文.docx> [--json]
退出码：0 = 标题形态无问题；1 = 检出虚标题或末章命名错误。
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

# 该刊标题字数画像（118 篇 2779 个标题实测）
LEN_P25, LEN_P75, LEN_P90 = 9, 19, 24
LONG_LIMIT = 30          # >30 字明显超长（P90=24，留余量）
SHORT_LIMIT = 4          # <4 字过短（「引言」「结语」等专用名除外）

# 虚标题（不带对象/方法）——该刊标题必含对象或方法词。
# 判定：整个标题由「泛词前缀 + 泛词」构成（如「若干思考」「有关研究」），
# 或整个标题就是一个泛词（「初探」「展望」）。
VAGUE_PREFIX = r"(?:有关|若干|一些|关于)"
VAGUE_CORE = (r"(?:浅谈|试论|初探|刍议|思考|探讨|研究|分析|综述|探索|实践|"
              r"展望|讨论|设想|考察|调研)")
VAGUE_TITLE = re.compile(rf"^(?:{VAGUE_PREFIX}?{VAGUE_CORE}|{VAGUE_PREFIX})$")
VAGUE_TAIL = re.compile(
    r"(研究|探讨|分析|思考|实践|探索|展望)\s*$")   # 仅作提示，不单独判

# 末章命名（官方序列）
CLOSING_OK = re.compile(r"结\s*语")
CLOSING_BAD = re.compile(r"结\s*论|结束语|余\s*论|总结与展望|研究结论")

# 专用短标题（不作为「过短」报）
SHORT_OK = {"引言", "结语", "结论", "绪论", "摘要", "前言", "方法", "结果",
            "讨论", "对策", "建议", "启示", "背景", "问题", "展望"}


def analyze(doc):
    paras, refs, abs_text, kw_text, titles, body = dp.extract_paras(doc)
    issues, notices = [], []

    long_t, short_t, vague_t = [], [], []
    for t in titles:
        s = re.sub(r"\s+", " ", (t or "").strip())
        if not s:
            continue
        # 去掉编号前缀，只量标题正文
        m = re.match(r"^\s*(\d+(?:\.\d+){0,3})\s*(.*)$", s)
        name = (m.group(2) if m else s).strip()
        if not name:
            continue
        L = len(dp.clean(name))
        if L > LONG_LIMIT:
            long_t.append({"title": s, "len": L})
        elif L < SHORT_LIMIT and dp.clean(name) not in SHORT_OK:
            short_t.append({"title": s, "len": L})
        # 虚标题：整个标题由「泛词前缀+泛词」构成，或本身就是泛词
        core = dp.clean(name)
        if VAGUE_TITLE.match(core):
            vague_t.append({"title": s, "len": L})

    for v in vague_t:
        issues.append(f"虚标题「{v['title'][:26]}」——不带对象或方法词，"
                      f"该刊标题必含具体对象（如「××群体××服务研究」）")
    for l in long_t:
        notices.append(f"标题过长（{l['len']} 字，该刊 P90=24）：「{l['title'][:26]}」"
                       f"——若用副标题，确认破折号衔接且不超 30 字")
    for s_ in short_t:
        notices.append(f"标题过短（{s_['len']} 字）：「{s_['title']}」——"
                       f"确认是否信息不足（该刊 P25=9）")

    # 末章命名
    last = titles[-1] if titles else ""
    if last and CLOSING_BAD.search(last) and not CLOSING_OK.search(last):
        issues.append(f"末章名为「{last[:20]}」——官方格式序列写「结语」，建议对齐")
    elif last and not CLOSING_OK.search(last) and not re.search(r"参考文献|注释", last):
        notices.append(f"末章名为「{last[:20]}」——该刊收口惯例为「结语」，"
                       f"请确认是否符合该栏目惯例")

    return {"issues": issues, "notices": notices,
            "n_titles": len(titles),
            "long": long_t, "short": short_t, "vague": vague_t,
            "last_title": last}


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
    print("标题形态核验（字数/虚标题/末章命名）")
    print(f"稿件：{os.path.basename(DOC)}")
    print(f"标题 {res['n_titles']} 个 ｜ 该刊画像：P25=9 字、P50=13、P75=19、P90=24")
    print("=" * 74)

    if res["issues"]:
        print(f"🔴 发现 {len(res['issues'])} 项标题问题：")
        for it in res["issues"]:
            print(f"   · {it}")
    else:
        print("✅ 标题形态：无虚标题、末章命名合规。")
    if res["notices"]:
        print(f"\n🟡 建议核对 {len(res['notices'])} 项：")
        for it in res["notices"]:
            print(f"   · {it}")

    if "--json" in flags:
        jf = os.path.splitext(DOC)[0] + ".heading_shape.json"
        with open(jf, "w", encoding="utf-8") as f:
            json.dump(res, f, ensure_ascii=False, indent=1)
        print(f"\nJSON 已写入：{jf}")

    sys.exit(1 if res["issues"] else 0)


if __name__ == "__main__":
    main()
