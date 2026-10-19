# -*- coding: utf-8 -*-
"""paragraph_shape.py —— 段落首尾形态与分号使用核验（三段合一）

paragraphing-guide.md 四记载的段内范式里，有三项是**确定性规则**，
适合合并为一个工具一次跑完（避免三个工具重复切段、重复读 docx）：

  · 段首句式（该刊 118 篇 4743 段实测：直接陈述 94.3%、连接词起句 1.4%、
    时间/背景起句 4.3%）——连接词起句偏离刊风；
  · 段末收束（实测套话收尾 0.02%，即 4743 段仅 1 处）——「综上所述」式
    收尾是该刊几乎不写的形态；
  · 分号使用（长段含分号率 25.1%）——并列信息用分号连接是该刊习惯，
    一段内多个并列句却零分号，往往是一口气到底的长句堆叠。

三项均为**提示级**：官方未对段落形态作规定，属「实际做法」统计，
用于风格拟合，不作为投稿资格问题。

用法:
    python paragraph_shape.py <论文.docx> [--top 20] [--json]
退出码：恒 0（提示级工具，不设硬伤退出码）。
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

MIN_PARA = 40          # 参与核验的最短段（字）
# 连接词起句（该刊仅 1.4%）
CONN_HEAD = re.compile(
    r"^\s*(然而|但是|但|此外|另外|同时|因此|所以|从而|进而|综上所述|总之|"
    r"与此同时|在此基础上|值得一提的是|不难看出|显而易见)")
# 段末套话（该刊 0.02%）
CLICHE_TAIL = re.compile(
    r"(综上所述|总而言之|总之|总的来说|由此可见|可见|"
    r"具有重要意义|值得深入探讨|提供有益参考|意义重大)[。！]?\s*$")
# 并列信号：段内多个「一是/第一/其一」式引导 或 多个顿号列举
PARALLEL_MARK = re.compile(
    r"(一是|二是|三是|四是|第一|第二|第三|第四|其一|其二|其三)")


def analyze(doc):
    paras, refs, abs_text, kw_text, titles, body = dp.extract_paras(doc)
    conn_head, cliche_tail, no_semi = [], [], []
    n_para = 0
    for pi, raw in enumerate(body):
        t = dp.clean(raw or "")
        if len(t) < MIN_PARA:
            continue
        n_para += 1
        pno = pi + 1
        sents = [s for s in re.split(r"(?<=[。？！])", t) if s.strip()]

        # ① 连接词起句
        if sents and CONN_HEAD.search(sents[0]):
            conn_head.append({"para": pno, "head": sents[0][:52]})

        # ② 段末套话
        if CLICHE_TAIL.search(t):
            cliche_tail.append({"para": pno, "tail": t[-42:]})

        # ③ 并列信息零分号：段内 ≥2 个并列引导 且 全段无分号
        n_par = len(PARALLEL_MARK.findall(t))
        if n_par >= 2 and ("；" not in t and ";" not in t):
            no_semi.append({"para": pno, "marks": n_par, "chars": len(t)})

    return {"n_para": n_para, "conn_head": conn_head, "cliche_tail": cliche_tail,
            "no_semi": no_semi}


def main():
    args = [a for a in sys.argv[1:] if not a.startswith("--")]
    flags = set(a for a in sys.argv[1:] if a.startswith("--"))
    if not args or args[0] in ("-h", "--help"):
        print(__doc__)
        sys.exit(0)
    DOC = args[0]
    if not os.path.isfile(DOC):
        sys.exit(f"找不到文件：{DOC}")
    TOP = 20
    for i, a in enumerate(sys.argv):
        if a == "--top" and i + 1 < len(sys.argv):
            try:
                TOP = int(sys.argv[i + 1])
            except ValueError:
                pass

    res = analyze(DOC)
    print("=" * 74)
    print("段落首尾形态与分号使用核验（提示级 · 风格拟合）")
    print(f"稿件：{os.path.basename(DOC)}")
    print(f"核验 {res['n_para']} 段（≥{MIN_PARA} 字）")
    print("=" * 74)
    print(f"合计：连接词起句 {len(res['conn_head'])} ｜ 段末套话 "
          f"{len(res['cliche_tail'])} ｜ 并列零分号 {len(res['no_semi'])}")
    print("该刊基线：连接词起句 1.4% ｜ 段末套话 0.02% ｜ 长段含分号 25.1%")

    if res["conn_head"]:
        print("\n【连接词起句】（该刊直接陈述占 94.3%，衔接靠内容回指）")
        for it in res["conn_head"][:TOP]:
            print(f"   段{it['para']}：{it['head']}")
    if res["cliche_tail"]:
        print("\n【段末套话】（该刊段末落在具体判断上，套话收尾仅 0.02%）")
        for it in res["cliche_tail"][:TOP]:
            print(f"   段{it['para']}段末：…{it['tail']}")
    if res["no_semi"]:
        print("\n【并列信息零分号】（该刊长段含分号率 25.1%，并列项宜用分号连接）")
        for it in res["no_semi"][:TOP]:
            print(f"   段{it['para']}（{it['marks']} 个并列标记，{it['chars']} 字）")

    total = len(res["conn_head"]) + len(res["cliche_tail"]) + len(res["no_semi"])
    print("\n" + "-" * 74)
    if total == 0:
        print("✅ 段落首尾形态与分号使用：未发现偏离。")
    else:
        print(f"🟡 共 {total} 处提示级偏离——不影响投稿资格，调整后更贴近刊风。")

    if "--json" in flags:
        jf = os.path.splitext(DOC)[0] + ".paragraph_shape.json"
        with open(jf, "w", encoding="utf-8") as f:
            json.dump(res, f, ensure_ascii=False, indent=1)
        print(f"\nJSON 已写入：{jf}")

    sys.exit(0)


if __name__ == "__main__":
    main()
