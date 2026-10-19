# -*- coding: utf-8 -*-
"""abstract_shape.py —— 摘要要素顺序与句式核验

diagnose_paper 的摘要体检核「四要素有无」（目的/方法/结果/价值）与忌词，
不核**要素顺序**与**句式形态**。而官方须知对摘要有明文要求：

  · 报道性摘要、结构严谨——「研究目的 → 研究方法 → 研究结果与结论」的
    顺序是该刊摘要的通行结构（见 references/section-guides.md 摘要节）；
  · 第三人称、不用「本文/我们」；
  · 忌「分析了/总结了/提出了/解决了」等字样；
  · 首句起笔形态：该刊摘要首句以背景/问题切入为主（「随着…」「针对…」），
    直接以「本研究…」开头占比低。

本脚本在前者基础上补三项**顺序与形态**核验：
  · 要素顺序：目的→方法→结果→价值 是否按序出现（乱序提示）；
  · 首句形态：以「本研究」直接起笔（该刊较少）→ 提示；
  · 末句收束：摘要末句应落在结论/价值，而非半截话或新问题。

用法:
    python abstract_shape.py <论文.docx> [--json]
退出码：0 = 摘要形态合规；1 = 要素乱序或官方忌词问题。
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

# 四要素信号（与 diagnose_paper.scan_abstract 同口径）
ELEM_PATS = {
    "目的": r"(旨在|为了|以期|目的|为填补|为探索|针对|面向)",
    "方法": r"(采用|运用|使用|以.{0,12}(为对象|为基础)|构建|设计|通过.{0,10}(方法|法))",
    "结果": r"(结果|表明|显示|发现|得出|实现|验证|建成|完成)",
    "价值": r"(为.{0,12}(提供|奠定|支撑)|具有|有助于|能够为|意义)",
}
ELEM_ORDER = ["目的", "方法", "结果", "价值"]
# 官方忌词（official-guidelines.md §5）
BAN_WORDS = ["分析了", "总结了", "提出了", "解决了", "本文", "我们"]
# 首句形态：直接以「本研究」起笔
FIRST_SELF = re.compile(r"^\s*本研究")


def analyze(doc):
    paras, refs, abs_text, kw_text, titles, body = dp.extract_paras(doc)
    abst = dp.scan_abstract(abs_text)
    issues, notices = [], []
    a = dp.clean(abs_text or "")
    if not a:
        return {"issues": ["未检出摘要"], "notices": [], "order": [],
                "elems": {}, "first": ""}

    # ① 要素顺序
    positions = {}
    for name in ELEM_ORDER:
        m = re.search(ELEM_PATS[name], a)
        if m:
            positions[name] = m.start()
    found = [n for n in ELEM_ORDER if n in positions]
    # 检查 found 是否按 ELEM_ORDER 的相对顺序出现
    seq = [positions[n] for n in found]
    disordered = any(seq[i] > seq[i + 1] for i in range(len(seq) - 1))
    if disordered:
        # 报出实际顺序
        actual = sorted(found, key=lambda n: positions[n])
        notices.append(f"摘要要素顺序为 {'→'.join(actual)}，"
                       f"该刊通行结构为 目的→方法→结果→价值，"
                       f"建议调整（不影响资格，但更符合报道性摘要规范）")

    # ② 官方忌词
    for w in BAN_WORDS:
        c = a.count(w)
        if c:
            issues.append(f"摘要含官方忌词「{w}」×{c}（须知明文：切忌「分析了/总结了/"
                          f"提出了/解决了」，不用「本文/我们」）")

    # ③ 首句形态
    first = ""
    for s in re.split(r"(?<=[。？！])", a):
        if s.strip():
            first = s.strip()
            break
    if FIRST_SELF.match(first):
        notices.append(f"摘要首句直接以「本研究」起笔——该刊摘要首句多为背景/问题"
                       f"切入（「随着…」「针对…」），可考虑前置问题语境")

    # ④ 末句收束：末句应含结论/价值信号
    last = ""
    sents = [s.strip() for s in re.split(r"(?<=[。？！])", a) if s.strip()]
    if sents:
        last = sents[-1]
        has_close = any(re.search(ELEM_PATS[k], last) for k in ("价值", "结果"))
        if not has_close:
            notices.append("摘要末句未见结论/价值收束信号——"
                           "报道性摘要通常以研究结论或应用价值收尾")

    return {"issues": issues, "notices": notices,
            "order": [n for n in ELEM_ORDER if n in positions],
            "elems": abst.get("elements", {}),
            "first": first[:60], "last": last[:60]}


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
    print("摘要要素顺序与句式核验（官方须知 §5 + 该刊通行结构）")
    print(f"稿件：{os.path.basename(DOC)}")
    print(f"检出要素顺序：{'→'.join(res['order']) or '（未检出）'}")
    print(f"首句：{res['first']}")
    print(f"末句：{res['last']}")
    print("=" * 74)

    if res["issues"]:
        print(f"🔴 发现 {len(res['issues'])} 项摘要问题：")
        for it in res["issues"]:
            print(f"   · {it}")
    else:
        print("✅ 摘要无官方忌词，四要素齐备。")
    if res["notices"]:
        print(f"\n🟡 建议核对 {len(res['notices'])} 项：")
        for it in res["notices"]:
            print(f"   · {it}")

    if "--json" in flags:
        jf = os.path.splitext(DOC)[0] + ".abstract_shape.json"
        with open(jf, "w", encoding="utf-8") as f:
            json.dump(res, f, ensure_ascii=False, indent=1)
        print(f"\nJSON 已写入：{jf}")

    sys.exit(1 if res["issues"] else 0)


if __name__ == "__main__":
    main()
