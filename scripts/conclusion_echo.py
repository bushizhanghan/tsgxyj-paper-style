# -*- coding: utf-8 -*-
"""conclusion_echo.py —— 摘要与结语的呼应核验（「结论与摘要对不上」专查）

审稿人读稿的常规动作：读摘要 → 读结论，看**两处承诺的是否一致**。
摘要说「构建了三类需求框架」，结语却说「构建了四维模型」；摘要称
「提出对策」，结语却只字不提对策——这类「摘要—结语脱节」在初审
与复审都容易被指出，属于低成本高损失的失分点。

本脚本把摘要与结语放在一起核：
  · 承诺抽取：从摘要抽「构建/提出/设计/实现/验证 + 数量 + 名词」式贡献句，
     从结语抽同样形态的句子；
  · 数量一致性：同类承诺的数量词（三类/四维/八个/两项…）是否一致；
  · 覆盖度：摘要承诺的每项贡献，在结语是否有对应表述（主题词重叠）；
  · 结语纪律：结语是否引入新数据/新引用/新论点（该刊要求结语只做收束）。

用法:
    python conclusion_echo.py <论文.docx> [--top 10] [--json]
退出码：0 = 摘要与结语呼应一致；1 = 检出数量矛盾或承诺未兑现。

局限（词面启发式）：
  · 同义改写（「框架」vs「模型」）检不出，需人工判断；
  · 贡献句的识别靠句式信号，非常规句式可能漏抽；
  · 本工具只核「摘要↔结语」，摘要承诺在正文的兑现由 argument_check 等核验。
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

HERE = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, HERE)
import diagnose_paper as dp  # noqa: E402
import number_consistency as nc  # noqa: E402

# 贡献承诺句式：动词 + （数量 + 名词）。
# 数量允许与动词间隔少量成分（如「构建了『感知—适配—帮扶—反馈』四环节服务机制」），
# 故在动词后窗口内就近匹配「数量词 + 类别名词」。
PROMISE_VERB = re.compile(
    r"(构建|提出|设计|实现|建立|形成|完成|验证|论证|揭示|归纳|厘清|界定|"
    r"搭建|打造|创建|给出|明确)")
QUANT_RE = re.compile(
    r"([一二两三四五六七八九十]\s*[类种项维个条方面度层步骤环节套]|"
    r"\d+\s*[类种项维个条方面度层步骤环节套])")
QUANT_UNIT = "类种项维个条方面度层步骤环节套"
VERB_WINDOW = 16   # 动词后最多向前找多少个字内的数量词
# 结语新内容信号（该刊要求结语只做收束，不引入新数据/新引用/新论点）
NEW_CONTENT_RE = [
    (r"\[\d{1,3}(?:\s*[-–—,，]\s*\d{1,3})*\]", "结语出现引用标注"),
    (r"\d+(?:\.\d+)?%", "结语出现新百分比"),
    (r"研究发现|结果表明|结果显示|调查发现", "结语引入新发现陈述"),
    (r"建议(?:后续|未来|今后|进一步)?(?:研究|工作)?(?:可|应|须)", "结语出现新建议"),
]


def clean_sent(s):
    return dp.clean(s or "")


def extract_promises(text):
    """从文本抽贡献承诺：[(动词, 数量+名词 or None, 原句)]。

    数量词在动词后 VERB_WINDOW 字窗口内就近取（允许中间隔引号/定语）。
    """
    out = []
    for s in re.split(r"[。；]", text or ""):
        s = s.strip()
        if len(s) < 8:
            continue
        for vm in PROMISE_VERB.finditer(s):
            verb = vm.group(1)
            window = s[vm.end():vm.end() + VERB_WINDOW]
            qm = QUANT_RE.search(window)
            quant = re.sub(r"\s+", "", qm.group(1)) if qm else None
            out.append({"verb": verb, "quant": quant, "sent": s})
    return out


def split_regions(path):
    paras, refs, abs_text, kw_text, titles, body = dp.extract_paras(path)
    concl, in_closing = [], False
    for style, t in paras:
        if not t:
            continue
        if t.startswith("参考文献"):
            break
        s = re.sub(r"\s+", " ", t.strip())
        is_title = (re.match(r"^(0|1|2|3|4|5|6|7)\s*[^\d]", s) and len(s) < 40) or \
                   (style.startswith("Heading") and len(s) < 40)
        if is_title:
            in_closing = bool(re.search(r"结\s*语|结\s*论|结束语|小\s*结", s))
            continue
        if in_closing:
            concl.append(t)
    return abs_text or "", "\n".join(concl), titles, body


def analyze(doc):
    abs_text, concl_text, titles, body = split_regions(doc)
    abs_c, concl_c = dp.clean(abs_text), dp.clean(concl_text)

    abs_p = extract_promises(abs_c)
    concl_p = extract_promises(concl_c)

    # ① 数量一致性：摘要有数量的承诺 vs 结语同动词的数量
    issues, notices = [], []
    for ap in abs_p:
        if not ap["quant"]:
            continue
        num = ap["quant"]
        for cp in concl_p:
            if cp["verb"] == ap["verb"] and cp["quant"] and cp["quant"] != num:
                # 只在该动词+名词族相同时报（粗粒度：同动词）
                issues.append(
                    f"数量不一致：摘要「{ap['verb']}…{num}」"
                    f"vs 结语「{cp['verb']}…{cp['quant']}」")
                break

    # ② 覆盖度：摘要带数量的承诺，在结语是否有对应句
    #    「对应」= 同动词 且（同数量 或 主题词有重叠）
    covered, uncovered = [], []
    import argument_check as _ac
    for ap in abs_p:
        if not ap["quant"]:
            continue
        key = f"{ap['verb']}…{ap['quant']}"
        hit = None
        for cp in concl_p:
            if cp["verb"] != ap["verb"]:
                continue
            if cp["quant"] == ap["quant"]:
                hit = cp
                break
            # 同动词但结语无数量 → 主题词重叠也只是「弱对应」，
            # 数量承诺未兑现，报提示级（不计 🔴，避免误伤合理简写）
            at = _ac.topic_vec(ap["sent"])
            ct = _ac.topic_vec(cp["sent"])
            if _ac.overlap_score(at, ct) >= 1.0:
                hit = "weak"
                break
        if hit and hit != "weak":
            covered.append(key)
        elif hit == "weak":
            covered.append(key)
            notices.append(
                f"摘要「{key}」在结语有同主题表述但未复述数量"
                f"（如「四环节」vs「服务机制」）——审稿人对照时可能视为简写，"
                f"若属强调重点建议补上")
        else:
            uncovered.append(ap)
    for ap in uncovered:
        issues.append(
            f"摘要承诺「{ap['verb']}…{ap['quant']}」在结语无对应表述——"
            f"审稿人对照摘要与结论时易指出「前后不一致」")

    # ③ 结语纪律：新数据/新引用/新论点
    for pat, desc in NEW_CONTENT_RE:
        if re.search(pat, concl_c):
            notices.append(f"{desc}——该刊结语只做收束，宜删或移入正文")

    # ④ 结语是否过短/缺失
    if not concl_c:
        issues.append("未检出结语段——该刊五段式收口必备")
    elif len(concl_c) < 80:
        notices.append(f"结语仅 {len(concl_c)} 字，偏短（该刊结语占正文中位 4.1%，"
                       f"通常 200–600 字）")

    return {"abs_promises": abs_p, "concl_promises": concl_p,
            "covered": covered, "uncovered": uncovered,
            "issues": issues, "notices": notices,
            "abs_len": len(abs_c), "concl_len": len(concl_c)}


def main():
    args = [a for a in sys.argv[1:] if not a.startswith("--")]
    flags = set(a for a in sys.argv[1:] if a.startswith("--"))
    if not args or args[0] in ("-h", "--help"):
        print(__doc__)
        sys.exit(0)
    DOC = args[0]
    if not os.path.isfile(DOC):
        sys.exit(f"找不到文件：{DOC}")
    TOP = 10
    for i, a in enumerate(sys.argv):
        if a == "--top" and i + 1 < len(sys.argv):
            try:
                TOP = int(sys.argv[i + 1])
            except ValueError:
                pass

    res = analyze(DOC)
    print("=" * 74)
    print("摘要与结语呼应核验（「结论与摘要对不上」专查）")
    print(f"稿件：{os.path.basename(DOC)}")
    print(f"摘要 {res['abs_len']} 字 ｜ 结语 {res['concl_len']} 字")
    print(f"贡献承诺：摘要 {len(res['abs_promises'])} 处 ／ 结语 "
          f"{len(res['concl_promises'])} 处")
    print("=" * 74)

    if res["abs_promises"]:
        print("【摘要中的贡献承诺】")
        for p in res["abs_promises"][:TOP]:
            q = p["quant"] or "（无数量）"
            print(f"  · {p['verb']} {q}")
    if res["concl_promises"]:
        print("\n【结语中的贡献承诺】")
        for p in res["concl_promises"][:TOP]:
            q = p["quant"] or "（无数量）"
            print(f"  · {p['verb']} {q}")

    print("\n" + "-" * 74)
    if res["issues"]:
        print(f"🔴 发现 {len(res['issues'])} 项呼应问题：")
        for it in res["issues"]:
            print(f"   · {it}")
    else:
        print("✅ 摘要与结语的贡献承诺无数量矛盾、结语存在。")
    if res["notices"]:
        print(f"\n🟡 建议核对 {len(res['notices'])} 项：")
        for it in res["notices"]:
            print(f"   · {it}")

    if "--json" in flags:
        jf = os.path.splitext(DOC)[0] + ".conclusion_echo.json"
        with open(jf, "w", encoding="utf-8") as f:
            json.dump(res, f, ensure_ascii=False, indent=1)
        print(f"\nJSON 已写入：{jf}")

    sys.exit(1 if res["issues"] else 0)


if __name__ == "__main__":
    main()
