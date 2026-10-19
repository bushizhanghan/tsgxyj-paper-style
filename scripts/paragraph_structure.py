# -*- coding: utf-8 -*-
"""paragraph_structure.py —— 段内结构范式核验（「一段只说一件事」专查）

paragraphing-guide.md 记载了该刊段内结构的实测范式，但 diagnose_paper 的
段落级扫描只查**长度**（句均偏短/超长/碎段），不查**结构**。而段落偏碎、
读起来散的深层原因，往往是「一段没说透一件事」：

  · 起句靠连接词开场（「然而……」「此外……」）——该刊衔接词起句仅 2%，
    主流是直接陈述（83%），衔接靠内容回指；
  · 段末用「综上所述」式套话收尾——该刊罕见，段末应落在具体判断上；
  · 一段塞进多个不相关主题（段内出现多个互异的「X 方面/层面」分主题标记，
    或多个并列小标题式引导）——该刊 2–6 句/段、一条线索一段；
  · 并列信息未用分号连接（句内分句过长，一口气到底）。

本脚本按上述范式逐段核验，报出**段号 + 原句起头 + 违反的范式 + 该刊实测比例**，
让你知道「不是字数问题，是段内结构问题」。

用法:
    python paragraph_structure.py <论文.docx> [--top 20] [--json]
退出码：0 = 未检出范式违反；1 = 检出（提示级为主，不影响「红线」判定）。

说明：范式比例取自 references/paragraphing-guide.md 四（118 篇实测），
      属「实际做法」统计；官方未对段内结构作规定，故本项一律为提示级。
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
import argument_check as ac  # noqa: E402

MIN_PARA = 40          # 参与核验的最短段（字）
# ④ 段内分句过长（该刊句内分句 16–21 字；实测该稿 P90 约 78 字，
#    阈值取 P90 上浮一档，避免把正常长句全判成问题）
LONG_CLAUSE = 95

# ③ 「一段多主题」确定性信号：段内出现多个互异的「X 方面/层面」式分主题标记，
#    或多个并列小标题式引导（「其一…其二…」）。
#    说明：曾用「主题词分散度」算法（Jaccard 距离）实现，实测在该稿上中位数即
#    1.0、无判别力——无分词器时任意两句中文的 n-gram 实体词都不重叠，算法路线
#    不可行，故改为确定性规则。
ASPECT_MARK = re.compile(
    r"(?:^|[。；，,、在从就]\s*)([\u4e00-\u9fff]{2,4})(?:方面|层面|维度|角度)\b")
ORDINAL_MARK = re.compile(
    r"^\s*(其一|其二|其三|其四|其五|第一|第二|第三|第四|第五|一是|二是|三是|四是|五是)")

# ① 连接词起句（该刊仅 2%）
CONN_HEAD = re.compile(
    r"^\s*(然而|但是|但|此外|另外|同时|因此|所以|从而|进而|综上所述|总之|"
    r"与此同期|与此同时|在此基础上|值得一提|不难看出|显而易见)")
# ② 套话收尾
CLICHE_TAIL = re.compile(
    r"(综上所述|总而言之|总之|总的来说|不难看出|由此可见|可见|"
    r"具有重要意义|值得深入探讨|提供有益参考)[。！]?\s*$")


def split_sents(text):
    return [s.strip() for s in re.split(r"(?<=[。？！])", text or "") if s.strip()]


def multi_topic_signals(sents):
    """检测「一段多主题」的确定性信号。

    返回 [(类型, 详情)]：
      · aspect：段内出现 ≥3 个互异的「X 方面/层面」标记
      · ordinal：段内 ≥3 句以并列小标题式引导（其一/第一/一是…）开头
    """
    out = []
    # ① 多个互异「X 方面」标记
    aspects = set()
    for s in sents:
        for m in ASPECT_MARK.finditer(s):
            aspects.add(m.group(1))
    if len(aspects) >= 3:
        out.append(("aspect", "、".join(sorted(aspects)[:4])))
    # ② 多句并列小标题式引导
    n_ord = sum(1 for s in sents if ORDINAL_MARK.match(s))
    if n_ord >= 3:
        out.append(("ordinal", f"{n_ord} 句以「其一/第一/一是」式开头"))
    return out


def analyze(doc):
    paras, refs, abs_text, kw_text, titles, body = dp.extract_paras(doc)

    findings = {
        "conn_head": [],      # 连接词起句
        "cliche_tail": [],    # 套话收尾
        "topic_drift": [],    # 主题分散
        "long_clause": [],    # 句内分句过长
    }
    for pi, raw in enumerate(body):
        t = dp.clean(raw or "")
        if len(t) < MIN_PARA:
            continue
        pno = pi + 1
        sents = split_sents(t)

        # ① 连接词起句
        if sents and CONN_HEAD.search(sents[0]):
            findings["conn_head"].append({"para": pno, "head": sents[0][:50]})

        # ② 套话收尾
        if sents and CLICHE_TAIL.search(t):
            findings["cliche_tail"].append({"para": pno, "tail": t[-40:]})

        # ③ 一段多主题（确定性信号）
        if len(sents) >= 4:
            sigs = multi_topic_signals(sents)
            if sigs:
                findings["topic_drift"].append(
                    {"para": pno, "n_sents": len(sents),
                     "signals": sigs, "head": sents[0][:40]})

        # ④ 句内分句过长（按分号切，单分句超 60 字）
        for si, s in enumerate(sents, 1):
            parts = [p for p in re.split(r"[；;]", s) if p.strip()]
            for p in parts:
                if len(dp.clean(p)) > LONG_CLAUSE:
                    findings["long_clause"].append(
                        {"para": pno, "sent": si, "clause": dp.clean(p)[:56]})
                    break

    # 汇总
    n_conn = len(findings["conn_head"])
    n_cliche = len(findings["cliche_tail"])
    n_drift = len(findings["topic_drift"])
    n_long = len(findings["long_clause"])
    n_paras = sum(1 for p in body if len(dp.clean(p or "")) >= MIN_PARA)
    return {"findings": findings, "n_paras": n_paras,
            "counts": {"conn_head": n_conn, "cliche_tail": n_cliche,
                       "topic_drift": n_drift, "long_clause": n_long}}


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
    c = res["counts"]
    print("=" * 74)
    print("段内结构范式核验（该刊 118 篇实测范式 · 提示级）")
    print(f"稿件：{os.path.basename(DOC)}")
    print(f"核验 {res['n_paras']} 段（≥{MIN_PARA} 字）")
    print("=" * 74)
    print(f"合计：连接词起句 {c['conn_head']} ｜ 套话收尾 {c['cliche_tail']} ｜ "
          f"主题分散 {c['topic_drift']} ｜ 分句过长 {c['long_clause']}")

    def show(title, items, fmt, hint):
        if not items:
            return
        print(f"\n【{title}】（{hint}）")
        for it in items[:TOP]:
            print("  " + fmt(it))

    show("连接词起句", res["findings"]["conn_head"],
         lambda it: f"段{it['para']}：{it['head']}",
         "该刊起句直接陈述占 83%，衔接词起句仅 2%——衔接靠内容回指，不靠连接词")
    show("套话收尾", res["findings"]["cliche_tail"],
         lambda it: f"段{it['para']}段末：…{it['tail']}",
         "该刊段末落在具体判断上，「综上所述」式套话罕见")
    show("一段多主题", res["findings"]["topic_drift"],
         lambda it: f"段{it['para']}（{it['n_sents']} 句，"
                    f"{'；'.join(f'{k}={v}' for k, v in it['signals'])}）：{it['head']}…",
         "该刊 2–6 句/段、一条线索一段；段内出现多个分主题标记说明塞了多主题，宜切段")
    show("句内分句过长", res["findings"]["long_clause"],
         lambda it: f"段{it['para']}第{it['sent']}句：{it['clause']}…",
         f"该刊句内分句 16–21 字；单分句 >{LONG_CLAUSE} 字宜用分号断开或另起一句")

    total = sum(c.values())
    print("\n" + "-" * 74)
    if total == 0:
        print("✅ 未检出段内结构范式违反。")
    else:
        print(f"🟡 共 {total} 处提示级问题——不影响投稿资格，但按该刊范式调整后")
        print("   文章会更贴近刊风（改写时优先处理「主题分散段」：切段比改词更有效）。")

    if "--json" in flags:
        jf = os.path.splitext(DOC)[0] + ".paragraph_structure.json"
        with open(jf, "w", encoding="utf-8") as f:
            json.dump(res, f, ensure_ascii=False, indent=1)
        print(f"\nJSON 已写入：{jf}")

    # 提示级工具：不设退出码 1（避免把风格建议当硬伤）
    sys.exit(0)


if __name__ == "__main__":
    main()
