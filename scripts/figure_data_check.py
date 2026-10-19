# -*- coding: utf-8 -*-
"""figure_data_check.py —— 图表数据与正文数字一致性核验

number_consistency.py 核的是「摘要 ↔ 正文 ↔ 结语」的数字，但读不到**表格单元格**
与**图注**里的数字。而图表与正文数字打架是最伤可信度的一类错误：
表格合计 698 条、正文写 751 条；图注说 20 道题、正文说 20 题次；表格
通过率 99.5%、正文写 99.6%——审稿人一对照就质疑数据可靠性。

本脚本把三处数字放在一起核对：
  · 表格单元格：读 docx 真实表格（含合并单元格文本），抽取带单位数值；
  · 图注/表注：图题、表题、资料来源行里的数值；
  · 正文：与上述「同名指标」（单位 + 前接名词）比对。

复用 number_consistency 的断言抽取与配对逻辑（import，不复制代码），
只把「数据来源区域」从纯段落扩展到表格单元格。

用法:
    python figure_data_check.py <论文.docx> [--top 20] [--json]
退出码：0 = 未发现图表与正文数字冲突；1 = 存在冲突（🔴）。

局限（词面配对启发式）：
  · 表格中的行列结构不做语义理解，同名指标靠「单位 + 前接名词」配对；
  · 同一表格内多行同单位数值（各模型通过率均 100%）不判冲突（正确情形）；
  · 抽取不到图表（纯文本排版）时明确提示，不静默通过。
"""
import os
import re
import sys
import json
from collections import defaultdict

# Windows 控制台 GBK 保护
if hasattr(sys.stdout, "reconfigure"):
    try:
        sys.stdout.reconfigure(encoding="utf-8", errors="replace")
    except Exception:
        pass

HERE = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, HERE)
import diagnose_paper as dp  # noqa: E402
import number_consistency as nc  # noqa: E402  # 复用断言抽取/配对逻辑

CAP_RE = re.compile(r"^\s*(图|表)\s*\d+")   # 图表标题行


def read_tables(doc_path):
    """读 docx 真实表格，返回 [{'idx', 'caption', 'rows':[[cell...]], 'flat':文本}]。"""
    from docx import Document
    doc = Document(doc_path)
    out = []
    # 先收集全文段落顺序，用于给表格找最近的表题
    paras = [p.text.strip() for p in doc.paragraphs]
    for ti, tbl in enumerate(doc.tables, 1):
        rows = []
        for row in tbl.rows:
            cells = []
            for c in row.cells:
                cells.append((c.text or "").strip())
            rows.append(cells)
        flat = "\n".join(" | ".join(r) for r in rows)
        # 找表题：段落中以「表 N」开头的最近一条（表格前后 3 段内）
        caption = ""
        for i, t in enumerate(paras):
            if re.match(r"^表\s*\d+\s*[\s　]", t) and len(t) < 80:
                caption = t
                break
        out.append({"idx": ti, "caption": caption, "rows": rows, "flat": flat})
    return out


def collect_fig_captions(doc_path):
    """收集图题/表题/资料来源注（段落中独立的图表标题行）。"""
    from docx import Document
    doc = Document(doc_path)
    caps = []
    for p in doc.paragraphs:
        t = (p.text or "").strip()
        if not t:
            continue
        if CAP_RE.match(t) and len(t) < 100:
            caps.append(t)
        elif re.match(r"^(资料来源|数据来源|注)[：:]", t) and len(t) < 200:
            caps.append(t)
    return caps


def extract_assertions_from_text(text, region, include_claim_units=False):
    """把一段任意文本当作一个「段落」做断言抽取。

    nc.extract_assertions 会跳过 CLAIM_UNITS（类/种/项/条/方面/维度）里的单位，
    因为正文里「三条建议」属条目数声称、不走数值比对。但**表格里的「698 条」
    是实测值**，必须参与比对——故本函数用 include_claim_units=True 时，
    把 CLAIM_UNITS 临时清空，让这些单位也抽成断言。
    """
    if not include_claim_units:
        return nc.extract_assertions([text], region)
    saved = nc.CLAIM_UNITS
    try:
        nc.CLAIM_UNITS = set()
        return nc.extract_assertions([text], region)
    finally:
        nc.CLAIM_UNITS = saved


def key_of(a):
    """断言的比对键：单位 + 归一化前接名词。"""
    h = nc.norm_head(a.get("head") or "")
    return (a.get("unit"), h)


def analyze(doc_path, top=20):
    tables = read_tables(doc_path)
    caps = collect_fig_captions(doc_path)
    reg = nc.split_regions(doc_path)

    body_as = nc.extract_assertions(reg["body"], "正文")
    # 正文侧补充：CLAIM_UNITS 单位（条/项/类…）在图表核验里也是实测值，
    # 但只在正文中「与图表同名指标」时才需要——先全量抽一遍供配对
    claim_body_as = extract_assertions_from_text("\n".join(reg["body"]), "正文",
                                                include_claim_units=True)
    abs_as = nc.extract_assertions([reg["abstract"]] if reg["abstract"] else [], "摘要")
    all_body = abs_as + body_as + claim_body_as

    # 图表侧断言（表格内「698 条」属实测值，须含 CLAIM_UNITS 单位）
    fig_as = []
    for tb in tables:
        fig_as += extract_assertions_from_text(tb["flat"], f"表{tb['idx']}",
                                              include_claim_units=True)
    for c in caps:
        fig_as += extract_assertions_from_text(c, "图表注", include_claim_units=True)

    # 按「单位+名词」建索引
    body_idx = defaultdict(list)
    for a in all_body:
        body_idx[key_of(a)].append(a)
    fig_idx = defaultdict(list)
    for a in fig_as:
        fig_idx[key_of(a)].append(a)

    conflicts, ok_pairs, weak = [], [], []
    for k, fa_list in fig_idx.items():
        if k not in body_idx:
            continue
        unit, head = k
        bvals = {nc.fmt_val(a["value"]) for a in body_idx[k]}
        fvals = {nc.fmt_val(a["value"]) for a in fa_list}
        if not bvals or not fvals:
            continue
        inter = bvals & fvals
        if inter:
            ok_pairs.append({"key": k, "values": sorted(inter)})
            continue
        # 数值集合完全不相交 → 冲突候选
        rec = {
            "unit": unit, "head": head or "（无明确名词）",
            "fig_vals": sorted(fvals), "body_vals": sorted(bvals),
            "fig_regions": sorted({a["region"] for a in fa_list}),
            "body_regions": sorted({a["region"] for a in body_idx[k]}),
            "fig_samples": [a.get("raw") or a.get("text") for a in fa_list[:3]],
            "body_samples": [a.get("raw") or a.get("text") for a in body_idx[k][:3]],
        }
        # 模糊量（余/多/左右）造成的差异 → 弱提示
        fuzzy = any(a.get("fuzzy") for a in fa_list + body_idx[k])
        (weak if fuzzy else conflicts).append(rec)

    return {"tables": len(tables), "captions": len(caps),
            "fig_assertions": len(fig_as), "body_assertions": len(all_body),
            "conflicts": conflicts, "weak": weak, "ok": ok_pairs,
            "table_details": [{"idx": t["idx"], "caption": t["caption"],
                               "rows": len(t["rows"])} for t in tables]}


def fmt_vals(vs, unit):
    return "、".join(f"{v}{unit}" for v in vs)


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

    res = analyze(DOC, top=TOP)
    print("=" * 72)
    print("图表数据与正文数字一致性核验")
    print(f"稿件：{os.path.basename(DOC)}")
    print(f"真实表格 {res['tables']} 个 ｜ 图表标题/注释 {res['captions']} 条 ｜ "
          f"图表侧数值断言 {res['fig_assertions']} 条 ｜ 正文侧 {res['body_assertions']} 条")
    print("=" * 72)

    if res["tables"] == 0 and res["captions"] == 0:
        print("未检出真实表格或图表标题（若为纯文本排版的图表，本项无法核验）。")
        print("建议：将表格转为 Word 三线表后再核，或人工对照图表与正文数字。")
        sys.exit(0)
    if res["fig_assertions"] == 0:
        print("图表侧未抽取到带单位的数值（可能仅为图形或纯文字表头），本项不适用。")
        sys.exit(0)

    if res["conflicts"]:
        print(f"🔴 发现 {len(res['conflicts'])} 项图表与正文数字冲突：")
        for i, c in enumerate(res["conflicts"][:TOP], 1):
            print(f"  {i}. 「{c['head']}」")
            print(f"     图表：{fmt_vals(c['fig_vals'], c['unit'])}"
                  f"（{('、'.join(c['fig_regions']))[:40]}）")
            print(f"     正文：{fmt_vals(c['body_vals'], c['unit'])}"
                  f"（{('、'.join(c['body_regions']))[:40]}）")
            if c["fig_samples"]:
                print(f"     图表原文：{nc.trunc(str(c['fig_samples'][0]), 70)}")
            if c["body_samples"]:
                print(f"     正文原文：{nc.trunc(str(c['body_samples'][0]), 70)}")
    else:
        print("✅ 未发现图表与正文的数字冲突"
              + (f"（{len(res['ok'])} 组同名指标数值一致）" if res["ok"] else ""))
    if res["weak"]:
        print(f"\n🟡 含模糊量（余/多/左右）的差异 {len(res['weak'])} 项，建议人工核对：")
        for c in res["weak"][:TOP]:
            print(f"   · 「{c['head']}」图表 {fmt_vals(c['fig_vals'], c['unit'])} "
                  f"vs 正文 {fmt_vals(c['body_vals'], c['unit'])}")

    if "--json" in flags:
        jf = os.path.splitext(DOC)[0] + ".figure_data.json"
        with open(jf, "w", encoding="utf-8") as f:
            json.dump(res, f, ensure_ascii=False, indent=1, default=str)
        print(f"\nJSON 已写入：{jf}")

    sys.exit(1 if res["conflicts"] else 0)


if __name__ == "__main__":
    main()
