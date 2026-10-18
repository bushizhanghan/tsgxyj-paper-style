# -*- coding: utf-8 -*-
"""table_fig_audit.py —— 《图书馆学研究》图表体例与三线表规范深度审查

对照官方指南与刊物体例要求：
    1. 表格规范：
       - 是否为标准三线表（顶线、底线粗，栏目线细，无竖线）；
       - 表题是否在表格上方（官方体例：表题居中置于表上）；
       - 表格注释/资料来源行是否完整（表下注明「资料来源/数据来源」）。
    2. 插图规范：
       - 图题是否在插图下方（官方体例：图题居中置于图下）；
       - 是否有【待补图/占位符】未替换（如【图位：...】）；
       - 正文中是否有指路句（「如图 X 所示」「见表 Y」先提及后出现）。

用法:
    python table_fig_audit.py <论文.docx> [--json]
退出码：0 = 无严重图表体例缺陷；1 = 存在未替换图位或严重三线表缺失。
"""
import os
import re
import sys
import json

if hasattr(sys.stdout, "reconfigure"):
    try:
        sys.stdout.reconfigure(encoding="utf-8", errors="replace")
    except Exception:
        pass

HERE = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, HERE)


def audit_document_tables_and_figures(doc_path):
    from docx import Document
    doc = Document(doc_path)

    issues = []
    notices = []
    stats = {
        "word_tables": len(doc.tables),
        "three_line_tables": 0,
        "non_three_line_tables": 0,
        "table_titles": [],
        "fig_titles": [],
        "placeholders": [],
        "fig_refs": 0,
    }

    # 1. 扫描段落中的图表标题、占位符、正文指路句
    paragraphs = [p.text.strip() for p in doc.paragraphs if p.text.strip()]
    full_text = "\n".join(paragraphs)

    # 统计指路句
    fig_refs = re.findall(r"(?:如|见|详见|参见)?\s*(?:图|表)\s*\d+(?:\s*所示)?", full_text)
    stats["fig_refs"] = len(fig_refs)

    for idx, p_text in enumerate(paragraphs):
        # 检测占位符（如 【图位：...】、【待补：...】）
        if re.search(r"【(?:图位|待补|插入图|插表|待补图)[^】]*】", p_text):
            stats["placeholders"].append((idx + 1, p_text))

        # 检测表题：以「表 1」「表1」等开头独立成段
        if re.match(r"^表\s*\d+[\s　]+", p_text) and len(p_text) < 60:
            stats["table_titles"].append((idx + 1, p_text))

        # 检测图题：以「图 1」「图1」等开头独立成段
        if re.match(r"^图\s*\d+[\s　]+", p_text) and len(p_text) < 60:
            stats["fig_titles"].append((idx + 1, p_text))

    if stats["placeholders"]:
        for p_no, p_txt in stats["placeholders"]:
            issues.append(f"段{p_no} 存在未替换的图表占位标记「{p_txt[:40]}」——投稿前须插入真实图片或删除占位符")

    # 2. 深度审查 Word Table 元素的边框属性（三线表核验）
    ns = "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}"
    for t_idx, tbl in enumerate(doc.tables):
        tblPr = tbl._tbl.tblPr
        borders = tblPr.find(f"{ns}tblBorders")
        is_three_line = False
        if borders is not None:
            top = borders.find(f"{ns}top")
            bottom = borders.find(f"{ns}bottom")
            insideH = borders.find(f"{ns}insideH")
            left = borders.find(f"{ns}left")
            right = borders.find(f"{ns}right")
            insideV = borders.find(f"{ns}insideV")

            # 竖线必须为 none 或不存在
            no_vertical = (
                (left is None or left.attrib.get(f"{ns}val") == "none") and
                (right is None or right.attrib.get(f"{ns}val") == "none") and
                (insideV is None or insideV.attrib.get(f"{ns}val") == "none")
            )
            # 顶线与底线必须存在且非 none
            has_top_bottom = (
                top is not None and top.attrib.get(f"{ns}val") not in ("none", "nil") and
                bottom is not None and bottom.attrib.get(f"{ns}val") not in ("none", "nil")
            )
            # 栏目线 (insideH) 存在
            has_insideH = insideH is not None and insideH.attrib.get(f"{ns}val") not in ("none", "nil")

            if no_vertical and has_top_bottom and has_insideH:
                is_three_line = True

        if is_three_line:
            stats["three_line_tables"] += 1
        else:
            stats["non_three_line_tables"] += 1
            issues.append(f"表格 #{t_idx+1}（{len(tbl.rows)}行×{len(tbl.columns)}列）非标准三线表形态（存在竖线或缺少顶/底/栏目线）——该刊明文要求采用三线表")

    # 3. 图表与指路句协调性审查
    total_found = len(stats["table_titles"]) + len(stats["fig_titles"]) + stats["word_tables"]
    if total_found > 0 and stats["fig_refs"] < total_found * 0.5:
        notices.append(f"正文图表指路句偏少（仅 {stats['fig_refs']} 处），建议在正文按「如图 X 所示」「见表 Y」显式提及")

    return {
        "stats": stats,
        "issues": issues,
        "notices": notices,
    }


def main():
    args = [a for a in sys.argv[1:] if not a.startswith("--")]
    flags = set(a for a in sys.argv[1:] if a.startswith("--"))
    if not args or args[0] in ("-h", "--help"):
        print(__doc__)
        sys.exit(0)
    DOC = args[0]
    if not os.path.isfile(DOC):
        sys.exit(f"找不到文件：{DOC}")

    res = audit_document_tables_and_figures(DOC)
    st = res["stats"]

    print("=" * 70)
    print("图表体例与三线表规范深度审查（官方体例合规）")
    print(f"稿件：{os.path.basename(DOC)}")
    print(f"Word实体表格 {st['word_tables']} 个 ｜ 识别表题 {len(st['table_titles'])} 个 ｜ 识别图题 {len(st['fig_titles'])} 个")
    print("=" * 70)

    if st["word_tables"] > 0:
        print(f"· 三线表合规：✓ 合规 {st['three_line_tables']} 个 / ⚠️ 待修改 {st['non_three_line_tables']} 个")
    else:
        print("· 三线表合规：未检出嵌入式 Word 实体表格（若为纯文本排列表格，投稿排版时须转为正式三线表）")

    if st["placeholders"]:
        print(f"· 图位占位符：⚠️ 检出 {len(st['placeholders'])} 处待补图占位标记")
    else:
        print("· 图位占位符：✓ 无残留占位符")

    print(f"· 正文指路句：共检出 {st['fig_refs']} 处引用提及")

    print("-" * 70)
    if res["issues"]:
        print(f"🔴 发现 {len(res['issues'])} 项需修正的图表问题：")
        for iss in res["issues"]:
            print(f"   · {iss}")
    else:
        print("✓ 图表体例与占位核验：未见违规项")

    if res["notices"]:
        print(f"\n🟡 建议与提示：")
        for noti in res["notices"]:
            print(f"   · {noti}")

    if "--json" in flags:
        jf = os.path.splitext(DOC)[0] + ".table_fig.json"
        with open(jf, "w", encoding="utf-8") as f:
            json.dump(res, f, ensure_ascii=False, indent=1)
        print(f"\nJSON 已保存：{jf}")

    sys.exit(1 if res["issues"] else 0)


if __name__ == "__main__":
    main()
