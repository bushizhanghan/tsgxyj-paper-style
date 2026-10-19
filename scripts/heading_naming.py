# -*- coding: utf-8 -*-
"""heading_naming.py —— 标题命名对位核验（该刊常见序列贴合度）

heading_shape.py 核标题的**形态**（字数/虚标题/末章命名），不核**命名习惯**。
该刊一级/二级标题有稳定的命名序列（见 references/section-guides.md 与
108 篇 heads 实测）。作者自创命名或过度西化命名（纯英文层级标题）会显得
不合刊风。

**词根表来自语料实测**：从 corpus_stats.json 的 heads 字段提取全部标题，
按 2–6 字 n-gram 统计频次，取出现在 ≥5 个标题里的高频词根作为
「该刊常见命名词根」——不是人工罗列的清单（人工清单必然覆盖不全，
实测初次版本漏掉「障碍类型」「需求类型」「检测」「测试」等常见词，
误报 5 处）。

本脚本把每个标题与该刊常见词根做匹配，报出「未含任何常见词根」的标题
供人工复核——**只作提示**（该刊亦有创新命名，不判错）。

用法:
    python heading_naming.py <论文.docx> [--corpus <dir>] [--json]
退出码：恒 0（提示级工具）。
"""
import os
import re
import sys
import json
from collections import Counter

# Windows 控制台 GBK 保护
if hasattr(sys.stdout, "reconfigure"):
    try:
        sys.stdout.reconfigure(encoding="utf-8", errors="replace")
    except Exception:
        pass

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import corpus  # noqa: E402
import diagnose_paper as dp  # noqa: E402

# 兜底词根（语料库缺失时用；已按实测校补过一轮）
FALLBACK_ROOTS = [
    "研究设计", "研究对象", "研究方法", "研究思路", "研究框架", "概念界定",
    "理论基础", "数据来源", "结果与分析", "实证分析", "结果分析", "检测结果",
    "需求分析", "现状", "问题", "差距", "特征", "演化", "比较", "影响因素",
    "对策", "策略", "路径", "启示", "建议", "提升", "优化", "完善", "构建",
    "保障", "机制", "结语", "结论", "讨论", "引言", "绪论", "综述", "述评",
    "研究进展", "背景", "设计", "实现", "应用", "实践", "探索", "分析",
    "研究", "评价", "服务", "体系", "模式", "模型", "框架", "技术", "方法",
    "调查", "调研", "统计", "可视化", "用户", "读者", "群体", "图书馆",
    "资源", "数据", "信息", "阅读", "素养", "空间", "数字", "智慧", "智能",
    "公共", "高校", "无障碍", "适老", "老年", "视障", "听障", "多元", "融合",
    "质量", "障碍", "类型", "场景", "检测", "测试", "验证", "功能", "模块",
    "系统", "平台", "流程", "环节", "主体", "维度", "适配", "档案",
]
MIN_TITLE_HITS = 5          # 词根至少出现在 5 个标题里才算「常见命名」
# 纯英文标题（该刊中文刊，英文标题仅用于篇名层）
EN_ONLY = re.compile(r"^[A-Za-z0-9\s:：\-—–,.'()]{8,}$")


def strip_num(t):
    m = re.match(r"^\s*(\d+(?:\.\d+){0,3})\s*(.*)$", t or "")
    return (m.group(2).strip() if m else (t or "").strip())


def build_roots(corpus_dir=None):
    """从语料 heads 字段构建「该刊常见命名词根」。

    取 2–6 字 n-gram 中出现在 ≥MIN_TITLE_HITS 个标题里的高频片段，
    并剔除被更长高频词包含的短片段（避免「研究」这类碎词刷屏）。
    """
    try:
        rows = corpus.load_corpus_stats(corpus_dir)
    except Exception:
        rows = []
    if not rows:
        return set(FALLBACK_ROOTS), "兜底词表（语料库缺失）"
    title_hits = Counter()
    for r in rows:
        heads = r.get("heads") or []
        seen = set()
        for h in heads:
            name = strip_num(re.sub(r"\s+", " ", str(h)))
            cn = re.sub(r"[^\u4e00-\u9fff]", "", name)
            if len(cn) < 2:
                continue
            for n in (6, 5, 4, 3, 2):
                for i in range(len(cn) - n + 1):
                    seen.add(cn[i:i + n])
        for w in seen:
            title_hits[w] += 1
    roots = {w for w, c in title_hits.items() if c >= MIN_TITLE_HITS}
    # 注意：这里**不做**「剔除被更长词包含的短词」——词根表是匹配用的词典，
    # 「研究方法」即使被「研究方法论」包含也应保留（标题含前者即算贴合）；
    # 早期的去碎片逻辑误删了大量有效词根，导致误报。
    if not roots:
        return set(FALLBACK_ROOTS), "兜底词表（语料无有效标题）"
    return roots, f"语料库实测（{len(roots)} 个高频词根，出现于 ≥{MIN_TITLE_HITS} 个标题）"


def analyze(doc, roots=None, roots_src=""):
    paras, refs, abs_text, kw_text, titles, body = dp.extract_paras(doc)
    odd, en_titles = [], []
    for t in titles:
        s = re.sub(r"\s+", " ", (t or "").strip())
        if not s:
            continue
        name = strip_num(s)
        if not name:
            continue
        core = dp.clean(name)
        if EN_ONLY.match(name) and re.search(r"[A-Za-z]{4,}", name):
            en_titles.append({"title": s, "name": name})
            continue
        if any(root in core for root in roots):
            continue
        odd.append({"title": s, "name": name, "len": len(core)})

    notices = []
    for o in odd:
        notices.append(f"标题「{o['title'][:26]}」未含该刊高频命名词根——"
                       f"请确认是否符合该刊命名习惯（不判错，仅提示）")
    for e in en_titles:
        notices.append(f"层级标题为纯英文「{e['name'][:24]}」——该刊为中文刊，"
                       f"层级标题宜用中文（英文仅用于篇名层）")

    return {"odd": odd, "en_titles": en_titles, "n_titles": len(titles),
            "roots_src": roots_src, "notices": notices}


def main():
    args = [a for a in sys.argv[1:] if not a.startswith("--")]
    flags = set(a for a in sys.argv[1:] if a.startswith("--"))
    if not args or args[0] in ("-h", "--help"):
        print(__doc__)
        sys.exit(0)
    DOC = args[0]
    if not os.path.isfile(DOC):
        sys.exit(f"找不到文件：{DOC}")
    _corpus_arg = None
    if "--corpus" in sys.argv:
        i = sys.argv.index("--corpus")
        _corpus_arg = sys.argv[i + 1] if i + 1 < len(sys.argv) else None
    try:
        cdir = corpus.find_corpus_dir(_corpus_arg)
    except FileNotFoundError as e:
        print(f"[错误] {e}")
        sys.exit(2)

    roots, roots_src = build_roots(cdir)
    res = analyze(DOC, roots, roots_src)
    print("=" * 74)
    print("标题命名对位核验（该刊常见命名序列贴合度 · 提示级）")
    print(f"稿件：{os.path.basename(DOC)}")
    print(f"标题 {res['n_titles']} 个 ｜ 未含高频词根 {len(res['odd'])} 个 ｜ "
          f"纯英文层级标题 {len(res['en_titles'])} 个")
    print(f"词根来源：{res['roots_src']}")
    print("=" * 74)

    if res["notices"]:
        print(f"🟡 建议核对 {len(res['notices'])} 项：")
        for it in res["notices"]:
            print(f"   · {it}")
    else:
        print("✅ 全部标题均含该刊高频命名词根。")

    if "--json" in flags:
        jf = os.path.splitext(DOC)[0] + ".heading_naming.json"
        with open(jf, "w", encoding="utf-8") as f:
            json.dump(res, f, ensure_ascii=False, indent=1)
        print(f"\nJSON 已写入：{jf}")

    sys.exit(0)


if __name__ == "__main__":
    main()
