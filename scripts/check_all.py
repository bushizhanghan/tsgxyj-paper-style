# -*- coding: utf-8 -*-
"""check_all.py —— 投稿前一键总检（模式 A + D + 润色 + E + G + 论证链 + 数字一致性 + 术语漂移 + 引文关联）

按投稿前自查的实用顺序，串起十二个工具，各自独立跑、互不影响：
    ① GB/T 7714 文献要素核验（ref_format.py）
    ② 引用标注与对应核验（verify_refs.py，模式 D）
    ③ 引文结构（时效/类型/来源）（citation_analysis.py）
    ④ 论文诊断（diagnose_paper.py，模式 A）
    ⑤ 整稿语言体检（language_audit.py，润色词典落地）
    ⑥ 模拟审稿 13 维打分（mock_review.py，模式 E）
    ⑦ 对标评估（benchmark_paper.py，模式 G）
    ⑧ 论证链对仗核对（argument_check.py，发现↔对策逐条匹配）
    ⑨ 图表体例与三线表规范（table_fig_audit.py）
    ⑩ 摘要—正文数字一致性（number_consistency.py，红线级）
    ⑪ 术语漂移检测（term_consistency.py，同义异形词混用，待办级）
    ⑫ 引文—论点关联（citation_quality.py，装饰性引用，待办级）
    ⑬ 稿内自我重复（self_repeat.py，查重前哨，待办级）
    ⑭ 英文摘要与关键词对称性（english_audit.py，官方中英文对应）
    ⑮ 标题编号体例（heading_audit.py，跳级/重号/断号）
    ⑯ 引用密度分布（citation_distribution.py，结构性失引，待办级）
    ⑰ 图表与正文数字一致性（figure_data_check.py，红线级）
    ⑱ 投稿材料包完整性（submission_pack.py，官方 §8 要素，待办级）
    ⑲ 与已见刊论文句面相合度（corpus_similarity.py，查重前哨，待办级）
    ⑳ 摘要与结语呼应核验（conclusion_echo.py，待办级）
    ㉑ 图表编号连续性（figure_numbering.py，断号/重号/悬空指路，待办级）
    ㉒ 关键词内容呼应（keyword_echo.py，装饰性关键词专查，待办级）
    ㉓ 段内结构范式（paragraph_structure.py，提示级，不计待办）
    ㉔ 句长分布（sentence_length.py，碎句化/臃肿句，待办级）
    ㉕ 段落首尾形态与分号（paragraph_shape.py，提示级，不计待办）
    ㉖ 标题形态（heading_shape.py，虚标题/末章命名，待办级）
    ㉗ 摘要要素顺序与句式（abstract_shape.py，官方忌词红线级）
    ㉘ 「本文」使用率（benwen_usage.py，摘要含「本文」红线级）
    ㉙ 篇章规模（structure_scale.py，章节数/字数，待办级）
    ㉚ 摘要字数/段长/图表对标（meta_scale.py，摘要字数越界待办级）
    ㉛ 引文年份越界（ref_year_check.py，未来文献红线级）
    ㉜ 文献类型标识（ref_type_check.py，标识缺失待办级）
    ㉝ 文献重复条目（ref_dup_check.py，重复给号待办级）
    ㉞ 单章字数失衡（chapter_balance.py，>40% 待办级）
    ㉟ 引言篇幅形态（intro_scale.py，缺缺口句红线级）
    㠀 标题命名对位（heading_naming.py，提示级）
    ㊱ 「了」密度（le_density.py，单段≥4 待办级）
    ㊲ 「即」解释句式密度（ji_density.py，提示级）
    ㊳ 程度词比率（hedge_ratio.py，强断言>75% 待办级）
    ㊴ 摘要—关键词关联（abstract_keyword_link.py，过半关键词未在摘要出现待办级）
    ㊵ 前置元数据序列（front_matter.py，官方 §4 序列，待办级）
    ㊶ 英文题名核验（en_title_check.py，缺失待办级）
最后给一份「总检小结」：红线数 / 待办数 / 模拟审稿分数。
每次运行把可量化指标写入 <论文>.checkall_snapshot.json，下次运行输出
「与上次总检对比」（▲/▼ 变化、改善✓/退步⚠）与 JSON 的结构化结果。

用法:
    python check_all.py <论文.docx> [--corpus <语料库目录>] [--no-benchmark]
    python check_all.py <论文.docx> --quick     # 只跑①②，快
    python check_all.py <论文.docx> --no-review # 跳过⑤模拟审稿
    python check_all.py <论文.docx> --json      # 追加结构化 JSON（含快照与上次对比）

退出码：0 = 无红线；1 = 有红线（需处理后再投）
说明：各工具的完整报告在各自小节输出；本脚本只做串联与小结，不替代细读。
"""
import os
import re
import sys
import json
import tempfile
import subprocess

# Windows 控制台 GBK 保护
if hasattr(sys.stdout, "reconfigure"):
    try:
        sys.stdout.reconfigure(encoding="utf-8", errors="replace")
    except Exception:
        pass

HERE = os.path.dirname(os.path.abspath(__file__))
_args = [a for a in sys.argv[1:] if not a.startswith("--")]
_flags = {a for a in sys.argv[1:] if a.startswith("--")}
if not _args or _args[0] in ("-h", "--help"):
    print(__doc__)
    sys.exit(0)
DOC = _args[0]
if not os.path.isfile(DOC):
    sys.exit(f"找不到文件：{DOC}")

QUICK = "--quick" in _flags
SKIP_BM = "--no-benchmark" in _flags or QUICK
SKIP_REVIEW = "--no-review" in _flags or QUICK
_corpus_arg = None
if "--corpus" in sys.argv:
    i = sys.argv.index("--corpus")
    _corpus_arg = sys.argv[i + 1] if i + 1 < len(sys.argv) else None

TOOLS = [
    ("① GB/T 7714 文献要素", "ref_format.py", False),
    ("② 引用标注核验（模式 D）", "verify_refs.py", False),
    ("③ 引文结构（时效/类型/来源）", "citation_analysis.py", False),
    ("④ 论文诊断（模式 A）", "diagnose_paper.py", False),
    ("⑤ 整稿语言体检（润色）", "language_audit.py", False),
    ("⑥ 模拟审稿（模式 E）", "mock_review.py", False),
    ("⑦ 对标评估（模式 G）", "benchmark_paper.py", True),
    ("⑧ 论证链对仗（发现↔对策）", "argument_check.py", False),
    ("⑨ 图表体例与三线表规范", "table_fig_audit.py", False),
    ("⑩ 摘要—正文数字一致性", "number_consistency.py", False),
    ("⑪ 术语漂移检测（同义异形词混用）", "term_consistency.py", False),
    ("⑫ 引文—论点关联（装饰性引用）", "citation_quality.py", False),
    ("⑬ 稿内自我重复（查重前哨）", "self_repeat.py", False),
    ("⑭ 英文摘要与关键词对称性", "english_audit.py", False),
    ("⑮ 标题编号体例", "heading_audit.py", False),
    ("⑯ 引用密度分布（结构性失引）", "citation_distribution.py", False),
    ("⑰ 图表与正文数字一致性", "figure_data_check.py", False),
    ("⑱ 投稿材料包完整性", "submission_pack.py", False),
    ("⑲ 与已见刊论文句面相合度", "corpus_similarity.py", False),
    ("⑳ 摘要与结语呼应核验", "conclusion_echo.py", False),
    ("㉑ 图表编号连续性", "figure_numbering.py", False),
    ("㉒ 关键词内容呼应", "keyword_echo.py", False),
    ("㉓ 段内结构范式", "paragraph_structure.py", False),
    ("㉔ 句长分布", "sentence_length.py", False),
    ("㉕ 段落首尾形态与分号", "paragraph_shape.py", False),
    ("㉖ 标题形态", "heading_shape.py", False),
    ("㉗ 摘要要素顺序与句式", "abstract_shape.py", False),
    ("㉘ 「本文」使用率", "benwen_usage.py", False),
    ("㉙ 篇章规模", "structure_scale.py", False),
    ("㉚ 摘要字数/段长/图表对标", "meta_scale.py", False),
    ("㉛ 引文年份越界", "ref_year_check.py", False),
    ("㉜ 文献类型标识", "ref_type_check.py", False),
    ("㉝ 文献重复条目", "ref_dup_check.py", False),
    ("㉞ 单章字数失衡", "chapter_balance.py", False),
    ("㉟ 引言篇幅形态", "intro_scale.py", False),
    ("㠀 标题命名对位", "heading_naming.py", False),
    ("㊱ 「了」密度", "le_density.py", False),
    ("㊲ 「即」解释句式密度", "ji_density.py", False),
    ("㊳ 程度词比率", "hedge_ratio.py", False),
    ("㊴ 摘要—关键词关联", "abstract_keyword_link.py", False),
    ("㊵ 前置元数据序列", "front_matter.py", False),
    ("㊶ 英文题名核验", "en_title_check.py", False),
]

# 跳过规则按脚本名判定，避免依赖显示序号（曾因编号错位跳过错误工具）
SKIP_SCRIPTS = set()
if SKIP_BM:
    SKIP_SCRIPTS.add("benchmark_paper.py")
if SKIP_REVIEW:
    SKIP_SCRIPTS.add("mock_review.py")
# --quick = 只跑前两项
QUICK_SCRIPTS = {"ref_format.py", "verify_refs.py"}

results = {}
for title, script, skip in TOOLS:
    if script in SKIP_SCRIPTS:
        continue
    if QUICK and script not in QUICK_SCRIPTS:
        continue
    print("\n" + "█" * 68)
    print(f"█ {title}")
    print("█" * 68)
    cmd = [sys.executable, os.path.join(HERE, script), DOC]
    if _corpus_arg:
        cmd += ["--corpus", _corpus_arg]
    r = subprocess.run(cmd, capture_output=True, encoding="utf-8", errors="replace")
    print(r.stdout.strip())
    if r.stderr.strip():
        print("[stderr]", r.stderr.strip()[-300:])
    results[title] = {"rc": r.returncode, "out": r.stdout or "", "err": r.stderr or ""}

# ---------- 指标快照：本轮可量化指标（作增量对比基准） ----------
SNAP = {}

# ---------- 总检小结 ----------
print("\n" + "═" * 68)
print("总检小结")
print("═" * 68)

red_flags = []
todos = []

# 从各工具输出抓关键信号
fmt_out = results.get("① GB/T 7714 文献要素", {}).get("out", "")
m = re.search(r"【要素与形态问题】共 (\d+) 处", fmt_out)
if m:
    n = int(m.group(1))
    if n:
        todos.append(f"文献要素/形态问题 {n} 处（见「文献要素」）")
    else:
        print("· 文献表 GB/T 7714 要素：✓ 无问题")

ver_out = results.get("② 引用标注核验（模式 D）", {}).get("out", "")
if "缺号: 无 | 重号: 无  ✓" not in ver_out:
    red_flags.append("参考文献序号有缺号/重号（见「引用核验」）")
if "引用但列表无: 无  ✓" not in ver_out:
    red_flags.append("正文有引用但文献表无对应条目（见「引用核验」）")
if "⚠️ 检出混合式" in ver_out:
    red_flags.append("存在混合式标注 [1-3,7]（本刊不用，见 ②）")
if "列表有但正文未引" in ver_out and "无" not in ver_out.split("列表有但正文未引")[1][:20]:
    todos.append("有文献表条目未被正文引用（见「引用核验」，建议核对）")
if red_flags:
    pass
else:
    print("· 引用标注与文献表对应：✓ 通过")

cit_out = results.get("③ 引文结构（时效/类型/来源）", {}).get("out", "")
if cit_out:
    m = re.search(r"【风险提示】\n((?:  ⚠️.*\n)+)", cit_out)
    if m:
        for line in m.group(1).strip().splitlines():
            todos.append(line.strip().replace("⚠️ ", "") + "（见「引文结构」）")
    else:
        print("· 引文结构与时效性：✓ 未见明显问题")
    m = re.search(r"近 5 年（\d+\+）占比：([\d.]+)%", cit_out)
    if m:
        print(f"· 近 5 年文献占比：{m.group(1)}%")

diag_out = results.get("④ 论文诊断（模式 A）", {}).get("out", "") if not QUICK else ""
m = re.search(r"\[红线合计\] (\d+) 项", diag_out)
if m:
    n_red = int(m.group(1))
    SNAP["diag_red"] = n_red
    if n_red:
        for m_ in re.finditer(r"  (\S+?): \d+ 处 = [\d.]+/万字（刊基线 [\d.]+）⚠️", diag_out):
            red_flags.append(f"反模式「{m_.group(1)}」超标（见「诊断」）")
        for m_ in re.finditer(r"未检出：([^\n（]+)", diag_out):
            red_flags.append(f"结构缺「{m_.group(1).strip()}」（见「诊断」）")
        if "忌词" in diag_out or "虚句" in diag_out:
            red_flags.append("摘要含忌词/虚句（见「诊断」）")
    else:
        print("· 诊断红线（反模式/结构/摘要）：✓ 0 项")
m = re.search(r"段落级扫描】共 (\d+) 个目标段", diag_out)
if m and int(m.group(1)):
    SNAP["diag_target_paras"] = int(m.group(1))
    todos.append(f"段落级改写目标 {m.group(1)} 段（见「诊断」）")
m = re.search(r"文献 (\d+)条 vs 刊中位 (\d+) ▼", diag_out)
if m:
    SNAP["ref_cnt"] = int(m.group(1))
    todos.append(f"参考文献 {m.group(1)} 条低于刊中位 {m.group(2)}（见「诊断」对照）")
m = re.search(r"段长中位 (\d+)字 vs 刊中位 (\d+) ▲ 低于", diag_out)
if m:
    SNAP["para_med"] = int(m.group(1))
    todos.append(f"段长中位 {m.group(1)} 字偏短（刊中位 {m.group(2)}，见「诊断」）")
# 句均与引用密度（诊断报告中的正文指标行）
m2 = re.search(r"句均 (\d+)(?:\.\d+)?字 vs 刊中位", diag_out)
if m2:
    SNAP["sent_mean"] = float(m2.group(1))
m3 = re.search(r"引用密度 ([\d.]+)/千字 vs 刊中位", diag_out)
if m3:
    SNAP["cite_rate"] = float(m3.group(1))

bm_out = results.get("⑤ 整稿语言体检（润色）", {}).get("out", "") if not QUICK else ""
m = re.search(r"【A】该刊 0 次口语词（必改）：(\d+) 处", bm_out)
if m:
    SNAP["lang_a"] = int(m.group(1))
    red_flags.append(f"含该刊 0 次口语词 {m.group(1)} 处（见「语言体检」，必改）")
elif "【A】该刊 0 次口语词（必改）：✓ 无" in bm_out:
    SNAP["lang_a"] = 0
    print("· 整稿语言体检 A 类口语词：✓ 无")
LANG_CATS = [
    ("B", "该刊罕用词", "该刊罕用词"),
    ("C", "无支撑强断言", "无支撑强断言"),
    ("D", "「了」属性陈述堆积", "了属性陈述"),
    ("E", "「被」字被动", "被字被动"),
    ("F", "解释未用「即」式", "解释未用即字式"),
    ("G", "反模式句子", "反模式句子"),
]
for cat, title_in_output, label in LANG_CATS:
    m = re.search(rf"【{cat}】{re.escape(title_in_output)}[^：]*：(\d+) 处", bm_out)
    if m and int(m.group(1)):
        todos.append(f"语言体检〔{label}〕{m.group(1)} 处（见「语言体检」）")

mr_out = results.get("⑥ 模拟审稿（模式 E）", {}).get("out", "") if not SKIP_REVIEW else ""
m = re.search(r"【总分】([\d.]+) / (\d+)", mr_out)
if m:
    score = float(m.group(1))
    SNAP["review_score"] = score
    print(f"· 模拟审稿总分：{score:.1f}/100")
    if score < 60:
        red_flags.append(f"模拟审稿 {score:.1f}/100 <60——结构性问题，需大修（见「模拟审稿」）")
    elif score < 75:
        todos.append(f"模拟审稿 {score:.1f}/100（60–74）——按路线图改一轮（见「模拟审稿」）")
    elif score < 85:
        todos.append(f"模拟审稿 {score:.1f}/100（75–84）——小修 1–2 处（见「模拟审稿」）")
    else:
        print(f"  形式层面优秀（≥85）")
    # 抓取失分最多的前 3 项
    for mm in re.finditer(r"^\s+\d+\. \[失 ([\d.]+) 分·([^\]]+)\] (.+)$", mr_out, re.M):
        if float(mm.group(1)) >= 3:
            todos.append(f"模拟审稿失分 Top：{mm.group(3)[:52]}（{mm.group(2)}，−{mm.group(1)} 分）")

g_out = results.get("⑦ 对标评估（模式 G）", {}).get("out", "")
m = re.search(r"文献条数\s+(\d+)条\s+刊中位\s+[\d.]+\s*\|\s*百分位\s+(\d+)%", g_out)
if m and int(m.group(2)) <= 25:
    todos.append(f"文献量仅高于 {m.group(2)}% 刊文（见「对标」）")

arg_out = results.get("⑧ 论证链对仗（发现↔对策）", {}).get("out", "")
m = re.search(r"⚠️ 疑似未对仗的发现 (\d+)/(\d+) 条", arg_out)
if m and int(m.group(1)):
    todos.append(f"论证链 {m.group(1)}/{m.group(2)} 条发现疑似缺对策回应（见 ⑧，逐条核对）")

tf_out = results.get("⑨ 图表体例与三线表规范", {}).get("out", "")
m = re.search(r"🔴 发现 (\d+) 项需修正的图表问题", tf_out)
if m and int(m.group(1)):
    todos.append(f"图表规范与占位符待办 {m.group(1)} 项（见 ⑨ 图表体例）")

nc_out = results.get("⑩ 摘要—正文数字一致性", {}).get("out", "")
m = re.search(r"🔴 摘要与正文数字硬伤 (\d+) 项", nc_out)
if m and int(m.group(1)):
    red_flags.append(f"摘要与正文数字不一致 {m.group(1)} 处（见 ⑩ 数字一致性，投稿前必须改）")
elif m:
    print("· 摘要—正文数字一致性：✓ 无硬伤")
m = re.search(r"🟡 数字待人工核对 (\d+) 项", nc_out)
if m and int(m.group(1)):
    todos.append(f"数字待人工核对 {m.group(1)} 项（模糊量/多候选，见 ⑩ 数字一致性）")

# ⑪ 术语漂移：待办级（不阻断投稿，建议统一；与 ⑩ 红线级区分开）
tc_out = results.get("⑪ 术语漂移检测（同义异形词混用）", {}).get("out", "")
m = re.search(r"🔴 术语混用 (\d+) 组", tc_out)
if m and int(m.group(1)):
    todos.append(f"术语漂移 {m.group(1)} 组（同义异形词混用，见 ⑪，建议全文统一）")
elif m:
    print("· 术语漂移（同义异形词混用）：✓ 未检出")
elif results.get("⑪ 术语漂移检测（同义异形词混用）", {}).get("rc") == 1:
    # 退出码 1 但组数解析失败 → 兜底记一条待核，防止静默漏报
    todos.append("术语混用待核（见 ⑪，组数解析失败）")

# ⑫ 引文—论点关联：零关联引用点 → 待办级（🟡 弱关联只提示，不占待办位）
cq_out = results.get("⑫ 引文—论点关联（装饰性引用）", {}).get("out", "")
# 分档统计行：🔴 零关联 12 ｜ 🟡 弱关联 3 ｜ ✅ 有关联 22
m = re.search(r"🔴 零关联 (\d+)\s*｜\s*🟡 弱关联 (\d+)", cq_out)
if m:
    _red_n, _weak_n = int(m.group(1)), int(m.group(2))
    if _red_n:
        todos.append(f"零关联引用点 {_red_n} 个（疑似装饰性引用，见 ⑫，逐条确认是否真支撑本句）")
    if _weak_n:
        todos.append(f"弱关联引用点 {_weak_n} 个（有交集但弱，见 ⑫，人工复核）")
    if not _red_n and not _weak_n:
        print("· 引文—论点关联（装饰性引用）：✓ 无疑似")
elif results.get("⑫ 引文—论点关联（装饰性引用）", {}).get("rc") == 1:
    todos.append("引文—论点关联待核（见 ⑫，计数解析失败）")

# ⑬ 稿内自我重复：跨区域重复 → 待办（查重率主因）；同区域内也提示
sr_out = results.get("⑬ 稿内自我重复（查重前哨）", {}).get("out", "")
m = re.search(r"🔴 跨区域 (\d+) 对\s*｜\s*🟡 同区域内 (\d+) 对", sr_out)
if m:
    _cross_n, _inner_n = int(m.group(1)), int(m.group(2))
    if _cross_n:
        todos.append(f"稿内跨区域重复 {_cross_n} 对（查重率主因，见 ⑬，须改写其中一处）")
    if _inner_n:
        todos.append(f"稿内同区域重复 {_inner_n} 对（套话复述，见 ⑬，建议合并）")
    if not _cross_n and not _inner_n and "未检出" in sr_out:
        print("· 稿内自我重复：✓ 未检出")
elif results.get("⑬ 稿内自我重复（查重前哨）", {}).get("rc") == 1:
    todos.append("稿内自我重复待核（见 ⑬，计数解析失败）")

# ⑭ 英文摘要与关键词：缺失/个数不一致 → 待办（官方格式明文）
en_out = results.get("⑭ 英文摘要与关键词对称性", {}).get("out", "")
m = re.search(r"🔴 发现 (\d+) 项需修改", en_out)
if m and int(m.group(1)):
    todos.append(f"英文摘要/关键词 {m.group(1)} 项问题（官方中英文对应要求，见 ⑭）")
elif m:
    print("· 英文摘要与关键词：✓ 对称性无问题")

# ⑮ 标题编号体例：跳级/重号/断号 → 待办（初审常见退回）
hd_out = results.get("⑮ 标题编号体例", {}).get("out", "")
m = re.search(r"🔴 发现 (\d+) 项编号体例问题", hd_out)
if m and int(m.group(1)):
    todos.append(f"标题编号体例 {m.group(1)} 项（跳级/重号/断号，见 ⑮）")
elif m:
    print("· 标题编号体例：✓ 合规")

# ⑯ 引用密度分布：结构性失引 → 待办（后半缺文献支撑，审稿高频意见）
cd_out = results.get("⑯ 引用密度分布（结构性失引）", {}).get("out", "")
m = re.search(r"🔴 发现 (\d+) 项结构性失引", cd_out)
if m and int(m.group(1)):
    todos.append(f"结构性失引 {m.group(1)} 项（引用集中在前部/论证空窗，见 ⑯）")
elif m:
    print("· 引用密度分布：✓ 未见结构性失引")

# ⑰ 图表与正文数字一致性：冲突 → 红线（数据可信度硬伤）
fd_out = results.get("⑰ 图表与正文数字一致性", {}).get("out", "")
m = re.search(r"🔴 发现 (\d+) 项图表与正文数字冲突", fd_out)
if m and int(m.group(1)):
    red_flags.append(f"图表与正文数字冲突 {m.group(1)} 项（见 ⑰，投稿前必须改）")
elif m:
    print("· 图表与正文数字：✓ 一致")

# ⑱ 投稿材料包：待补要素 → 待办（实名投稿前须补齐）
sp_out = results.get("⑱ 投稿材料包完整性", {}).get("out", "")
m = re.search(r"🔴 材料包待补 (\d+) 项", sp_out)
if m and int(m.group(1)):
    todos.append(f"投稿材料包待补 {m.group(1)} 项（作者简介/联系方式/定稿日期，见 ⑱）")
elif m:
    print("· 投稿材料包：✓ 要素齐全")

# ⑲ 与已见刊论文句面相合度：高度相合 → 待办（查重前哨）
cs_out = results.get("⑲ 与已见刊论文句面相合度", {}).get("out", "")
m = re.search(r"🔴 高度相合 ≥[\d.]+：(\d+) 句", cs_out)
if m and int(m.group(1)):
    todos.append(f"与已刊论文高度相合 {m.group(1)} 句（查重前哨，见 ⑲，须改写或删除）")
elif m:
    print("· 句面相合度（vs 118 篇见刊）：✓ 无高度相合")
m = re.search(r"🟡 中度相合 [\d.–]+：(\d+) 句", cs_out)
if m and int(m.group(1)):
    todos.append(f"与已刊论文中度相合 {m.group(1)} 句（可能计入复制比，见 ⑲，建议改写）")

# ⑳ 摘要与结语呼应：数量矛盾/承诺未兑现 → 待办
ce_out = results.get("⑳ 摘要与结语呼应核验", {}).get("out", "")
m = re.search(r"🔴 发现 (\d+) 项呼应问题", ce_out)
if m and int(m.group(1)):
    todos.append(f"摘要—结语呼应 {m.group(1)} 项（数量矛盾/承诺未兑现，见 ⑳）")
elif m:
    print("· 摘要—结语呼应：✓ 一致")

# ㉑ 图表编号连续性：断号/重号/悬空指路 → 待办（编辑部校稿必查）
fg_out = results.get("㉑ 图表编号连续性", {}).get("out", "")
m = re.search(r"🔴 发现 (\d+) 项编号问题", fg_out)
if m and int(m.group(1)):
    todos.append(f"图表编号问题 {m.group(1)} 项（断号/重号/悬空指路，见 ㉑）")
elif m:
    print("· 图表编号连续性：✓ 无问题")

# ㉒ 关键词内容呼应：装饰性关键词 → 待办
kw_out = results.get("㉒ 关键词内容呼应", {}).get("out", "")
m = re.search(r"🔴 发现 (\d+) 项问题", kw_out)
if m and int(m.group(1)):
    todos.append(f"关键词问题 {m.group(1)} 项（装饰性关键词/个数不符，见 ㉒）")
elif m:
    print("· 关键词内容呼应：✓ 无装饰性问题")

# ㉓ 段内结构范式：提示级，只汇总不进待办（避免风格建议淹没硬问题）
ps_out = results.get("㉓ 段内结构范式", {}).get("out", "")
m = re.search(r"合计：连接词起句 (\d+) ｜ 套话收尾 (\d+) ｜ 主题分散 (\d+) ｜ 分句过长 (\d+)", ps_out)
if m:
    _n = sum(int(x) for x in m.groups())
    if _n:
        print(f"· 段内结构范式：{_n} 处提示级（连接词起句 {m.group(1)}/"
              f"主题分散 {m.group(3)}/分句过长 {m.group(4)}，见 ㉓）")
    else:
        print("· 段内结构范式：✓ 无违反")

# ㉔ 句长分布：短句/长句率超限 → 待办（碎句化是显性文风问题）
sl_out = results.get("㉔ 句长分布", {}).get("out", "")
m = re.search(r"🔴 发现 (\d+) 项分布问题", sl_out)
if m and int(m.group(1)):
    _m2 = re.search(r"短句率（<30字）([\d.]+)%", sl_out)
    _m3 = re.search(r"长句率（>120字）([\d.]+)%", sl_out)
    _detail = f"短句 {_m2.group(1)}%/长句 {_m3.group(1)}%" if (_m2 and _m3) else ""
    todos.append(f"句长分布偏离（{_detail}，见 ㉔）")
elif m:
    print("· 句长分布：✓ 在刊容差内")

# ㉕ 段落首尾形态：提示级，只汇总不进待办
ph_out = results.get("㉕ 段落首尾形态与分号", {}).get("out", "")
m = re.search(r"合计：连接词起句 (\d+) ｜ 段末套话 (\d+) ｜ 并列零分号 (\d+)", ph_out)
if m:
    _n = sum(int(x) for x in m.groups())
    if _n:
        print(f"· 段落首尾形态：{_n} 处提示级（连接词起句 {m.group(1)}/"
              f"段末套话 {m.group(2)}/并列零分号 {m.group(3)}，见 ㉕）")
    else:
        print("· 段落首尾形态：✓ 无偏离")

# ㉖ 标题形态：虚标题/末章命名 → 待办
hs_out = results.get("㉖ 标题形态", {}).get("out", "")
m = re.search(r"🔴 发现 (\d+) 项标题问题", hs_out)
if m and int(m.group(1)):
    todos.append(f"标题形态 {m.group(1)} 项（虚标题/末章命名，见 ㉖）")
elif m:
    print("· 标题形态：✓ 无虚标题、末章合规")

# ㉗ 摘要要素顺序与句式：官方忌词 → 红线；乱序/首句形态 → 提示
as_out = results.get("㉗ 摘要要素顺序与句式", {}).get("out", "")
m = re.search(r"🔴 发现 (\d+) 项摘要问题", as_out)
if m and int(m.group(1)):
    # 官方忌词属红线级
    if "忌词" in as_out:
        red_flags.append(f"摘要含官方忌词（见 ㉗，须知明文禁止）")
    else:
        todos.append(f"摘要形态 {m.group(1)} 项（见 ㉗）")
elif m:
    print("· 摘要要素顺序与句式：✓ 合规")

# ㉘ 「本文」使用率：摘要含「本文」→ 红线（官方明文）；密度超 P75 → 待办
bw_out = results.get("㉘ 「本文」使用率", {}).get("out", "")
m = re.search(r"🔴 发现 (\d+) 项问题", bw_out)
if m and int(m.group(1)):
    if "摘要含「本文/我们」" in bw_out:
        red_flags.append("摘要含「本文/我们」（官方须知明文禁止，见 ㉘）")
    else:
        todos.append(f"「本文」密度超该刊 P90（见 ㉘）")
elif m:
    print("· 「本文」使用率：✓ 符合官方与刊内密度")

# ㉙ 篇章规模：章节过少/字数越界 → 待办
ss_out = results.get("㉙ 篇章规模", {}).get("out", "")
m = re.search(r"🔴 发现 (\d+) 项规模问题", ss_out)
if m and int(m.group(1)):
    todos.append(f"篇章规模偏离（章节数/字数，见 ㉙）")
elif m:
    print("· 篇章规模：✓ 在常规区间")

# ㉚ 摘要字数/段长/图表对标：摘要字数越界 → 待办
ms_out = results.get("㉚ 摘要字数/段长/图表对标", {}).get("out", "")
m = re.search(r"🔴 发现 (\d+) 项问题", ms_out)
if m and int(m.group(1)):
    _m2 = re.search(r"摘要 (\d+) 字", ms_out)
    todos.append(f"摘要字数越界（{_m2.group(1)} 字，见 ㉚）" if _m2 else "摘要字数越界（见 ㉚）")
elif m:
    print("· 摘要字数/段长/图表：✓ 在常规区间")

# ㉛ 引文年份越界：未来文献 → 红线（引用不存在文献是硬伤）
ry_out = results.get("㉛ 引文年份越界", {}).get("out", "")
m = re.search(r"🔴 发现 (\d+) 项越界年硬伤", ry_out)
if m and int(m.group(1)):
    red_flags.append(f"引文含未来文献 {m.group(1)} 条（见 ㉛，出版年晚于当前年）")
elif m:
    print("· 引文年份：✓ 无越界年")

# ㉜ 文献类型标识：缺失 → 待办（GB/T 7714 必填）
rt_out = results.get("㉜ 文献类型标识", {}).get("out", "")
m = re.search(r"🔴 发现 (\d+) 项类型标识缺失", rt_out)
if m and int(m.group(1)):
    todos.append(f"文献类型标识缺失 {m.group(1)} 条（见 ㉜，GB/T 7714 必填）")
elif m:
    print("· 文献类型标识：✓ 全部带标识")

# ㉝ 文献重复条目：完全/同题名重复 → 待办
rd_out = results.get("㉝ 文献重复条目", {}).get("out", "")
m = re.search(r"🔴 发现 (\d+) 项重复", rd_out)
if m and int(m.group(1)):
    todos.append(f"文献表重复条目 {m.group(1)} 项（见 ㉝，须合并共用序号）")
elif m:
    print("· 文献表重复：✓ 无重复")

# ㉞ 单章字数失衡：>40% → 待办
cb_out = results.get("㉞ 单章字数失衡", {}).get("out", "")
m = re.search(r"🔴 发现 (\d+) 项严重失衡", cb_out)
if m and int(m.group(1)):
    todos.append(f"单章字数失衡 {m.group(1)} 项（>40%，见 ㉞）")
elif m:
    print("· 单章字数失衡：✓ 无")

# ㉟ 引言篇幅形态：缺缺口句 → 红线（问题意识未立）
is_out = results.get("㉟ 引言篇幅形态", {}).get("out", "")
m = re.search(r"🔴 发现 (\d+) 项问题", is_out)
if m and int(m.group(1)):
    if "未检出缺口句" in is_out:
        red_flags.append("引言缺缺口句（问题意识未立，见 ㉟）")
    else:
        todos.append(f"引言篇幅形态 {m.group(1)} 项（见 ㉟）")
elif m:
    print("· 引言篇幅形态：✓ 正常")

# 㠀 标题命名对位：提示级，只汇总
hno_out = results.get("㠀 标题命名对位", {}).get("out", "")
m = re.search(r"未含高频词根 (\d+) 个", hno_out)
if m and int(m.group(1)):
    print(f"· 标题命名对位：{m.group(1)} 个未含高频词根（提示级，见 㠀）")
elif m:
    print("· 标题命名对位：✓ 均含高频词根")

# ㊱ 「了」密度：单段 ≥4 → 待办
ld_out = results.get("㊱ 「了」密度", {}).get("out", "")
m = re.search(r"🔴 发现 (\d+) 项「了」堆积", ld_out)
if m and int(m.group(1)):
    todos.append(f"「了」堆积 {m.group(1)} 段（单段≥4，见 ㊱）")
elif m:
    print("· 「了」密度：✓ 无堆积段")

# ㊲ 「即」密度：提示级
jd_out = results.get("㊲ 「即」解释句式密度", {}).get("out", "")
m = re.search(r"「即」(\d+) 处 = ([\d.]+)/万字", jd_out)
if m:
    _r = float(m.group(2))
    if _r < 1.5 or _r > 8.0:
        print(f"· 「即」密度：{_r}/万字（提示级偏离，见 ㊲）")
    else:
        print("· 「即」密度：✓ 常规区间")

# ㊳ 程度词比率：>75% → 待办；>60% → 提示
hr_out = results.get("㊳ 程度词比率", {}).get("out", "")
m = re.search(r"🔴 发现 (\d+) 项问题", hr_out)
if m and int(m.group(1)):
    todos.append(f"强断言词占比过高（审慎性不足，见 ㊳）")
elif m:
    _m2 = re.search(r"强断言占比 ([\d.]+)%", hr_out)
    if _m2 and float(_m2.group(1)) > 60:
        print(f"· 程度词比率：强断言 {_m2.group(1)}%（提示级偏离，见 ㊳）")
    else:
        print("· 程度词比率：✓ 均衡")

# ㊴ 摘要—关键词关联：过半未出现 → 待办
ak_out = results.get("㊴ 摘要—关键词关联", {}).get("out", "")
m = re.search(r"🔴 发现 (\d+) 项关联问题", ak_out)
if m and int(m.group(1)):
    _m2 = re.search(r"(\d+)/(\d+) 个关键词未在摘要出现", ak_out)
    todos.append(f"摘要—关键词脱节（{_m2.group(1)}/{_m2.group(2)} 未在摘要出现，见 ㊴）"
                 if _m2 else "摘要—关键词脱节（见 ㊴）")
elif m:
    print("· 摘要—关键词关联：✓ 良好")

# ㊵ 前置元数据序列：缺英文题名/顺序错 → 待办（官方 §4 序列）
fm_out = results.get("㊵ 前置元数据序列", {}).get("out", "")
m = re.search(r"🔴 发现 (\d+) 项序列问题", fm_out)
if m and int(m.group(1)):
    todos.append(f"前置元数据序列 {m.group(1)} 项（官方 §4 要求，见 ㊵）")
elif m:
    print("· 前置元数据序列：✓ 符合官方序列")

# ㊶ 英文题名：缺失 → 待办；形态问题 → 提示
et_out = results.get("㊶ 英文题名核验", {}).get("out", "")
if "未见英文题名" in et_out:
    todos.append("缺英文题名（官方 §4 序列要求，见 ㊶）")
elif "🔴" in et_out:
    _m2 = re.search(r"🔴 发现 (\d+) 项", et_out)
    todos.append(f"英文题名形态 {_m2.group(1) if _m2 else ''} 项（见 ㊶）")
elif et_out:
    print("· 英文题名：✓ 合规")

print()
if red_flags:
    print(f"🔴 红线 {len(red_flags)} 项——投稿前必须处理：")
    for f in red_flags:
        print(f"   · {f}")
else:
    print("🔴 红线：0 项")
if todos:
    print(f"🟡 待办 {len(todos)} 项——建议处理：")
    for t in todos:
        print(f"   · {t}")
else:
    print("🟡 待办：0 项")

# ---------- 增量对比：与上次总检（.checkall_snapshot.json）比较 ----------
SNAP_FILE = os.path.splitext(DOC)[0] + ".checkall_snapshot.json"
_DELTA_KEYS = [
    ("review_score", "模拟审稿总分", 0.1, "higher", "分", 85.0),
    ("ref_cnt", "文献条数", 1, "higher", "条", 40.0),
    ("para_med", "段长中位", 1, "higher", "字", 220.0),
    ("sent_mean", "句均长度", 1, "higher", "字", 63.0),
    ("cite_rate", "引用密度", 0.01, "higher", "/千字", 3.23),
    ("diag_red", "诊断红线", 1, "lower", "项", 0.0),
    ("diag_target_paras", "改写目标段", 1, "lower", "段", 0.0),
    ("lang_a", "必改口语词", 1, "lower", "处", 0.0),
]
_last = {}
if os.path.isfile(SNAP_FILE):
    try:
        with open(SNAP_FILE, encoding="utf-8") as f:
            _last = json.load(f)
    except (ValueError, OSError):
        _last = {}
if _last:
    print("\n" + "─" * 68)
    print("与上次总检对比（指标变化与刊中位可视化对照）")
    print("─" * 68)
    _changed = 0
    for _k, _label, _eps, _dir, _unit, _base_med in _DELTA_KEYS:
        if _k not in SNAP or _k not in _last:
            continue
        cur, prev = SNAP[_k], _last[_k]
        if abs(cur - prev) < _eps:
            continue
        _changed += 1
        df = cur - prev
        arrow = "▲" if df > 0 else "▼"
        good = (df > 0) == (_dir == "higher")
        mark = "✓" if good else "⚠"

        # 基准可视化条：显示当前值相对刊中位的差距 (针对于有刊基准的指标)
        bar_str = ""
        if _base_med > 0 and _dir == "higher":
            pct = min(1.0, max(0.0, cur / _base_med))
            filled = int(round(pct * 10))
            bar = "■" * filled + "□" * (10 - filled)
            diff_med = round(_base_med - cur, 2)
            diff_str = f"{diff_med:g}"
            status_tag = "达标" if cur >= _base_med else f"距中位还差 {diff_str}{_unit}"
            bar_str = f" | 刊中位 {_base_med:g}{_unit} [{bar}] ({status_tag})"
        elif _dir == "lower":
            status_tag = "清洁" if cur == 0 else f"仍存 {cur:g}{_unit}"
            bar_str = f" | 目标 0{_unit} ({status_tag})"

        print(f"  {arrow} {_label}：{prev:g}{_unit} → {cur:g}{_unit}（{df:+g}）{mark}{bar_str}")
    if not _changed:
        print("  （本轮指标与上次持平）")
    print(f"  快照：{os.path.basename(SNAP_FILE)}")
# 写回本轮快照（无论是否 --json 都记录，供下次对比）
try:
    with open(SNAP_FILE, "w", encoding="utf-8") as f:
        json.dump(SNAP, f, ensure_ascii=False, indent=1)
except OSError:
    pass

if "--json" in _flags:
    out = {
        "doc": os.path.basename(DOC),
        "red_flags": red_flags, "todos": todos,
        "snapshot": SNAP,
        "prev_snapshot": _last,
        "rc": {k: v["rc"] for k, v in results.items()},
    }
    jf = os.path.splitext(DOC)[0] + ".checkall.json"
    with open(jf, "w", encoding="utf-8") as f:
        json.dump(out, f, ensure_ascii=False, indent=1)
    print(f"\nJSON 已写入：{jf}")

sys.exit(1 if red_flags else 0)
