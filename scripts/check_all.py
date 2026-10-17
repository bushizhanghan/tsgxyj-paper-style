# -*- coding: utf-8 -*-
"""check_all.py —— 投稿前一键总检（模式 A + D + 润色 + E + G）

按投稿前自查的实用顺序，串起六个工具，各自独立跑、互不影响：
    ① GB/T 7714 文献要素核验（ref_format.py）
    ② 引用标注与对应核验（verify_refs.py，模式 D）
    ③ 论文诊断（diagnose_paper.py，模式 A）
    ④ 整稿语言体检（language_audit.py，润色词典落地）
    ⑤ 模拟审稿 7 维打分（mock_review.py，模式 E）
    ⑥ 对标评估（benchmark_paper.py，模式 G）
最后给一份「总检小结」：红线数 / 待办数 / 模拟审稿分数。

用法:
    python check_all.py <论文.docx> [--corpus <语料库目录>] [--no-benchmark]
    python check_all.py <论文.docx> --quick     # 只跑①②，快
    python check_all.py <论文.docx> --no-review # 跳过⑤模拟审稿

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
    todos.append(f"段落级改写目标 {m.group(1)} 段（见「诊断」）")
m = re.search(r"文献 (\d+)条 vs 刊中位 (\d+) ▼", diag_out)
if m:
    todos.append(f"参考文献 {m.group(1)} 条低于刊中位 {m.group(2)}（见「诊断」对照）")
m = re.search(r"段长中位 (\d+)字 vs 刊中位 (\d+) ▲ 低于", diag_out)
if m:
    todos.append(f"段长中位 {m.group(1)} 字偏短（刊中位 {m.group(2)}，见「诊断」）")

bm_out = results.get("⑤ 整稿语言体检（润色）", {}).get("out", "") if not QUICK else ""
m = re.search(r"【A】该刊 0 次口语词（必改）：(\d+) 处", bm_out)
if m:
    red_flags.append(f"含该刊 0 次口语词 {m.group(1)} 处（见「语言体检」，必改）")
elif "【A】该刊 0 次口语词（必改）：✓ 无" in bm_out:
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

if "--json" in _flags:
    out = {
        "doc": os.path.basename(DOC),
        "red_flags": red_flags, "todos": todos,
        "rc": {k: v["rc"] for k, v in results.items()},
    }
    jf = os.path.splitext(DOC)[0] + ".checkall.json"
    with open(jf, "w", encoding="utf-8") as f:
        json.dump(out, f, ensure_ascii=False, indent=1)
    print(f"\nJSON 已写入：{jf}")

sys.exit(1 if red_flags else 0)
