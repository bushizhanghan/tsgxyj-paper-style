# -*- coding: utf-8 -*-
"""mock_review.py -- 模式 E 模拟审稿：13 维 × 100 分制自动打分 + 审稿意见

性质声明（与 references/reviewer-simulation.md 一致）：
    这是基于本刊 118 篇刊文特征构造的「模拟审稿人」，不是该刊真实审稿标准。
    能自动评的是**可核验的形式特征**；选题价值、创新性、论证深度只能给「待人工」并标【仅参考】。

12 维度权重（合计 100 分）——权重按该刊真实审稿关注度设定（见 reviewer-simulation.md）：
    1 选题与定位        12   缺口句/研究设计/摘要目的；增量【仅参考】
    2 结构与篇幅        10   五段式完整度 + 章节占比 + 结语篇幅
    3 方法与数据         9   样本量/时间/工具/可复现
    4 论证链             9   对策↔发现对仗 + 收口
    5 文献质量          10   文献量分位 + 时效性 + 来源集中度
    6 引文规范           8   GB/T 7714 要素 + 序号 + 标注形态
    7 体例规范           8   摘要/关键词/标题层级
    8 语言与文风        10   段长/句均/引用密度/术语
    9 润色与语感         8   「了」「被」「即」/罕用口语词/强断言
   10 净度（AI味）       8   反模式超标 + 数字一致性
   11 图表质量           5   图表数量与有效引用
   12 结构创新度         3   章节序列是否贴合刊型【仅参考】

用法:
    python mock_review.py <论文.docx> [--corpus <dir>] [--json]
    python mock_review.py <论文.docx> --strict    # 从严（阈值收紧、扣分加倍）

判定线：
    >=85  形式层面优秀，可直接投（研究价值另论）
    75-84 良好，小修 1 处即可
    60-74 合格，按路线图改一轮
    <60   大修，结构性问题

退出码：0 = >=60 分；1 = <60 分
"""
import os
import re
import sys
import json
import statistics as st

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import corpus  # noqa: E402

if hasattr(sys.stdout, "reconfigure"):
    try:
        sys.stdout.reconfigure(encoding="utf-8", errors="replace")
    except Exception:
        pass

_args = [a for a in sys.argv[1:] if not a.startswith("--")]
_flags = {a for a in sys.argv[1:] if a.startswith("--")}
if not _args or _args[0] in ("-h", "--help"):
    print(__doc__)
    sys.exit(0)
DOC = _args[0]
if not os.path.isfile(DOC):
    sys.exit(f"找不到文件：{DOC}")
STRICT = "--strict" in _flags
_cli_corpus = None
if "--corpus" in sys.argv:
    i = sys.argv.index("--corpus")
    _cli_corpus = sys.argv[i + 1] if i + 1 < len(sys.argv) else None

import diagnose_paper as dp  # noqa: E402

corpus_dir = corpus.find_corpus_dir(_cli_corpus)
base = corpus.load_reference_stats(corpus_dir)
paras, refs, abs_text, kw_text, titles, body = dp.extract_paras(DOC)
m = dp.compute_metrics(body, refs, abs_text, kw_text, titles)
abst = dp.scan_abstract(abs_text)
struct = dp.scan_structure(m["titles"])
secs = dp.scan_sections(paras, m["titles"])
anti = dp.scan_anti_patterns("".join(dp.clean(p) for p in body))
para_issues = dp.scan_paragraphs(m["body"])
alltext = dp.clean("".join(m["body"]))

# ---------- 打分基础设施 ----------
# scores[dim] = {"got": 得分, "max": 满分, "ev": [...], "op": [...]}
scores = {}
K = 1.5 if STRICT else 1.0     # 严格模式扣分倍数


def add(dim, max_pt, got, ev, op):
    """记录一维得分。got 会被夹在 [0, max_pt]。"""
    scores[dim] = {
        "got": round(max(0.0, min(float(max_pt), float(got))), 1),
        "max": max_pt, "ev": ev, "op": op,
    }


def pct_score(value, stat, tol=0.25):
    """把一个指标值换算成「相对该刊基准的达标率」（0~1），供按比例扣分。"""
    med = stat.get("med") or 0
    if not med:
        return 1.0
    if value >= med:
        return 1.0
    if value <= med * (1 - tol):
        return 0.0
    return 1 - (med - value) / (med * tol)


def count_pat(p, text=None):
    return len(re.findall(p, text if text is not None else alltext))


def sents_of(text):
    return [s for s in re.split(r"[。？！]", text) if len(s.strip()) >= 6]


# ==================== 1. 选题与定位（12） ====================
ev, op = [], []
sc = 12.0
has_gap = bool(re.search(r"(尚未|尚缺乏|缺乏|不足|有待|仍需|未得到充分|鲜有).{0,30}(研究|探讨|讨论|关注)", alltext)) \
    or bool(re.search(r"(既有|现有|已有)研究.{0,40}(多聚焦|主要关注|较少)", alltext))
ev.append("缺口句：" + ("有" if has_gap else "未检出"))
if not has_gap:
    sc -= 4 * K
    op.append("未检出「既有研究…尚缺乏…」式缺口句。审稿人常因此认为问题意识不足（该刊引言首段多为政策/现象切入 + 缺口）。")
has_design = "研究设计/对象与方法" in struct["found"]
ev.append("研究设计章：" + ("有" if has_design else "无"))
if not has_design:
    sc -= 3 * K
    op.append("未单列「研究设计/对象与方法」章（该刊 37% 篇目单列）。若并入引言，须确保样本与手段写明。")
if abst.get("present") and not abst["elements"].get("目的"):
    sc -= 2 * K
    ev.append("摘要目的句：无")
    op.append("摘要未检出「目的」要素句（该刊 75% 摘要含目的句）。")
if abst.get("present") and not abst["elements"].get("价值"):
    sc -= 1 * K
    op.append("摘要未检出「价值/启示」句（该刊 62% 含）。")
op.append("【仅参考】选题增量与栏目对位需人工比对刊内同主题论文（工具不判断创新性）。")
if STRICT:
    sc -= 2
    op.append("【仅参考】严格模式默认扣 2 分：创新性未经验证。")
add("选题与定位", 12, sc, ev, op)

# ==================== 2. 结构与篇幅（10） ====================
ev, op = [], []
sc = 10.0
must = ["引言", "对策/策略/路径/启示（收口标配）", "结语"]
missing_must = [x for x in must if x in struct["missing"]]
if missing_must:
    sc -= min(6, len(missing_must) * 2.5) * K
    ev.append("缺章：" + "、".join(missing_must))
    op.append(f"缺标配章：{'、'.join(missing_must)}——该刊收口标配（命中率 52%–67%）。")
else:
    ev.append("三标配章齐全")
if "文献综述/相关研究" in struct["missing"]:
    sc -= 1 * K
    ev.append("文献综述：未单列")
    op.append("未单列「文献综述/相关研究」章（该刊 29% 单列，其余并入引言）——须确认综述内容存在且带引用。")
# 章节占比
if secs:
    for key, med, hi, lo in [
        ("结语", 4.1, 8, 0), ("引言", 7.1, 14, 0), ("对策/策略/路径/启示", 21.9, 99, 10),
    ]:
        if key in secs["merged"]:
            v = secs["merged"][key]
            ev.append(f"{key} {v:.1f}%（刊中位 {med}%）")
            if hi < 99 and v > hi:
                sc -= 1.2 * K
                op.append(f"{key}占正文 {v:.1f}%，超该刊中位（{med}%）偏多——{key}应克制。")
            elif lo and v < lo:
                sc -= 1.2 * K
                op.append(f"{key}仅占 {v:.1f}%（该刊中位 {med}%）——该刊最看重收口章，建议补充篇幅与条款。")
add("结构与篇幅", 10, sc, ev, op)

# ==================== 3. 方法与数据（9） ====================
ev, op = [], []
sc = 9.0
sample_hit = count_pat(r"(样本|被试|调查对象|研究对象|共计|共回收|有效样本|纳入).{0,25}\d")
time_hit = count_pat(r"(20\d{2}\s*年|20\d{2}\s*[-—至]\s*20\d{2})")
tool_hit = count_pat(r"(问卷|访谈|实验|编码|量表|统计|模型|算法|测试|检测|抽样)")
ev += [f"样本信号 {sample_hit}", f"时间信号 {time_hit}", f"方法信号 {tool_hit}"]
if sample_hit == 0:
    sc -= 4 * K
    op.append("未见样本量陈述（该刊要求写清样本量与本数）——审稿高频意见「样本描述缺失」。")
elif sample_hit == 1:
    sc -= 2 * K
    op.append("样本量信号仅 1 处——建议明确「共××个，有效××个，有效率××%」。")
if time_hit == 0:
    sc -= 2 * K
    op.append("未见时间范围陈述（数据采集/调研时间段）。")
if tool_hit < 3:
    sc -= 2 * K
    op.append("方法工具信号偏少——该刊方法段要求写清手段与过程（可复现）。")
# 可复现细节（有效率/信度/编码一致性）
repro = count_pat(r"(有效率|回收率|一致性|信度|效度|Kappa|Cronbach|α系数|编码者)")
ev.append(f"可复现细节信号 {repro}")
if repro == 0:
    sc -= 1.5 * K
    op.append("未见可复现细节（有效率/信度/编码一致性等）——该刊实证论文常给出这类指标。")
add("方法与数据", 9, sc, ev, op)

# ==================== 4. 论证链（9） ====================
ev, op = [], []
sc = 9.0
find_n = max(count_pat(r"一是"), count_pat(r"其一"), count_pat(r"发现"))
pol_n = max(count_pat(r"（1）|（2）|\(1\)|\(2\)"),
            count_pat(r"应[^\n]{0,6}(完善|健全|构建|优化|加强|建立|推进)"))
ev += [f"发现侧列举 {find_n}", f"对策侧条款 {pol_n}"]
if find_n == 0 and pol_n == 0:
    sc -= 4 * K
    op.append("未检出「一是/其二」式发现列举与「应…」式对策句——审稿人难以核对「对策—发现」对仗。")
elif abs(find_n - pol_n) >= 3:
    sc -= 2.5 * K
    op.append(f"发现列举 {find_n} 处 vs 对策条款 {pol_n} 处，差 ≥3——逐条画对应表核对是否一一对仗。")
# 缺口句兑现（第 2–4 章是否回应引言问题）
if has_gap and secs and len(secs["secs"]) >= 3:
    ev.append("章节数 ≥3（有承接空间）")
else:
    sc -= 1
    op.append("章节结构过简，缺口句难以在正文兑现。")
if "对策/策略/路径/启示（收口标配）" in struct["missing"]:
    sc -= 1.5 * K
    op.append("缺对策章，论证链无收口。")
op.append("【仅参考】缺口句是否在第 2–4 章兑现，需人工逐章核对（工具只看信号）。")
add("论证链", 9, sc, ev, op)

# ==================== 5. 文献质量（10） ====================
ev, op = [], []
sc = 10.0
n_refs = m["ref_n"]
ev.append(f"文献 {n_refs} 条（刊中位 {base['refs']['med']}）")
# 数量分档（该刊四分位 32–57）
if n_refs >= 57:
    ev.append("数量：超 P75")
elif n_refs >= 40:
    ev.append("数量：达中位")
elif n_refs >= 32:
    sc -= 1.5 * K
    op.append(f"文献 {n_refs} 条，处于该刊 P25–中位（32–40）——可再补 5–8 条进入中位区间。")
elif n_refs >= 20:
    sc -= 3.5 * K
    op.append(f"文献 {n_refs} 条，低于该刊 P25（32）——「文献支撑不足」是常见退修理由，建议补至 32 条以上。")
else:
    sc -= 5 * K
    op.append(f"文献仅 {n_refs} 条，远低于该刊区间（32–57）——须大幅补充。")
# 时效性 + 来源集中度（citation_analysis）
try:
    import citation_analysis as ca
    cbase = ca.corpus_stats(corpus_dir) if corpus_dir else None
    if cbase:
        yrs, srcs = [], []
        for r0 in refs:
            b0 = ca.norm(r0)
            y0 = ca.ref_year(b0)
            if y0:
                yrs.append(y0)
            s0 = ca.ref_source(b0, ca.ref_type(b0))
            if s0:
                srcs.append(s0)
        if yrs:
            recent = sum(1 for y in yrs if y >= 2021) / len(yrs) * 100
            ev.append(f"近5年 {recent:.0f}%（刊中位 {cbase['recent5_med']:.0f}%）")
            if recent < cbase["recent5_p25"]:
                sc -= 2.5 * K
                op.append(f"近 5 年文献仅 {recent:.0f}%，低于该刊 P25（{cbase['recent5_p25']:.0f}%）——「文献陈旧」。")
            elif recent < cbase["recent5_med"]:
                sc -= 1.2 * K
                op.append(f"近 5 年文献 {recent:.0f}%，低于刊中位（{cbase['recent5_med']:.0f}%）——补 2–3 篇新文献。")
        if srcs:
            uniq = len(set(srcs))
            ev.append(f"来源 {uniq} 种")
            if uniq <= 5 and n_refs >= 15:
                sc -= 1.5 * K
                op.append(f"来源仅 {uniq} 种，集中度过高——审稿意见常见「文献来源单一」。")
except Exception:
    pass
add("文献质量", 10, sc, ev, op)

# ==================== 6. 引文规范（8） ====================
ev, op = [], []
sc = 8.0
bad_type = [r0 for r0 in refs if not re.search(r"\[(J|M|D|C|S|N|Z|R|P|EB/OL|S/OL|Z/OL|J/OL|C/OL)\]", r0, re.I)]
if refs:
    ev.append(f"缺类型标识 {len(bad_type)} 条")
    if bad_type:
        sc -= min(3, len(bad_type) * 0.8) * K
        op.append(f"{len(bad_type)} 条文献缺文献类型标识（GB/T 7714 要素不全）。")
else:
    sc -= 4 * K
    op.append("未解析到参考文献表。")
# 序号连续 + 正文对应
nums = []
for r0 in refs:
    mm = re.match(r"^[\[［]?\s*(\d{1,3})", r0)
    if mm:
        nums.append(int(mm.group(1)))
if nums:
    expect = set(range(min(nums), max(nums) + 1))
    missing = sorted(expect - set(nums))
    dup = [x for x in set(nums) if nums.count(x) > 1]
    ev.append(f"序号 {'连续' if not missing and not dup else '异常'}")
    if missing or dup:
        sc -= 2.5 * K
        op.append(f"参考文献序号异常：缺号 {missing[:5] or '无'} / 重号 {dup[:5] or '无'}。")
# 引用形态
mixed = count_pat(r"\[\d{1,3}\s*[-–—]\s*\d{1,3}[\s*,，]\s*\d{1,3}")
adjacent = count_pat(r"\[\d+\]\s*\[\d+\]")
ev.append(f"混合式 {mixed} / 连排 {adjacent}")
if mixed:
    sc -= 1.5 * K
    op.append(f"混合式标注 {mixed} 处 [1-3,7]——该刊 0 例，拆分或改写句子。")
if adjacent:
    sc -= 0.8 * K
    op.append(f"连排单引 {adjacent} 处 [5][6]——改范围式或列举式。")
add("引文规范", 8, sc, ev, op)

# ==================== 7. 体例规范（8） ====================
ev, op = [], []
sc = 8.0
if abst.get("present"):
    ev.append(f"摘要 {abst['chars']} 字")
    if abst["chars"] < 180:
        sc -= 1.5 * K
        op.append(f"摘要 {abst['chars']} 字偏短（该刊 200–300 主流，中位 262）。")
    elif abst["chars"] > 320:
        sc -= 1.5 * K
        op.append(f"摘要 {abst['chars']} 字偏长（该刊 300 字以上仅 30 篇）。")
    if any("忌词" in i for i in abst["issues"]):
        sc -= 2 * K
        op.append("摘要含忌词（本文/分析了/提出了）——该刊须知明令。")
    if any("虚句" in i for i in abst["issues"]):
        sc -= 1 * K
        op.append("摘要含虚句（「取得良好效果」类）——结果句需具体内容。")
else:
    sc -= 4 * K
    op.append("未提取到摘要（需独立「摘要」段）。")
if m["kw_n"]:
    ev.append(f"关键词 {m['kw_n']} 个")
    if not (3 <= m["kw_n"] <= 6):
        sc -= 1.2 * K
        op.append(f"关键词 {m['kw_n']} 个（该刊 4–5 为主，均值 4.4）。")
else:
    sc -= 1.5 * K
    op.append("未提取到关键词。")
# 标题层级
if struct["h3_count"] > dp.H3_ALERT_THRESHOLD:
    sc -= 1 * K
    op.append(f"三级标题 {struct['h3_count']} 个，超实测 P90（约 {dp.H3_ALERT_THRESHOLD}）——层级可能过碎。")
ev.append(f"三级标题 {struct['h3_count']} 个")
# 摘要数字与正文一致
abs_nums = set(re.findall(r"\d+(?:\.\d+)?", dp.clean(abs_text or "")))
body_nums = set(re.findall(r"\d+(?:\.\d+)?", alltext))
miss_nums = {x for x in abs_nums if x not in body_nums and len(x) >= 2}
if miss_nums:
    sc -= 1.2 * K
    op.append("摘要数字未在正文复现：" + "、".join(sorted(miss_nums)[:5]) + "——须逐项核对。")
add("体例规范", 8, sc, ev, op)

# ==================== 8. 语言与文风（10） ====================
ev, op = [], []
sc = 8.0
sents = sents_of(alltext)
smean = st.mean([len(s) for s in sents]) if sents else 0
_mp, _ms, _mc = base["paras"]["med"], base["sents"]["med"], base["cites"]["med"]
ev += [f"段长 {m['para_med']:.0f}（刊 {_mp}）", f"句均 {smean:.0f}（刊 {_ms:.0f}）",
       f"引用密度 {m['cite_rate']:.2f}（刊 {_mc:.2f}）"]
# 段长：低于 P25 扣 3，低于中位按比例扣
if m["para_med"] < base["paras"]["q25"]:
    sc -= 3 * K
    op.append(f"段长中位 {m['para_med']:.0f} 字，低于刊 P25（{base['paras']['q25']}）——段落偏碎，"
              f"按 paragraphing-guide 合并（目标接近刊中位 {_mp}）。")
elif m["para_med"] < _mp * 0.85:
    sc -= 1.5 * K
    op.append(f"段长中位 {m['para_med']:.0f} 字，低于刊中位（{_mp}）较多——适度合并。")
# 句均
if smean and smean < _ms * 0.85:
    sc -= 2.5 * K
    op.append(f"句均 {smean:.0f} 字明显短于刊中位（{_ms:.0f}）——短句多，合并分句（该刊 40–90 字长句为主）。")
elif smean and smean > _ms * 1.35:
    sc -= 1.5 * K
    op.append(f"句均 {smean:.0f} 字明显长于刊中位（{_ms:.0f}）——句子过长，适当断句。")
# 引用密度
if m["cite_rate"] < base["cites"]["q25"]:
    sc -= 2 * K
    op.append(f"引用密度 {m['cite_rate']:.2f}/千字，低于刊 P25（{base['cites']['q25']:.1f}）——论述支撑偏少。")
elif m["cite_rate"] < _mc * 0.8:
    sc -= 1 * K
    op.append(f"引用密度 {m['cite_rate']:.2f}/千字，略低于刊中位（{_mc:.2f}）。")
add("语言与文风", 8, sc, ev, op)

# ==================== 9. 润色与语感（8） ====================
ev, op = [], []
sc = 7.0
le_per_1k = alltext.count("了") / max(len(alltext), 1) * 1000
bei_per_1k = count_pat(r"被") / max(len(alltext), 1) * 1000
ev += [f"「了」{le_per_1k:.1f}/千字（刊 2.7）", f"「被」{bei_per_1k:.2f}/千字（刊 0.53）"]
if le_per_1k > 6:
    sc -= 2.5 * K
    op.append(f"「了」{le_per_1k:.1f}/千字，超该刊（2.7）一倍以上——属性陈述改「能够/得到」或删。")
elif le_per_1k > 4:
    sc -= 1.2 * K
    op.append(f"「了」{le_per_1k:.1f}/千字，高于该刊（2.7）——控制属性陈述中的「了」。")
if bei_per_1k > 1.5:
    sc -= 1.5 * K
    op.append(f"「被」{bei_per_1k:.2f}/千字偏多（刊 0.53）——被动改主动或「得以」。")
# 罕用口语词
rare = [(w, f) for w, f in dp.COLLOQUIAL_HINTS if w in alltext]
ev.append(f"罕用词 {len(rare)} 类")
if rare:
    sc -= min(2.5, len(rare) * 1.2) * K
    op.append("含该刊罕用口语词：" + "、".join(f"{w}→{f}" for w, f in rare[:5]) + "（见 polish-lexicon.md）。")
# 必改口语词（该刊 0 次）
zero_words = ["差不多", "大概", "搞得", "咱们", "老百姓", "靠谱", "给力", "硬核", "走心", "一下子", "有时候"]
hit_zero = [w for w in zero_words if w in alltext]
if hit_zero:
    sc -= 3 * K
    op.append("含该刊 0 次口语词：" + "、".join(hit_zero) + "——必须替换（polish-lexicon.md §1）。")
# 无支撑强断言
strong_hit = count_pat(r"(?<!显著)(?<!明显)(彻底|完全|绝对|必然|极大)") if False else \
    len([x for x in re.finditer(r"(彻底|完全|绝对|必然|极大)", alltext)
         if not re.search(r"(彻底|完全|绝对|必然|极大)(按|依照|基于|一致|相同|结果|趋势|产物|要求)", alltext[max(0, x.start() - 2):x.end() + 6])])
if strong_hit:
    sc -= min(2, strong_hit * 0.6) * K
    op.append(f"无数据支撑的强断言 {strong_hit} 处（彻底/完全/绝对/必然/极大）——改「有所/一定程度上/相对」。")
add("润色与语感", 7, sc, ev, op)

# ==================== 9b. 官方合规（7） ====================
ev, op = [], []
sc = 7.0
_official = dp.scan_official(kw_text, refs, m["titles"], alltext)
if _official["columns"]:
    ev.append("栏目对位：" + _official["columns"][0])
else:
    sc -= 2 * K
    op.append("未能判定所属栏目（官方九栏）——投稿时须注明栏目，无法对位易被判「与刊物定位不符」。")
if _official["issues"]:
    sc -= min(5, len(_official["issues"]) * 1.6) * K
    for it in _official["issues"]:
        op.append("[官方] " + it)
    ev.append(f"官方明文问题 {len(_official['issues'])} 项")
else:
    ev.append("官方明文项：无问题")
# 摘要字数按官方 200 字左右（而非实测 262）
if abst.get("present"):
    _n = abst["chars"]
    if _n > 300:
        sc -= 1.0 * K
        op.append(f"摘要 {_n} 字，超官方「200 字左右」较多（实测中位 262，但官方目标更短）。")
    ev.append(f"摘要 {_n} 字（官方 200 左右）")
add("官方合规", 7, sc, ev, op)

# ==================== 10. 净度（AI 味）（8） ====================
ev, op = [], []
sc = 6.0
flagged = [a for a in anti if a["flag"]]
if flagged:
    sc -= min(5, len(flagged) * 2.0) * K
    for a in flagged:
        op.append(f"反模式超标「{a['name']}」{a['n']} 处 = {a['rate']}/万字（刊基线 {a['base']}）。")
    ev.append("反模式超标：" + "、".join(a["name"] for a in flagged))
else:
    ev.append("反模式 0 超标")
# A 级红线句（该刊 0–2 次）
red_hits = count_pat(r"显而易见|众所周知|毋庸置疑|总而言之|不难看出|值得一提|总的来说")
if red_hits:
    sc -= min(3, red_hits * 1.2) * K
    op.append(f"含该刊 0–2 次的 A 级红线句 {red_hits} 处（显而易见/众所周知等）。")
# 万能收尾句
wan = count_pat(r"具有重要意义|提供有益参考|有益参考|重要指导意义|值得深入探讨")
if wan:
    sc -= min(2.5, wan * 1.2) * K
    op.append(f"万能收尾句 {wan} 处（该刊 121 万字仅 8 次「具有重要意义」）。")
ev.append(f"红线句 {red_hits} / 万能句 {wan} / 反模式 {len(flagged)}")
add("净度（AI味）", 6, sc, ev, op)

# ==================== 11. 图表质量（5） ====================
ev, op = [], []
sc = 3.0
fig_n, tab_n = m["fig_n"], m["tab_n"]
ev += [f"图 {fig_n} 张（刊中位 {base['figs']['med']}）", f"表 {tab_n} 张（刊中位 {base['tabs']['med']}）"]
if fig_n + tab_n == 0:
    sc -= 3.5 * K
    op.append("全文无图表——该刊实证论文图 ~2 张、表 ~4 张，无图表易被视为「论述单薄」。")
else:
    if fig_n < base["figs"]["med"] and tab_n < base["tabs"]["med"]:
        sc -= 1 * K
        op.append(f"图表偏少（图 {fig_n}/表 {tab_n}，刊中位 图 {base['figs']['med']}/表 {base['tabs']['med']}）。")
# 图表是否有正文指路句
fig_ref = count_pat(r"(如图|见表|图\s*\d|表\s*\d)")
ev.append(f"图表指路句 {fig_ref}")
if fig_n + tab_n > 0 and fig_ref < (fig_n + tab_n) * 0.6:
    sc -= 1.2 * K
    op.append("图表在正文中的指路句偏少——该刊惯例是「如图 1 所示」「见表 2」显式引出。")
add("图表质量", 3, sc, ev, op)

# ==================== 12. 结构创新度（3，仅参考） ====================
ev, op = [], []
sc = 3.0
seq = " › ".join(t[:10] for t in m["titles"] if re.match(r"^[0-9]\s", t))
ev.append(f"一级序列：{seq[:48] or '未识别'}")
n_h1 = len([t for t in m["titles"] if re.match(r"^\d+\s", t) and not re.match(r"^\d+\.\d", t)])
if 5 <= n_h1 <= 7:
    ev.append(f"一级章 {n_h1} 个（该刊常见 5–7）")
elif n_h1 < 5:
    sc -= 1
    op.append(f"一级章仅 {n_h1} 个，结构偏简（该刊常见 5–7 章）。")
else:
    sc -= 0.5
    op.append(f"一级章 {n_h1} 个，偏多（该刊常见 5–7 章）——考虑合并。")
op.append("【仅参考】章节命名与组织方式的创新性需人工判断；工具只核对是否贴合该刊常见序列。")
add("结构创新度", 3, sc, ev, op)

# ==================== 输出 ====================
DIMS = [
    ("选题与定位", 12), ("结构与篇幅", 10), ("方法与数据", 9), ("论证链", 9),
    ("文献质量", 10), ("引文规范", 8), ("体例规范", 8), ("语言与文风", 8),
    ("润色与语感", 7), ("净度（AI味）", 6), ("官方合规", 7), ("图表质量", 3), ("结构创新度", 3),
]
total = round(sum(scores[d]["got"] for d, _ in DIMS), 1)
max_total = sum(scores[d]["max"] for d, _ in DIMS)

print("=" * 70)
print("模拟审稿报告（模式 E · 100 分制）")
print(f"稿件：{os.path.basename(DOC)}")
print(f"基准：《图书馆学研究》{len(corpus.load_corpus_stats(corpus_dir) or [])} 篇刊文特征"
      + ("（严格模式，扣分 ×1.5）" if STRICT else ""))
print("=" * 70)
print(f"\n【总分】{total} / {max_total}   ｜   {total/max_total*100:.1f} 分（百分制）")
if total >= 85:
    verdict = "优秀：形式层面就绪，可直接进入投稿流程（研究价值另论）"
elif total >= 75:
    verdict = "良好：小修 1–2 处即可投稿"
elif total >= 60:
    verdict = "合格：按修改路线图改一轮"
else:
    verdict = "不合格：结构性问题，先大修"
print(f"【判定】{verdict}")

print("\n【十三维得分】（权重×得分）")
for d, mx in DIMS:
    got = scores[d]["got"]
    ratio = got / mx
    filled = int(ratio * 10)
    bar = "█" * filled + "░" * (10 - filled)
    print(f"  {d:<12} {got:>5.1f}/{mx:<2}  {bar}  {'｜'.join(scores[d]['ev'][:3])}")

print("\n" + "─" * 70)
print("逐维意见")
print("─" * 70)
for d, mx in DIMS:
    got = scores[d]["got"]
    ops = scores[d]["op"]
    print(f"\n■ {d}（{got}/{mx}）")
    if ops:
        for o in ops:
            print(f"  · {o}")
    else:
        print("  · 本维未发现形式问题。")

print("\n" + "─" * 70)
print("修改路线图（按失分绝对值排序）")
print("─" * 70)
prio = []
for d, mx in DIMS:
    got = scores[d]["got"]
    loss = mx - got
    if loss >= 1.0:
        for o in scores[d]["op"]:
            if o.startswith("【仅参考】"):
                continue
            prio.append((loss, d, o))
prio.sort(key=lambda x: -x[0])
if prio:
    seen = set()
    n = 0
    for loss, d, o in prio:
        if o in seen:
            continue
        seen.add(o)
        n += 1
        print(f"  {n}. [失 {loss:.1f} 分·{d}] {o}")
        if n >= 12:
            break
    total_loss = sum(max(0, mx - scores[d]["got"]) for d, mx in DIMS)
    print(f"\n  合计失分 {total_loss:.1f} 分。按上表逐项处理即可接近 90+。")
else:
    print("  无显著失分项。建议进入语言润色环节（polish-lexicon.md）。")

print("\n【免责】本报告为刊文特征模拟，非该刊真实审稿标准；选题价值与创新性判断标【仅参考】。")

if "--json" in _flags:
    out = {
        "doc": os.path.basename(DOC), "total": total, "max": max_total,
        "percent": round(total / max_total * 100, 1),
        "verdict": verdict,
        "dims": {d: {"got": scores[d]["got"], "max": mx,
                     "evidence": scores[d]["ev"], "opinions": scores[d]["op"]}
                 for d, mx in DIMS},
    }
    jf = os.path.splitext(DOC)[0] + ".review.json"
    with open(jf, "w", encoding="utf-8") as f:
        json.dump(out, f, ensure_ascii=False, indent=1)
    print(f"\nJSON 已写入：{jf}")

sys.exit(0 if total >= 60 else 1)
