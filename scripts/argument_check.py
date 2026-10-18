# -*- coding: utf-8 -*-
"""argument_check.py —— 论证链「对策—发现」对仗核对（模式 E 第 4 维的语义层补充）

诊断/审稿现有的思路是信号匹配：数「发现侧列举」「对策侧条款」的数量，
只能知道数量差，不能回答「每条发现是否被一条对策一一对仗回应」。
本脚本做一层轻量语义核对：

  · 发现侧：提取「一是/其二/最为明显/占比最高/研究发现」等结果信号句
  · 对策侧：提取「（1）（2）…」「其一…其二…」「应…」「建议…」等条款句
  · 对仗评分：用主题词共现（常见名词短语交集）估计一条发现与一条对策的相关度，
    输出「发现 → 有回应/疑似缺回应」的逐条核对表，并指出可能的缺口句子。

局限（务必在报告里写明）：
  · 是基于词面共现的启发式，不是语义理解；「呼应」判断需人工复核
  · 发现列在与对策章分离的章节时（如分析章→对策章），跨章匹配有效
  · 同一发现被多条对策覆盖、或一条对策回应多条发现，均按「有回应」计

用法:
    python argument_check.py <论文.docx> [--corpus <dir>] [--top 5] [--json]
退出码：0 = 无「疑似缺回应」的发现；1 = 存在疑似未对仗的发现。
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
import diagnose_paper as dp  # noqa: E402

# ---------- 信号模式 ----------
# 发现侧：研究结果、问题、现状的陈述信号
FIND_SIGNS = [
    r"(?<!第)一是", r"(?<!第)二是", r"(?<!第)三是", r"(?<!第)四是",
    r"其[一二三四五](?!他)",  # 其一…其五（排除"其他"）
    r"最为明显", r"最为突出", r"占比(最高|最大)", r"比例(最高|最大)",
    r"研究发现", r"结果表明", r"结果显示", r"调查发现", r"分析发现", r"数据表明",
    r"存在以下", r"主要表现", r"集中体现", r"困境(主要|集中)", r"问题(主要|集中)",
    r"尤为突出", r"亟待", r"不足(主要|集中)", r"短板",
]
POLICY_SIGNS = [
    r"（1）", r"（2）", r"（3）", r"（4）", r"（5）", r"\(1\)", r"\(2\)", r"\(3\)",
    r"其[一二三四五](?!他)", r"第一，", r"第二，",
    r"(?<![相对呼适])应(?!用|对|急|变|景|酬|聘|邀|诊|援|征|届|允|许)",  # 应…（排除应用/应对/适应/相应…）
    r"应当", r"需(?![求要])", r"需要", r"建议", r"须", r"着力", r"注重",
    r"对策", r"路径", r"策略", r"措施", r"机制(建设|构建|完善)", r"体系(建设|构建|完善)",
]
# 主题词停用字：含这些字的 2-gram 无区分度
STOP_CHARS = set("的了和在是为对从于以而中上下后前内外等与被将把让向往由这那之也都又再更最很较比越显先后就同同")

# 主题词：忽略的泛化词（无区分度）
STOPWORDS = set("""的 了 和 与 及 或 在 是 为 对 从 于 以 而 中 上 下 后 前 内 外 等 与 被 将 把 让 向 往 由
本研究 本文 论文 文章 研究 方面 层面 维度 角度 问题 现状 情况 过程 内容 方式 方法 手段 途径
其 该 这 那 之 也 还 就 都 又 再 更 最 很 较 比 越 显 先 后 同时 此外 另外 综上 可见 表明 显示
以及 通过 基于 针对 围绕 按照 依据 借助 利用 使用 采用 提出 开展 推进 加强 完善 健全 构建
建设 优化 提升 增强 强化 促进 推动 实现 形成 建立 打造 深化 拓展 支持 保障 服务 管理 工作
图书馆 公共图书馆 高校图书馆 用户 读者 群体 人员 机构 部门 系统 平台 资源 数据 信息 服务""".split())

MIN_TOPIC_LEN = 2    # 主题词最短长度（字符）
TOP_TOPICS = 6       # 每句取前 N 个主题词
MAX_PAIRS = 8        # 每条发现最多检查的对策句数


def cjk_words(s):
    """候选主题词：2–4 字 n-gram，过滤含停用字的无区分度片段。
    返回 [(词, 权重)]，权重 = 词长（4字=4，3字=3，2字=2），让完整语义词优先。"""
    s = dp.clean(s)
    blocks = re.findall(r"[\u4e00-\u9fff]{2,}", s)
    cand = []
    for blk in blocks:
        for n in (2, 3, 4):
            for i in range(len(blk) - n + 1):
                w = blk[i:i + n]
                if any(c in STOP_CHARS for c in w):
                    continue
                if w in STOPWORDS:
                    continue
                cand.append((w, n))
    return cand


def topic_vec(s, top=TOP_TOPICS):
    """主题词袋：长 n-gram（≥3字）与短 2-gram 各取一半，长短混合保证既有语义词又有桥接词。"""
    words = cjk_words(s)
    from collections import Counter
    long_c, short_c = Counter(), Counter()
    for w, n in words:
        (long_c if len(w) >= 3 else short_c)[w] += n
    vec = dict(long_c.most_common(top // 2 + 1))
    short_n = top - len(vec)
    if short_n > 0:
        vec.update(dict(short_c.most_common(short_n)))
    return vec


def overlap_score(a_vec, b_vec):
    """主题词重叠得分：同词记 1；互相为子串（如「供给」⊂「服务供给」）记 0.5。"""
    score = 0.0
    for ka in a_vec:
        for kb in b_vec:
            if ka == kb:
                score += 1.0
            elif len(ka) >= 2 and len(kb) >= 2 and (ka in kb or kb in ka):
                score += 0.5
    return score


def split_sents(text):
    return [s for s in re.split(r"[。？！]", dp.clean(text)) if len(s) >= 8]


# 展望/结语式列举（不是研究发现的缺陷，不要求对仗）
OUTLOOK_SIGNS = [r"后续研究可从", r"未来(研究|可|将)", r"展望", r"拓展深化", r"值得进一步",
                 r"下一步(研究|工作)", r"有待(今后|未来)"]
# 测试/数据列举（如「（1）Hy4 Preview 模型，20 道题全部通过」），非对策条款
TEST_ENUM_SIGNS = [r"通过率", r"道题", r"全部通过", r"模型，", r"题次"]


def is_outlook(s):
    return any(re.search(p, s) for p in OUTLOOK_SIGNS)


def is_test_enum(s):
    # 列举行 `（1）…` 且句内是测试结果
    return (re.match(r"^（\d）", s) or re.match(r"^\(", s)) and any(
        re.search(p, s) for p in TEST_ENUM_SIGNS)


def is_find_sentence(s):
    if is_outlook(s):
        return False
    # 方法/设计依据列举（「依据有二」「划分依据是」）不是研究发现的缺陷，不算发现句
    if re.search(r"(依据|划分依据|设计依据|判定依据|标准依据)[有为是：:]{0,3}[一二三]", s):
        return False
    return any(re.search(p, s) for p in FIND_SIGNS)


def is_policy_sentence(s):
    if is_test_enum(s):
        return False
    return any(re.search(p, s) for p in POLICY_SIGNS)


def extract_findings(body_paras, titles, sec_map):
    """按段落扫描，收集发现句（带章节定位）。"""
    finds = []
    for i, p in enumerate(body_paras):
        for s in split_sents(p):
            if is_find_sentence(s):
                finds.append({"par_i": i, "text": s, "sec": sec_map.get(i, "")})
    return finds


def extract_policies(body_paras, titles, sec_map):
    """收集对策句：优先收口章（对策/策略/优化/建议/措施/路径/启示），再补全文条款句。"""
    # 收口章标题模式
    CLOSING = re.compile(r"(对策|策略|优化|建议|措施|路径|启示|展望|解决)")
    closing_pars = set()
    cur = None
    for i, t in enumerate(titles):
        if CLOSING.search(t):
            # 该标题到下一标题之间的段落归属收口章——用 sec_map 反查范围
            for pi, sec in sec_map.items():
                if sec == t:
                    closing_pars.add(pi)
    pols = []
    for i, p in enumerate(body_paras):
        for s in split_sents(p):
            if is_policy_sentence(s):
                pols.append({"par_i": i, "text": s, "sec": sec_map.get(i, ""),
                             "closing": i in closing_pars})
    # 收口章内的对策句排前面（优先匹配）
    pols.sort(key=lambda x: (0 if x["closing"] else 1, x["par_i"]))
    return pols


def sec_map_of(body_paras, titles):
    """段号 → 所属一级/二级标题名（就近原则）。"""
    mapping = {}
    cur = ""
    ti = 0
    for i in range(len(body_paras)):
        # 检查该段是否为标题（与 titles 匹配）——用简单规则：段文本与某标题相同
        txt = body_paras[i].strip()
        if any(t == txt for t in titles):
            cur = txt
        mapping[i] = cur
    return mapping


def match_findings(finds, pols):
    """逐条发现做全局匹配：遍历全部对策句取最高分（跳过同段自匹配）。"""
    rows = []
    for k, f in enumerate(finds):
        fv = topic_vec(f["text"])
        best, best_score = None, 0.0
        for p in pols:
            if p["par_i"] == f["par_i"] and p["text"] == f["text"]:
                continue  # 跳过自匹配（同一句被同时识别为发现与对策）
            pv = topic_vec(p["text"])
            sc = overlap_score(fv, pv)
            if sc > best_score:
                best, best_score = p, sc
        # 分档：≥2 强回应；1–1.5 弱回应（提示人工）；<1 疑似缺
        if best_score >= 2:
            level = "strong"
        elif best_score >= 1:
            level = "weak"
        else:
            level = "orphan"
        rows.append({
            "no": k + 1,
            "find": f["text"],
            "find_para": f["par_i"] + 1,
            "find_sec": f["sec"],
            "level": level,
            "matched": level != "orphan",
            "response": best["text"] if best else "",
            "resp_sec": best["sec"] if best else "",
            "score": round(best_score, 2),
        })
    return rows


def main():
    args = [a for a in sys.argv[1:] if not a.startswith("--")]
    flags = set(a for a in sys.argv[1:] if a.startswith("--"))
    if not args or args[0] in ("-h", "--help"):
        print(__doc__)
        sys.exit(0)
    DOC = args[0]
    if not os.path.isfile(DOC):
        sys.exit(f"找不到文件：{DOC}")
    top = 5
    if "--top" in sys.argv:
        i = sys.argv.index("--top")
        try:
            top = int(sys.argv[i + 1])
        except (IndexError, ValueError):
            pass

    paras, refs, abs_text, kw_text, titles, body = dp.extract_paras(DOC)
    mets = dp.compute_metrics(body, refs, abs_text, kw_text, titles)
    alltext = mets.get("alltext", "")
    if not alltext:
        alltext = dp.clean("".join(body))

    sm = sec_map_of(body, titles)
    finds = extract_findings(body, titles, sm)
    pols = extract_policies(body, titles, sm)

    print("=" * 70)
    print("论证链对仗核对（模式 E · 第 4 维语义补充）")
    print(f"稿件：{os.path.basename(DOC)}")
    print(f"发现句 {len(finds)} 条 ｜ 对策句 {len(pols)} 条（收口章优先）｜ 正文 {len(alltext):,} 字")
    print("=" * 70)
    if not finds and not pols:
        print("本稿未检出「一是/其二/占比最高」类发现句与「（1）（2）/其一/应…」类对策句。")
        print("若属思辨型稿（无实证结果），论证链的「发现—对策」对仗不适用，仅核对章节齐备即可。")
        sys.exit(0)

    print("\n【逐条对仗】强回应≥2分｜弱回应1–1.5（需人工复核）｜<1 疑似缺回应")
    print("-" * 70)
    orphans, weaks = 0, 0
    rows = match_findings(finds, pols)
    for r in rows_sorted(rows, top):
        if r["level"] == "orphan":
            mark, orphans = "⚠️ 疑似缺回应", orphans + 1
        elif r["level"] == "weak":
            mark, weaks = "🟡 弱回应(人工复核)", weaks + 1
        else:
            mark = "✅ 有回应"
        print(f"发现{r['no']:>2}（{r['find_para']}节{r['find_sec']}）{mark} [匹配 {r['score']:.1f}]")
        print(f"       {trunc(r['find'], 90)}")
        if r["matched"]:
            print(f"   ↳  {trunc(r['response'], 90)}  （{r['resp_sec']}）")
        else:
            print(f"   ↳  （未找到主题词重叠 ≥1 的对策句）")
    print("-" * 70)
    if orphans:
        print(f"⚠️ 疑似未对仗的发现 {orphans}/{len(finds)} 条——建议逐条补充对应对策，"
              f"或在对策章显式回指（如「针对上述第一类问题…」）。")
        sys.exit(1)
    if weaks:
        print(f"🟡 弱回应 {weaks}/{len(finds)} 条——词面有呼应但关联弱，请人工判定语义是否对仗。")
        sys.exit(0)
    print(f"✅ 全部 {len(finds)} 条发现均有强对策回应（词面启发式，请人工复核语义）。")
    sys.exit(0)


def rows_sorted(rows, top):
    """输出顺序：缺回应在前、弱回应次之，最多 top 条。"""
    rk = {"orphan": 0, "weak": 1, "strong": 2}
    s = sorted(rows, key=lambda r: (rk.get(r["level"], 3), r["score"]))
    if top and len(s) > top:
        head = [r for r in s if r["level"] != "strong"]
        tail = [r for r in s if r["level"] == "strong"]
        return head + tail[: max(0, top - len(head))]
    return s


def trunc(s, n):
    return s if len(s) <= n else s[:n] + "…"


if __name__ == "__main__":
    main()
