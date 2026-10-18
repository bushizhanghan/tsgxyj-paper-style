# -*- coding: utf-8 -*-
"""number_consistency.py —— 摘要—正文数字一致性核验器（数字硬伤专查）

诊断/审稿工具链此前只核文字与结构，不核数字：摘要写「回收有效问卷 286 份」、
正文写「回收有效问卷 300 份」，或摘要称「三类需求」、正文只列两类，这类
**数字硬伤**人工校对极易漏掉，被审稿人抓到就是原则性质疑。
本脚本把摘要 / 正文 / 结语中的数值断言抽出来做交叉比对：

  · 断言抽取：阿拉伯数字与中文数字（三百余份 / 百分之三十五 / 三成）＋ 计量单位
    （份 / 名 / 人 / 篇 / 条 / 种 / 题 / 道 / 项 / 类 / 个 / % …），
    记录数值、单位、前接名词串（head）与前文关键词（kw）；
  · 交叉比对：摘要 ↔（正文＋结语）、结语 ↔ 正文。先按「单位＋前接名词」精确配对，
    配不上再按「单位＋前文关键词」唯一兜底；同名指标数值不等即报硬伤；
  · 稿内逻辑：回收/有效样本大于发放样本、百分比越界（>100%）、
    文中「有效率」与「有效÷发放」算不符、百分比枚举组合计偏离 100%；
  · 条目数核对：摘要称「三类/四项/三个方面」与正文声称值、正文实际列举
    （一是…其二…／（1）（2）…）条数是否一致。

判定分档：
    🔴 硬伤   —— 同名指标数值不等，或样本量逻辑矛盾，投稿前必须改
    🟡 待核对 —— 模糊量（三百余/约）、多候选、枚举组差异、条目数与列举数不符
    ✅ 一致   —— 同名指标数值相同

用法:
    python number_consistency.py <论文.docx> [--top 12] [--json]
退出码：0 = 无数字硬伤；1 = 存在数字硬伤（摘要—正文不一致或稿内逻辑矛盾）。

局限（词面配对的启发式，不是语义理解）：
  · 前接名词被改写（「有效问卷」→「有效回收问卷」）时可能漏比，属安全侧漏报；
  · 同一维度多值并列（分类占比）只在「前接名词完全相同」时判硬伤，
    枚举组整体差异只报「待核对」，避免把序指代误判为矛盾；
  · 引述他人的统计结果若与本文数据同名同单位，可能误配，需人工复核。
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

# ---------------- 常量 ----------------
CJK = r"\u4e00-\u9fff"
CN_DIGITS = {"零": 0, "一": 1, "二": 2, "两": 2, "三": 3, "四": 4, "五": 5,
             "六": 6, "七": 7, "八": 8, "九": 9}
CN_UNITS = {"十": 10, "百": 100, "千": 1000}

# 参与跨区比对的计量单位（年月日/版次/页码等非事实性数字不参与，避免噪声）
UNITS = ["%", "份", "名", "人", "篇", "条", "册", "种", "家", "所", "题", "道",
         "项", "类", "次", "个", "部"]
UNIT_SET = set(UNITS)
CN_UNIT_RE = re.compile(
    r"([零一二三四五六七八九十百千万两]{1,7})(余|多|左右|以上|以下)?\s*("
    + "|".join(UNITS) + r")")
AR_NUM_RE = re.compile(r"(?<![\d.])(\d{1,5}(?:,\d{3})*(?:\.\d+)?)(?![\d])")
CN_PCT_RE = re.compile(r"百分之([零一二三四五六七八九十百]{1,6})(?:点([零一二三四五六七八九]))?")
CHENG_RE = re.compile(r"([一二三四五六七八九两]{1,2})成")
PUNCT_RE = re.compile(r"[。；;，,、：:（）()\[\]【】…\s]")

# 泛化词：数字前的这些字不构成区分度。
# 目的：防止「发放问卷 300」↔「回收问卷 286」「有效率」↔「发放率」被当成同名指标。
# 注意：「对策/建议/策略/路径/机制」等收口清单名词不在此列——它们是「摘要 N 类对策
# ↔ 正文 M 类对策」这类条目数核对的锚点，若一并泛化则声称数通道失效。
GENERIC = set("""的 了 和 与 及 或 在 是 为 对 从 于 以 而 中 上 下 后 前 内 外 等 被 将 把 让 向 由
占比 比例 占到 占 达到 约为 约 近 超过 高达 仅 只 为 有 共 共计 总共 合计 累计 分别 其中
结果 表明 显示 发现 可见 可知 样本 问卷 数量 有效 回收 发放 调查 统计 分析 显示 出
一类 两类 三类 四类 五类 六类 第一类 第二类 第三类 第四类 第五类
第一 第二 第三 第四 第五 其一 其二 其三 其四 其五 一方面 另一方面
需求 服务 用户 读者 图书馆 群体 数据 信息 资源 建设 管理 发展 现状 问题 模式
体系 平台 技术 水平 能力 质量 效果 影响 因素 层面 维度 方面
项 类 种 个 条 篇 人次 人数 篇数 条数 本 该 各 每 多数 部分 全部 整体 总体 不同 各种""".split())

# 前接名词串归一化：剥掉「共/约/高达/为」类前后缀，使「共回收有效问卷」与「回收有效问卷」同键。
# 注意：前缀不含「有」——否则「有效率」会被削成「效率」。
FILLER_PREFIX = re.compile(r"^(共计|总共|合计|累计|共|约|近|超过|达到|高达|仅为|只有|大约)+")
FILLER_SUFFIX = re.compile(r"(数量|个数|数|分别为|分别是|分别|为|达|是|有|占)+$")

# 条目数声称：「三类需求」「四个方面」「两项机制」
CLAIM_RE = re.compile(r"([一二三四五六七八九十两\d]{1,3})\s*(类|种|项|条|方面|维度|层面|轮)")
# 这些分类量单位由「条目数声称」通道负责（带名词），不再进数值断言跨区比对，避免重复报
CLAIM_UNITS = {"类", "种", "项", "条", "方面", "维度", "层面", "轮"}
# 正文实际列举标记（用于「声称 N 项 vs 实际列 M 项」）
ORD_MARKERS = ["一是", "二是", "三是", "四是", "五是", "六是", "七是", "八是", "九是", "十是",
               "其一", "其二", "其三", "其四", "其五", "其六",
               "第一，", "第二，", "第三，", "第四，", "第五，", "第六，",
               "第一,", "第二,", "第三,", "第四,", "第五,",
               "（1）", "（2）", "（3）", "（4）", "（5）", "（6）", "（7）", "（8）",
               "（9）", "（10）", "(1)", "(2)", "(3)", "(4)", "(5)", "(6)", "(7)", "(8)"]

# 收口章标题（结语/结论段单列，参与「摘要—结语」比对）
CLOSING_TITLE = re.compile(r"结\s*语|结\s*论|结束语|小\s*结|研究结论|结语与展望")


# ---------------- 基础工具 ----------------
def norm(s):
    """全角数字/百分号/逗号归一为半角（与 citation_analysis.norm 同口径）。"""
    if not s:
        return s
    tbl = {ord("０") + i: ord("0") + i for i in range(10)}
    tbl[ord("％")] = ord("%")
    tbl[ord("，")] = ord(",")
    tbl[ord("　")] = ord(" ")
    return s.translate(tbl)


norm_text = norm  # 兼容性别名（与 diagnose_paper.clean 口径区分：此处保留原文不做去空白）


def cn_to_int(s):
    """中文数字→整数（支持 十/百/千/万 与「两」）；无法解析返回 None。"""
    if not s:
        return None
    total, num = 0, 0
    for ch in s:
        if ch in CN_DIGITS:
            num = CN_DIGITS[ch]
        elif ch in CN_UNITS:
            total += (num if num else 1) * CN_UNITS[ch]
            num = 0
        elif ch == "万":
            total = (total + num) * 10000
            num = 0
        else:
            return None
    return total + num


def parse_ar(raw):
    """阿拉伯数字串→数值（去千分位；整数取 int）。"""
    v = raw.replace(",", "")
    try:
        f = float(v)
    except ValueError:
        return None
    return int(f) if f == int(f) else f


def fmt_val(v):
    return f"{v:g}"


def unit_after(t, pos):
    """数字后（容忍一个空格）的计量单位；非计量单位返回 None。"""
    if pos < len(t) and t[pos] == " ":
        pos += 1
    if t.startswith("个百分点", pos):
        return "个百分点"
    if pos < len(t) and t[pos] in UNIT_SET:
        return t[pos]
    return None


def head_before(text, pos, maxlen=12):
    """数字前最近的 CJK 串（遇标点/数字/拉丁字母即止）——前接名词串。"""
    i, buf = pos - 1, []
    while i >= 0 and len(buf) < maxlen:
        if re.match(r"[" + CJK + r"]", text[i]):
            buf.append(text[i])
            i -= 1
        else:
            break
    return "".join(reversed(buf))


def norm_head(h):
    """前接名词串归一化：剥前后缀填充词，保证「共回收有效问卷」与「回收有效问卷」同键。"""
    if not h:
        return ""
    s = FILLER_PREFIX.sub("", h)
    s2 = FILLER_SUFFIX.sub("", s)
    return s2 if len(s2) >= 2 else s


def ctx_keywords(text, pos, maxlen=14):
    """前文窗口（自上一个标点起）的区分性双字词——关键词兜底配对用。"""
    i, buf = pos - 1, []
    while i >= 0 and len(buf) < maxlen:
        if PUNCT_RE.match(text[i]):
            break
        buf.append(text[i])
        i -= 1
    win = "".join(reversed(buf))
    kws = set()
    for run in re.findall(r"[" + CJK + r"]{2,}", win):
        for j in range(len(run) - 1):
            bg = run[j:j + 2]
            if bg not in GENERIC:
                kws.add(bg)
    return kws


def snip(text, pos, back=26, fwd=16):
    """证据片段：数字前后各取一段，便于人工定位。"""
    a = max(0, pos - back)
    b = min(len(text), pos + fwd)
    s = text[a:b].replace("\n", " ")
    return ("…" if a > 0 else "") + s + ("…" if b < len(text) else "")


def trunc(s, n):
    return s if len(s) <= n else s[:n] + "…"


# ---------------- 断言抽取 ----------------
def mk_assert(region, par_i, text, start, value, unit, fuzzy, raw):
    head_raw = head_before(text, start)
    return {
        "region": region, "par": par_i + 1, "value": value, "unit": unit,
        "head_raw": head_raw, "head": norm_head(head_raw),
        "kw": ctx_keywords(text, start), "fuzzy": fuzzy, "raw": raw,
        "snip": snip(text, start),
    }


def extract_assertions(paras, region):
    """从段落列表抽取数值断言（阿拉伯数字 + 中文数字 + 百分之X + X成）。"""
    out = []
    for pi, raw_p in enumerate(paras):
        t = norm_text(raw_p or "")
        if not t:
            continue
        # ① 阿拉伯数字 + 单位（单位在数字后）
        for m in AR_NUM_RE.finditer(t):
            unit = unit_after(t, m.end())
            if unit is None or unit in CLAIM_UNITS:
                continue  # 「4 类」走条目数声称通道，避免与数值断言重复报
            val = parse_ar(m.group(1))
            if val is None:
                continue
            tail = t[m.end() + len(unit):]
            head = head_before(t, m.start(), 6)
            fuzzy = bool(re.match(r"^(余|多|左右|以上|以下)", tail)) or \
                (m.start() > 0 and t[m.start() - 1] in "约近超≈~") or \
                re.search(r"[约近超]", head[:3]) is not None
            out.append(mk_assert(region, pi, t, m.start(), val, unit, fuzzy, m.group(0)))
        # ② 中文数字 + 单位（三百余份 / 两类 / 三个方面）
        for m in CN_UNIT_RE.finditer(t):
            if m.start() > 0 and t[m.start() - 1] == "第":
                continue  # 序数指代（第三类）不是数量断言
            if m.group(3) in CLAIM_UNITS:
                continue  # 「三类需求」走条目数声称通道（带名词），避免与数值断言重复报
            val = cn_to_int(m.group(1))
            if val is None:
                continue
            fuzzy = bool(m.group(2)) or (m.start() > 0 and t[m.start() - 1] in "约近超")
            out.append(mk_assert(region, pi, t, m.start(), val, m.group(3), fuzzy,
                                 m.group(0)))
        # ③ 百分之三十五（点二）
        for m in CN_PCT_RE.finditer(t):
            base = cn_to_int(m.group(1))
            if base is None:
                continue
            val = float(base)
            if m.group(2):
                val += CN_DIGITS[m.group(2)] / 10.0
            fuzzy = m.start() > 0 and t[m.start() - 1] in "约近超"
            out.append(mk_assert(region, pi, t, m.start(), val, "%", fuzzy, m.group(0)))
        # ④ 三成 → 30%
        for m in CHENG_RE.finditer(t):
            n = cn_to_int(m.group(1))
            if n is None:
                continue
            fuzzy = m.start() > 0 and t[m.start() - 1] in "约近超"
            out.append(mk_assert(region, pi, t, m.start(), float(n * 10), "%", fuzzy,
                                 m.group(0)))
    return out


# ---------------- 区域切分 ----------------
def split_regions(path):
    """切出摘要 / 正文 / 结语三段文本（结语按收口章标题定位）。"""
    paras, refs, abs_text, kw_text, titles, body = dp.extract_paras(path)
    concl, in_closing = [], False
    for style, t in paras:
        if not t:
            continue
        if t.startswith("参考文献"):
            break
        is_title = (re.match(r"^(0|1|2|3|4|5|6|7)\s*[^\d]", t) and len(t) < 40) or \
                   (style.startswith("Heading") and len(t) < 40)
        if is_title:
            in_closing = bool(CLOSING_TITLE.search(t))
            continue
        if in_closing:
            concl.append(t)
    return {"abstract": dp.clean(abs_text or ""), "body": body,
            "conclusion": concl, "titles": titles, "refs": refs}


# ---------------- 交叉比对 ----------------
def _pair_label(src_region, pool_region):
    return f"{src_region}↔{pool_region}"


def compare_regions(src_as, pool_as, src_region):
    """把 src 区断言逐个与 pool 区比对，返回 (hard, todo, ok)。"""
    hard, todo, ok = [], [], []
    by_key = defaultdict(list)
    for b in pool_as:
        by_key[(b["unit"], b["head"])].append(b)
    for a in src_as:
        if len(a["head"]) < 2:
            continue  # 无前接名词（如「3 个城市」中的裸数字）不做精确配对
        cands = by_key.get((a["unit"], a["head"]), [])
        if cands:
            vals = {c["value"] for c in cands}
            b = cands[0]
            rec = {"pair": _pair_label(src_region, b["region"]), "head": a["head"],
                   "unit": a["unit"], "src_val": a["value"], "pool_val": b["value"],
                   "src_par": a["par"], "pool_par": b["par"], "src_snip": a["snip"],
                   "pool_snip": b["snip"], "match": "同名指标"}
            if a["value"] in vals:
                ok.append(rec)
            elif len(vals) == 1:
                (todo if (a["fuzzy"] or any(c["fuzzy"] for c in cands)) else hard).append(rec)
            else:
                rec["pool_vals"] = sorted(vals)
                todo.append(rec)
            continue
        # 关键词唯一兜底：前接名词被改写时的第二道防线
        kwc = [b for b in pool_as if b["unit"] == a["unit"] and (a["kw"] & b["kw"])]
        if len(kwc) == 1:
            b = kwc[0]
            rec = {"pair": _pair_label(src_region, b["region"]), "head": a["head"],
                   "unit": a["unit"], "src_val": a["value"], "pool_val": b["value"],
                   "src_par": a["par"], "pool_par": b["par"], "src_snip": a["snip"],
                   "pool_snip": b["snip"], "match": "关键词匹配（需人工确认）"}
            if abs(b["value"] - a["value"]) < 1e-9:
                ok.append(rec)
            else:
                (todo if (a["fuzzy"] or b["fuzzy"]) else hard).append(rec)
        elif len(kwc) > 1:
            vals = {b["value"] for b in kwc}
            if a["value"] not in vals:
                todo.append({"pair": _pair_label(src_region, kwc[0]["region"]),
                             "head": a["head"], "unit": a["unit"],
                             "src_val": a["value"], "pool_vals": sorted(vals),
                             "src_par": a["par"], "pool_par": kwc[0]["par"],
                             "src_snip": a["snip"], "pool_snip": kwc[0]["snip"],
                             "match": "多候选（人工判定）"})
    return hard, todo, ok


# ---------------- 稿内逻辑核验 ----------------
def check_pct_bounds(assertions):
    """百分比越界（>100% 或 <0）——明显的笔误。"""
    out = []
    for a in assertions:
        if a["unit"] == "%" and (a["value"] > 100 or a["value"] < 0):
            out.append({"where": f"{a['region']}·段{a['par']}",
                        "head": re.sub(r"^(其中|超过|近|约)+$", "", a["head"]) or a["head"],
                        "src_val": a["value"], "pool_val": None,
                        "match": "百分比越界（>100% 或 <0）", "snip": a["snip"]})
    return out


def check_sample_logic(assertions):
    """同段内「回收/有效样本 > 发放样本」——样本量逻辑矛盾。"""
    out = []
    by_par = defaultdict(list)
    for a in assertions:
        by_par[(a["region"], a["par"])].append(a)
    for _, lst in by_par.items():
        for f in lst:
            if "发放" not in f["head"]:
                continue
            for o in lst:
                if o is f or o["unit"] != f["unit"]:
                    continue
                if not (("有效" in o["head"]) or ("回收" in o["head"])):
                    continue
                if f["head"][-2:] != o["head"][-2:]:
                    continue  # 不是同一对象（如发放问卷 vs 回收试卷）
                if o["value"] > f["value"]:
                    out.append({"where": f"{o['region']}·段{o['par']}",
                                "head": o["head"], "unit": o["unit"],
                                "src_val": o["value"], "pool_val": f["value"],
                                "match": "回收/有效样本大于发放样本",
                                "src_snip": o["snip"], "pool_snip": f["snip"]})
    return out


def check_rate_consistency(assertions):
    """文中「有效率/回收率」与「有效÷发放」算不符（容差 0.6 个百分点）。"""
    out = []
    by_par = defaultdict(list)
    for a in assertions:
        by_par[(a["region"], a["par"])].append(a)
    for _, lst in by_par.items():
        f = next((x for x in lst if "发放" in x["head"]), None)
        v = next((x for x in lst if "有效" in x["head"] and f and
                  x["unit"] == f["unit"]), None)
        if not f or not v or not f["value"]:
            continue
        r = next((x for x in lst if x["unit"] == "%" and "率" in x["head"]), None)
        if not r:
            continue
        calc = v["value"] / f["value"] * 100
        if abs(calc - r["value"]) > 0.6:
            out.append({"where": f"{r['region']}·段{r['par']}", "head": r["head"],
                        "src_val": r["value"], "pool_val": round(calc, 2),
                        "match": "有效率与「有效÷发放」算不符",
                        "src_snip": r["snip"]})
    return out

def pct_groups(assertions):
    """同段多个百分比聚为「枚举组」（分类占比并列），用于整体差异与合计核验。"""
    by_par = defaultdict(list)
    for a in assertions:
        if a["unit"] == "%":
            by_par[(a["region"], a["par"])].append(a)
    groups = []
    for (region, par), lst in sorted(by_par.items()):
        if len(lst) < 2:
            continue
        kw = set()
        for a in lst:
            kw |= a["kw"]
        groups.append({"region": region, "par": par,
                       "vals": sorted({a["value"] for a in lst}), "kw": kw,
                       "snip": lst[0]["snip"]})
    return groups


def check_enum_groups(src_groups, pool_groups, src_region):
    """摘要的百分比枚举组与正文枚举组整体不一致（可能序指代不同，只报待核对）。"""
    out = []
    for g in src_groups:
        cand = [p for p in pool_groups if p["kw"] & g["kw"]]
        if not cand:
            continue
        pool_vals = set()
        for p in cand:
            pool_vals |= set(p["vals"])
        if not set(g["vals"]) <= pool_vals:
            out.append({"pair": f"{src_region}↔{cand[0]['region']}",
                        "par": g["par"], "src_vals": g["vals"],
                        "pool_vals": sorted(pool_vals), "snip": g["snip"]})
    return out


def check_pct_sum(groups):
    """含「分别/合计」信号的枚举组，百分比合计应约等于 100%（容差按四舍五入推断）。"""
    out = []
    for g in groups:
        if not (2 <= len(g["vals"]) <= 6):
            continue
        if not re.search(r"分别|合计|总共|累计|共计|总和|共占", g["snip"]):
            continue
        s = sum(g["vals"])
        tol = 0.6 * len(g["vals"]) + 0.5
        if abs(s - 100) > tol:
            out.append({"where": f"{g['region']}·段{g['par']}", "sum": round(s, 2),
                        "vals": g["vals"], "snip": g["snip"]})
    return out


# ---------------- 条目数（题次）核对 ----------------
# 「三类需求 / 四个方面 / 两项机制」的后续名词：取到量词后最近的语义名词
CLAIM_NOUN_TAIL = re.compile(r"[" + CJK + r"]{1,8}")
CLAIM_NOUN_SPLIT = re.compile(r"[的之与和及或、，。；：？（）()]")
CLAIM_NOUN_STOP = set("""的 之 与 和 及 或 在 是 为 对 从 于 以 而 并 提出 形成 归纳 划分 分为
结果 表明 显示 发现 可知 可见 框架 体系 结论 影响 因素 之一 其二 需求 服务
用户 读者 图书馆 群体 数据 信息 资源 建设 管理 工作 本 该 各 每 多 部分
全部 整体 总体 不同 各种 优化 提升 完善 健全 加强 推进 构建""".split())

# 条目名词前的填充动词（逐条剥；单字不在表内，避免误剥「对策」的「对」）
CLAIM_NOUN_PREFIX = re.compile(
    r"^(?:提出|形成|归纳|划分|分为|构建|确立|凝练|概括|得出|总结|发现|包括|包含|呈现|表现为|归为)")
# 名词后的截断词（碰到就停）
CLAIM_NOUN_CUT = re.compile(r"(框架|体系|机制|措施|维度|层面|方面|路径|对策|策略|建议|结论|调查)$")
# 纯泛词名词：只抽出这些词说明无区分度（「三类需求」），放弃该声称
CLAIM_NOUN_VAGUE = re.compile(r"^(需求|服务|用户|读者|群体|数据|信息|资源|工作|问题|现状|情况|内容)$")


def claim_noun(t, m):
    """从量词后提取条目名词（剔除泛化词与动词），无法确定返回 ""。

    「对策/建议/策略/路径/机制」是该刊对仗清单的通用量词名词，本身无区分度
    （摘要「四类对策」↔ 正文「三类对策」正是要抓的差异），因此保留；
    只剥「提出/形成/构建」类动词前缀与「需求/服务」类泛名词。
    """
    tail = CLAIM_NOUN_TAIL.match(t, m.end())
    if not tail:
        return ""
    seg = tail.group(0)
    noun = CLAIM_NOUN_SPLIT.split(seg)[0]
    noun = re.sub(r"(方面|层面|维度|类|种|项|条)+$", "", noun)
    # 截到最近的清单名词（「需求框架并提出三类对策」→ 取「对策」）。
    # 「三类需求」这类无清单名词的声称本身区分度低，直接放弃——摘要称「三类需求」、
    # 正文写「三类需求框架」时靠关键词兜底通道也能发现差异。
    cut = CLAIM_NOUN_CUT.search(noun)
    if not cut:
        return ""
    noun = noun[cut.start():]
    prev = None
    while prev != noun:
        prev = noun
        noun = CLAIM_NOUN_PREFIX.sub("", noun)
    if len(noun) < 2 or noun in GENERIC or noun in CLAIM_NOUN_STOP:
        return ""
    if CLAIM_NOUN_VAGUE.match(noun):
        return ""
    return noun


def extract_claims(text, region, par_i=1):
    """抽取「三类需求 / 四个方面 / 两项机制」类条目数声称。"""
    t = norm_text(text or "")
    out = []
    for m in CLAIM_RE.finditer(t):
        raw_n = m.group(1)
        n = int(raw_n) if raw_n.isdigit() else cn_to_int(raw_n)
        if n is None or n < 2 or n > 20:
            continue
        noun = claim_noun(t, m)
        if not noun:
            continue
        kws = set()
        win = head_before(t, m.start(), 10)
        for run in re.findall(r"[" + CJK + r"]{2,}", win):
            for j in range(len(run) - 1):
                bg = run[j:j + 2]
                if bg not in GENERIC:
                    kws.add(bg)
        out.append({"region": region, "par": par_i, "n": n, "noun": noun, "kw": kws,
                    "snip": snip(t, m.start())})
    return out


def compare_claims(src_claims, pool_claims, src_region):
    """摘要声称条目数 ↔ 正文声称条目数（同名名词不同值即硬伤）。"""
    hard, todo = [], []
    by_noun = defaultdict(list)
    for c in pool_claims:
        by_noun[c["noun"]].append(c)
    for a in src_claims:
        cands = by_noun.get(a["noun"], [])
        if cands:
            ns = {c["n"] for c in cands}
            if a["n"] in ns:
                continue
            rec = {"pair": f"{src_region}↔{cands[0]['region']}", "noun": a["noun"],
                   "src_val": a["n"], "pool_val": sorted(ns)[0] if len(ns) == 1 else None,
                   "pool_vals": sorted(ns) if len(ns) > 1 else None,
                   "src_par": a["par"], "pool_par": cands[0]["par"],
                   "src_snip": a["snip"], "pool_snip": cands[0]["snip"],
                   "match": "同名名词条目数不一致"}
            (hard if len(ns) == 1 else todo).append(rec)
            continue
        kwc = [c for c in pool_claims if c["kw"] & a["kw"]]
        if len(kwc) == 1 and kwc[0]["n"] != a["n"]:
            todo.append({"pair": f"{src_region}↔{kwc[0]['region']}", "noun": a["noun"],
                         "src_val": a["n"], "pool_val": kwc[0]["n"],
                         "src_par": a["par"], "pool_par": kwc[0]["par"],
                         "src_snip": a["snip"], "pool_snip": kwc[0]["snip"],
                         "match": "关键词匹配（需人工确认）"})
    return hard, todo


def count_ordinals(paras):
    """正文实际列举条数（一是/其一/（1）…去重计数）。"""
    found = set()
    for p in paras:
        t = norm_text(p or "")
        for mk_ in ORD_MARKERS:
            if mk_ in t:
                found.add(mk_)
    return len(found)


def check_claim_vs_enum(src_claims, pool_paras, pool_region, src_region):
    """摘要声称 N 项 ↔ 正文实际列举 M 项。"""
    out = []
    for a in src_claims:
        hit = [p for p in pool_paras if a["kw"] and
               any(k in norm_text(p) for k in a["kw"])]
        if len(hit) != 1:
            continue
        m = count_ordinals(hit)
        if m and m != a["n"]:
            out.append({"pair": f"{src_region}↔{pool_region}", "noun": a["noun"],
                        "src_val": a["n"], "pool_val": m, "src_par": a["par"],
                        "pool_par": None, "src_snip": a["snip"],
                        "pool_snip": snip(norm_text(hit[0]), 0, 0, 40),
                        "match": "声称条数 vs 实际列举"})
    return out


# ---------------- 输出 ----------------
def main():
    args = [a for a in sys.argv[1:] if not a.startswith("--")]
    flags = set(a for a in sys.argv[1:] if a.startswith("--"))
    if not args or args[0] in ("-h", "--help"):
        print(__doc__)
        sys.exit(0)
    DOC = args[0]
    if not os.path.isfile(DOC):
        sys.exit(f"找不到文件：{DOC}")
    TOP = 12
    if "--top" in sys.argv:
        i = sys.argv.index("--top")
        try:
            TOP = int(sys.argv[i + 1])
        except (IndexError, ValueError):
            pass

    reg = split_regions(DOC)
    abs_paras = [reg["abstract"]] if reg["abstract"] else []
    body_paras = reg["body"]
    concl_paras = reg["conclusion"]

    abs_as = extract_assertions(abs_paras, "摘要")
    body_as = extract_assertions(body_paras, "正文")
    concl_as = extract_assertions(concl_paras, "结语")
    pool_as = body_as + concl_as

    print("=" * 70)
    print("摘要—正文数字一致性核验（数字硬伤专查）")
    print(f"稿件：{os.path.basename(DOC)}")
    print(f"摘要断言 {len(abs_as)} 条 ｜ 正文 {len(body_as)} 条 ｜ 结语 {len(concl_as)} 条")
    print("=" * 70)

    hard, todo, ok = [], [], []
    if not abs_as:
        print("未检出摘要数值断言（若无摘要或摘要不含数字，本项不适用）。")
    else:
        h1, t1, o1 = compare_regions(abs_as, pool_as, "摘要")
        hard += h1
        todo += t1
        ok += o1
        # 百分比枚举组：摘要 vs 正文+结语
        ag = pct_groups(abs_as)
        pg = pct_groups(pool_as)
        todo += check_enum_groups(ag, pg, "摘要")
        todo += check_pct_sum(ag)
        # 结语 vs 正文
        h2, t2, o2 = compare_regions(concl_as, body_as, "结语")
        hard += h2
        todo += t2
        ok += o2
        todo += check_pct_sum(pct_groups(concl_as))

    # 稿内逻辑（不分区域）
    hard += check_sample_logic(abs_as + body_as + concl_as)
    hard += check_pct_bounds(abs_as + body_as + concl_as)
    todo += check_rate_consistency(abs_as + body_as + concl_as)

    # 条目数（题次）核对
    abs_claims = extract_claims(reg["abstract"], "摘要")
    body_claims = []
    for i, p in enumerate(body_paras):
        body_claims += extract_claims(p, "正文", i + 1)
    concl_claims = extract_claims("\n".join(concl_paras), "结语")
    ch, ct = compare_claims(abs_claims, body_claims + concl_claims, "摘要")
    hard += ch
    todo += ct
    todo += check_claim_vs_enum(abs_claims, body_paras, "正文", "摘要")

    print(f"\n【交叉比对】🔴 硬伤 {len(hard)} 项 ｜ 🟡 待人工核对 {len(todo)} 项 ｜ "
          f"✅ 一致 {len(ok)} 组")
    print("-" * 70)
    if hard:
        print(f"🔴 摘要与正文数字硬伤 {len(hard)} 项（投稿前必须处理）：")
        for r in hard[:TOP]:
            _show(r)
        if len(hard) > TOP:
            print(f"   …另 {len(hard) - TOP} 项，用 --top N 调整显示条数")
    else:
        print("🔴 摘要与正文数字硬伤 0 项")

    if todo:
        print(f"\n🟡 数字待人工核对 {len(todo)} 项（模糊量/多候选/枚举组差异，非硬伤）：")
        for r in todo[:TOP]:
            _show(r, hard=False)
        if len(todo) > TOP:
            print(f"   …另 {len(todo) - TOP} 项，用 --top N 调整显示条数")

    if not hard and not todo:
        print("✅ 摘要、正文、结语的数值断言未检出矛盾（词面启发式，仍建议抽查关键数字）。")

    if "--json" in flags:
        jf = os.path.splitext(DOC)[0] + ".number_consistency.json"
        with open(jf, "w", encoding="utf-8") as f:
            json.dump({"doc": os.path.basename(DOC),
                       "counts": {"abstract": len(abs_as), "body": len(body_as),
                                  "conclusion": len(concl_as)},
                       "hard": hard, "todo": todo, "ok_n": len(ok),
                       "claims": {"abstract": abs_claims, "body": body_claims}},
                      f, ensure_ascii=False, indent=1)
        print(f"\nJSON 已写入：{jf}")

    sys.exit(1 if hard else 0)


def _show(r, hard=True):
    """单条结果输出（硬伤/待核对共用；稿内逻辑行无 pair，用 where 定位）。"""
    if r.get("pool_vals"):
        pv = "／".join(fmt_val(x) for x in r["pool_vals"])
    else:
        pv = fmt_val(r.get("pool_val")) if r.get("pool_val") is not None else "?"
    name = r.get("head") or r.get("noun") or "数值"
    unit = r.get("unit", "")
    loc = r.get("pool_par") or r.get("src_par") or r.get("par")
    if r.get("pair"):
        head = f"   · [{r['pair']}·段{loc}]"
    else:
        head = f"   · [{r.get('where', '本稿')}]"
    if "calc" in r:
        print(f"{head} {name}：文中值 {fmt_val(r['stated'])}% ≠ 算式 "
              f"{fmt_val(r['calc'])}%（{r.get('match', '有效率与除式不符')}）")
    elif "率" in (r.get("head") or ""):
        print(f"{head} {name}：{fmt_val(r['src_val'])}{unit} ≠ 算式 "
              f"{fmt_val(r['pool_val'])}%（{r.get('match')}）")
    elif r.get("pool_val") is None:
        print(f"{head} {name}：{fmt_val(r['src_val'])}{unit}（{r.get('match')}）")
    else:
        print(f"{head} {name}：{fmt_val(r['src_val'])}{unit} ≠ {pv}{unit}"
              f"（{r.get('match', '稿内逻辑矛盾')}）")
    if r.get("src_snip"):
        print(f"       {trunc(r['src_snip'], 74)}")
    if r.get("pool_snip"):
        print(f"       ↳ {trunc(r['pool_snip'], 74)}")
    elif "src_snip" not in r and r.get("snip"):
        print(f"       {trunc(r['snip'], 74)}")


if __name__ == "__main__":
    main()
