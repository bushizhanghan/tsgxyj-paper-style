# -*- coding: utf-8 -*-
"""diagnose_paper.py —— 《图书馆学研究》论文诊断（对照 108 篇实测基线）

用法（推荐在技能 scripts/ 目录或任一处运行）:
    python diagnose_paper.py <论文.docx>
    python diagnose_paper.py <论文.docx> --corpus <语料库目录>   # 显式指定语料库
    python diagnose_paper.py <论文.docx> --json                  # 附加输出机器可读 JSON
    python diagnose_paper.py <论文.docx> --brief                 # 只输出问题清单（红线优先）
    python diagnose_paper.py <论文.docx> --paragraphs            # 只看段落级扫描

输出四段：
    ① 篇目概览（结构 / 摘要 / 关键词 / 文献 / 图表）
    ② 量化对照（段长 / 句均 / 引用 / 「本文」——对照语料库实测基线，
       语料库缺失时用 style_stats.md 静态值并明确标注）
    ③ 反模式扫描（对比该刊基线，超 2 倍标 ⚠️）
    ④ 段落级扫描（改写定位：句均<45 / 序数口语词 / 超长段 / 碎段）

基线来源：
    - 语料库 corpus_stats.json（108 篇逐篇统计）→ 动态汇总
    - 缺语料库 → corpus.py 内置静态基线（style_stats.md 记录值）
"""
import re
import sys
import json
import os
import statistics as st

# Windows 控制台默认 GBK 无法输出部分 UTF-8 字符，统一走 UTF-8
if hasattr(sys.stdout, "reconfigure"):
    try:
        sys.stdout.reconfigure(encoding="utf-8", errors="replace")
    except Exception:
        pass

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import corpus  # noqa: E402


def extract_paras(path):
    """从 docx 提取段落，返回 (paras, refs, abs_text, kw_text, titles, body)。
    paras: [(style, text)] 全量；其余为分类结果。"""
    from docx import Document
    doc = Document(path)
    paras = [(p.style.name if p.style else "", p.text.strip()) for p in doc.paragraphs]
    abs_text = kw_text = ""
    refs, body, titles = [], [], []
    in_ref = False
    for style, t in paras:
        if not t:
            continue
        if t.startswith("参考文献"):
            in_ref = True
            continue
        if in_ref:
            if re.match(r"^\[?\d{1,3}\]", t) or re.match(r"^[\[［]\d+[\]］]", t):
                refs.append(t)
            continue
        if re.search(r"摘\s*要", t[:10]) and not abs_text:
            abs_text = re.sub(r"^【?摘\s*要】?[：:]?", "", t)
            continue
        if re.search(r"关\s*键\s*词", t[:10]) and not kw_text:
            kw_text = t
            continue
        if (re.match(r"^(0|1|2|3|4|5|6|7)\s*[^\d]", t) and len(t) < 40) or \
           (style.startswith("Heading") and len(t) < 40):
            titles.append(t)
            continue
        if len(t) < 45 and not re.search(r"[。；，]", t) and len(body) > 0:
            continue  # 图题/表题
        if len(t) > 45 or (not body):
            if "摘" in t[:6] or "Abstract" in t[:12]:
                continue
        body.append(t)
    return paras, refs, abs_text, kw_text, titles, body


def clean(s):
    return re.sub(r"\s+", "", s)


def cjk_len(s):
    return len(clean(s))


# ---- 结构完备度（对照该刊 108 篇一级标题命中率）----
# style_stats.md §1：对策/策略/启示/路径 67%、引言 52%、结语 52%、
# 研究设计/对象与方法 37%、文献综述/相关研究 29%。
# 三级标题：corpus_stats.json heads 实测 85/108 篇含三级（79%），
# 每篇中位 6.5 个、均值 8.9、最大 64——三级标题是该刊常态，不判为问题；
# 仅当单篇三级数显著超过实测 P90（约 22 个）时提示「是否过碎」。
H3_ALERT_THRESHOLD = 22
STRUCT_PATTERNS = [
    ("引言", r"^\s*0?\s*引\s*言"),
    ("文献综述/相关研究", r"^\s*\d+\s*(文献综述|相关研究|研究综述|研究述评|国内外)"),
    ("研究设计/对象与方法", r"^\s*\d+\s*(研究设计|研究对象|研究方法|研究思路|概念界定|理论基础)"),
    ("对策/策略/路径/启示（收口标配）", r"^\s*\d+\s*.*(对策|策略|路径|启示|建议|提升|优化)"),
    ("结语", r"^\s*\d+\s*结\s*语"),
]


def scan_structure(titles):
    """结构完备度体检：返回 {found: [...], missing: [...], h3_count, note}。"""
    norm_titles = [re.sub(r"\s+", " ", t.strip()) for t in titles]
    found, missing = [], []
    for name, pat in STRUCT_PATTERNS:
        if any(re.search(pat, t) for t in norm_titles):
            found.append(name)
        else:
            missing.append(name)
    # 三级标题（该刊常态：85/108 篇含三级，中位 6.5 个/篇；仅过量提示）
    h3 = [t for t in norm_titles if re.match(r"^\d+\.\d+\.\d+", t)]
    # 一级标题数（0/1/2...开头）
    h1 = [t for t in norm_titles if re.match(r"^\d+\s", t) and not re.match(r"^\d+\.\d+", t)]
    return {"found": found, "missing": missing, "h3_count": len(h3),
            "h1_count": len(h1), "note": ""}


# ---- 章节字数占比（对照 style_stats.md §1 实测中位）----
# 仅统计单列该章的篇目：引言 7.1%(56篇)、文献综述 15.5%(24篇)、
# 研究设计/方法 14.9%(26篇)、对策/策略 21.9%(16篇)、结语 4.1%(23篇)
SEC_BASE = {
    "引言": 7.1, "文献综述/相关研究": 15.5, "研究设计/对象与方法": 14.9,
    "结果/分析": None,  # 该刊占比方差大，不设基准
    "对策/策略/路径/启示": 21.9, "结语": 4.1,
}
SEC_CLASS = [
    ("引言", r"^\s*0?\s*引\s*言"),
    ("文献综述/相关研究", r"^\s*\d+\s*(文献综述|相关研究|研究综述|研究述评|国内外)"),
    ("研究设计/对象与方法", r"^\s*\d+\s*(研究设计|研究对象|研究方法|研究思路|概念界定|理论基础|模型|框架)"),
    ("对策/策略/路径/启示", r"^\s*\d+\s*.*(对策|策略|路径|启示|建议|提升|优化)"),
    ("结果/分析", r"^\s*\d+\s*(结果|分析|实证|发现|识别|评价|测度|比较|调研|调查)"),
    ("结语", r"^\s*\d+\s*结\s*语"),
]


# ---- 官方合规检查（references/official-guidelines.md 明文要求）----
# 关键词泛词（官方：「不应出现无检索意义的词，如研究/概述/发展等包含范围过广的词语」）
KW_BAN = ["研究", "概述", "发展", "探讨", "分析", "初探", "浅析", "综述", "思考", "实践"]
# 官方九类栏目（用于对位提示）
OFFICIAL_COLUMNS = {
    "基础理论（研究进展）": r"理论|学科|范式|知识体系|研究进展|述评",
    "专业教育": r"教育|培养|课程|素养|教学|人才",
    "图书馆管理": r"管理|治理|绩效|评价|评估|制度",
    "数字化、网络化": r"数字化|智慧图书馆|人工智能应用|大语言模型|算法模型|技术架构|网络化",
    "信息资源建设": r"资源建设|馆藏|采购|开放获取|文献资源|资源整合",
    "信息组织": r"标引|分类法|元数据|知识组织|本体构建|编目",
    "信息服务": r"知识服务|信息服务|咨询服务|信息推送|参考咨询",
    "读者工作": r"读者|用户|障碍|视障|听障|老年|弱势|可及性|无障碍|需求|使用行为",
    "国外图书馆": r"国外|美国|英国|日本|加拿大|澳大利亚|海外|国际比较",
}


def scan_official(kw_text, refs, titles, alltext, paras=None):
    """官方须知合规检查。返回 {issues: [...], columns: [...], package_notices: [...] 等}。
    paras: [(style, text)] 全量段落，供检查英文题名/摘要/关键词与文末作者元数据。"""
    issues = []
    package_notices = []  # 投稿必备包要素提醒（作者简介/定稿日期/通信方式等）

    full_paras_text = "\n".join(t for _, t in paras) if paras else alltext

    # ① 关键词泛词
    kws = [k.strip() for k in re.split(r"[\s;；,，、]+",
                                       re.sub(r"^【?关\s*键\s*词】?[：:]?", "", kw_text or ""))
           if k.strip()]
    for k in kws:
        for b in KW_BAN:
            if k == b or (b in k and len(k) <= len(b) + 2):
                issues.append(f"关键词「{k}」属官方点名的泛词（无检索意义），建议替换")
                break
    # ② 英文文献标题实意词首字母大写（官方明文）
    bad_en = []
    for r in refs:
        # 英文条目：含 [J]/[C]/[M] 且题名以拉丁字母为主
        m = re.match(r"^[\[［]?\s*\d+\s*[\]］]?\s*(.+?)[\.．]\s*(.+?)\[(J|C|M|D|N|S)",
                     r)
        if not m:
            continue
        title = m.group(2).strip()
        if not re.search(r"[A-Za-z]{4,}", title):
            continue
        if not re.search(r"[\u4e00-\u9fff]", title):
            # 纯英文题名 → 检查实意词是否首字母大写
            words = [w for w in re.findall(r"[A-Za-z][A-Za-z\-]{2,}", title)
                     if w.lower() not in {"the", "and", "for", "with", "from", "into",
                                          "of", "in", "on", "to", "a", "an", "by", "or"}]
            lower = [w for w in words if w[0].islower()]
            if lower:
                bad_en.append((r[:36], lower[:3]))
    if bad_en:
        issues.append(f"{len(bad_en)} 条英文文献题名实意词未首字母大写"
                      f"（官方明文要求），例：{bad_en[0][1]}")
    # ③ 栏目对位（按命中次数排序，取最匹配者）
    col_hits = []
    for name, pat in OFFICIAL_COLUMNS.items():
        c = len(re.findall(pat, alltext))
        if c:
            col_hits.append((c, name))
    col_hits.sort(reverse=True)
    cols = [name for _, name in col_hits]
    if not cols:
        issues.append("未能判定所属栏目——官方九栏为：基础理论/专业教育/图书馆管理/"
                      "数字化网络化/信息资源建设/信息组织/信息服务/读者工作/国外图书馆")
    # ④ 注释序号递增（官方：按递增顺序，同一文献不重复给号）
    nums = []
    for r in refs:
        m = re.match(r"^[\[［]?\s*(\d{1,3})", r)
        if m:
            nums.append(int(m.group(1)))
    if nums and nums != sorted(nums):
        issues.append("参考文献序号非递增排列（官方要求按递增顺序）")
    if nums and len(nums) != len(set(nums)):
        issues.append("参考文献有重复序号（官方要求同一文献用同一序号、不重复给号）")
    # ⑤ 末章名称（官方序列写「结语」）
    has_jieyu = any(re.search(r"结\s*语", t) for t in titles)
    has_jielun = any(re.search(r"结\s*论|结束语|余论", t) for t in titles)
    if has_jielun and not has_jieyu:
        issues.append("末章名为「结论/结束语」——官方格式序列写「结语」，建议对齐")

    # ⑥ 官方排版序列与对称性检查（Abstract: 与 Key words:）
    has_en_abs = bool(re.search(r"(?:Abstract|ABSTRACT)\s*[:：\s]", full_paras_text))
    has_en_kws = bool(re.search(r"Key\s*words?\s*[:：\s]", full_paras_text, re.I))
    if not has_en_abs:
        issues.append("缺少英文摘要（官方明文格式：Abstract:，须中英文对称）")
    if not has_en_kws:
        issues.append("缺少英文关键词（官方明文格式：Key words:，须中英文对应）")

    # ⑦ 投稿必备包完整性（文末作者简介、联系方式、定稿日期）
    has_bio = bool(re.search(r"作者简介|第一作者|通讯作者|【作者简介】", full_paras_text))
    has_contact = bool(re.search(r"电话|邮箱|E-?mail|通信地址|通讯地址|邮编|邮政编码", full_paras_text, re.I))
    has_date = bool(re.search(r"定稿日期|收稿日期", full_paras_text))

    if not has_bio:
        package_notices.append("文末未见「作者简介」（姓名、单位、职称、学位、研究方向；官方要求初审完备）")
    if not has_contact:
        package_notices.append("文末未见「联系方式」（详细通信地址、邮编、电话、电子邮箱；官方投稿必备）")
    if not has_date:
        package_notices.append("文末未见「定稿日期」（官方格式序列要求末尾附定稿日期）")

    return {
        "issues": issues, "columns": cols, "kws": kws,
        "en_title_issues": len(bad_en),
        "package_notices": package_notices,
        "has_en_abs": has_en_abs,
        "has_en_kws": has_en_kws,
    }


def scan_sections(paras, titles):
    """按标题切分正文，统计各章字数与占比。paras 为 [(style, text)]。"""
    # 标题在全文段中的位置
    title_pos = []
    tset = {t for t in titles}
    for idx, (style, t) in enumerate(paras):
        s = re.sub(r"\s+", " ", t.strip())
        if s in tset or (len(s) < 40 and (
                re.match(r"^\d+(\.\d+)?\s", s) or style.startswith("Heading"))):
            title_pos.append((idx, s))
    if not title_pos:
        return None
    bounds = [(p, s) for p, s in title_pos]
    secs = []
    total = 0
    for i, (pos, name) in enumerate(bounds):
        end = bounds[i + 1][0] if i + 1 < len(bounds) else len(paras)
        chars = 0
        for j in range(pos + 1, end):
            txt = paras[j][1]
            if re.match(r"^\s*(参考文献|摘\s*要|关\s*键\s*词)", txt):
                continue
            if len(txt) < 45 and not re.search(r"[。；，]", txt):
                continue  # 图题表题
            chars += cjk_len(txt)
        if chars:
            cls = None
            for cname, pat in SEC_CLASS:
                if re.search(pat, name):
                    cls = cname
                    break
            secs.append({"name": name, "chars": chars, "cls": cls})
            total += chars
    if not total:
        return None
    for s in secs:
        s["pct"] = s["chars"] / total * 100
    # 同类合并占比
    merged = {}
    for s in secs:
        if s["cls"]:
            merged[s["cls"]] = merged.get(s["cls"], 0) + s["chars"]
    return {"secs": secs, "total": total,
            "merged": {k: v / total * 100 for k, v in merged.items()}}


def scan_abstract(abs_text, title_heads=""):
    """摘要四要素体检（section-guides.md 四要素公式 + official-guidelines.md 官方忌词/字数）。

    为什么做 → 怎么做 → 发现什么 → 什么价值。仅做形态提示，不判定内容质量。
    """
    t = clean(re.sub(r"^【?摘\s*要】?[：:]?\n?", "", abs_text or ""))
    if not t:
        return {"present": False}
    issues = []
    # 忌词（官方明文：分析了/总结了/提出了/解决了/我们/本文）
    OFFICIAL_TABOO = {
        "本文": "官方明文忌用（须第三人称）",
        "我们": "官方明文忌用（须第三人称）",
        "分析了": "官方明文忌用",
        "总结了": "官方明文忌用",
        "提出了": "官方明文忌用",
        "解决了": "官方明文忌用",
    }
    for w, why in OFFICIAL_TABOO.items():
        if w in t:
            issues.append(f"含忌词「{w}」——{why}")
    # 字数：官方「200 字左右」，实测中位 262
    n = len(t)
    if n < 180:
        issues.append(f"{n} 字偏短（官方要求 200 字左右；实测中位 262）")
    elif n > 320:
        issues.append(f"{n} 字偏长（官方要求 200 字左右；该刊 300 字以上仅 30 篇）")
    # 四要素启发式：目的/方法/结果/价值
    sents = [s for s in re.split(r"[。；;]", t) if s]
    joined = t
    pools = {
        "目的": r"为[了服]?|旨在|面向|针对|以期",
        "方法": r"采用|利用|以.{1,18}为对象|基于|借助|通过.{1,16}(方法|模型|框架)|构建",
        "结果": r"表明|发现|显示|构建|形成|提出|落在|划分为|包含|得出",
        "价值": r"提供|启示|参考|依据|路径|策略|对策|支撑",
    }
    has = {k: bool(re.search(v, joined)) for k, v in pools.items()}
    for k in ("目的", "方法", "结果", "价值"):
        if not has[k]:
            issues.append(f"未检出「{k}」要素信号（{pools[k].strip()}）——请人工确认")
    # 空洞评语与自我评价（官方：无空洞评语，不进行自我评价）
    for w in ("取得了良好效果", "起到了积极作用", "效果显著", "具有重要意义",
              "具有重要价值", "填补了空白", "具有推广价值"):
        if w in t:
            issues.append(f"空洞评语「{w}」——官方要求无空洞评语、不自我评价")
    return {"present": True, "chars": n, "n_sents": len(sents),
            "elements": has, "issues": issues}


def compute_metrics(body, refs, abs_text, kw_text, titles):
    bk = [p for p in body if len(clean(p)) >= 30]
    lens = [cjk_len(p) for p in bk]
    sents = []
    for p in bk:
        sents += [s for s in re.split(r"[。？！]", clean(p)) if len(s) >= 6]
    alltext = clean("".join(bk))
    cites = len(re.findall(r"\[\d+([-,，]\s*\d+)*\]", alltext))
    benwen = alltext.count("本文")
    abs_len = cjk_len(re.sub(r"^【?摘\s*要】?[：:]?", "", abs_text))
    kw_n = len([k for k in re.split(r"[\s;；,，、]+",
                                    re.sub(r"^【?关\s*键\s*词】?[：:]?", "", kw_text)) if k])
    figs = [int(m) for m in re.findall(r"图\s*(\d{1,2})", alltext)]
    tabs = [int(m) for m in re.findall(r"表\s*(\d{1,2})", alltext)]
    fig_n = max([x for x in figs if x <= 20], default=0)
    tab_n = max([x for x in tabs if x <= 20], default=0)
    return {
        "body_chars": len(alltext), "para_n": len(lens), "para_med": st.median(lens) if lens else 0,
        "sent_mean": st.mean([len(s) for s in sents]) if sents else 0,
        "cite_rate": cites / max(len(alltext), 1) * 1000,
        "ref_n": len(refs), "abs_len": abs_len, "kw_n": kw_n,
        "fig_n": fig_n, "tab_n": tab_n, "benwen": benwen,
        "benwen_rate": benwen / max(len(alltext), 1) * 10000,
        "titles": titles, "body": bk,
    }


# ---- 反模式基线（该刊 118 篇 / 121 万字原文实测，次/万字）----
# 来源：references/anti-patterns.md A/B 级清单；B 级为该刊亦用、控密度即可。
ANTI_BASE = {
    "A级红线句（显而易见/众所周知/总而言之等）": 0.05,
    "万能句（具有重要意义/值得深入探讨/提供有益参考）": 0.10,
    "机械序数（首先/其次/再次+逗号）": 1.44,
    "「不仅…而且」（该刊罕用，主流是更/还）": 0.08,
    "空转连接（综上/在此基础上/与此同时）": 1.16,
    "三连「×化」排比": 0.13,
}
ANTI_PATS = {
    "A级红线句（显而易见/众所周知/总而言之等）": (
        r"显而易见|众所周知|毋庸置疑|总而言之|不难看出|值得一提|总的来说|值得深入探讨"),
    "万能句（具有重要意义/值得深入探讨/提供有益参考）": (
        r"具有重要意义|提供有益参考|有益参考|重要指导意义|意义重大|具有重要价值"),
    "机械序数（首先/其次/再次+逗号）": r"(首先|其次|再次)[，,]",
    "「不仅…而且」（该刊罕用，主流是更/还）": r"不仅[^。]{2,50}而且",
    "空转连接（综上/在此基础上/与此同时）": r"综上所述|在此基础上|与此同时",
    "三连「×化」排比": (
        r"[\u4e00-\u9fff]{1,4}化[、,，][\u4e00-\u9fff]{1,4}化[、,，][\u4e00-\u9fff]{1,4}化"),
}


def scan_anti_patterns(alltext):
    n = max(len(alltext), 1)
    out = []
    for name, pat in ANTI_PATS.items():
        ms = list(re.finditer(pat, alltext))
        rate = len(ms) / n * 10000
        flag = rate > ANTI_BASE[name] * 2
        sample = ""
        if ms:
            i = ms[0].start()
            sample = alltext[max(0, i - 12):i + 22]
        out.append({"name": name, "n": len(ms), "rate": round(rate, 2),
                    "base": ANTI_BASE[name], "flag": flag, "sample": sample})
    return out


# ---- 该刊罕见表达（润色词典，依据 references/polish-lexicon.md 118 篇实证）----
# 每条：(该刊罕用词, 该刊对应书面词)。仅在段落级扫描命中时报，给替换方向。
COLLOQUIAL_HINTS = [
    ("差不多", "基本"), ("大概", "约"), ("越来越好", "日益/持续"), ("一下子", "迅速"),
    ("有时候", "偶尔"), ("搞得", "导致"), ("事儿", "工作"), ("咱们", "我国"),
    ("老百姓", "公众"), ("活儿", "任务"), ("靠谱", "可信"), ("给力", "有效"),
    ("走心", "精准"), ("硬核", "核心"), ("弄", "厘清"), ("东西", "内容/要素"),
    ("大家", "用户/公众"), ("什么样", "特征/何种形态"), ("怎么", "如何"),
    ("其实", "事实上/本质上"), ("当然", "固然/诚然"), ("看来", "由此可见"),
]
# 动态助词「了」的润色提示（依据句法特征统计：该刊 0.25 次/千字）


def scan_paragraphs(body, min_len=80, over_len=450, short_len=110):
    """段落级扫描（改写定位用，不按全篇均值动手）。"""
    issues = []
    for idx, para in enumerate(body):
        c = clean(para)
        if len(c) < min_len:
            continue
        sents = [s for s in re.split(r"[。？！]", c) if len(s) >= 6]
        sj = st.mean([len(s) for s in sents]) if sents else 0
        ordn = len(re.findall(r"排在首位|排在最后|其次|再次|更进一层|合起来看|各有一套", c))
        oral = len(re.findall(r"很多|非常|特别|大概|差不多|其实|当然|看来|越来越好", c))
        tags = []
        if sj < 45:
            tags.append(f"句均 {sj:.0f} 偏短")
        if ordn >= 2:
            tags.append(f"序数/口语 ×{ordn}")
        if oral >= 2:
            tags.append(f"口语词 ×{oral}")
        # 该刊罕用词（润色词典实证：刊内 0–17 次/121 万字）
        hits = [(w, f) for w, f in COLLOQUIAL_HINTS if w in c]
        if hits:
            tags.append("罕用词：" + "、".join(f"{w}→{f}" for w, f in hits[:3]))
        # 「了」密度：该刊 4380 段实测——含「了」的段仅 37%，每段中位 0 个，
        # 「了」≥3 的段仅 8.4%。故段内 ≥3 个提示（属该刊少数形态）
        n_le = c.count("了")
        if n_le >= 3:
            tags.append(f"「了」×{n_le}（该刊 37% 段含「了」、中位 0 个，考虑改「能够/得到」或删）")
        if len(c) > over_len:
            tags.append(f"超长段 {len(c)}字·考虑切分")
        if len(c) < short_len:
            tags.append(f"碎段 {len(c)}字·考虑合并")
        if tags:
            issues.append({"para_no": idx + 1, "chars": len(c),
                           "tags": "｜".join(tags), "head": c[:36]})
    return issues


def fmt_pos(v, stat, higher_better=True, unit="", fmt="{:.0f}"):
    """输出「 vs 刊中位 + 位置判断」（仅对比部分，值由调用方前缀打印）。"""
    med = stat.get("med", 0)
    n = stat.get("n", 0)
    if not n:
        return "（无基线可比）"
    tag = ""
    if higher_better and v > stat.get("q75", med):
        tag = " ▲ 高于 75% 刊文"
    elif not higher_better and v < stat.get("q25", med):
        tag = " ▲ 低于 25% 刊文"
    elif higher_better and v < stat.get("q25", med):
        tag = " ▼ 低于 25% 刊文"
    elif not higher_better and v > stat.get("q75", med):
        tag = " ▼ 高于 75% 刊文"
    elif med and abs(v - med) / med < 0.25:
        tag = " ≈ 居中"
    return f" vs 刊中位 {fmt.format(med)}{unit}{tag}"


def main():
    args = sys.argv[1:]
    if not args or args[0] in ("-h", "--help"):
        print(__doc__)
        return 0
    # 解析参数
    cli_corpus = None
    flags = set()
    rest = []
    for a in args:
        if a.startswith("--corpus"):
            cli_corpus = args[args.index(a) + 1] if args.index(a) + 1 < len(args) else None
        elif a.startswith("--"):
            flags.add(a)
        else:
            rest.append(a)
    path = rest[0] if rest else None
    if not path:
        print("用法: python diagnose_paper.py <论文.docx> [--corpus <dir>] [--json] [--brief] [--paragraphs]")
        return 2
    if not os.path.isfile(path):
        print(f"找不到文件：{path}")
        return 2

    corpus_dir = None
    try:
        corpus_dir = corpus.find_corpus_dir(cli_corpus)
    except FileNotFoundError as e:
        print(f"[警告] {e}，将使用内置静态基线。")

    base = corpus.load_reference_stats(corpus_dir)
    paras, refs, abs_text, kw_text, titles, body = extract_paras(path)
    m = compute_metrics(body, refs, abs_text, kw_text, titles)
    abst = scan_abstract(abs_text)
    ant = scan_anti_patterns("".join(clean(p) for p in body))
    para_issues = scan_paragraphs(m["body"])
    struct = scan_structure(m["titles"])
    secs = scan_sections(paras, m["titles"])
    official = scan_official(kw_text, refs, m["titles"], "".join(clean(p) for p in body), paras=paras)

    lines = []
    lines.append("=" * 62)
    lines.append("论文诊断：对照《图书馆学研究》实测基线")
    lines.append(f"文档：{os.path.basename(path)}")
    lines.append(f"基线：{base['source']}" + ("（注意：内置静态值）" if base["static"] else ""))
    lines.append("=" * 62)

    lines.append(f"\n【① 篇目概览】正文 {m['body_chars']} 字 | 段 {m['para_n']} | "
                 f"摘要 {m['abs_len']} 字 | 关键词 {m['kw_n']} 个 | "
                 f"文献 {m['ref_n']} 条 | 图 {m['fig_n']} | 表 {m['tab_n']}")
    lines.append(f"  结构标题 {len(m['titles'])} 个：" + " › ".join(t[:14] for t in m['titles'][:12]))
    lines.append(f"  完备度（刊实测命中率）：" + "；".join(
        f"{n}{'✓' if n in struct['found'] else '✗'}" for n, _ in STRUCT_PATTERNS))
    if struct["missing"]:
        lines.append(f"    · 未检出：{('、'.join(struct['missing']))}（命中率 37%–67% 的标配章，"
                     f"缺则需确认是并入他章还是遗漏）")
    if struct["h3_count"]:
        tag = (f"· 三级标题 {struct['h3_count']} 个——该刊常态（85/108 篇含三级，中位 6.5 个/篇）"
               if struct["h3_count"] <= H3_ALERT_THRESHOLD
               else f"· 三级标题 {struct['h3_count']} 个——超过实测 P90（约 {H3_ALERT_THRESHOLD} 个），考虑是否过碎")
        lines.append("    " + tag)
    if abs_text:
        lines.append(f"  摘要开头80字：{clean(abs_text)[:80]}")
        lines.append(f"  摘要忌词：本文×{abs_text.count('本文')} 分析了×{abs_text.count('分析了')} 提出了×{abs_text.count('提出了')}")

    # 【官方合规】——官方须知明文要求，优先级高于所有实证基线
    lines.append("\n【⓪ 官方须知合规】（references/official-guidelines.md 明文要求）")
    if official["columns"]:
        lines.append(f"  栏目对位：{'、'.join(official['columns'][:3])}"
                     f"（官方九栏之一，投稿时须注明）")
    else:
        lines.append("  栏目对位：未判定")
    if official["issues"]:
        for it in official["issues"]:
            lines.append(f"  ⚠️ {it}")
    else:
        lines.append("  ✓ 官方明文项：未发现问题")
    if official.get("package_notices"):
        lines.append("  [投稿包要素提醒（实名投稿必备，盲审阶段可略）]：")
        for pn in official["package_notices"]:
            lines.append(f"    · ℹ️ {pn}")
    lines.append("  （官方要点：摘要约 200 字、忌 本文/我们/分析了/总结了/提出了/解决了、"
                 "投稿邮箱 tsgxyj@vip.163.com、审稿时限 7 周、2022 年起月刊）")
    if abst.get("present"):
        elem_str = "".join("✓" if abst["elements"][k] else "✗" for k in ("目的", "方法", "结果", "价值"))
        lines.append(f"  四要素体检（目的/方法/结果/价值）：{elem_str}"
                     + (f"（{abst['n_sents']} 句）" if abst.get("n_sents") else ""))
        for it in abst.get("issues", []):
            lines.append(f"    · {it}")

    lines.append("\n【② 量化对照】")
    lines.append(f"  段长中位 {m['para_med']:.0f}字" + fmt_pos(m["para_med"], base["paras"], higher_better=False))
    lines.append(f"  句均 {m['sent_mean']:.0f}字" + fmt_pos(m["sent_mean"], base["sents"], higher_better=False))
    lines.append(f"  引用密度 {m['cite_rate']:.2f}/千字" + fmt_pos(m["cite_rate"], base["cites"]))
    lines.append(f"  摘要 {m['abs_len']}字" + fmt_pos(m["abs_len"], base["abs"], higher_better=False))
    lines.append(f"  关键词 {m['kw_n']}个" + fmt_pos(m["kw_n"], base["kws"]))
    lines.append(f"  文献 {m['ref_n']}条" + fmt_pos(m["ref_n"], base["refs"]))
    lines.append(f"  图 {m['fig_n']}张" + fmt_pos(m["fig_n"], base["figs"]))
    lines.append(f"  表 {m['tab_n']}张" + fmt_pos(m["tab_n"], base["tabs"]))
    lines.append(f"  「本文」{m['benwen']}次 = {m['benwen_rate']:.2f}/万字" +
                 fmt_pos(m["benwen_rate"], base["benwen"]))
    # 红线合计（供 check_all.py 抓取；官方合规 / 反模式超标 / 结构缺标配 / 摘要红线）
    n_anti = sum(1 for a in ant if a["flag"])
    n_struct = sum(1 for x in ("对策/策略/路径/启示（收口标配）", "结语") if x in struct["missing"])
    n_abs = sum(1 for i in abst.get("issues", []) if "忌词" in i or "虚句" in i or "偏长" in i)
    n_off = len(official.get("issues", []))
    n_red = n_anti + n_struct + n_abs + n_off
    lines.append(f"  [红线合计] {n_red} 项（官方合规 {n_off} / 反模式超标 {n_anti} / "
                 f"结构缺标配 {n_struct} / 摘要红线 {n_abs}）"
                 + ("——0 红线，形式层可投" if n_red == 0 else "——见下"))

    if secs and not "--brief" in flags:
        lines.append("\n【②b 章节占比】")
        for cname, base_pct in SEC_BASE.items():
            if cname in secs["merged"]:
                v = secs["merged"][cname]
                if base_pct:
                    tag = "偏高" if v > base_pct * 1.6 else ("偏低" if v < base_pct * 0.5 else "≈")
                    lines.append(f"  {cname}: {v:.1f}%（刊单列章中位 {base_pct}%）{tag}")
                else:
                    lines.append(f"  {cname}: {v:.1f}%")
        shown = {c for c in secs["merged"]}
        missing = [c for c in SEC_BASE if c not in shown and c != "结果/分析"]
        if missing:
            lines.append(f"  （未单列：{'、'.join(missing)}）")

    red = [a for a in ant if a["flag"]]
    if not "--brief" in flags:
        lines.append("\n【③ 反模式扫描】(对比该刊基线，>2x 标 ⚠️)")
        for a in ant:
            mark = "⚠️ 明显高于刊基线" if a["flag"] else "正常"
            lines.append(f"  {a['name']}: {a['n']} 处 = {a['rate']:.2f}/万字（刊基线 {a['base']}）{mark}")
            if a["sample"]:
                lines.append(f"      示例：{a['sample']}")
        if not red:
            lines.append("  ✓ 无超标反模式")

        # 连接词频率（次/千字，对照 style_stats.md §3）
        conn_base = base.get("conn") or {}
        if conn_base:
            alltext2 = clean("".join(body))
            n2 = max(len(alltext2), 1)
            conn_hits = {}
            for w in ("同时", "因此", "一方面", "从而", "此外", "另一方面", "进而", "然而"):
                c = alltext2.count(w)
                if c:
                    conn_hits[w] = c / n2 * 1000
            if conn_hits:
                lines.append("\n【③b 连接词频率】(次/千字；该刊 108 篇实测基线)")
                for w, v in sorted(conn_hits.items(), key=lambda x: -x[1]):
                    b = conn_base.get(w, 0)
                    tag = ""
                    if b:
                        tag = (" ▲ 偏高频" if v > b * 1.5 else
                               " ▼ 偏低频" if v < b * 0.4 else " ≈")
                    lines.append(f"  {w}: {v:.1f}（刊基线 {b:.1f}）{tag}")

    if "--paragraphs" in flags or not "--brief" in flags:
        lines.append(f"\n【④ 段落级扫描】共 {len(para_issues)} 个目标段（句均<45／序数口语词／超长>450／碎段<110）")
        for it in para_issues[:30]:
            lines.append(f"  段{it['para_no']} [{it['chars']}字] {it['tags']} —— {it['head']}…")
        if len(para_issues) > 30:
            lines.append(f"  （仅显示前 30 段，共 {len(para_issues)} 段）")
        if not para_issues:
            lines.append("  ✓ 段落健康")

    # 优先输出红线问题
    brief_lines = []
    if "--brief" in flags:
        brief_lines.append("【优先处理（红线）】")
        red_items = []
        # 官方明文项优先级最高
        for it in official.get("issues", []):
            red_items.append(f"  ⚠️ [官方] {it}")
        for a in red:
            red_items.append(f"  ⚠️ {a['name']}：{a['n']} 处（基线 {a['base']}）{a['sample']}")
        for must in ("对策/策略/路径/启示（收口标配）", "结语"):
            if must in struct["missing"]:
                red_items.append(f"  ⚠️ 结构缺「{must}」——该刊收口标配（命中率 52%–67%）")
        if abst.get("issues"):
            for it in abst["issues"]:
                if "忌词" in it or "虚句" in it or "偏长" in it:
                    red_items.append(f"  ⚠️ 摘要：{it}")
        if m["benwen_rate"] > base["benwen"].get("q75", 100):
            red_items.append(f"  ⚠️ 「本文」{m['benwen_rate']:.1f}次/万字 高于 75% 刊文")
        brief_lines.extend(red_items if red_items else ["  无红线问题"])
        if m["benwen_rate"] > base["benwen"].get("q75", 100):
            brief_lines.append(f"  ⚠️ 「本文」{m['benwen_rate']:.1f}次/万字 高于 75% 刊文")
        if m["abs_len"] and "本文" in clean(abs_text or ""):
            brief_lines.append("  ⚠️ 摘要含「本文」——该刊摘要忌用")
        brief_lines.append("\n【段落级重点】")
        for it in para_issues[:10]:
            brief_lines.append(f"  段{it['para_no']} [{it['chars']}字] {it['tags']} —— {it['head']}…")
        if not para_issues:
            brief_lines.append("  无目标段")

    print("\n".join(lines))
    if brief_lines:
        print("\n" + "\n".join(brief_lines))
    print("\n" + "-" * 62)
    print("说明：量化基线取自语料库统计（或内置静态值）；黑体判断为汇总归纳，")
    print("非期刊编辑部口径；改稿前请按 skill 的 style-rules / paragraphing-guide 执行。")

    if "--json" in flags:
        out = {
            "doc": os.path.basename(path), "metrics": m, "abstract": abst,
            "structure": struct, "sections": secs, "official": official,
            "anti": ant, "paragraph_issues": para_issues,
            "baseline": {"source": base["source"], "static": base["static"],
                         "med": {k: v.get("med") for k, v in base.items() if isinstance(v, dict)}},
        }
        jf = os.path.splitext(path)[0] + ".diagnose.json"
        with open(jf, "w", encoding="utf-8") as f:
            json.dump(out, f, ensure_ascii=False, indent=1)
        print(f"\nJSON 报告已写入：{jf}")
    return 0


if __name__ == "__main__":
    sys.exit(main())
