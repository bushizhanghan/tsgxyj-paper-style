# -*- coding: utf-8 -*-
"""language_audit.py -- 整稿语言体检：把润色词典落实到「每个词的真实位置」

一次扫完全文，按类别列出**可直接改的位置**（段落号 + 原句 + 建议），
避免"通篇说口语词多，但不知道在哪"。

检查项（全部依据 references/polish-lexicon.md 与 anti-patterns.md 的原文实证）：

    A 该刊 0 次口语词     差不多/大概/搞得/咱们/老百姓/靠谱/给力/硬核…（必改）
    B 该刊罕用（<20 次）  很多/其实/当然/看来/东西/怎么/大家…（视语境）
    C 无支撑强断言        显著/全面/充分/彻底/完全/绝对/必然（该改审慎词）
    D 「了」属性陈述堆积   该刊 63% 段不用「了」；叙述过程可留
    E 「被」字被动        该刊 0.53/千字
    F 解释不用「即」      破折号/括号解释可改「即」式
    G 反模式句子          三连×化 / 不仅…而且 / 万能句 / 红线句（带原文位置）

用法:
    python language_audit.py <论文.docx>                 # 全部检查
    python language_audit.py <论文.docx> --only A,B       # 只查指定类别
    python language_audit.py <论文.docx> --limit 10       # 每类最多显示 N 条
    python language_audit.py <论文.docx> --json           # 结构化输出

退出码：0 = 无 A 类问题；1 = 有 A 类必改项
"""
import os
import re
import sys
import json

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import corpus  # noqa: E402
import diagnose_paper as dp  # noqa: E402

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

LIMIT = 8
if "--limit" in sys.argv:
    i = sys.argv.index("--limit")
    try:
        LIMIT = int(sys.argv[i + 1])
    except (ValueError, IndexError):
        pass
ONLY = None
if "--only" in sys.argv:
    i = sys.argv.index("--only")
    if i + 1 < len(sys.argv):
        ONLY = {x.strip().upper() for x in sys.argv[i + 1].split(",") if x.strip()}

# ---- 检查词表（分级依据：polish-lexicon.md 的原文次数）----
A_WORDS = ["差不多", "大概", "越来越好", "一下子", "有时候", "搞得", "事儿", "咱们",
           "老百姓", "活儿", "靠谱", "给力", "走心", "硬核"]          # 该刊 0 次
B_WORDS = ["很多", "其实", "当然", "看来", "东西", "怎么", "什么样", "大家", "弄"]  # <20 次
C_WORDS = ["显著", "全面", "充分", "彻底", "完全", "绝对", "必然"]                 # 强断言
C_SAFE = re.compile(r"显著(正|负)?(相关|影响|差异|提升|作用|效应)|显著性")

paras, refs, abs_text, kw_text, titles, body = dp.extract_paras(DOC)
body_clean = [dp.clean(p) for p in body]
paras_c = [(i + 1, c) for i, c in enumerate(body_clean) if len(c) >= 40]

findings = {}   # cat -> [(para_no, 原句, 建议)]


def sentences_with_pos(para_no, text):
    """返回 (句, 句内偏移) 列表。"""
    out = []
    pos = 0
    for m in re.finditer(r"[^。？！]*[。？！]?", text):
        s = m.group(0)
        if s and len(s.strip()) >= 4:
            out.append((s.strip(), pos))
        pos = m.end()
    return out


def scan_level():
    for cat, words, advice in [
        ("A", A_WORDS, "该刊 121 万字原文中 0 次出现，必改（见 polish-lexicon.md §1）"),
        ("B", B_WORDS, "该刊 <20 次，视语境替换（见 polish-lexicon.md §1）"),
    ]:
        findings[cat] = []
        for pno, text in paras_c:
            for sent, _ in sentences_with_pos(pno, text):
                hit = [w for w in words if w in sent]
                if hit:
                    findings[cat].append((pno, sent, f"含「{'、'.join(hit)}」→ {advice}"))


def scan_strong():
    findings["C"] = []
    for pno, text in paras_c:
        for sent, _ in sentences_with_pos(pno, text):
            # 排除统计口径（显著相关/显著影响/显著差异…）
            s_clean = C_SAFE.sub("", sent)
            hit = []
            for w in C_WORDS:
                if w not in s_clean:
                    continue
                # 「完全」在不表程度断言的搭配中放过（完全按/完全一致/完全基于）
                if w == "完全" and re.search(r"完全(按|依照|依据|基于|一致|相同|照)", s_clean):
                    continue
                # 「必然」在「必然结果/必然趋势」中属名词性，放过
                if w == "必然" and re.search(r"必然(结果|趋势|产物|要求)", s_clean):
                    continue
                hit.append(w)
            if hit:
                findings["C"].append(
                    (pno, sent, f"强断言「{'、'.join(hit)}」无数据支撑→ 改「有所/一定程度上/相对」"))


def scan_le():
    findings["D"] = []
    for pno, text in paras_c:
        n = text.count("了")
        if n >= 2:
            # 判断是否属性陈述（…了…。 且后接名词短语）
            attr = len(re.findall(r"了[\u4e00-\u9fff]{1,8}[。，；]", text))
            if attr >= 2:
                sent = next((s for s, _ in sentences_with_pos(pno, text) if "了" in s), text[:50])
                findings["D"].append(
                    (pno, sent,
                     f"段内「了」×{n}（属性陈述 {attr} 处）→ 该刊 63% 段不用「了」，改「能够/得到」或删"))


def scan_bei():
    """只报真正的被动句：被 + 动词，排除「被动模式/被动接受」等名词搭配。"""
    findings["E"] = []
    for pno, text in paras_c:
        for sent, _ in sentences_with_pos(pno, text):
            m = re.search(r"被(?!动)[\u4e00-\u9fff]{0,6}(广泛|普遍|认为|视为|称为|用于|应用|采用|接受|解读|感知|区别|忽视|纳入|排除|证明|证实|发现|显示|改写|改造|替代|忽略)", sent)
            if m:
                findings["E"].append(
                    (pno, sent, "被动式→ 改主动/无主/「得以××」（该刊 0.53/千字）"))


def scan_explain():
    """只报「术语后跟长解释」且未用「即」的情形；排除引号内列举、范围号、待补标记。"""
    findings["F"] = []
    for pno, text in paras_c:
        for sent, _ in sentences_with_pos(pno, text):
            # 去掉引号内内容（引号里多为列举/术语，不属解释）
            s = re.sub(r"[“”「」][^“”「」]*[“”「」]", "", sent)
            # 括号解释：括号内 ≥12 字且像完整句子（含逗号或句号）
            m_par = re.search(r"（([^）]{12,80})）", s)
            # 破折号解释：破折号后 ≥10 字，且不是「A—B—C」式三元列举
            m_dash = re.search(r"——([^—]{10,80})", s)
            m_trip = re.search(r"[^—]{2,}—[^—]{2,}—[^—]{2,}", s)
            if not (m_par or m_dash):
                continue
            if m_trip and not m_par:
                continue          # 三元破折号列举，跳过
            target = (m_par or m_dash).group(1)
            if not re.search(r"[，。；]", target):
                continue          # 不是完整解释句（多为术语注释/范围）
            if re.search(r"待补|截图|图\d|表\d", sent):
                continue          # 编辑标记
            if re.search(r"即[^，。]{2,}", s):
                continue          # 已用「即」
            findings["F"].append(
                (pno, sent, "长括号/破折号解释→ 该刊 418 处「即」句更常见，可改「××，即××」"))


def scan_anti():
    findings["G"] = []
    pats = [
        (r"[\u4e00-\u9fff]{1,4}化[、,，][\u4e00-\u9fff]{1,4}化[、,，][\u4e00-\u9fff]{1,4}化", "三连×化排比"),
        (r"不仅[^。]{2,50}而且", "「不仅…而且」（该刊主流是更/还）"),
        (r"具有重要意义|提供有益参考|值得深入探讨", "万能句"),
        (r"显而易见|众所周知|毋庸置疑|总而言之|不难看出|值得一提", "A 级红线句"),
        (r"(首先|其次|再次)[，,]且.{0,40}(首先|其次|再次)[，,]", "同段机械序数堆叠"),
    ]
    for pno, text in paras_c:
        for sent, _ in sentences_with_pos(pno, text):
            for pat, name in pats:
                if re.search(pat, sent):
                    findings["G"].append((pno, sent, f"反模式：{name}"))


# A/B 由 scan_level 一次产出
if ONLY is None or "A" in ONLY or "B" in ONLY:
    scan_level()
for cat, fn in [("C", scan_strong), ("D", scan_le), ("E", scan_bei),
                ("F", scan_explain), ("G", scan_anti)]:
    if ONLY is None or cat in ONLY:
        fn()

TITLES = {"A": "该刊 0 次口语词（必改）", "B": "该刊罕用词（视语境）",
          "C": "无支撑强断言", "D": "「了」属性陈述堆积",
          "E": "「被」字被动", "F": "解释未用「即」式", "G": "反模式句子"}

print("=" * 68)
print(f"整稿语言体检：{os.path.basename(DOC)}")
print("=" * 68)

total_a = 0
shown = 0
for cat in ["A", "B", "C", "D", "E", "F", "G"]:
    if cat not in findings:
        continue
    items = findings[cat]
    if not items:
        print(f"\n【{cat}】{TITLES[cat]}：✓ 无")
        continue
    print(f"\n【{cat}】{TITLES[cat]}：{len(items)} 处")
    for pno, sent, advice in items[:LIMIT]:
        print(f"  段{pno}｜{sent[:56]}…")
        print(f"       → {advice}")
    if len(items) > LIMIT:
        print(f"   …另有 {len(items) - LIMIT} 处（用 --limit 调整）")
    if cat == "A":
        total_a += len(items)

print("\n" + "─" * 68)
if total_a:
    print(f"⚠️ 有 {total_a} 处该刊 0 次口语词，投稿前必须处理。")
else:
    print("✓ 无该刊 0 次口语词。")
print("依据：references/polish-lexicon.md（118 篇原文实证）｜细则：anti-patterns.md")

if "--json" in _flags:
    out = {"doc": os.path.basename(DOC),
           "findings": {k: [{"para": p, "sent": s, "advice": a} for p, s, a in v]
                        for k, v in findings.items()},
           "A_count": total_a}
    jf = os.path.splitext(DOC)[0] + ".langaudit.json"
    with open(jf, "w", encoding="utf-8") as f:
        json.dump(out, f, ensure_ascii=False, indent=1)
    print(f"\nJSON 已写入：{jf}")

sys.exit(1 if total_a else 0)
