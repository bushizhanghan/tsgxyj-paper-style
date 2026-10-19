# -*- coding: utf-8 -*-
"""submission_email.py —— 投稿邮件要素生成器（官方邮箱投稿）

用途：
    读 references/response-template.md 的「投稿邮件正文（模板骨架）」，从
    用户稿件自动提取题名、字数、栏目对位、摘要核心发现，生成一封填好的
    投稿邮件草稿：含官方投稿邮箱、随稿附件清单与查重要求提示。

依据：
    · 邮件骨架：references/response-template.md「二、投稿材料（该刊邮箱
      投稿）→ 投稿邮件正文（模板骨架）」——填充槽位为「论文《××××××》
      （约 ×× 字）」「围绕____问题，采用____方法开展研究，主要发现为
      ____，并提出____」「与本刊「信息服务」（或实际栏目）方向相关」
      「随邮件附：论文稿件（Word）、查重报告、作者简介」；
    · 官方邮箱与投稿方式：references/official-guidelines.md §3「投稿方式：
      电子邮箱 tsgxyj@vip.163.com（Email 投稿）；回复方式：邮箱回复」；
    · 栏目九类：official-guidelines.md §2（对位用 dp.scan_official 的
      columns，无障碍/弱势群体类选题 → 读者工作 或 信息服务）；
    · 字数口径：该刊 checklists.md「全文字数 10000 以上」，实测投稿稿
      集中 10000–15000 字；
    · 查重要求：checklists.md「官方口径 <10%，投稿者实操反馈要求总文字
      复制比 <7%、去除引用文献复制比接近 0——按 7% 准备更稳」「AIGC 检测
      <10%」；官方 §3「确保原创，文责自负」；
    · 审稿时限：official-guidelines.md §3「审稿时限 7 周」。

用法:
    python submission_email.py <论文.docx> [--out <文件>] [--author <姓名>]
                               [--contact <联系方式>] [--no-attachments]

    默认打印到 stdout；--out 写文件（UTF-8）。
    --author/--contact 填写文末「作者：×××（单位、联系方式）」槽位；
    未提供时该槽位以占位符保留，提示发送前补齐。

退出码:
    0 = 邮件草稿生成成功
    2 = 文件不存在 / 稿件不可解析 / 未找到摘要（无法生成发现句）
"""
import os
import re
import sys

# Windows 控制台 GBK 保护
try:
    sys.stdout.reconfigure(encoding="utf-8", errors="replace")
except Exception:
    pass

HERE = os.path.dirname(os.path.abspath(__file__))
SKILL_ROOT = os.path.dirname(HERE)
sys.path.insert(0, HERE)
import diagnose_paper as dp  # noqa: E402  （复用 extract_paras / scan_official）

# ---- 官方口径常量（official-guidelines.md §2/§3；checklists.md）----
OFFICIAL_EMAIL = "tsgxyj@vip.163.com"      # 官方 §3：唯一投稿渠道
REVIEW_WEEKS = 7                            # 官方 §3：审稿时限
WORD_MIN = 10000                            # checklists.md：全文字数 10000 以上
SIM_BENCH = 7.0                             # checklists.md：总文字复制比按 7% 准备
AIGC_BENCH = 10.0                           # checklists.md：AIGC 检测 <10%

# 官方九类栏目（official-guidelines.md §2）
OFFICIAL_COLUMNS = [
    "基础理论（研究进展）", "专业教育", "图书馆管理", "数字化、网络化",
    "信息资源建设", "信息组织", "信息服务", "读者工作", "国外图书馆",
]

# 摘要核心发现抽取信号（官方 §5 摘要四要素之「研究结果与结论」）
FIND_SIG = re.compile(r"(发现|表明|显示|结果|达到|通过|构建|提出|设计|实现|"
                      r"验证|证明|提升|改善|覆盖|满足)")
METHOD_SIG = re.compile(r"(采用|运用|使用|基于|依托|借助|设计并实现|构建|"
                        r"通过.{0,12}方法|方法|问卷|调研|测试|检测|实验)")


def title_of(paras):
    """题名：正文首两段（主标题 + 副标题）。"""
    texts = [t for _, t in paras if t]
    if not texts:
        return ""
    head = [t for t in texts[:3] if not re.match(r"^【?摘\s*要", t)]
    if not head:
        return ""
    main = head[0].strip()
    sub = head[1].strip() if len(head) > 1 and head[1].startswith(
        ("——", "--", "-")) else ""
    return main + sub if sub else main


def word_count(body):
    """全文字数（去空白计，正文区）。"""
    return sum(len(re.sub(r"\s+", "", t)) for t in body)


def split_sentences(text):
    return [s.strip() for s in re.split(r"[。；？！]", text or "") if len(s.strip()) >= 8]


def core_finding(abs_text):
    """抽取摘要核心发现：优先取含结果信号的最长一句，否则取末句。"""
    sents = split_sentences(abs_text)
    if not sents:
        return ""
    scored = []
    for s in sents:
        score = len(FIND_SIG.findall(s)) * 2 + len(s) / 50.0
        if re.search(r"\d", s):
            score += 1.0
        scored.append((score, s))
    scored.sort(reverse=True)
    best = scored[0][1]
    # 取该句前后各补一点上下文，避免单句过短
    idx = sents.index(best)
    seg = best
    if len(best) < 30 and idx + 1 < len(sents):
        seg = best + "；" + sents[idx + 1]
    return seg[:160]


def method_phrase(body, abs_text):
    """方法短语：从摘要中找「采用/基于…」片段。"""
    src = abs_text or ""
    m = re.search(r"(?:采用|运用|使用|依托|基于)([^，。；]{2,26})", src)
    if m:
        return m.group(1).strip()
    for p in body:
        m = re.search(r"(?:采用|运用|使用|基于)([^，。；]{2,26})", p)
        if m:
            return m.group(1).strip()
    return "混合研究方法"


def column_line(cols):
    """栏目对位行：按 dp.scan_official 的 columns 排序取首选 + 备选。"""
    if not cols:
        return ("投稿栏目：未能自动判定，请对照官方九类栏目人工确认"
                "（基础理论（研究进展）/专业教育/图书馆管理/数字化、网络化/"
                "信息资源建设/信息组织/信息服务/读者工作/国外图书馆）")
    first = cols[0]
    rest = [c for c in cols[1:4]]
    line = f"投稿栏目：{first}"
    if rest:
        line += f"（备选：{'、'.join(rest)}）"
    line += "。"
    # 官方 §2 明文：无障碍/弱势群体类选题 → 读者工作 或 信息服务
    if first not in ("读者工作", "信息服务") and \
            any(c in ("读者工作", "信息服务") for c in cols):
        line += ("（按官方 §2，无障碍/弱势群体类选题宜对位"
                 "「读者工作」或「信息服务」）")
    return line


def build_email(info):
    """按 response-template.md 骨架填槽，生成邮件正文。"""
    wc = info["word_count"]
    lines = []
    lines.append(f"邮件主题：投稿-{info['title']}（{info['author_tag']}）")
    lines.append(f"收件人：{OFFICIAL_EMAIL}（官方唯一投稿渠道）")
    lines.append("")
    lines.append("尊敬的编辑老师：")
    lines.append("")
    lines.append("您好！现向贵刊投递论文《%s》（约 %d 字），恳请审阅。"
                 % (info["title"], wc))
    lines.append("")
    lines.append("论文围绕%s问题，采用%s开展研究，主要发现为%s，"
                 "并提出相应优化策略。%s论文选题与本刊「%s」方向相关。"
                 % (info["topic"], info["method"], info["finding"],
                    info["column_line"],
                    info["columns"][0] if info["columns"] else "（待定）"))
    lines.append("")
    lines.append("随邮件附：")
    for i, a in enumerate(info["attachments"], 1):
        lines.append(f"  {i}. {a}")
    lines.append("")
    lines.append("查重提示：%s" % info["dup_note"])
    lines.append("全体作者均阅读并同意投稿，无一稿多投。")
    lines.append("")
    lines.append("盼复为感！")
    lines.append("")
    lines.append("%s" % info["author_tag"])
    lines.append("%s" % info["contact_tag"])
    lines.append("%s" % info["date_tag"])
    lines.append("")
    lines.append("——")
    lines.append("【发送前核对】")
    for i, c in enumerate(info["precheck"], 1):
        lines.append(f"  {i}. {c}")
    return "\n".join(lines)


def analyze(doc, author="", contact=""):
    paras, refs, abs_text, kw_text, titles, body = dp.extract_paras(doc)
    alltext = "\n".join(t for _, t in paras)
    off = dp.scan_official(kw_text, refs, titles, alltext, paras=paras)
    cols = off.get("columns", [])

    title = title_of(paras)
    wc = word_count(body)
    finding = core_finding(abs_text)
    if not title or not abs_text:
        raise ValueError("未解析到题名或摘要，无法生成邮件要素")

    # 主题词：取题名主干（去「研究」类泛词）
    topic = re.sub(r"(研究|分析|探讨|初探|浅析|综述|思考)$", "",
                   re.sub(r"^面向|^基于", "", title.split("——")[0]))
    topic = topic[:34] if topic else title[:34]

    author_tag = (f"作者：{author}" if author else "作者：×××（姓名待填）")
    contact_tag = (f"联系方式：{contact}" if contact
                   else "联系方式：×××（单位、电话、邮箱、通信地址待填）")
    date_tag = f"定稿日期：{info_date()}"
    return {
        "title": title, "word_count": wc, "finding": finding,
        "method": method_phrase(body, abs_text), "topic": topic,
        "columns": cols, "column_line": column_line(cols),
        "attachments": attachments(wc),
        "dup_note": dup_note(wc),
        "precheck": precheck(wc, cols),
        "author_tag": author_tag, "contact_tag": contact_tag,
        "date_tag": date_tag,
    }


def info_date():
    import datetime
    return datetime.date.today().strftime("%Y 年 %m 月 %d 日")


def attachments(wc):
    """随稿附件清单（response-template.md 骨架 + checklists.md 材料单）。"""
    atts = [f"论文稿件（Word 版，正文约 {wc} 字）",
            "查重报告（总文字复制比建议 <7%，去除引用后接近 0）",
            "AIGC 检测报告（<10%）",
            "作者简介（姓名、单位、职称、学位、研究方向；"
            "如有通讯作者请标注）",
            "联系方式页（邮编、详细通信地址、电话、电子邮箱、定稿日期）"]
    return atts


def dup_note(wc):
    """查重要求提示（官方 §3 原创要求 + checklists.md 实操口径）。"""
    return (f"按该刊要求，全文字数 {WORD_MIN} 以上（本文约 {wc} 字）；"
            f"总文字复制比建议压至 {SIM_BENCH:.0f}% 以下、去除引用文献复制比"
            f"接近 0，AIGC 检测低于 {AIGC_BENCH:.0f}%；官方 §3 要求确保原创、"
            f"文责自负，请随稿附查重报告。")


def precheck(wc, cols):
    """发送前核对项。"""
    items = [
        f"附件四件齐全（稿件 / 查重报告 / AIGC 报告 / 作者简介），"
        f"共 5 项请按本文清单核对",
        f"查重：总文字复制比 <{SIM_BENCH:.0f}%（去除引用接近 0），"
        f"AIGC <{AIGC_BENCH:.0f}%",
        f"栏目对位已确认（当前判定：{cols[0] if cols else '待定'}）",
        f"正文 {wc} 字，符合该刊 {WORD_MIN} 字以上要求",
        f"文末已附作者简介与联系方式（官方 §8）——本文检出"
        f"缺失项时须先补齐",
        f"署名与顺序已在全体作者间确认（录用后不可改）",
        f"审稿时限 {REVIEW_WEEKS} 周（官方 §3），回复方式为邮箱回复",
    ]
    return items


def main():
    args = sys.argv[1:]
    if not args or args[0] in ("-h", "--help"):
        print(__doc__)
        return 0
    DOC = args[0]
    if not os.path.isfile(DOC):
        print(f"找不到文件：{DOC}")
        return 2

    def opt(name, default=""):
        if name in args:
            i = args.index(name)
            if i + 1 < len(args):
                return args[i + 1]
        return default

    out_path = opt("--out")
    author = opt("--author")
    contact = opt("--contact")

    try:
        info = analyze(DOC, author=author, contact=contact)
    except Exception as e:
        print(f"稿件解析失败：{e}")
        return 2

    text = build_email(info)
    if out_path:
        with open(out_path, "w", encoding="utf-8") as fh:
            fh.write(text + "\n")
        print(f"投稿邮件草稿已写入：{out_path}")
        print("-" * 60)
    print(text)
    return 0


if __name__ == "__main__":
    sys.exit(main())
