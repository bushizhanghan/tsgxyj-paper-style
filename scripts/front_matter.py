# -*- coding: utf-8 -*-
"""front_matter.py —— 稿件前置元数据序列核验（官方 §4 格式序列）

官方投稿须知 §4「格式顺序（官方明文）」给出的前置序列为：

    文章标题（中文）→ 作者姓名（中文）→ 文章标题（英文）→
    作者姓名（英文）→ [摘要]（中文）→ [关键词]（中文）→
    Abstract:（英文）→ Key words:（英文）→ 正文

已有工具链只查英文摘要/关键词「是否存在」（diagnose_paper 的官方合规
维度），不查**先后次序**。次序错误（如英文题名排在中文摘要之后、英文
摘要在中文摘要之前、作者姓名跑到题名后面很远）在初审里是显性的
体例问题，本脚本专门按段落顺序核验：

  · 七类前置元素（中文题名/作者名/英文题名/英文作者名/中文摘要/
    中文关键词/英文摘要/英文关键词）各自首次出现的段落号；
  · 依官方序列比对实际次序，报出逆序对（如「英文题名在中文摘要后」）；
  · 英文题名缺失（官方序列明文要求，投稿稿常漏）；
  · 英文作者姓名缺失；
  · 作者姓名行是否紧随中文题名之后（相隔过远＝疑似署名位置不当）；
  · 中文题名本身是否缺失或被「摘要」抢先。

用法:
    python front_matter.py <论文.docx> [--json]
退出码：0 = 前置序列无次序问题；1 = 检出次序错误或缺英文题名/中文题名。

说明：只核**顺序与存在性**，不核元素内部格式（摘要字数见 meta_scale、
英文题名大小写见 en_title_check.py、作者简介形态见 author_meta.py）。
"""
import os
import re
import sys
import json

# Windows 控制台 GBK 保护
if hasattr(sys.stdout, "reconfigure"):
    try:
        sys.stdout.reconfigure(encoding="utf-8", errors="replace")
    except Exception:
        pass

HERE = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, HERE)
import diagnose_paper as dp  # noqa: E402

# 官方 §4 前置序列的键序（顺序即要求顺序）
SEQ = ["cn_title", "cn_author", "en_title", "en_author",
       "cn_abs", "cn_kw", "en_abs", "en_kw"]
SEQ_LABEL = {
    "cn_title": "中文题名", "cn_author": "作者姓名（中文）",
    "en_title": "文章标题（英文）", "en_author": "作者姓名（英文）",
    "cn_abs": "摘要（中文）", "cn_kw": "关键词（中文）",
    "en_abs": "Abstract:（英文）", "en_kw": "Key words:（英文）",
}

CJK_RE = re.compile(r"[\u4e00-\u9fff]")


def _is_cjk_line(s):
    """该段是否以中文为主（中文题名/摘要/关键词/作者名）。"""
    return len(CJK_RE.findall(s or "")) >= 2


def _looks_en(s):
    """该段是否以拉丁字母为主（英文题名/英文摘要/英文关键词）。"""
    lat = len(re.findall(r"[A-Za-z]", s or ""))
    cjk = len(CJK_RE.findall(s or ""))
    return lat >= 3 and lat > cjk


def _name_line(s):
    """作者姓名行：短行（去空白后 ≤12 字）、以中文为主、无句读。"""
    c = dp.clean(s or "")
    if not (2 <= len(c) <= 12):
        return False
    if not re.match(r"^[\u4e00-\u9fff]+$", c):
        return False
    if re.search(r"[。；，,;:]", s or ""):
        return False
    return True


def find_front_matter(paras):
    """按段落顺序定位前置元素。

    返回 (found, detail)：found 为 {键: 段落号}，detail 为 {键: 原文摘要}。
    """
    found, detail = {}, {}
    seen_cn_title = False
    seen_abs = False
    for idx, (style, raw) in enumerate(paras):
        t = (raw or "").strip()
        if not t:
            continue
        head = t[:14]
        # 中文题名：摘要之前的第一个非空中文长段（≥8 字）
        if "cn_title" not in found and not seen_abs and _is_cjk_line(t):
            c = dp.clean(t)
            if len(c) >= 8 and not re.search(r"^【?\s*摘\s*要", head) \
                    and not re.search(r"^【?\s*关\s*键\s*词", head):
                found["cn_title"] = idx
                detail["cn_title"] = t[:60]
                seen_cn_title = True
                continue
        # 中文摘要 / 关键词
        if re.search(r"^【?\s*摘\s*要】?[：:]?", head):
            if "cn_abs" not in found:
                found["cn_abs"] = idx
                detail["cn_abs"] = t[:60]
                seen_abs = True
            continue
        if re.search(r"^【?\s*关\s*键\s*词】?[：:]?", head):
            if "cn_kw" not in found:
                found["cn_kw"] = idx
                detail["cn_kw"] = t[:60]
            continue
        # 英文摘要 / 关键词
        if re.match(r"^(?:Abstract|ABSTRACT)\s*[:：]", t):
            if "en_abs" not in found:
                found["en_abs"] = idx
                detail["en_abs"] = t[:60]
            continue
        if re.match(r"^(?:Key\s*words?|KEY\s*WORDS?)\s*[:：]", t, re.I):
            if "en_kw" not in found:
                found["en_kw"] = idx
                detail["en_kw"] = t[:60]
            continue
        # 作者姓名（中文）：题名后、摘要前的姓名行
        if seen_cn_title and "cn_abs" not in found and _name_line(t):
            found["cn_author"] = idx
            detail["cn_author"] = t[:30]
            continue
        # 英文题名：摘要前的拉丁长行
        if seen_cn_title and "cn_abs" not in found and "en_abs" not in found \
                and _looks_en(t) and len(t) >= 15:
            if "en_title" not in found:
                found["en_title"] = idx
                detail["en_title"] = t[:70]
                continue
            # 英文题名之后若还有短拉丁行 → 英文作者姓名
            if "en_author" not in found and len(dp.clean(t)) <= 40:
                found["en_author"] = idx
                detail["en_author"] = t[:40]
    return found, detail


def order_issues(found):
    """按官方序列比对实际次序，返回 (issues, notices)。"""
    issues, notices = [], []
    # 逐对比对相邻键的实际段落号是否递增
    for a, b in zip(SEQ, SEQ[1:]):
        if a in found and b in found and found[b] < found[a]:
            issues.append(
                f"次序错误：{SEQ_LABEL[b]}（段{found[b]}）排在"
                f"{SEQ_LABEL[a]}（段{found[a]}）之前——官方序列要求"
                f"「{SEQ_LABEL[a]}」在前")
    # 英文题名缺失（官方序列明文要求）
    if "en_title" not in found and "cn_title" in found:
        issues.append("缺少英文题名——官方格式序列要求中文题名后紧跟英文题名")
    # 英文作者姓名缺失
    if "en_author" not in found and "cn_title" in found:
        notices.append("未见英文作者姓名——官方序列为「文章标题（英文）→"
                       "作者姓名（英文）」；若作者名只有中文，请确认刊物是否接受")
    # 作者姓名与题名的间距
    if "cn_title" in found and "cn_author" in found:
        gap = found["cn_author"] - found["cn_title"]
        if gap > 2:
            notices.append(f"作者姓名（段{found['cn_author']}）与中文题名"
                           f"（段{found['cn_title']}）相隔 {gap} 段——"
                           f"官方序列要求作者姓名紧随题名之后")
    return issues, notices


def analyze(doc):
    paras, refs, abs_text, kw_text, titles, body = dp.extract_paras(doc)
    found, detail = find_front_matter(paras)
    issues, notices = order_issues(found)
    if "cn_title" not in found:
        issues.append("未检出中文题名——稿件应以文章标题开头")
    return {"found": found, "detail": detail,
            "order": [k for k in SEQ if k in found],
            "issues": issues, "notices": notices}


def main():
    args = [a for a in sys.argv[1:] if not a.startswith("--")]
    flags = set(a for a in sys.argv[1:] if a.startswith("--"))
    if not args or args[0] in ("-h", "--help"):
        print(__doc__)
        sys.exit(0)
    DOC = args[0]
    if not os.path.isfile(DOC):
        sys.exit(f"找不到文件：{DOC}")

    res = analyze(DOC)
    print("=" * 72)
    print("稿件前置元数据序列核验（官方须知 §4 格式顺序）")
    print(f"稿件：{os.path.basename(DOC)}")
    print("=" * 72)

    print("\n【官方序列实测】（段落号为 python-docx 全量段落索引）")
    for k in SEQ:
        if k in res["found"]:
            print(f"  ✓ {SEQ_LABEL[k]:<16} 段{res['found'][k]:>3}"
                  f"  {res['detail'][k][:52]}")
        else:
            need = "🔴 缺失" if k in ("cn_title", "en_title") else "— 未见"
            print(f"  ✗ {SEQ_LABEL[k]:<16} {need}")

    print("\n" + "-" * 72)
    if res["issues"]:
        print(f"🔴 发现 {len(res['issues'])} 项序列问题：")
        for it in res["issues"]:
            print(f"   · {it}")
    else:
        print("✅ 前置元素齐全且次序符合官方序列。")
    if res["notices"]:
        print(f"\n🟡 建议核对 {len(res['notices'])} 项：")
        for it in res["notices"]:
            print(f"   · {it}")

    if "--json" in flags:
        jf = os.path.splitext(DOC)[0] + ".front_matter.json"
        with open(jf, "w", encoding="utf-8") as f:
            json.dump(res, f, ensure_ascii=False, indent=1)
        print(f"\nJSON 已写入：{jf}")

    sys.exit(1 if res["issues"] else 0)


if __name__ == "__main__":
    main()
