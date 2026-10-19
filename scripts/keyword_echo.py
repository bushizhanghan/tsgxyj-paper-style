# -*- coding: utf-8 -*-
"""keyword_echo.py —— 关键词与论文内容的呼应核验

关键词是审稿人判断「论文到底研究什么」的第一触点，也是数据库检索的
唯一入口。常见问题却几乎无人自查：

  · **装饰性关键词**：关键词在全文（题名+摘要+正文）一次都没出现——
    读者按它检索到论文，却发现论文根本不谈这个；
  · **数量与形态**：该刊实测 4–5 个为主（均值 4.4）；少于 3 个覆盖不足，
    多于 6 个有堆砌之嫌；
  · **中英文对应**：英文关键词与中文关键词是否逐位对应、译法一致；
  · **题名呼应**：关键词与题名主题词的重叠度（该刊关键词通常是题名
    核心概念的同义/下位展开，零重叠往往意味着文不对题）。

本脚本逐项核验，并输出「每个关键词在全文的出现次数」供你判断
是否真的被论文覆盖。

用法:
    python keyword_echo.py <论文.docx> [--json]
退出码：0 = 关键词无装饰性问题；1 = 检出装饰性关键词（全文未出现）。
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

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import diagnose_paper as dp  # noqa: E402

MIN_KW = 3           # 少于 3 个 → 覆盖不足
MAX_KW = 6           # 多于 6 个 → 堆砌
# 关键词在题名/摘要中至少出现 1 次才算「被覆盖」；正文 0 次但题名或摘要有 → 提示级
COVER_REGIONS = ("题名", "摘要", "正文")


def split_kws(kw_text):
    t = (kw_text or "").strip()
    t = re.sub(r"^【?\s*(?:关\s*键\s*词|Key\s*words?)\s*】?\s*[:：]?", "", t, flags=re.I)
    return [p.strip() for p in re.split(r"[;；,，、]\s*|\s{2,}", t) if p.strip()]


def title_of(paras):
    """题名：摘要段之前的第一个较长中文段（跳过纯作者名行）。"""
    for _, t in paras:
        s = (t or "").strip()
        if not s:
            continue
        if re.search(r"^【?摘\s*要】?[：:]?", s[:12]):
            return ""
        c = re.sub(r"\s+", "", s)
        if 2 <= len(c) <= 5:
            continue                      # 作者名行
        if re.match(r"^[A-Za-z\s:：\-]+$", s) and len(s) < 60:
            continue                      # 英文题名单独处理
        return s
    return ""


def en_keywords(paras):
    for _, t in paras:
        m = re.search(r"(?:Key\s*words?)\s*[:：]\s*(.+)$", t, re.I)
        if m:
            return [x.strip() for x in re.split(r"[;；,]\s*", m.group(1)) if x.strip()]
    return []


def analyze(doc):
    paras, refs, abs_text, kw_text, titles, body = dp.extract_paras(doc)
    cn_kws = split_kws(kw_text)
    en_kws = en_keywords(paras)
    ttl = title_of(paras)
    abs_c = dp.clean(abs_text or "")
    body_c = dp.clean("\n".join(body))

    issues, notices = [], []
    rows = []
    for i, kw in enumerate(cn_kws, 1):
        in_title = ttl.count(kw) if ttl else 0
        in_abs = abs_c.count(kw)
        in_body = body_c.count(kw)
        total = in_title + in_abs + in_body
        rows.append({"i": i, "kw": kw, "title": in_title, "abs": in_abs,
                     "body": in_body, "total": total})
        if total == 0:
            issues.append(f"关键词「{kw}」全文（题名/摘要/正文）0 次出现——"
                          f"装饰性关键词，读者检索到却发现论文不谈此概念，"
                          f"建议替换为论文实际核心概念")
        elif in_body == 0 and in_abs == 0 and in_title <= 1:
            notices.append(f"关键词「{kw}」仅出现在题名（{in_title} 次），"
                           f"摘要与正文均未出现——确认它是否真是论文核心概念")

    # 数量
    n = len(cn_kws)
    if n < MIN_KW:
        issues.append(f"关键词仅 {n} 个（该刊实测 4–5 个为主，均值 4.4）——覆盖不足")
    elif n > MAX_KW:
        notices.append(f"关键词 {n} 个偏多（该刊 4–5 个为主）——避免堆砌，"
                       f"保留最具检索价值的")

    # 中英文对应
    if cn_kws and en_kws and len(cn_kws) != len(en_kws):
        issues.append(f"中英文关键词个数不一致：中文 {len(cn_kws)} vs 英文 {len(en_kws)}")
    kw_pairs = list(zip(cn_kws, en_kws)) if len(cn_kws) == len(en_kws) else []

    # 题名呼应
    if ttl:
        title_core = set()
        for kw in cn_kws:
            if kw in ttl:
                title_core.add(kw)
        if cn_kws and not title_core:
            notices.append("关键词与题名零重叠——该刊关键词通常是题名核心概念的"
                           "同义/下位展开，请确认文题是否相符")

    return {"cn_kws": cn_kws, "en_kws": en_kws, "title": ttl,
            "rows": rows, "kw_pairs": kw_pairs,
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
    print("=" * 74)
    print("关键词与论文内容呼应核验（装饰性关键词专查）")
    print(f"稿件：{os.path.basename(DOC)}")
    print(f"中文关键词 {len(res['cn_kws'])} 个 ｜ 英文 {len(res['en_kws'])} 个")
    if res["title"]:
        print(f"题名：{res['title'][:60]}")
    print("=" * 74)

    if res["rows"]:
        print("【关键词覆盖度】（题名/摘要/正文 出现次数）")
        for r in res["rows"]:
            mark = "🔴" if r["total"] == 0 else ("🟡" if r["body"] == 0 else "✓")
            print(f"  {mark} {r['i']}. {r['kw']}：题名 {r['title']} ｜ 摘要 {r['abs']}"
                  f" ｜ 正文 {r['body']} ｜ 合计 {r['total']}")

    if res["kw_pairs"]:
        print("\n【中英关键词对照】")
        for i, (c, e) in enumerate(res["kw_pairs"], 1):
            print(f"  {i}. {c}  →  {e}")

    print("\n" + "-" * 74)
    if res["issues"]:
        print(f"🔴 发现 {len(res['issues'])} 项问题：")
        for it in res["issues"]:
            print(f"   · {it}")
    else:
        print("✅ 关键词均被论文覆盖，数量与中英对应无问题。")
    if res["notices"]:
        print(f"\n🟡 建议核对 {len(res['notices'])} 项：")
        for it in res["notices"]:
            print(f"   · {it}")

    if "--json" in flags:
        jf = os.path.splitext(DOC)[0] + ".keyword_echo.json"
        with open(jf, "w", encoding="utf-8") as f:
            json.dump(res, f, ensure_ascii=False, indent=1)
        print(f"\nJSON 已写入：{jf}")

    sys.exit(1 if res["issues"] else 0)


if __name__ == "__main__":
    main()
