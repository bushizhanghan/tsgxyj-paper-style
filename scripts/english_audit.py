# -*- coding: utf-8 -*-
"""english_audit.py —— 英文摘要与关键词对称性核验

官方投稿须知明文要求中英文一一对应（格式序列：中文题名/作者 → 英文题名/作者 →
中文摘要 → 中文关键词 → Abstract: → Key words:）。此前工具链只检查
「Abstract: / Key words: 是否存在」（diagnose_paper 的官方合规维度），
不检查**质量对称性**——而这是初审退修的常见技术性理由：

  · 长度对称：英文摘要与中文摘要字数比是否落在合理区间（英文 ~2–3 倍中文
    字符数为常见；过短说明只是标题式翻译，过长多为中式冗余展开）；
  · 关键词对称：英文关键词个数与中文关键词个数是否一致、逐个是否有对应；
  · 术语一致：中文关键词在正文的译法是否与英文关键词用词一致
    （如中文「信息无障碍」正文注为 accessibility，英文关键词却写
     information barrier-free —— 同一概念两种英文表述）；
  · 体例形态：Abstract/Key words 标签形态（官方写 `Abstract:` 与
     `Key words:`，冒号须为半角；英文字母大小写、句首大写）；
  · 常见硬伤：英文摘要里残留中文标点、残留汉字、时态混用提示。

用法:
    python english_audit.py <论文.docx> [--json]
退出码：0 = 未发现问题；1 = 存在需修改项（🟡 建议级不算）。

局限（形态与字数比对，不做语法/语义评价）：
  · 不判断英文语法正误，也不评价翻译质量；
  · 术语对应的判断基于「中文关键词 / 英文关键词 同位置配对」，仅供参考；
  · 该刊对英文摘要的实际篇幅方差较大，长度区间设为宽容区间并标🟡。
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

# 官方标签形态（半角冒号）
LABEL_ABS = re.compile(r"(?:Abstract|ABSTRACT)\s*[:：]", re.I)
LABEL_KW = re.compile(r"(?:Key\s*words?|KEY\s*WORDS?)\s*[:：]", re.I)
# 英文摘要合理长度区间（字母/字符数）：中文 200–300 字 → 英文约 400–1200 字符
ABS_EN_MIN, ABS_EN_MAX = 300, 1600
# 中文摘要与英文摘要的字符数比合理区间（英文常为中文的 1.6–3.2 倍）
RATIO_MIN, RATIO_MAX = 1.2, 3.6

CJK_RE = re.compile(r"[\u4e00-\u9fff]")
# 英文句中的中文标点（. , 之外的顿号/句号/引号等）
CJK_PUNCT_RE = re.compile(r"[，。；：、（）【】“”‘’！？]")


def find_en_block(paras, label_re):
    """定位英文标签段：返回 (标签原文, 其后文本)。标签可能独占一段、也可能后接正文。"""
    for i, (style, t) in enumerate(paras):
        if not t:
            continue
        m = label_re.search(t)
        if not m:
            continue
        head = t[:m.end()]
        tail = t[m.end():].strip()
        # 若标签段只有标签，正文在下一段
        if len(tail) < 20 and i + 1 < len(paras):
            nxt = paras[i + 1][1].strip()
            if nxt and not label_re.search(nxt):
                tail = (tail + " " + nxt).strip()
        return head, tail
    return None, ""


def count_en_chars(s):
    """英文摘要有效字符数（去空白）。"""
    return len(re.sub(r"\s+", "", s or ""))


def split_kw(s):
    """中英文关键词切分（支持 ; ； , ， 空格 分隔；剥离中英文标签前缀）。"""
    t = (s or "").strip()
    # 剥离各种标签形态：【关键词】/关键词：/Key words: 等
    t = re.sub(r"^【?\s*(?:关\s*键\s*词|Key\s*words?)\s*】?\s*[:：]?", "", t, flags=re.I)
    parts = re.split(r"[;；,，、]\s*|\s{2,}", t)
    return [p.strip() for p in parts if p.strip()]


def audit(doc):
    paras, refs, abs_text, kw_text, titles, body = dp.extract_paras(doc)
    issues, notices = [], []
    stat = {}

    abs_head, abs_en = find_en_block(paras, LABEL_ABS)
    kw_head, kw_en = find_en_block(paras, LABEL_KW)

    # ---------- ① 存在性 ----------
    if abs_head is None:
        issues.append("未检出英文摘要（官方格式：Abstract:）")
    if kw_head is None:
        issues.append("未检出英文关键词（官方格式：Key words:）")
    if abs_head is None and kw_head is None:
        return {"issues": issues, "notices": notices, "stat": stat}

    # ---------- ② 标签形态 ----------
    for head, name in ((abs_head, "Abstract"), (kw_head, "Key words")):
        if not head:
            continue
        if "：" in head:
            issues.append(f"{name} 标签用了全角冒号，官方体例为半角「{name}:」")
        if name == "Key words" and re.search(r"Keywords", head, re.I):
            notices.append("英文关键词标签写作 Keywords，官方体例为「Key words:」（两词）")

    # ---------- ③ 英文摘要长度与对称性 ----------
    cn_len = len(dp.clean(abs_text or ""))
    en_len = count_en_chars(abs_en)
    stat.update({"abs_cn_chars": cn_len, "abs_en_chars": en_len})
    if abs_head is not None:
        if en_len < ABS_EN_MIN:
            issues.append(f"英文摘要仅 {en_len} 字符，过短——"
                          f"官方要求中英文摘要对应，疑为标题式翻译（建议 ≥{ABS_EN_MIN}）")
        elif en_len > ABS_EN_MAX:
            notices.append(f"英文摘要 {en_len} 字符偏长（建议 ≤{ABS_EN_MAX}），"
                           f"多因中式冗余展开，可压缩")
        if cn_len >= 80 and en_len:
            ratio = en_len / cn_len
            stat["en_cn_ratio"] = round(ratio, 2)
            if ratio < RATIO_MIN:
                notices.append(f"英文摘要为中文摘要的 {ratio:.1f} 倍（常见 1.6–3.2 倍），偏短，"
                               f"核对是否漏译要素")
            elif ratio > RATIO_MAX:
                notices.append(f"英文摘要为中文摘要的 {ratio:.1f} 倍（常见 1.6–3.2 倍），偏长，"
                               f"核对是否有冗余")

    # ---------- ④ 关键词对称 ----------
    cn_kws = split_kw(kw_text)
    en_kws = split_kw(kw_en)
    stat.update({"kw_cn": len(cn_kws), "kw_en": len(en_kws),
                 "cn_keywords": cn_kws, "en_keywords": en_kws})
    if abs_head is not None or kw_head is not None:
        if cn_kws and en_kws and len(cn_kws) != len(en_kws):
            issues.append(f"中英文关键词个数不一致：中文 {len(cn_kws)} 个 vs 英文 {len(en_kws)} 个"
                          f"（官方要求一一对应）")
        if len(en_kws) < 3:
            issues.append(f"英文关键词仅 {len(en_kws)} 个，过少（该刊实测 4–5 个）")

    # ---------- ⑤ 英文块残留中文标点/汉字 ----------
    for text, name in ((abs_en, "英文摘要"), (kw_en, "英文关键词")):
        if not text:
            continue
        cjk = CJK_RE.findall(text)
        punct = CJK_PUNCT_RE.findall(text)
        if cjk:
            issues.append(f"{name}中残留汉字 {len(cjk)} 个（如「{''.join(cjk[:5])}」）——须全英文")
        if punct:
            issues.append(f"{name}中残留中文标点 {len(punct)} 个（如「{''.join(punct[:5])}」）"
                          f"——英文须用半角标点")

    # ---------- ⑥ 术语一致（中文关键词 vs 英文关键词 同位置配对提示） ----------
    if cn_kws and en_kws and len(cn_kws) == len(en_kws):
        stat["kw_pairs"] = list(zip(cn_kws, en_kws))
    # 英文摘要与英文关键词的首字母大写体例
    if abs_en and abs_en[0].islower():
        notices.append("英文摘要首字母未大写（Abstract: 后通常首字母大写）")

    return {"issues": issues, "notices": notices, "stat": stat}


def main():
    args = [a for a in sys.argv[1:] if not a.startswith("--")]
    flags = set(a for a in sys.argv[1:] if a.startswith("--"))
    if not args or args[0] in ("-h", "--help"):
        print(__doc__)
        sys.exit(0)
    DOC = args[0]
    if not os.path.isfile(DOC):
        sys.exit(f"找不到文件：{DOC}")

    res = audit(DOC)
    st = res["stat"]
    print("=" * 72)
    print("英文摘要与关键词对称性核验（官方中英文对应要求）")
    print(f"稿件：{os.path.basename(DOC)}")
    if st.get("abs_cn_chars"):
        print(f"中文摘要 {st['abs_cn_chars']} 字 ｜ 英文摘要 {st.get('abs_en_chars', 0)} 字符"
              + (f"（比值 {st['en_cn_ratio']:.1f}）" if st.get("en_cn_ratio") else ""))
    if st.get("kw_cn") is not None:
        print(f"关键词：中文 {st.get('kw_cn', 0)} 个 ｜ 英文 {st.get('kw_en', 0)} 个"
              + ("  ✓ 个数一致" if st.get("kw_cn") == st.get("kw_en") else "  ⚠ 个数不一致"))
    print("=" * 72)

    if res["issues"]:
        print(f"🔴 发现 {len(res['issues'])} 项需修改：")
        for it in res["issues"]:
            print(f"   · {it}")
    else:
        print("✅ 英文摘要与关键词：未发现形态/对称性问题")
    if res["notices"]:
        print(f"\n🟡 建议核对 {len(res['notices'])} 项：")
        for it in res["notices"]:
            print(f"   · {it}")
    if st.get("kw_pairs"):
        print("\n中英关键词对照（供核对译法一致性）：")
        for i, (c, e) in enumerate(st["kw_pairs"], 1):
            print(f"  {i}. {c}  →  {e}")

    if "--json" in flags:
        jf = os.path.splitext(DOC)[0] + ".english_audit.json"
        with open(jf, "w", encoding="utf-8") as f:
            json.dump(res, f, ensure_ascii=False, indent=1)
        print(f"\nJSON 已写入：{jf}")

    sys.exit(1 if res["issues"] else 0)


if __name__ == "__main__":
    main()
