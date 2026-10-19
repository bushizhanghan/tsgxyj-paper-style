# -*- coding: utf-8 -*-
"""en_title_check.py —— 英文题名核验（官方 §4 序列 + §7 英文题名体例）

官方投稿须知两处与英文题名直接相关：

  · §4 格式顺序：中文题名之后**必须**有「文章标题（英文）」，
    投稿稿漏英文题名是常见退修理由（存在性由 front_matter.py 负责）；
  · §7 参考文献格式：「文末注释或参考文献中所有**英文文献标题中的
    实意词首字母大写**」——diagnose_paper 只把这条落到**参考文献条目**上，
    稿件自身的英文题名同样应遵守同一口径（Title Case）。

本脚本核英文题名本身的三项形态：

  · 实意词首字母大写（Title Case）：与官方对英文文献题名的要求同口径，
    逐词列出未大写的实意词；冠词/连词/介词等虚词不算；
  · 与中文题名的对应性：按词数比 0.5–2.5 判定明显过短/过长
    （中文题名按语义块粗分），提示「英文题名与中文题名不对应」；
  · 副标题衔接：中文题名用破折号「——」带副标题时，英文题名是否用
    冒号或破折号同样带出副标题（副标题丢失＝英文题名信息缺位）。

用法:
    python en_title_check.py <论文.docx> [--json]
退出码：0 = 英文题名形态合规；1 = 实意词未大写或中英题名明显不对应。

说明：不做语法与翻译质量评价；中文题名分块为启发式切分，仅用于词数比。
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

# 虚词（Title Case 下不要求首字母大写）——与 diagnose_paper 的英文文献口径一致
FUNC_WORDS = {"the", "and", "for", "with", "from", "into", "of", "in", "on",
              "to", "a", "an", "by", "or", "as", "at", "but", "nor", "so",
              "yet", "per", "via", "up", "off", "out", "than", "that",
              "if", "is", "are", "was", "were", "be", "been", "being"}
# 常见复合虚词（如 "based on" 中的 based on）
FUNC_BIGRAMS = {"based on", "according to", "due to", "in terms of",
                "with regard to", "in order to", "as well as"}
# 单个字母缩略词（如 A、AI）保留原样
ACRONYM_RE = re.compile(r"^[A-Z0-9]{1,4}$")

# 中文题名分块信号（按语义块粗分，仅用于词数比）
CN_SPLIT_RE = re.compile(r"——|—|－|[-:：]|、|与|和|及|基于|面向|下的|中的|"
                         r"对|从|在|为|以|的|了|着|及其")
# 词数比合理区间（中文块数 vs 英文词数）
RATIO_MIN, RATIO_MAX = 0.5, 2.5
# 中文题名最短块数（低于此不判对应性）
CN_BLOCK_MIN = 2


def content_words(title):
    """英文题名的实意词（去虚词、去纯数字/符号）。"""
    words = []
    for w in re.findall(r"[A-Za-z][A-Za-z\-']*", title or ""):
        lw = w.lower().strip("-'")
        if not lw:
            continue
        if lw in FUNC_WORDS:
            continue
        # 复合虚词的首词也跳过（如 based / according / due）
        if lw in {"based", "according", "due", "terms", "regard", "order",
                  "well", "rather"}:
            continue
        words.append(w)
    return words


def cn_blocks(title):
    """中文题名的语义块粗分（去掉「的」等结构词后再切）。"""
    t = re.sub(r"[的了着及其与和及]", "", title or "")
    parts = [p for p in CN_SPLIT_RE.split(t) if p and p.strip()]
    return parts


def find_titles(paras):
    """返回 (中文题名, 英文题名, found)。口径与 front_matter.py 一致。"""
    import front_matter as fm
    found, detail = fm.find_front_matter(paras)
    return detail.get("cn_title", ""), detail.get("en_title", ""), found


def analyze(doc):
    paras, refs, abs_text, kw_text, titles, body = dp.extract_paras(doc)
    cn_title, en_title, found = find_titles(paras)
    issues, notices = [], []
    stat = {"cn_title": cn_title, "en_title": en_title}

    if not cn_title:
        issues.append("未检出中文题名——无法核验英文题名对应性")
    if not en_title:
        # 存在性由 front_matter.py 报（同一官方条款），这里只作形态侧提示
        notices.append("未见英文题名——官方 §4 格式序列要求中文题名后有"
                       "「文章标题（英文）」；英文题名的实意词大写、与中文"
                       "题名对应性、副标题衔接在此无法核验")
        return {"issues": issues, "notices": notices, "stat": stat}

    # ---------- ① 实意词首字母大写（官方 §7 同口径） ----------
    cw = content_words(en_title)
    lower = [w for w in cw if w[0].islower()]
    stat["content_words"] = len(cw)
    if lower:
        issues.append(f"英文题名有 {len(lower)} 个实意词未首字母大写："
                      f"{'、'.join(lower[:6])}"
                      f"——官方 §7 要求英文文献标题实意词首字母大写，"
                      f"稿件英文题名应同口径（Title Case）")
    elif cw:
        notices.append(f"英文题名实意词首字母大写：✓（{len(cw)} 个实意词）")

    # ---------- ② 中英题名对应性（词数比） ----------
    if cn_title:
        blocks = cn_blocks(cn_title)
        n_cn = max(len(blocks), 1)
        n_en = len(re.findall(r"[A-Za-z][A-Za-z\-']*", en_title))
        stat.update({"cn_blocks": n_cn, "en_words": n_en})
        if n_en and n_cn >= CN_BLOCK_MIN:
            ratio = n_en / n_cn
            stat["word_ratio"] = round(ratio, 2)
            if ratio < RATIO_MIN:
                issues.append(f"英文题名词数 {n_en} 明显少于中文题名块数 "
                              f"{n_cn}（比值 {ratio:.2f} < {RATIO_MIN}）——"
                              f"英文题名疑过短，与中文题名不对应")
            elif ratio > RATIO_MAX:
                notices.append(f"英文题名词数 {n_en} 明显多于中文题名块数 "
                               f"{n_cn}（比值 {ratio:.2f} > {RATIO_MAX}）——"
                               f"核对是否译出了中文题名没有的限定语")
        elif n_cn < CN_BLOCK_MIN:
            notices.append(f"中文题名仅 {n_cn} 个语义块，过短不判对应性")

    # ---------- ③ 副标题衔接 ----------
    cn_sub = bool(re.search(r"——|—|－", cn_title or ""))
    en_sub = bool(re.search(r":|—|–| - ", en_title or ""))
    stat.update({"cn_subtitle": cn_sub, "en_subtitle": en_sub})
    if cn_sub and not en_sub:
        issues.append("中文题名含副标题（破折号衔接），英文题名未见冒号/破折号"
                      "衔接的副标题——英文题名信息缺位，需补出副标题")
    if en_sub and not cn_sub:
        notices.append("英文题名含副标题而中文题名没有——核对是否为中文题名漏排"
                       "副标题")

    # ---------- ④ 常见形态提示 ----------
    if re.match(r"^[a-z]", en_title):
        notices.append("英文题名首字母小写——题名首词应大写")
    if en_title.rstrip().endswith((".", "。", "，", ",")):
        notices.append("英文题名以句号/逗号结尾——题名通常不加句末标点")

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

    res = analyze(DOC)
    st = res["stat"]
    print("=" * 72)
    print("英文题名核验（官方 §4 序列 + §7 英文题名大写口径）")
    print(f"稿件：{os.path.basename(DOC)}")
    print("=" * 72)

    print(f"\n中文题名：{st.get('cn_title') or '（未检出）'}")
    print(f"英文题名：{st.get('en_title') or '（未检出）'}")
    if st.get("cn_blocks"):
        print(f"中文题名语义块 {st['cn_blocks']} 个 ｜ 英文题名词数 "
              f"{st['en_words']} 个"
              + (f"（比值 {st['word_ratio']:.2f}，区间 "
                 f"{RATIO_MIN}–{RATIO_MAX}）" if st.get("word_ratio") else ""))
    print(f"副标题：中文 {'有' if st.get('cn_subtitle') else '无'} ｜ "
          f"英文 {'有' if st.get('en_subtitle') else '无'}")

    print("\n" + "-" * 72)
    if res["issues"]:
        print(f"🔴 发现 {len(res['issues'])} 项英文题名问题：")
        for it in res["issues"]:
            print(f"   · {it}")
    else:
        print("✅ 英文题名形态合规（实意词大写、与中文题名对应、副标题衔接）。")
    if res["notices"]:
        print(f"\n🟡 建议核对 {len(res['notices'])} 项：")
        for it in res["notices"]:
            print(f"   · {it}")

    if "--json" in flags:
        jf = os.path.splitext(DOC)[0] + ".en_title_check.json"
        with open(jf, "w", encoding="utf-8") as f:
            json.dump(res, f, ensure_ascii=False, indent=1)
        print(f"\nJSON 已写入：{jf}")

    sys.exit(1 if res["issues"] else 0)


if __name__ == "__main__":
    main()
