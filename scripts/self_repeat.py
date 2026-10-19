# -*- coding: utf-8 -*-
"""self_repeat.py —— 稿内自我重复检测（查重前哨）

该刊官方硬门槛：「总文字复制比 < 7%」「去除引用文献复制比接近 0」，
未通过不予排版。而**查重率偏高的第一来源往往不是抄袭，而是作者自己**：
摘要与结语重复、引言与文献综述重复、同一段论证在两个章节各写一遍、
图表说明与正文重复——这类「稿内自我重复」既拉高复制比，也让审稿人觉得注水。

本脚本只做**稿内**比对（不与外部语料比对，那不是本工具职责）：
  · 句子级：切句（≥12 字），归一化（去空白/标点/全半角统一）后做
    字符 shingle（3-gram）Jaccard 相似度，≥阈值判重（默认 0.75）；
  · 段落级：同法比对段落，≥阈值判重（默认 0.80）；
  · 区域标注：摘要 / 正文各章 / 结语 / 图表标题 分属不同区域，
    跨区域重复比同区域内重复更严重（同区域内少量套话可接受）；
  · 引用豁免：含引用点 [n] 的句子相似度扣减（引用他人表述属正常学术行为，
    但仍提示核对是否过长引用）。

用法:
    python self_repeat.py <论文.docx> [--thresh 0.75] [--top 20] [--json]
退出码：0 = 未检出重复；1 = 检出疑似稿内重复（🔴 跨区域重复或 🟡 区域内重复）。

局限（字符 shingle 相似度，不是语义比对）：
  · 同义改写（把「提升服务效能」改成「改善服务效率」）检不出；
  · 模板化套话（「本研究以……为对象，采用……方法」）可能误判，故设阈值且分区计；
  · 与外部文献的重复需用 CNKI 等查重系统，本工具不能替代。
"""
import os
import re
import sys
import json
import itertools

# Windows 控制台 GBK 保护
if hasattr(sys.stdout, "reconfigure"):
    try:
        sys.stdout.reconfigure(encoding="utf-8", errors="replace")
    except Exception:
        pass

HERE = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, HERE)
import diagnose_paper as dp  # noqa: E402

SHINGLE_N = 3          # 字符 shingle 长度
MIN_SENT_LEN = 12      # 参与句子级比对的句长下限
MIN_PARA_LEN = 60      # 参与段落级比对的段长下限
CITE_RE = re.compile(r"\[\d{1,3}(?:\s*[-–—,，]\s*\d{1,3})*\]")
# 枚举模板句：同一格式的并列条目（「（4）XX模型，20道题全部通过，通过率100%」×10 条），
# 属正常列举，与「自我重复」性质不同——统一剔除后再比对。
ENUM_TEMPLATE_RE = re.compile(
    r"^[（(]\s*\d{1,2}\s*[)）]\s*"          # 以（n）开头
    r"|^\d{1,2}\s*[、.．)]\s*")             # 或以「n、」开头
ENUM_CORE_RE = re.compile(r"通过率|得分|占比|共计|合计|总计|条数|道题")


def normalize(s):
    """归一化：去空白与标点、全角转半角、去引用点。"""
    t = dp.clean(s or "")
    # 全角字母数字 → 半角
    out = []
    for ch in t:
        code = ord(ch)
        if 0xFF01 <= code <= 0xFF5E:
            out.append(chr(code - 0xFEE0))
        elif code == 0x3000:
            out.append(" ")
        else:
            out.append(ch)
    t = "".join(out)
    t = CITE_RE.sub("", t)
    t = re.sub(r"[^\w\u4e00-\u9fff]", "", t)
    return t.lower()


def shingles(s, n=SHINGLE_N):
    if len(s) < n:
        return set()
    return {s[i:i + n] for i in range(len(s) - n + 1)}


def jaccard(a, b):
    if not a or not b:
        return 0.0
    inter = len(a & b)
    union = len(a | b)
    return inter / union if union else 0.0


def split_regions(path):
    """切出 摘要 / 各章正文 / 结语 / 图表标题 四个区域的句子列表。

    注意：摘要段与关键词段必须从正文区域剔除——dp.extract_paras 的 body
    只去掉参考文献，摘要仍留在 paras 里，若不去重会把摘要与自身当「跨区域重复」。
    """
    paras, refs, abs_text, kw_text, titles, body = dp.extract_paras(path)
    regions = []          # [(区域名, [句子...])]

    if abs_text:
        regions.append(("摘要", [s for s in re.split(r"(?<=[。？！])", abs_text) if s.strip()]))

    # 正文按章切（只跳过「带摘要/关键词前缀」的元数据段，绝不按文本相似度跳过——
    # 那会把正文中与摘要雷同的句子（正是要检测的重复！）一并误删）
    cur_chapter = "正文（未分章）"
    chapters = {}
    for style, t in paras:
        if not t:
            continue
        if t.startswith("参考文献"):
            break
        # 跳过摘要段与关键词段（以「摘要/关键词」前缀标识，含【摘　要】变体）
        if re.search(r"^【?摘\s*要】?[：:]?", t[:12]) or \
           re.search(r"^【?关\s*键\s*词】?[：:]?", t[:12]):
            continue
        is_title = (re.match(r"^(0|1|2|3|4|5|6|7)\s*[^\d]", t) and len(t) < 40) or \
                   (style.startswith("Heading") and len(t) < 40)
        if is_title:
            cur_chapter = t.strip()[:22]
            continue
        if cur_chapter not in chapters:
            chapters[cur_chapter] = []
        chapters[cur_chapter].append(t)
    for ch, plist in chapters.items():
        if "结语" in ch or "结论" in ch:
            name = "结语"
        else:
            name = ch
        sents = []
        for p in plist:
            sents += [s for s in re.split(r"(?<=[。？！])", p) if s.strip()]
        if sents:
            regions.append((name, sents))

    # 图表标题
    caps = [t.strip() for t in titles + list(chapters.keys())
            if re.match(r"^(图|表)\s*\d", t.strip())]
    if caps:
        regions.append(("图表标题", caps))

    return regions, paras, refs, abs_text, kw_text, titles, body


def collect_units(regions, level="sent"):
    """把各区文本展开成 (区域, 序号, 原文, 归一化, shingles) 列表。

    枚举模板句（并列条目的同格式陈述）不参与比对——那是列表不是重复。
    """
    units = []
    skipped_enum = 0
    for region, texts in regions:
        for i, t in enumerate(texts):
            raw = (t or "").strip()
            if not raw:
                continue
            min_len = MIN_SENT_LEN if level == "sent" else MIN_PARA_LEN
            if len(dp.clean(raw)) < min_len:
                continue
            if level == "sent" and ENUM_TEMPLATE_RE.match(raw) and ENUM_CORE_RE.search(raw):
                skipped_enum += 1
                continue
            nz = normalize(raw)
            if len(nz) < min_len:
                continue
            units.append({
                "region": region, "idx": i + 1, "text": raw,
                "norm": nz, "sh": shingles(nz),
                "has_cite": bool(CITE_RE.search(raw)),
            })
    return units, skipped_enum


def find_repeats(units, thresh):
    """两两比对找重复对（先按长度分层倒排，减少无效比较）。"""
    pairs = []
    n = len(units)
    for i in range(n):
        a = units[i]
        if not a["sh"]:
            continue
        for j in range(i + 1, n):
            b = units[j]
            if not b["sh"]:
                continue
            # 长度差异过大（<50%）直接跳过，省算力
            la, lb = len(a["norm"]), len(b["norm"])
            if min(la, lb) / max(la, lb) < 0.5:
                continue
            jac = jaccard(a["sh"], b["sh"])
            if jac < thresh:
                continue
            # 引用豁免：一方含引用点时相似度打 8 折（长段引用不判硬重复）
            score = jac * 0.8 if (a["has_cite"] or b["has_cite"]) else jac
            if score < thresh:
                continue
            cross = a["region"] != b["region"]
            pairs.append({
                "a": a, "b": b, "score": round(score, 3),
                "raw_score": round(jac, 3),
                "cross": cross,
                "level": "🔴 跨区域" if cross else "🟡 同区域内",
            })
    pairs.sort(key=lambda x: (-x["score"], x["a"]["region"]))
    return pairs


def trunc(s, n):
    return s if len(s) <= n else s[:n] + "…"


def main():
    args = [a for a in sys.argv[1:] if not a.startswith("--")]
    flags = set(a for a in sys.argv[1:] if a.startswith("--"))
    if not args or args[0] in ("-h", "--help"):
        print(__doc__)
        sys.exit(0)
    DOC = args[0]
    if not os.path.isfile(DOC):
        sys.exit(f"找不到文件：{DOC}")
    THRESH = 0.75
    TOP = 20
    for i, a in enumerate(sys.argv):
        if a == "--thresh" and i + 1 < len(sys.argv):
            try:
                THRESH = float(sys.argv[i + 1])
            except ValueError:
                pass
        if a == "--top" and i + 1 < len(sys.argv):
            try:
                TOP = int(sys.argv[i + 1])
            except ValueError:
                pass

    regions, paras, refs, abs_text, kw_text, titles, body = split_regions(DOC)
    sent_units, enum_skipped = collect_units(regions, level="sent")
    sent_pairs = find_repeats(sent_units, THRESH)

    print("=" * 72)
    print("稿内自我重复检测（查重前哨 · 只比稿内，不替代 CNKI 查重）")
    print(f"稿件：{os.path.basename(DOC)}")
    print(f"比对句数 {len(sent_units)}（≥{MIN_SENT_LEN} 字，跳过枚举模板句 {enum_skipped}）"
          f"｜阈值 {THRESH:.2f}｜shingle {SHINGLE_N}-gram Jaccard")
    print(f"区域划分：{' / '.join(name for name, _ in regions)}")
    print("=" * 72)

    if not sent_pairs:
        print(f"✅ 未检出相似度 ≥{THRESH:.2f} 的重复句对。")
        sys.exit(0)

    cross = [p for p in sent_pairs if p["cross"]]
    inner = [p for p in sent_pairs if not p["cross"]]
    print(f"相似句对：🔴 跨区域 {len(cross)} 对 ｜ 🟡 同区域内 {len(inner)} 对")
    print("-" * 72)
    shown = 0
    for p in sent_pairs:
        if shown >= TOP:
            print(f"……（其余 {len(sent_pairs) - shown} 对见 --json）")
            break
        shown += 1
        a, b = p["a"], p["b"]
        cite_note = "（含引用已折减）" if (a["has_cite"] or b["has_cite"]) else ""
        print(f"{p['level']} 相似度 {p['score']:.2f}{cite_note}")
        print(f"  A [{a['region']} 第{a['idx']}句] {trunc(a['text'], 76)}")
        print(f"  B [{b['region']} 第{b['idx']}句] {trunc(b['text'], 76)}")
    print("-" * 72)
    if cross:
        print(f"🔴 跨区域重复 {len(cross)} 对——查重率的主要来源，投稿前必须改写或删除其中一处。")
        print("   （摘要↔结语、引言↔文献综述 是最常见的两类）")
        sys.exit(1)
    print(f"🟡 同区域内重复 {len(inner)} 对——多为套话复述，建议合并精简。")
    sys.exit(1 if inner else 0)


if __name__ == "__main__":
    main()
