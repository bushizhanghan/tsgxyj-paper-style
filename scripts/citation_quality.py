# -*- coding: utf-8 -*-
"""citation_quality.py —— 引文—论点关联核验（「装饰性引用」专查）

审稿人退修意见的高频理由：「引用与论点无关，只挂在段尾凑格式」。
本工具链此前核引文的**形态**（ref_format.py：GB/T 7714 要素）与**结构**
（citation_analysis.py：时效/类型/来源），不核**引用点与所在句论点的关联**。
本脚本补上这一环：

  · 对正文每个引用点 [n]（范围式 [2-5] 展开为多目标、列举式 [2,3] 逐条、
    混合式 [1-3,7] 跳过并明示），取「所在句 + 前一句」作为论点上下文；
  · 从文献表第 n 条提取题名主题词（中文题名 n-gram + 英文题名实意词），
    复用 argument_check 的主题词机制（import，不复制代码）；
  · 关联度评分：论点上下文主题词 ↔ 文献题名主题词的重叠分
    （同词 1.0、互为子串 0.5）；分档：
      🔴 零关联（0 分）  —— 论点与文献题名无任何主题词交集
      🟡 弱关联（<1）   —— 有交集但很弱，需人工确认
      ✅ 有关联（≥1）
  · 规范挂靠识别：「某某[n]认为/指出/发现/表明」式引用算显性支撑证据，
    在输出中标注，避免把规范写法误判为问题。

用法:
    python citation_quality.py <论文.docx> [--top 20] [--json]
退出码：0 = 无 🔴 零关联引用点；1 = 存在零关联引用点（🟡 不阻断）。

局限（词面启发式，不是语义理解）：
  · 引用支撑的是**方法/数据/结论**而非题名字面时（如引「SUS 量表原文」支撑
    「应采用 SUS」），题名词与论点词可能不重叠——故 🔴 也只是「待核」，
     最终以人工判断为准；
  · 英文文献题名的实意词集合与中文论点本就少交集，跨语种引用系统性偏低，
    输出按语种分列，中文条目从严、英文条目只报 🔴；
  · 「某某[n]指出…」句式本身已证明作者挂靠意图，仍可能装饰性，
    工具只标注不豁免。
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
import corpus  # noqa: E402
import diagnose_paper as dp  # noqa: E402
import argument_check as ac  # noqa: E402  # 复用主题词机制

# 引用点正则：单个 [n] / 范围式 [2-5] / 列举式 [2,3] / 混合式 [1-3,7]（该刊不用但需识别跳过）
CITE_RE = re.compile(r"\[(\d{1,3})(?:\s*[-–—]\s*(\d{1,3}))?(?:\s*[,，]\s*(\d{1,3}))*\]")

# 规范挂靠：引用前有作者名 + 认知动词
ATTACH_RE = re.compile(r"[\u4e00-\u9fffA-Za-z]{2,12}(?:等)?\s*\[(?:\d[^\]]*)\]\s*"
                       r"(?:认为|指出|发现|表明|提出|强调|建议|构建|设计|验证|显示|综述|分析)")

# 英文停用词（题名实意词过滤）
EN_STOP = {"the", "and", "for", "with", "from", "into", "of", "in", "on", "to",
           "a", "an", "by", "or", "as", "at", "is", "are", "was", "were", "based",
           "study", "research", "analysis", "using", "towards", "between", "how"}


def parse_citations(text):
    """从一段文本提取引用点列表：[(原文字符串, [目标文献号...], 是否混合式)]。"""
    out = []
    for m in CITE_RE.finditer(text):
        raw = m.group(0)
        inner = raw[1:-1]
        if "," in inner or "，" in inner:
            parts = re.split(r"[,，]", inner)
            nums = []
            mixed = False
            for p in parts:
                p = p.strip()
                rm = re.match(r"^(\d{1,3})\s*[-–—]\s*(\d{1,3})$", p)
                if rm:
                    a, b = int(rm.group(1)), int(rm.group(2))
                    if b >= a and b - a <= 30:
                        nums.extend(range(a, b + 1))
                        mixed = True     # 含范围段 → 混合式
                    else:
                        nums.append(a)
                elif p.isdigit():
                    nums.append(int(p))
            if not nums:
                continue
            out.append((raw, nums, mixed))
        elif "-" in inner or "–" in inner or "—" in inner:
            rm = re.match(r"^\s*(\d{1,3})\s*[-–—]\s*(\d{1,3})\s*$", inner)
            if rm:
                a, b = int(rm.group(1)), int(rm.group(2))
                if b >= a and b - a <= 30:
                    out.append((raw, list(range(a, b + 1)), False))
                else:
                    out.append((raw, [a], False))
            # 反向范围（如 [5-2]）忽略
        else:
            out.append((raw, [int(inner)], False))
    return out


def ref_title_terms(ref_text, vocab=None):
    """从一条文献条目提取题名主题词集合（中文块切分 + 英文实意词）。"""
    # 取「作者. 题名[类型]」中的题名段
    m = re.match(r"^[\[［]?\s*\d{1,3}\s*[\]］]?\s*(?:.+?)[\.．]\s*(.+?)\s*\[(?:[A-Z/]+)\]",
                 ref_text)
    title = m.group(1) if m else ref_text
    title = re.sub(r"^[\[［]?\s*\d{1,3}\s*[\]］]?\s*", "", title)
    terms = _terms_from_blocks(re.findall(r"[\u4e00-\u9fff]+", title), vocab)
    # 英文部分：实意词（≥4 字母、非停用）
    for w in re.findall(r"[A-Za-z][A-Za-z\-]{3,}", title):
        if w.lower() not in EN_STOP:
            terms.add(w.lower())
    return terms


def _terms_from_blocks(blocks, vocab=None):
    """从若干中文块提取主题词：语料真实词表优先，块内 n-gram 兜底。

    vocab: 语料库高频术语词表（set）——命中词表的切片直接作为主题词，
    避免 n-gram 跨词边界产生「视障老年」这类粘接碎片。
    块内 n-gram 只保留不被更长候选包含的片段，且不跨块切分。
    """
    terms = set()
    for blk in blocks:
        if len(blk) < 2:
            continue
        cands = []
        if vocab:
            # 词表优先：长词先匹配，命中区域从块中剔除，避免重复计数
            rest = blk
            for w in sorted(vocab, key=len, reverse=True):
                if len(w) < 2 or w not in rest:
                    continue
                terms.add(w)
                rest = rest.replace(w, "\x00" * len(w))
            blk = rest.replace("\x00", "")
            if len(blk) < 2:
                continue
        for n in (4, 3, 2):
            for i in range(len(blk) - n + 1):
                w = blk[i:i + n]
                if any(c in ac.STOP_CHARS for c in w):
                    continue
                if w in ac.STOPWORDS:
                    continue
                cands.append(w)
        kept = [w for w in cands if not any(w != w2 and w in w2 for w2 in cands)]
        terms.update(kept)
    return terms


def arg_terms(para_text, vocab=None):
    """论点上下文的主题词集合（按中文块独立切分 + 语料词表优先）。"""
    blocks = re.findall(r"[\u4e00-\u9fff]+", para_text)
    return _terms_from_blocks(blocks, vocab)


def score_pair(arg_t, ref_t):
    """主题词重叠分：同词 1.0；互为子串 0.5（只计一次）。"""
    score, matched = 0.0, []
    used = set()
    for a in sorted(arg_t, key=len, reverse=True):
        for b in sorted(ref_t, key=len, reverse=True):
            if (a, b) in used:
                continue
            if a == b:
                score += 1.0
                matched.append(a)
                used.add((a, b))
                break
            if len(a) >= 2 and len(b) >= 2 and (a in b or b in a):
                score += 0.5
                matched.append(f"{a}≈{b}")
                used.add((a, b))
                break
    return score, matched


def is_attached(sent):
    """句内是否有规范挂靠形态。"""
    return bool(ATTACH_RE.search(sent))


def build_vocab(corpus_dir=None):
    """从语料库构建主题词表：108 篇 terms 字段中出现篇数 ≥20 的词。

    词表用于把「视障、老年」正确切成「视障」「老年」，而不是 n-gram 粘接碎片。
    语料库缺失时返回 None（工具降级为纯 n-gram，仍可用）。
    """
    try:
        rows = corpus.load_corpus_stats(corpus_dir)
    except Exception:
        return None
    if not rows:
        return None
    from collections import Counter
    cnt = Counter()
    for r in rows:
        for t in (r.get("terms") or {}):
            if isinstance(t, str) and len(t) >= 2:
                cnt[t] += 1
    vocab = {t for t, n in cnt.items() if n >= 20 and len(t) <= 8}
    return vocab or None


def audit(doc_path, top=20, vocab=None):
    paras, refs, abs_text, kw_text, titles, body = dp.extract_paras(doc_path)
    # 文献号 → 条目文本
    ref_map = {}
    for r in refs:
        m = re.match(r"^[\[［]?\s*(\d{1,3})", r)
        if m:
            ref_map[int(m.group(1))] = r

    # 论点上下文 = 所在句 + 前一句（按段聚合后切句）
    rows = []
    skipped_mixed = 0
    missing_ref = 0
    for pi, p in enumerate(body):
        t = dp.clean(p or "")
        if not t or "[" not in t:
            continue
        sents = [s for s in re.split(r"(?<=[。？！])", t) if s.strip()]
        cites = parse_citations(t)
        if not cites:
            continue
        for raw, nums, mixed in cites:
            if mixed:
                skipped_mixed += 1
                continue
            # 定位引用点所在句（enumerate 取真实下标，避免 sents.index
            # 在重复句上取错索引导致论点句错位）
            pos = t.find(raw)
            sent = ""
            prev = ""
            acc = 0
            for si, s in enumerate(sents):
                if acc <= pos < acc + len(s):
                    sent = s.strip()
                    if si > 0:
                        prev = sents[si - 1].strip()
                    break
                acc += len(s)
            if not sent:
                sent = t[max(0, pos - 40):pos + len(raw) + 20]
            ctx = (prev + sent) if prev else sent
            ctx_t = arg_terms(ctx, vocab)
            for n in nums:
                rt = ref_map.get(n)
                if rt is None:
                    missing_ref += 1
                    continue
                ref_t = ref_title_terms(rt, vocab)
                sc, matched = score_pair(ctx_t, ref_t)
                lang = "EN" if re.search(r"[A-Za-z]{4,}", rt) and not re.search(
                    r"[\u4e00-\u9fff]", ref_title_text(rt)) else "CN"
                rows.append({
                    "para": pi + 1,
                    "cite": raw,
                    "ref_no": n,
                    "ref_text": rt,
                    "ref_lang": lang,
                    "sent": sent[:110],
                    "score": round(sc, 2),
                    "matched": matched[:6],
                    "attached": is_attached(sent),
                })

    # 分档：🔴 零关联（中文从严；英文只报零关联）｜🟡 弱关联（中文 <1）
    reds, yellows, oks = [], [], []
    for r in rows:
        if r["score"] == 0:
            if r["ref_lang"] == "EN" or not r["attached"]:
                reds.append(r)
            else:
                reds.append(r)     # 中文零关联一律报（挂靠形态只作证据标注）
        elif r["score"] < 1:
            if r["ref_lang"] == "CN":
                yellows.append(r)
            else:
                oks.append(r)
        else:
            oks.append(r)
    reds.sort(key=lambda r: r["para"])
    yellows.sort(key=lambda r: r["para"])
    return {
        "rows": rows, "reds": reds, "yellows": yellows, "oks": oks,
        "n_cites": len(rows), "skipped_mixed": skipped_mixed,
        "missing_ref": missing_ref,
    }


def ref_title_text(ref_text):
    m = re.match(r"^[\[［]?\s*\d{1,3}\s*[\]］]?\s*(?:.+?)[\.．]\s*(.+?)\s*\[(?:[A-Z/]+)\]",
                 ref_text)
    return m.group(1) if m else ref_text


def main():
    args = [a for a in sys.argv[1:] if not a.startswith("--")]
    flags = set(a for a in sys.argv[1:] if a.startswith("--"))
    if not args or args[0] in ("-h", "--help"):
        print(__doc__)
        sys.exit(0)
    DOC = args[0]
    if not os.path.isfile(DOC):
        sys.exit(f"找不到文件：{DOC}")
    TOP = 20
    if "--top" in sys.argv:
        i = sys.argv.index("--top")
        try:
            TOP = int(sys.argv[i + 1])
        except (IndexError, ValueError):
            pass
    _corpus_arg = None
    if "--corpus" in sys.argv:
        i = sys.argv.index("--corpus")
        _corpus_arg = sys.argv[i + 1] if i + 1 < len(sys.argv) else None

    try:
        _cdir = corpus.find_corpus_dir(_corpus_arg)
    except FileNotFoundError as e:
        print(f"[错误] {e}")
        sys.exit(2)
    vocab = build_vocab(_cdir)

    res = audit(DOC, top=TOP, vocab=vocab)
    n = res["n_cites"]
    print("=" * 70)
    print("引文—论点关联核验（装饰性引用专查）")
    print(f"稿件：{os.path.basename(DOC)}")
    if vocab:
        print(f"主题词表：语料库高频术语 {len(vocab)} 词（108 篇 terms 字段，出现 ≥20 篇）")
    else:
        print("主题词表：（未找到语料库，降级为纯 n-gram 切分）")
    print(f"可评引用点 {n} 个（混合式跳过 {res['skipped_mixed']}，"
          f"文献表缺号 {res['missing_ref']}）")
    print("=" * 70)
    if n == 0:
        print("正文未检出引用点，或文献表不可解析。")
        sys.exit(0)

    print(f"分档：🔴 零关联 {len(res['reds'])} ｜ 🟡 弱关联 {len(res['yellows'])} ｜ "
          f"✅ 有关联 {len(res['oks'])}")
    print("-" * 70)
    shown = 0
    for r in (res["reds"] + res["yellows"]):
        if shown >= TOP:
            print(f"……（其余 {len(res['reds']) + len(res['yellows']) - shown} 条见 --json）")
            break
        shown += 1
        mark = "🔴" if r in res["reds"] else "🟡"
        att = "｜有规范挂靠「…[n]指出」" if r["attached"] else "｜无挂靠动词"
        print(f"{mark} 段{r['para']} {r['cite']} → 文献[{r['ref_no']}]"
              f"（{r['ref_lang']}·关联 {r['score']}）{att}")
        print(f"    论点句：{r['sent']}")
        print(f"    文献题名：{ac.trunc(r['ref_text'], 66)}")
        if r["matched"]:
            print(f"    命中主题词：{'、'.join(r['matched'])}")
        else:
            print(f"    命中主题词：（无交集）")
    print("-" * 70)
    if res["reds"]:
        print(f"🔴 零关联引用点 {len(res['reds'])} 个——论点与文献题名无任何主题词交集，")
        print("   请逐条确认：该文献是否真支撑本句？若否，改挂到相关论点句或删除。")
        sys.exit(1)
    if res["yellows"]:
        print(f"🟡 弱关联引用点 {len(res['yellows'])} 个——有交集但弱，人工复核是否装饰性引用。")
    else:
        print("✅ 未检出零关联/弱关联引用点。")
    sys.exit(0)


if __name__ == "__main__":
    main()
