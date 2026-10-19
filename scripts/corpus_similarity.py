# -*- coding: utf-8 -*-
"""corpus_similarity.py —— 与已见刊论文的句面相合度预检（外部查重前哨）

官方硬门槛：「总文字复制比 < 7%」「去除引用文献复制比接近 0」，未通过不予排版。
CNKI 查重报告投稿前必须备一份，但**在校对阶段发现问题再改，比重查到再改从容得多**。

本脚本不替代 CNKI（那需要其授权库），只做**本地可做的预检**：
把稿件句子与该刊 118 篇见刊论文（121 万字）的句子做字符 shingle 比对，
找出「与已刊论文高度相似的句子」——这些是查重系统最可能标红的地方。

  · 索引：118 篇全部段落切句，建 3-gram 倒排索引（一次建好，内存内复用）；
  · 匹配：稿件每个句子先查索引找候选（同 shingle 数 ≥ 阈值），再算 Jaccard；
  · 分档：
      🔴 ≥0.75  高度相合——查重极可能标红，必须改写或删除
      🟡 0.55–0.75 中度相合——可能被计入复制比，建议改写
      ✅ <0.55  可接受（术语/标准表述的少量重合属正常）
  · 输出：逐条列出相似句、命中的见刊论文（篇号|年期|作者）、相似度、
     该刊原文片段，供直接对照改写。

用法:
    python corpus_similarity.py <论文.docx> [--corpus <dir>] [--top 20] [--json]
退出码：0 = 无高度相合句；1 = 检出 🔴 高度相合句。

局限（字符 shingle 相似度，不是语义/查重系统口径）：
  · CNKI 的库比本刊 118 篇大得多（覆盖期刊/学位论文/会议/网络资源），
    本工具的「未检出」不等于「查重能过」——真正的底线是 CNKI 报告；
  · 同义改写检不出；标准条文、法律名称、通用术语的重复属合理，
    脚本已在报告中标明命中篇目，由你判断是否属「必要引用式重复」；
  · 只比**见刊论文**，不覆盖学位论文与网络资源，这是该刊最易撞车的两块。
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

SHINGLE_N = 3
MIN_SENT_LEN = 16          # 参与比对的句长下限（短句无判别力）
CANDIDATE_MIN = 2          # 候选句须共用的最少 shingle 数
RED_THRESH = 0.75
YELLOW_THRESH = 0.55
MAX_CAND = 300             # 单句候选上限（防热门 shingle 爆炸）


def normalize(s):
    t = dp.clean(s or "")
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
    t = re.sub(r"\[\d{1,3}(?:\s*[-–—,，]\s*\d{1,3})*\]", "", t)
    t = re.sub(r"[^\w\u4e00-\u9fff]", "", t)
    return t.lower()


def shingles(s, n=SHINGLE_N):
    if len(s) < n:
        return set()
    return {s[i:i + n] for i in range(len(s) - n + 1)}


def sentences_of(text):
    return [s.strip() for s in re.split(r"(?<=[。？！])", text or "") if s.strip()]


def build_index(corpus_dir):
    """建 shingle 倒排索引。返回 (index, sentences)。

    index: {shingle: [句子下标...]}
    sentences: [{norm, sh, no, issue, author, para_no, raw}]
    """
    papers = corpus.load_paragraphs(corpus_dir)
    index = defaultdict(list)
    sentences = []
    for paper in papers:
        no = paper.get("no", "?")
        issue = paper.get("issue", "?")
        fname = paper.get("file", "")
        author = os.path.splitext(fname)[0].split("_")[-1] if "_" in fname else ""
        for pi, para in enumerate(paper.get("paras") or []):
            for s in sentences_of(para):
                nz = normalize(s)
                if len(nz) < MIN_SENT_LEN:
                    continue
                sh = shingles(nz)
                if not sh:
                    continue
                si = len(sentences)
                sentences.append({
                    "norm": nz, "sh": sh, "no": no, "issue": issue,
                    "author": author, "para_no": pi + 1, "raw": s,
                })
                for g in sh:
                    index[g].append(si)
    return index, sentences


def match_sentence(nz, sh, index, sentences):
    """为稿件句子找最相似的见刊句。返回 (score, 命中句 or None)。"""
    if not sh:
        return 0.0, None
    # 候选：按共用 shingle 数计数
    cand_count = defaultdict(int)
    for g in sh:
        lst = index.get(g)
        if not lst:
            continue
        if len(lst) > MAX_CAND:
            # 过热 shingle（如「的研」「图书馆」）跳过，避免拖慢且无判别力
            continue
        for si in lst:
            cand_count[si] += 1
    if not cand_count:
        return 0.0, None
    # 取共用最多的前 50 个候选
    cands = sorted(cand_count.items(), key=lambda x: -x[1])[:50]
    best_score, best = 0.0, None
    la = len(nz)
    for si, _ in cands:
        cand = sentences[si]
        if cand_count[si] < CANDIDATE_MIN:
            continue
        lb = len(cand["norm"])
        if min(la, lb) / max(la, lb) < 0.4:
            continue
        inter = len(sh & cand["sh"])
        union = len(sh | cand["sh"])
        jac = inter / union if union else 0.0
        if jac > best_score:
            best_score, best = jac, cand
    return best_score, best


def analyze(doc, corpus_dir=None, top=20):
    paras, refs, abs_text, kw_text, titles, body = dp.extract_paras(doc)
    index, sentences = build_index(corpus_dir)

    # 稿件句子：摘要 / 正文（按章标注）/ 结语
    doc_sents = []
    if abs_text:
        for s in sentences_of(abs_text):
            doc_sents.append(("摘要", s))
    cur_ch = "正文"
    for style, t in paras:
        if not t:
            continue
        if t.startswith("参考文献"):
            break
        s = re.sub(r"\s+", " ", t.strip())
        if len(s) < 40 and ((re.match(r"^(0|1|2|3|4|5|6|7)\s*[^\d]", s))
                            or style.startswith("Heading")):
            cur_ch = s[:22]
            continue
        if re.search(r"^【?摘\s*要】?[：:]?", s[:12]) or \
           re.search(r"^【?关\s*键\s*词】?[：:]?", s[:12]) or \
           re.search(r"^(?:Abstract|Key\s*words?)\s*[:：]", s[:20], re.I):
            continue
        for x in sentences_of(t):
            doc_sents.append((cur_ch, x))

    hits = []
    n_compared = 0
    for region, raw in doc_sents:
        nz = normalize(raw)
        if len(nz) < MIN_SENT_LEN:
            continue
        n_compared += 1
        sh = shingles(nz)
        score, best = match_sentence(nz, sh, index, sentences)
        if score >= YELLOW_THRESH and best:
            hits.append({
                "region": region, "score": round(score, 3),
                "doc_sent": raw,
                "corpus_no": best["no"], "issue": best["issue"],
                "author": best["author"], "para_no": best["para_no"],
                "corpus_sent": best["raw"],
            })
    hits.sort(key=lambda h: -h["score"])
    return {"hits": hits, "n_compared": n_compared,
            "n_corpus_sents": len(sentences),
            "red": sum(1 for h in hits if h["score"] >= RED_THRESH),
            "yellow": sum(1 for h in hits if YELLOW_THRESH <= h["score"] < RED_THRESH)}


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
    TOP = 20
    for i, a in enumerate(sys.argv):
        if a == "--top" and i + 1 < len(sys.argv):
            try:
                TOP = int(sys.argv[i + 1])
            except ValueError:
                pass
    _corpus_arg = None
    if "--corpus" in sys.argv:
        i = sys.argv.index("--corpus")
        _corpus_arg = sys.argv[i + 1] if i + 1 < len(sys.argv) else None
    try:
        cdir = corpus.find_corpus_dir(_corpus_arg)
    except FileNotFoundError as e:
        print(f"[错误] {e}")
        sys.exit(2)

    res = analyze(DOC, cdir, top=TOP)
    print("=" * 74)
    print("与已见刊论文的句面相合度预检（外部查重前哨）")
    print(f"稿件：{os.path.basename(DOC)}")
    print(f"比对：稿件 {res['n_compared']} 句 ↔ 见刊论文 {res['n_corpus_sents']} 句"
          f"（{len(corpus.load_paragraphs(cdir))} 篇）")
    print("=" * 74)

    if not res["hits"]:
        print(f"✅ 未检出相似度 ≥{YELLOW_THRESH} 的句子。")
        print("   （注意：本工具只比本刊 118 篇，不覆盖 CNKI 全库；")
        print("     「未检出」≠「查重能过」，底线仍是 CNKI 报告）")
        sys.exit(0)

    print(f"分档：🔴 高度相合 ≥{RED_THRESH}：{res['red']} 句 ｜ "
          f"🟡 中度相合 {YELLOW_THRESH}–{RED_THRESH}：{res['yellow']} 句")
    print("-" * 74)
    for i, h in enumerate(res["hits"][:TOP], 1):
        mark = "🔴" if h["score"] >= RED_THRESH else "🟡"
        src = f"[{h['corpus_no']}|{h['issue']}|{h['author']}]"
        print(f"{mark} {i}. 相似度 {h['score']:.2f}（{h['region']}）命中 {src} 段{h['para_no']}")
        print(f"    稿件：{trunc(h['doc_sent'], 70)}")
        print(f"    刊文：{trunc(h['corpus_sent'], 70)}")
    print("-" * 74)
    if res["red"]:
        print(f"🔴 {res['red']} 句高度相合——查重极可能标红，必须改写或删除。")
        print("   注意：标准条文、法律名称、通用术语的重复属合理必要，")
        print("         「命中篇目」供你判断是否属这类；不确定的优先改写。")
        sys.exit(1)
    print(f"🟡 {res['yellow']} 句中度相合——可能被计入复制比，建议改写。")
    sys.exit(1 if res["yellow"] else 0)


if __name__ == "__main__":
    main()
