# -*- coding: utf-8 -*-
"""citation_distribution.py —— 引用密度分布核验（「结构性失引」专查）

diagnose_paper 只报**全篇平均**引用密度（如 2.31/千字 vs 刊中位 3.23），
但平均值会掩盖一种更严重的问题：**结构性失引**——引用全堆在引言与文献综述，
论证章、结果章、对策章几乎没有引用。审稿人常因此质疑「后半部分缺文献支撑」
「对策凭经验提出而非基于研究证据」。

本脚本把引用密度拆到**逐章 + 逐段**两个层面看分布：

  · 章节级：按一级标题切章，统计每章引用密度（处/千字）与零引用段占比；
  · 分布形态：识别「前置集中型」（前 2 章占全部引用的 ≥70%）、
    「尾部空窗」（末 40% 段落零引用）等模式；
  · 段级空窗：连续 ≥N 段（默认 8 段）无任何引用即报「论证空窗」，
    并定位起止段号与所属章节；
  · 对照该刊基线：全篇密度对照分栏目基线，章节级只做形态判断
    （该刊各章引用密度方差大，不设硬阈值，只报形态异常）。

用法:
    python citation_distribution.py <论文.docx> [--corpus <dir>] [--gap 8] [--json]
退出码：0 = 分布形态正常；1 = 检出结构性失引（前置集中或尾部大空窗）。

局限（按段位切分的形态判断，不做内容相关性评估）：
  · 该刊确实存在「结果章少引用」的合理情形（自建数据本身即证据），
    故章级只报形态、不判对错，由作者判断是否属合理情形；
  · 引用的**质量**（是否装饰性）由 citation_quality.py 核验，本脚本只看数量分布。
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
import corpus  # noqa: E402
import diagnose_paper as dp  # noqa: E402

CITE_RE = re.compile(r"\[\d{1,3}(?:\s*[-–—,，]\s*\d{1,3})*\]")
MIN_PARA = 40          # 参与统计的最短段（字）
FRONT_RATIO = 0.70     # 前 2 章占全部引用 ≥70% → 前置集中
GAP_DEFAULT = 8        # 连续零引用段数阈值


def cjk_len(s):
    return len(dp.clean(s))


def build_chapters(paras, titles):
    """按**一级标题**切章：返回 [{name, paras:[(段号,文本)]}]。

    只在一级（0/1/2/3… 或章级命名）切，二级及以下标题留在段序列里——
    否则「1 研究对象与方法」会被它的 1.1/1.2 拆成多个"章"，章节级密度
    统计单元被打碎，空窗判断也会失真。
    """
    tset = {re.sub(r"\s+", " ", t.strip()) for t in titles}
    chapters = []
    cur = {"name": "（章前内容）", "paras": []}
    for idx, (style, t) in enumerate(paras):
        if not t:
            continue
        if t.startswith("参考文献"):
            break
        s = re.sub(r"\s+", " ", t.strip())
        # 判定是否一级标题：数字编号无点号，或纯章名（引言/结语/结论）
        lvl1 = bool(re.match(r"^\d+\s*[^\d.]", s)) or \
               bool(re.match(r"^\s*(引\s*言|结\s*语|结\s*论|结束语|余\s*论)\s*$", s))
        is_heading = (s in tset) or (len(s) < 40 and (
            lvl1 or style.startswith("Heading")))
        if is_heading and lvl1:
            if cur["paras"]:
                chapters.append(cur)
            cur = {"name": s, "paras": []}
            continue
        if is_heading:
            # 二级及以下标题：不切章（保留在章内，不入正文统计）
            continue
        # 跳过摘要/关键词段（中英文均跳过；英文段落长，误入会拉低全篇密度）
        if re.search(r"^【?摘\s*要】?[：:]?", s[:12]) or \
           re.search(r"^【?关\s*键\s*词】?[：:]?", s[:12]) or \
           re.search(r"^(?:Abstract|ABSTRACT)\s*[:：]", s[:20]) or \
           re.search(r"^(?:Key\s*words?)\s*[:：]", s[:20], re.I):
            continue
        # 题名/作者行（无中文实义且短，属正文前元数据）
        if cjk_len(s) <= 0 and len(s) < 40:
            continue
        cur["paras"].append((idx + 1, t))
    if cur["paras"]:
        chapters.append(cur)
    # 剔除无内容的空章（如纯元数据的「章前内容」）
    chapters = [c for c in chapters if c["paras"] and
                any(cjk_len(t) >= MIN_PARA for _, t in c["paras"])]
    return chapters


def analyze(doc, gap_n=GAP_DEFAULT):
    paras, refs, abs_text, kw_text, titles, body = dp.extract_paras(doc)
    chapters = build_chapters(paras, titles)

    total_cite, total_chars = 0, 0
    ch_rows = []
    for ch in chapters:
        c_sum, n_sum, n_par, n_zero = 0, 0, 0, 0
        for pno, txt in ch["paras"]:
            n = cjk_len(txt)
            if n < MIN_PARA:
                continue
            c = len(CITE_RE.findall(txt))
            c_sum += c
            n_sum += n
            n_par += 1
            if c == 0:
                n_zero += 1
        total_cite += c_sum
        total_chars += n_sum
        ch_rows.append({
            "name": ch["name"], "paras": n_par, "chars": n_sum, "cites": c_sum,
            "rate": round(c_sum / n_sum * 1000, 2) if n_sum else 0.0,
            "zero_paras": n_zero,
            "zero_ratio": round(n_zero / n_par, 2) if n_par else 0.0,
        })

    # 前置集中：前 2 个有内容章占全部引用比例
    real_chs = [r for r in ch_rows if r["paras"] >= 1]
    front = real_chs[:2]
    front_c = sum(r["cites"] for r in front)
    front_ratio = round(front_c / total_cite, 2) if total_cite else 0.0

    # 段级连续零引用空窗
    seq = []
    for ch in chapters:
        for pno, txt in ch["paras"]:
            if cjk_len(txt) < MIN_PARA:
                continue
            seq.append((pno, ch["name"], len(CITE_RE.findall(txt))))
    gaps, run = [], []
    for pno, cname, c in seq:
        if c == 0:
            run.append((pno, cname))
        else:
            if len(run) >= gap_n:
                gaps.append({"from": run[0][0], "to": run[-1][0],
                             "n": len(run), "chapter": run[0][1]})
            run = []
    if len(run) >= gap_n:
        gaps.append({"from": run[0][0], "to": run[-1][0],
                     "n": len(run), "chapter": run[0][1]})

    # 尾部空窗：末 40% 段落中零引用占比
    tail = seq[int(len(seq) * 0.6):] if len(seq) >= 5 else []
    tail_zero = sum(1 for _, _, c in tail if c == 0)
    tail_ratio = round(tail_zero / len(tail), 2) if tail else 0.0

    overall = round(total_cite / total_chars * 1000, 2) if total_chars else 0.0
    return {
        "chapters": ch_rows, "overall_rate": overall,
        "total_cites": total_cite, "total_chars": total_chars,
        "front_ratio": front_ratio, "gaps": gaps, "tail_zero_ratio": tail_ratio,
        "n_paras": len(seq),
    }


def main():
    args = [a for a in sys.argv[1:] if not a.startswith("--")]
    flags = set(a for a in sys.argv[1:] if a.startswith("--"))
    if not args or args[0] in ("-h", "--help"):
        print(__doc__)
        sys.exit(0)
    DOC = args[0]
    if not os.path.isfile(DOC):
        sys.exit(f"找不到文件：{DOC}")
    gap_n = GAP_DEFAULT
    for i, a in enumerate(sys.argv):
        if a == "--gap" and i + 1 < len(sys.argv):
            try:
                gap_n = max(3, int(sys.argv[i + 1]))
            except ValueError:
                pass
    _corpus_arg = None
    if "--corpus" in sys.argv:
        i = sys.argv.index("--corpus")
        _corpus_arg = sys.argv[i + 1] if i + 1 < len(sys.argv) else None

    res = analyze(DOC, gap_n)
    try:
        cdir = corpus.find_corpus_dir(_corpus_arg)
    except FileNotFoundError:
        cdir = None
    base = corpus.load_reference_stats(cdir) if cdir else None
    base_rate = (base or {}).get("cites", {}).get("med")

    print("=" * 72)
    print("引用密度分布核验（结构性失引专查）")
    print(f"稿件：{os.path.basename(DOC)}")
    print(f"全篇引用密度 {res['overall_rate']:.2f}/千字（{res['total_cites']} 处 / "
          f"{res['total_chars']} 字）"
          + (f"｜刊中位 {base_rate:.2f}" if base_rate else ""))
    print(f"统计 {res['n_paras']} 段（≥{MIN_PARA} 字）")
    print("=" * 72)

    print("【章节级分布】")
    print(f"  {'章节':<26}{'段':>4}{'字':>7}{'引用':>5}{'密度':>7}{'零引用段':>9}")
    for r in res["chapters"]:
        name = r["name"][:24]
        print(f"  {name:<26}{r['paras']:>4}{r['chars']:>7}{r['cites']:>5}"
              f"{r['rate']:>7.2f}{r['zero_paras']:>6}({r['zero_ratio']:.0%})")
    print("-" * 72)

    issues, notices = [], []
    # ① 前置集中
    if res["total_cites"] >= 5 and res["front_ratio"] >= FRONT_RATIO:
        issues.append(f"引用前置集中：前 2 章占全部引用的 {res['front_ratio']:.0%}"
                      f"（阈值 {FRONT_RATIO:.0%}）——后半部分缺文献支撑，"
                      f"审稿人常质疑「对策凭经验而非证据」")
    # ② 连续空窗
    for g in res["gaps"]:
        issues.append(f"论证空窗：第 {g['from']}–{g['to']} 段连续 {g['n']} 段零引用"
                      f"（{g['chapter'][:18]}）")
    # ③ 尾部空窗
    if res["tail_zero_ratio"] >= 0.85 and res["total_cites"] >= 5:
        issues.append(f"尾部失引：末 40% 段落中 {res['tail_zero_ratio']:.0%} 为零引用")
    # ④ 全篇密度
    if base_rate:
        if res["overall_rate"] < base_rate * 0.6:
            issues.append(f"全篇引用密度 {res['overall_rate']:.2f} 显著低于刊中位 "
                          f"{base_rate:.2f}（<60%）")
        elif res["overall_rate"] > base_rate * 2.5:
            notices.append(f"全篇引用密度 {res['overall_rate']:.2f} 远高于刊中位 "
                           f"{base_rate:.2f}（>2.5 倍），核对是否堆砌引用")

    if issues:
        print(f"🔴 发现 {len(issues)} 项结构性失引：")
        for it in issues:
            print(f"   · {it}")
        print("\n  提示：结果章/系统实现章少引用属合理情形（自建数据本身即证据），")
        print("       但论证章与对策章应有文献支撑——请按上表定位并补充关键引用。")
    else:
        print("✅ 引用分布形态：未见结构性失引")
    if notices:
        print(f"\n🟡 建议核对 {len(notices)} 项：")
        for it in notices:
            print(f"   · {it}")

    if "--json" in flags:
        jf = os.path.splitext(DOC)[0] + ".citation_distribution.json"
        with open(jf, "w", encoding="utf-8") as f:
            json.dump(res, f, ensure_ascii=False, indent=1)
        print(f"\nJSON 已写入：{jf}")

    sys.exit(1 if issues else 0)


if __name__ == "__main__":
    main()
