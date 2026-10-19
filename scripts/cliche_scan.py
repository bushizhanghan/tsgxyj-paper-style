# -*- coding: utf-8 -*-
"""cliche_scan.py —— 学术套话全表扫描（该刊真实次数 vs 本文次数）

用途
----
把 anti-patterns.md 的 A/B 两级清单做成**全文逐词扫描 + 对照表**，
每个词同时给出「该刊 121 万字实测次数」与「本文次数」，并标出命中位置
（段号 + 原句片段），便于直接定位删除或改写。

  · **A 级红线**（该刊 0–2 次）：显而易见 / 众所周知 / 毋庸置疑 / 总而言之 /
    不难看出 / 值得一提 / 总的来说 / 首先需要指出 / 重要指导意义 /
    值得深入探讨 / 提供有益参考。出现 1 次即按该刊口径判问题。
  · **B 级控密度**（该刊也用，AI 稿常超量 5–10 倍）：具有重要意义（8 次）/
    深远影响（3 次）/ 三连×化排比（16 处）/「不仅…而且」（10 处，
    该刊主流是「不仅…更」122 处 /「不仅…还」106 处）/ 综上所述（23 处）/
    在此基础上（66 处）/ 与此同时（36 处）。
  · B 级按**密度倍数**判：>3 倍该刊单位篇幅基准 → 🔴；1.5–3 倍 → 🟡。
    （该刊基准按 121 万字换算为「次/万字」，与本文按同口径比较。）

复用 language_audit 的句子切分逻辑（该模块 import 时即解析 sys.argv，
故本脚本复刻其 `sentences_with_pos` 切句实现而非 import，避免抢跑命令行）。

用法:
    python cliche_scan.py <论文.docx> [--level A|B|ALL] [--top 12] [--json]

退出码：0 = 无 A 级命中且 B 级未超 3 倍；1 = 检出 A 级命中或 B 级 >3 倍。
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
import corpus  # noqa: E402
import diagnose_paper as dp  # noqa: E402

# language_audit 在 import 时就会解析 sys.argv 并可能 sys.exit()（其 -h/文件
# 不存在分支），直接 import 会抢跑本脚本。故不 import 它，改为复刻其
# sentences_with_pos 的切句逻辑（见 _sentences_with_pos，与源实现等价）。
_SENT_SPLIT = re.compile(r"[^。？！]*[。？！]?")


def _sentences_with_pos(para_no, text):
    """返回 (句, 句内偏移) 列表——与 language_audit.sentences_with_pos 等价。"""
    out = []
    pos = 0
    for m in _SENT_SPLIT.finditer(text):
        s = m.group(0)
        if s and len(s.strip()) >= 4:
            out.append((s.strip(), pos))
        pos = m.end()
    return out

# ---- 词表：A 级红线（该刊 0–2 次）----
# (词, 该刊次数, 处理建议)
A_LEVEL = [
    ("显而易见", 0, "删，直接陈述（该刊 0 次）"),
    ("首先需要指出", 0, "删（该刊 0 次）"),
    ("重要指导意义", 0, "改「为××提供参考」并补对象（该刊 0 次）"),
    ("众所周知", 1, "删——该刊不靠共识句铺垫（1 次）"),
    ("毋庸置疑", 1, "删（1 次）"),
    ("总而言之", 1, "改「总体而言」或删（1 次）"),
    ("不难看出", 1, "删，直接给结论（1 次）"),
    ("值得深入探讨", 1, "删（1 次）"),
    ("提供有益参考", 1, "改「提供参考/依据」并补对象（1 次）"),
    ("值得一提", 2, "删或并入正文（2 次）"),
    ("总的来说", 2, "改「总体而言」（该刊 36 处）（2 次）"),
]

# ---- 词表：B 级控密度（该刊次数，按 121 万字基准换算次/万字）----
CORPUS_CHARS = 1210000        # 该刊 118 篇 121 万字正文（polish-lexicon.md 口径）
B_LEVEL = [
    # (词或正则说明, 匹配函数名, 该刊次数, 纪律)
    ("具有重要意义", "plain", 8, "极少独立成句；收束句必须带具体对象"),
    ("深远影响", "plain", 3, "罕用；改具体影响描述"),
    ("综上所述", "plain", 23, "可用，一段最多一次；结语才用"),
    ("在此基础上", "plain", 66, "可用，密度别超（空转连接专项见 conn_density.py）"),
    ("与此同时", "plain", 36, "可用（空转连接专项见 conn_density.py）"),
    ("不仅…而且", "notonly", 10, "该刊主流是「不仅…更」122 处 /「不仅…还」106 处"),
    ("三连×化排比", "triple_hua", 16, "该刊确有，但一段别超过一次"),
]

# B 级判密度阈值（相对该刊单位篇幅基准的倍数）
B_RED_MULT = 3.0
B_YELLOW_MULT = 1.5
MIN_SAMPLE = 800              # 全文不足此字数不判密度


def _count_plain(text, word):
    return text.count(word)


def _count_notonly(text, _):
    return len(re.findall(r"不仅[^。]{2,50}而且", text))


def _count_triple_hua(text, _):
    return len(re.findall(
        r"[\u4e00-\u9fff]{1,4}化[、,，][\u4e00-\u9fff]{1,4}化[、,，][\u4e00-\u9fff]{1,4}化",
        text))


COUNTERS = {"plain": _count_plain, "notonly": _count_notonly,
            "triple_hua": _count_triple_hua}


def _locs(body, word, kind, limit=4):
    """返回命中位置 [(段号, 片段)]。"""
    out = []
    for pi, b in enumerate(body):
        t = dp.clean(b or "")
        if len(t) < 40:
            continue
        if kind == "plain":
            hit = word in t
        elif kind == "notonly":
            hit = bool(re.search(r"不仅[^。]{2,50}而且", t))
        else:
            hit = bool(re.search(
                r"[\u4e00-\u9fff]{1,4}化[、,，][\u4e00-\u9fff]{1,4}化[、,，][\u4e00-\u9fff]{1,4}化",
                t))
        if hit:
            # 片段：取命中处前后文
            if kind == "plain":
                i = t.find(word)
                s = max(0, i - 16)
                out.append((pi + 1, t[s:i + len(word) + 22]))
            else:
                out.append((pi + 1, t[:46]))
        if len(out) >= limit:
            break
    return out


def analyze(doc, corpus_dir=None):
    """全文扫描套话，返回 A/B 两级对照结果。"""
    paras, refs, abs_text, kw_text, titles, body = dp.extract_paras(doc)
    text = dp.clean(abs_text or "") + "".join(dp.clean(b or "") for b in body)
    n = len(text)

    a_rows, b_rows = [], []
    issues, notices = [], []

    for word, c_cnt, advice in A_LEVEL:
        c = _count_plain(text, word)
        a_rows.append({"word": word, "corpus": c_cnt, "doc": c,
                       "locs": _locs(body, word, "plain") if c else [],
                       "advice": advice})
        if c:
            loc = "、".join(f"段{p}" for p, _ in a_rows[-1]["locs"][:4]) or "摘要"
            issues.append(f"A 级红线「{word}」本文 {c} 处（该刊 {c_cnt} 处）"
                          f"[{loc}] → {advice}")

    for word, kind, c_cnt, disc in B_LEVEL:
        c = COUNTERS[kind](text, word)
        rate_doc = c / max(n, 1) * 10000          # 次/万字
        rate_base = c_cnt / CORPUS_CHARS * 10000  # 次/万字
        mult = rate_doc / rate_base if rate_base else 0.0
        tag = "≈"
        if mult > B_RED_MULT:
            tag = "🔴 超 3 倍"
        elif mult > B_YELLOW_MULT:
            tag = "🟡 超 1.5 倍"
        b_rows.append({"word": word, "corpus": c_cnt, "doc": c,
                       "rate_doc": round(rate_doc, 3), "rate_base": round(rate_base, 3),
                       "mult": round(mult, 2), "tag": tag,
                       "locs": _locs(body, word, kind) if c else [],
                       "discipline": disc})
        if n >= MIN_SAMPLE and mult > B_RED_MULT:
            loc = "、".join(f"段{p}" for p, _ in b_rows[-1]["locs"][:4])
            issues.append(f"B 级「{word}」本文 {c} 次 = {rate_doc:.2f}/万字，"
                          f"为该刊基准的 {mult:.1f} 倍 [{loc}] → {disc}")
        elif n >= MIN_SAMPLE and mult > B_YELLOW_MULT:
            notices.append(f"B 级「{word}」{c} 次 = {rate_doc:.2f}/万字"
                           f"（该刊基准的 {mult:.1f} 倍）→ {disc}")

    return {"chars": n, "a_rows": a_rows, "b_rows": b_rows,
            "a_hits": sum(1 for r in a_rows if r["doc"]),
            "issues": issues, "notices": notices}


def main():
    args = [a for a in sys.argv[1:] if not a.startswith("--")]
    flags = set(a for a in sys.argv[1:] if a.startswith("--"))
    if not args or args[0] in ("-h", "--help"):
        print(__doc__)
        sys.exit(0)

    LEVEL = "ALL"
    if "--level" in sys.argv:
        i = sys.argv.index("--level")
        if i + 1 < len(sys.argv):
            LEVEL = sys.argv[i + 1].upper()
    TOP = 12
    for i, a in enumerate(sys.argv):
        if a == "--top" and i + 1 < len(sys.argv):
            try:
                TOP = int(sys.argv[i + 1])
            except (ValueError, IndexError):
                pass

    DOC = args[0]
    if not os.path.isfile(DOC):
        sys.exit(f"找不到文件：{DOC}")

    res = analyze(DOC)
    print("=" * 78)
    print("学术套话全表扫描（该刊 121 万字实测次数 vs 本文次数）")
    print(f"稿件：{os.path.basename(DOC)}")
    print(f"统计口径：摘要+正文共 {res['chars']} 字（clean 后）")
    print("=" * 78)

    if LEVEL in ("ALL", "A"):
        print("\n【A 级红线】该刊 0–2 次，出现即按该刊口径处理")
        print(f"  {'表达':<10}{'该刊':>5}{'本文':>5}  位置与处理")
        for r in res["a_rows"]:
            mark = "🔴" if r["doc"] else "  "
            loc = "、".join(f"段{p}" for p, _ in r["locs"][:3])
            print(f"  {mark}{r['word']:<9}{r['corpus']:>5}{r['doc']:>5}"
                  f"  {loc if loc else '—'}")
            if r["doc"]:
                print(f"      → {r['advice']}")
                for p, frag in r["locs"][:2]:
                    print(f"      段{p}：…{frag}…")

    if LEVEL in ("ALL", "B"):
        print("\n【B 级控密度】该刊也用，AI 稿常超量 5–10 倍（密度倍数判）")
        print(f"  {'表达':<12}{'该刊次数':>7}{'本文':>5}{'本文/万字':>9}"
              f"{'该刊/万字':>9}{'倍数':>7}  判定")
        for r in res["b_rows"]:
            print(f"  {r['word']:<11}{r['corpus']:>7}{r['doc']:>5}"
                  f"{r['rate_doc']:>9.2f}{r['rate_base']:>9.2f}"
                  f"{r['mult']:>7.2f}  {r['tag']}")
        print("\n【B 级命中位置】")
        any_loc = False
        for r in res["b_rows"]:
            if not r["locs"]:
                continue
            any_loc = True
            print(f"  {r['word']}（{r['doc']} 次）：")
            for p, frag in r["locs"][:TOP // 3 + 1]:
                print(f"      段{p}：…{frag}…")
            print(f"      → {r['discipline']}")
        if not any_loc:
            print("  （无命中）")

    print("\n" + "-" * 78)
    if res["issues"]:
        print(f"🔴 发现 {len(res['issues'])} 项套话问题：")
        for it in res["issues"]:
            print(f"   · {it}")
    else:
        print("✅ 未检出 A 级红线套话，B 级密度未超 3 倍。")
    if res["notices"]:
        print(f"\n🟡 建议核对 {len(res['notices'])} 项：")
        for it in res["notices"]:
            print(f"   · {it}")

    if "--json" in flags:
        jf = os.path.splitext(DOC)[0] + ".cliche_scan.json"
        with open(jf, "w", encoding="utf-8") as f:
            json.dump(res, f, ensure_ascii=False, indent=1)
        print(f"\nJSON 已写入：{jf}")

    sys.exit(1 if res["issues"] else 0)


if __name__ == "__main__":
    main()
