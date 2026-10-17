# -*- coding: utf-8 -*-
"""verify_refs.py —— 参考文献核验：序号连续 / 正文-列表对应 / 引用形态 / 语料库溯源

用法（任一处运行）:
    python verify_refs.py <论文.docx>
    python verify_refs.py <论文.docx> --corpus <语料库目录>   # 显式指定（溯源用）
    python verify_refs.py <论文.docx> --json                  # 附加输出 JSON

输出四项核验：
    ① 序号连续性（缺号 / 重号）
    ② 正文引用 ↔ 文献列表对应（引用但列表无 / 列表有但正文未引）
    ③ 引用标注形态（对照本刊 118 篇实测：范围式为主、列举式次之、
       混合式 0 处不用、连排单引不推荐）
    ④ 语料库溯源（能匹配到语料库篇目的文献标记 ✅，未匹配需人工核验 ⚠️）

依赖：python-docx；语料库仅用于溯源，缺失时该项自动跳过并提示。
"""
import re
import sys
import os
import json

# Windows 控制台默认 GBK 无法输出部分 UTF-8 字符，统一走 UTF-8
if hasattr(sys.stdout, "reconfigure"):
    try:
        sys.stdout.reconfigure(encoding="utf-8", errors="replace")
    except Exception:
        pass

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import corpus  # noqa: E402


def norm(s):
    return s.translate(str.maketrans("０１２３４５６７８９［］", "0123456789[]"))


def clean(s):
    return re.sub(r"\s+", "", s)


def split_body_refs(path):
    """从 docx 拆分正文与参考文献。"""
    from docx import Document
    doc = Document(path)
    paras = [p.text.strip() for p in doc.paragraphs if p.text.strip()]
    refs, body = [], []
    in_ref = False
    for t in paras:
        if t.startswith("参考文献"):
            in_ref = True
            continue
        (refs if in_ref else body).append(t)
    body_txt = clean(norm("".join(body)))
    return body_txt, refs


def parse_ref_items(refs):
    items = []
    for t in refs:
        m = re.match(r"^[\[［]?(\d{1,3})[\]］]?\s*[:：.]?\s*(.+)", t)
        if m and len(m.group(2)) > 8:
            items.append((int(m.group(1)), m.group(2)))
    return items


def collect_cited(body_txt):
    """收集正文全部被引序号（展开范围式/列举式）。"""
    cited_all = set()
    for m in re.finditer(r"\[(\d{1,3}(?:\s*[-,，]\s*\d{1,3})*)\]", body_txt):
        seg = m.group(1)
        for part in re.split(r"[-,，]", seg):
            if part.strip().isdigit():
                cited_all.add(int(part.strip()))
    # 范围式 [2-5] 展开
    for m in re.finditer(r"\[(\d{1,3})\s*[-–—]\s*(\d{1,3})\]", body_txt):
        a, b = int(m.group(1)), int(m.group(2))
        if 0 < a <= b <= 200:
            cited_all.update(range(a, b + 1))
    return cited_all


def main():
    args = sys.argv[1:]
    if not args or args[0] in ("-h", "--help"):
        print(__doc__)
        return 0
    cli_corpus = None
    flags = set()
    rest = []
    i = 0
    while i < len(args):
        a = args[i]
        if a == "--corpus":
            cli_corpus = args[i + 1] if i + 1 < len(args) else None
            i += 2
            continue
        if a.startswith("--"):
            flags.add(a)
        else:
            rest.append(a)
        i += 1
    path = rest[0] if rest else None
    if not path or not os.path.isfile(path):
        print("用法: python verify_refs.py <论文.docx> [--corpus <dir>] [--json]")
        return 2

    body_txt, refs = split_body_refs(path)
    items = parse_ref_items(refs)
    nums_list = [n for n, _ in items]

    print("=" * 62)
    print(f"参考文献核验：{os.path.basename(path)}")
    print("=" * 62)
    if not items:
        print("未解析到参考文献条目（请确认正文含「参考文献」段且条目有编号）")
        return 1
    print(f"参考文献条目：{len(items)} 条（编号 {min(nums_list)}–{max(nums_list)}）")

    # ① 序号连续性
    expect = set(range(min(nums_list), max(nums_list) + 1))
    missing = sorted(expect - set(nums_list))
    dup = [n for n in set(nums_list) if nums_list.count(n) > 1]
    print(f"  [序号] 缺号: {missing or '无'} | 重号: {dup or '无'}"
          + ("  ⚠️ 需修复" if missing or dup else "  ✓"))

    # ② 正文引用 ↔ 列表
    cited_all = collect_cited(body_txt)
    if items:
        listed = set(nums_list)
        ghost = sorted(cited_all - listed)
        unused = sorted(listed - cited_all)
        print(f"  [对应] 正文引用 {len(cited_all)} 个序号 | "
              f"引用但列表无: {ghost or '无'}"
              + ("  ⚠️" if ghost else "  ✓")
              + f" | 列表有但正文未引: {unused or '无'}"
              + ("  ⚠️ 建议核对" if unused else ""))

    # ③ 引用标注形态
    singles_n = len(re.findall(r"\[\d{1,3}\]", body_txt))
    ranges_n = len(re.findall(r"\[\d{1,3}\s*[-–—]\s*\d{1,3}\]", body_txt))
    commas_n = len(re.findall(r"\[\d{1,3}(?:\s*[,，]\s*\d{1,3})+\]", body_txt))
    mixed_n = len(re.findall(r"\[\d{1,3}\s*[-–—]\s*\d{1,3}[\s*,，]\s*\d{1,3}(?:[\s*,，]\s*\d{1,3})*\]", body_txt))
    adjacent_n = len(re.findall(r"\[\d+\]\s*\[\d+\]", body_txt))
    print("  [形态] （刊文基线：范围式为主流、列举式 38 处、混合式不用、连排单引不推荐）")
    print(f"    ↑ 单个 [5]: {singles_n} | 范围式 [2-5]: {ranges_n} | "
          f"列举式 [2,3]: {commas_n} | 混合式: {mixed_n} | 连排单引: {adjacent_n}")
    if mixed_n:
        print("    ⚠️ 检出混合式——本刊无此先例，建议拆分标注或改写句子使编号连续")
    if adjacent_n:
        print("    ⚠️ 检出连排单引——建议改为范围式或列举式")

    # ④ 语料库溯源
    print("\n  [溯源] (匹配语料库 118 篇；未匹配需人工核验 ⚠️)")
    corpus_dir = None
    try:
        corpus_dir = corpus.find_corpus_dir(cli_corpus)
    except FileNotFoundError as e:
        print(f"    [提示] {e}；跳过溯源（不影响前三项）")
    hits, misses = [], []
    if corpus_dir:
        pj = os.path.join(corpus_dir, "paragraphs.json")
        corpus_titles = []
        try:
            for x in json.load(open(pj, encoding="utf-8")):
                m = re.match(r"(.+)_([^_]+)\.pdf$", x.get("file", ""))
                if m:
                    corpus_titles.append((x.get("no"), x.get("issue"), clean(norm(m.group(1)))))
        except Exception:
            corpus_titles = []
        for n, txt in items:
            ct = clean(norm(txt))
            matched = None
            for no, iss, title in corpus_titles:
                key = title[:10]
                if len(key) >= 8 and key in ct:
                    matched = f"{no}|{iss}"
                    break
            (hits if matched else misses).append((n, matched, ct[:42]))
        for n, m, t in hits:
            print(f"    ✅ [{n}] → {m} | {t}")
        for n, m, t in misses[:40]:
            print(f"    ⚠️ [{n}] 人工核验 | {t}")
        if not hits:
            print("    （语料库无一命中——若文献均为外刊/专著属正常）")
        print(f"    —— 命中 {len(hits)} 条 / 需人工核验 {len(misses)} 条（未匹配≠错，只是不在本语料范围）")

    if "--json" in flags:
        out = {
            "doc": os.path.basename(path), "ref_n": len(items),
            "missing_nums": missing, "dup_nums": dup,
            "cited_not_listed": sorted(cited_all - set(nums_list)),
            "listed_not_cited": sorted(set(nums_list) - cited_all),
            "forms": {"single": singles_n, "range": ranges_n, "comma": commas_n,
                      "mixed": mixed_n, "adjacent": adjacent_n},
            "corpus_hits": [n for n, _, _ in hits] if corpus_dir else [],
            "corpus_misses": [n for n, _, _ in misses] if corpus_dir else [],
        }
        jf = os.path.splitext(path)[0] + ".refs.json"
        with open(jf, "w", encoding="utf-8") as f:
            json.dump(out, f, ensure_ascii=False, indent=1)
        print(f"\nJSON 已写入：{jf}")
    return 0


if __name__ == "__main__":
    sys.exit(main())
