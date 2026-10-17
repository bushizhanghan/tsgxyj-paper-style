# -*- coding: utf-8 -*-
"""search.py —— 《图书馆学研究》语料库段落检索（覆盖全 118 篇）

用法（任一处运行，自动定位语料库；也可 --corpus 指定）:
    python search.py 适老化                    # 搜索所有含「适老化」的段落
    python search.py 适老化 服务               # AND：同时含两词
    python search.py --or 适老化 适老          # OR：含任一词
    python search.py --count 适老化            # 只统计命中段数/篇数
    python search.py --ctx 适老化              # 关键词前后 100 字上下文（默认 60 字）
    python search.py --re "适老.{0,4}服务"     # 正则搜索
    python search.py --top 100 阅读推广        # 最多输出 100 段（默认 60）
    python search.py 适老化 --out hits.md      # 结果写入 Markdown 文件
    python search.py --topic 无障碍            # 限定主题篇目范围（按篇名/keywords 粗筛）

输出每段标注：篇号 | 年期 | 作者 | 段落序号。
"""
import json
import sys
import os
import re

# Windows 控制台默认 GBK 无法输出部分 UTF-8 字符，统一走 UTF-8
if hasattr(sys.stdout, "reconfigure"):
    try:
        sys.stdout.reconfigure(encoding="utf-8", errors="replace")
    except Exception:
        pass

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import corpus  # noqa: E402


def author_of(fname):
    stem = os.path.splitext(fname)[0]
    parts = stem.split("_")
    return parts[-1] if len(parts) > 1 else ""


def load(corpus_dir):
    d = corpus_dir or corpus.find_corpus_or_exit(need="段落检索")
    data = corpus.load_paragraphs(d)
    if not data:
        print(f"[错误] 语料库 {d} 中没有 paragraphs.json")
        sys.exit(2)
    return data


def topic_filter(items, topic):
    """按主题粗筛：篇号 >= 109（扩展区）或篇名含主题词。"""
    if not topic:
        return items
    out = []
    for it in items:
        no = it.get("no", "")
        name = os.path.splitext(it.get("file", ""))[0]
        if topic in name or topic in it.get("issue", "") or \
           (no.isdigit() and int(no) >= 109) or topic in json.dumps(it.get("paras", []), ensure_ascii=False)[:600]:
            out.append(it)
    return out


def match(args, mode_re, p):
    if mode_re:
        try:
            return bool(re.search(args[0], p))
        except re.error as e:
            print(f"[错误] 正则无效：{e}")
            sys.exit(2)
    return all(k in p for k in args) if not _or_mode else any(k in p for k in args)


def build_hit_block(item, i, p, width, mode_re):
    """构造 (位置标注, 带上下文片段)。"""
    if mode_re:
        m = re.search(sys_re_args[0], p)
        s = max(0, m.start() - width)
        ctx = p[s:m.end() + width]
    else:
        # 取首个命中关键词的最短者作为锚点，尽量覆盖多个关键词
        anchor = None
        for k in keys:
            pos = p.find(k)
            if pos >= 0 and (anchor is None or pos < anchor[1]):
                anchor = (k, pos)
        if anchor:
            k, pos = anchor
            s = max(0, pos - width)
            ctx = p[s:pos + len(k) + width]
        else:
            ctx = p[:width * 2]
    loc = f"[{item.get('no', '?')}|{item.get('issue', '?')}|{author_of(item.get('file', ''))}]"
    return loc, ctx


def main():
    global _or_mode, sys_re_args, keys  # noqa: PLW0603（供 match 使用）
    args = [a for a in sys.argv[1:]]
    if not args or args[0] in ("-h", "--help", "--"):
        print(__doc__)
        return 0

    # ---- 参数解析 ----
    flags = {a for a in args if a.startswith("--") and not a.startswith("--corpus")}
    cli_corpus = None
    rest = []
    i = 0
    while i < len(args):
        a = args[i]
        if a == "--corpus":
            cli_corpus = args[i + 1] if i + 1 < len(args) else None
            i += 2
            continue
        if a.startswith("--"):
            rest.append(a)
        else:
            rest.append(a)
        i += 1

    move_out = None
    if "--out" in flags:
        # 从 rest 中取 --out 的值
        idx = rest.index("--out")
        move_out = rest[idx + 1] if idx + 1 < len(rest) else "hits.md"
        del rest[idx:idx + 2]
        flags.discard("--out")
    mode_re = "--re" in flags
    count_only = "--count" in flags
    or_mode = "--or" in flags
    _or_mode = or_mode
    width = 100 if "--ctx" in flags else 60
    max_show = 60
    for f in list(flags):
        if f.startswith("--top"):
            idx = rest.index(f)
            try:
                max_show = int(rest[idx + 1])
            except (ValueError, IndexError):
                max_show = 60
            del rest[idx:idx + 2]
            flags.discard(f)
    topic = None
    if "--topic" in flags:
        idx = rest.index("--topic")
        topic = rest[idx + 1] if idx + 1 < len(rest) else None
        del rest[idx:idx + 2]
        flags.discard("--topic")

    keys = [a for a in rest if not a.startswith("--")]
    if not keys:
        print(__doc__)
        return 0
    sys_re_args = keys

    data = load(cli_corpus)
    data = topic_filter(data, topic)

    hits = []
    for item in data:
        for i, p in enumerate(item.get("paras", [])):
            if not p.strip():
                continue
            if match(keys, mode_re, p):
                hits.append((item, i, p))
    n_papers = len(set(h[0].get("no") for h in hits))
    head = f"命中 {len(hits)} 段，涉及 {n_papers} 篇" + (f"（主题筛选：{topic}）" if topic else "")

    if count_only:
        print(head)
        # 按篇细目
        from collections import Counter
        by_paper = Counter(f"[{it.get('no')}|{it.get('issue')}|{author_of(it.get('file'))}]"
                           for it, _, _ in hits)
        for loc, cnt in by_paper.most_common(10):
            print(f"  {loc} ×{cnt}")
        return 0

    out_lines = [head, ""]
    for item, i, p in hits[:max_show]:
        loc, ctx = build_hit_block(item, i, p, width, mode_re)
        out_lines.append(f"{loc} §{i + 1} …{ctx}…")
        out_lines.append("")
    if len(hits) > max_show:
        out_lines.append(f"（仅显示前 {max_show} 段，共 {len(hits)} 段；用 --top 增加）")

    text = "\n".join(out_lines)
    print(text)
    if move_out:
        with open(move_out, "w", encoding="utf-8") as f:
            f.write(text)
        print(f"\n已写入：{os.path.abspath(move_out)}")
    return 0


_or_mode = False
sys_re_args = []
keys = []


if __name__ == "__main__":
    sys.exit(main())
