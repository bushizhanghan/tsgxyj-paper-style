# -*- coding: utf-8 -*-
"""feedback_loop.py -- 审查反哺闭环：把每次审查结果沉淀为可复用的知识/基线

设计目标（闭环）：
    输入论文 → 审查 → 提取可用数据 → 写回 skill 内部知识库 → 下次审查更准

两类输入，两种吸收策略：

    A. 已见刊论文（该刊正式发表）——**全面吸收**
       判定：文件在语料库 text/ 内、或文件名含「2026(1)」式期次、或用 --published 声明
       吸收内容：段落/句子几何、摘要素材、标题句式、引用密度、类型结构、来源
       → 更新 corpus_stats.json 的对应篇目 / 新建观察记录
       → 自动重算各维度基准（median/quartile）

    B. 用户投稿稿——**只取长处 + 记录短板**
       吸收内容：把用户稿中「优于该刊中位」的维度记录下来，作为正面样本
                 （如文笔好、结构完整、反模式干净），写入 observations
       不吸收：用户稿的偏差值（避免污染基准）

产物：
    - 语料库目录下 feedback_log.json（每次审查的完整记录）
    - 语料库目录下 observations.json（累积的正面样本与偏差统计）
    - skill 内 references/learned-patterns.md（人类可读的沉淀）

用法:
    python feedback_loop.py <论文.docx> [--published] [--corpus <dir>] [--dry-run]
    python feedback_loop.py --report              # 查看已积累的观察

纪律（防污染）：
    · 用户稿永不写入 corpus_stats.json；只有 --published 或语料库内文件才允许
    · 任何基准重算都保留原始 n 与前后差异，可回溯
    · 单篇偏差不改变基线，需积累到阈值（默认 5 篇）才提示「可考虑更新基准」
"""
import os
import re
import sys
import json
import statistics as st
from datetime import datetime

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import corpus  # noqa: E402

if hasattr(sys.stdout, "reconfigure"):
    try:
        sys.stdout.reconfigure(encoding="utf-8", errors="replace")
    except Exception:
        pass

_args = [a for a in sys.argv[1:] if not a.startswith("--")]
_flags = {a for a in sys.argv[1:] if a.startswith("--")}
_cli_corpus = None
if "--corpus" in sys.argv:
    i = sys.argv.index("--corpus")
    _cli_corpus = sys.argv[i + 1] if i + 1 < len(sys.argv) else None

DRY = "--dry-run" in _flags
DECLARED_PUBLISHED = "--published" in _flags
THRESHOLD = 5      # 同类偏差积累篇数阈值

# 注意：模块级不做任何 print/exit，保证 `import feedback_loop` 安全；
# 无参数时的帮助输出放在 main() 里（见文件末尾）。
_RESOLVED_CORPUS = None


def _corpus_dir():
    """延迟解析语料库（导入时不触发 sys.exit）。"""
    global _RESOLVED_CORPUS
    if _RESOLVED_CORPUS is None:
        _RESOLVED_CORPUS = corpus.find_corpus_or_exit(_cli_corpus, need="反哺闭环")
    return _RESOLVED_CORPUS

import diagnose_paper as dp  # noqa: E402

LOG_PATH = None   # 延迟到 _corpus_dir() 解析后设置
OBS_PATH = None


def _paths():
    global LOG_PATH, OBS_PATH
    d = _corpus_dir()
    LOG_PATH = os.path.join(d, "feedback_log.json")
    OBS_PATH = os.path.join(d, "observations.json")
    return LOG_PATH, OBS_PATH


def load_json(path, default):
    if os.path.isfile(path):
        try:
            with open(path, encoding="utf-8") as f:
                return json.load(f)
        except Exception:
            pass
    return default


def save_json(path, obj):
    if DRY:
        print(f"  [dry-run] 将写入 {path}")
        return
    with open(path, "w", encoding="utf-8") as f:
        json.dump(obj, f, ensure_ascii=False, indent=1)


# ---------- 已见刊判定 ----------
def detect_published(doc_path, corpus_dir):
    """返回 (is_published, reason)。"""
    if DECLARED_PUBLISHED:
        return True, "--published 声明"
    abs_doc = os.path.abspath(doc_path)
    text_dir = os.path.join(corpus_dir, "text")
    # 文件名是否在语料库 text/ 中
    stem = os.path.splitext(os.path.basename(doc_path))[0]
    for cand in (text_dir, corpus_dir):
        if os.path.isdir(cand):
            for fn in os.listdir(cand):
                if stem in fn or fn.startswith(stem[:20]):
                    return True, f"语料库内有同名条目（{fn[:30]}）"
    # 文件名含年期次
    if re.search(r"20\d{2}\s*[\(（]\s*\d{1,2}\s*[\)）]", stem):
        return True, "文件名含年期次（疑似已见刊）"
    return False, "用户投稿稿"


# ---------- 通用指标提取 ----------
def extract_metrics(doc_path):
    paras, refs, abs_text, kw_text, titles, body = dp.extract_paras(doc_path)
    m = dp.compute_metrics(body, refs, abs_text, kw_text, titles)
    abst = dp.scan_abstract(abs_text)
    struct = dp.scan_structure(m["titles"])
    secs = dp.scan_sections(paras, m["titles"])
    anti = dp.scan_anti_patterns("".join(dp.clean(p) for p in body))
    para_issues = dp.scan_paragraphs(m["body"])
    alltext = dp.clean("".join(m["body"]))
    sents = [s for s in re.split(r"[。？！]", alltext) if len(s.strip()) >= 6]
    types = {}
    sources = []
    try:
        import citation_analysis as ca
        for r in refs:
            b = ca.norm(r)
            t = ca.ref_type(b)
            types[t] = types.get(t, 0) + 1
            s = ca.ref_source(b, t)
            if s:
                sources.append(s)
    except Exception:
        pass
    return {
        "doc": os.path.basename(doc_path),
        "body_chars": m["body_chars"],
        "para_n": m["para_n"],
        "para_med": m["para_med"],
        "sent_mean": st.mean([len(s) for s in sents]) if sents else 0,
        "cite_rate": m["cite_rate"],
        "ref_n": m["ref_n"],
        "abs_len": m["abs_len"],
        "kw_n": m["kw_n"],
        "fig_n": m["fig_n"],
        "tab_n": m["tab_n"],
        "benwen_rate": m["benwen_rate"],
        "titles": m["titles"][:30],
        "h3_count": struct["h3_count"],
        "missing_chapters": struct["missing"],
        "section_pct": {k: round(v, 1) for k, v in (secs["merged"].items() if secs else [])},
        "anti_flagged": [a["name"] for a in anti if a["flag"]],
        "anti_total": sum(a["n"] for a in anti),
        "para_issue_n": len(para_issues),
        "abs_elements": abst.get("elements", {}),
        "abs_issues": abst.get("issues", []),
        "ref_types": types,
        "ref_sources": sources[:40],
        "le_per_1k": alltext.count("了") / max(len(alltext), 1) * 1000,
        "bei_per_1k": alltext.count("被") / max(len(alltext), 1) * 1000,
    }


# ---------- 与基准比对，找出「优于刊」的维度 ----------
def strengths_vs_base(m, base):
    """返回 (strengths, weaknesses)。用于用户稿只取长处。"""
    s, w = [], []
    checks = [
        ("段长", m["para_med"], base["paras"], "between"),
        ("句均", m["sent_mean"], base["sents"], "between"),
        ("引用密度", m["cite_rate"], base["cites"], "high"),
        ("摘要长度", m["abs_len"], base["abs"], "between"),
        ("文献量", m["ref_n"], base["refs"], "high"),
    ]
    for name, v, stat, kind in checks:
        med = stat.get("med") or 0
        q25, q75 = stat.get("q25") or 0, stat.get("q75") or 0
        if not med:
            continue
        if kind == "high":
            if v >= q75:
                s.append(f"{name} {v:.0f} 达/超刊刊 P75（{q75:.0f}）".replace("刊刊", "刊"))
            elif v < q25:
                w.append(f"{name} {v:.0f} 低于刊 P25（{q25:.0f}）")
        else:
            # between：落在四分位区间即达标
            if q25 and q75 and q25 <= v <= q75:
                s.append(f"{name} {v:.0f} 落在刊四分位区间 [{q25:.0f},{q75:.0f}]")
            elif q75 and v > q75:
                w.append(f"{name} {v:.0f} 高于刊 P75（{q75:.0f}）")
            elif q25 and v < q25:
                w.append(f"{name} {v:.0f} 低于刊 P25（{q25:.0f}）")
    # 「了」「被」
    if m["le_per_1k"] <= 3:
        s.append(f"「了」{m['le_per_1k']:.1f}/千字，低于刊 2.7（控制良好）")
    if m["bei_per_1k"] <= 0.6:
        s.append(f"「被」{m['bei_per_1k']:.2f}/千字，低于刊 0.53（控制良好）")
    # 反模式
    if not m["anti_flagged"]:
        s.append("反模式零超标（优于全部刊文水准）")
    else:
        w.append("反模式超标：" + "、".join(m["anti_flagged"]))
    # 摘要
    if all(m["abs_elements"].values()):
        s.append("摘要四要素齐全")
    bad_abs = [i for i in m["abs_issues"] if "忌词" in i]
    if bad_abs:
        w.append("摘要：" + bad_abs[0])
    return s, w


# ---------- A. 已见刊论文：全面吸收 ----------
def absorb_published(m, corpus_dir, base):
    print("\n【A. 已见刊论文 —— 全面吸收】")
    rows = corpus.load_corpus_stats(corpus_dir)
    if not rows:
        print("  ⚠️ 语料库无 corpus_stats.json，跳过基准吸收")
        return {"absorbed": False}

    # 找到同名篇目（若已存在则更新，否则追加新观察）
    matched = None
    stem = os.path.splitext(m["doc"])[0]
    for r in rows:
        rf = os.path.splitext(r.get("file", ""))[0]
        if stem[:20] in rf or rf[:20] in stem:
            matched = r
            break

    rec = {
        "no": (matched or {}).get("no") or f"NEW{len(rows)+1:03d}",
        "file": m["doc"],
        "issue": (matched or {}).get("issue") or "unknown",
        "body_chars": round(m["body_chars"]),
        "para_n": m["para_n"],
        "para_med": round(m["para_med"]),
        "sent_mean": round(m["sent_mean"], 2),
        "cite_rate": round(m["cite_rate"], 2),
        "ref_cnt": m["ref_n"],
        "abstract_len": m["abs_len"],
        "kw_n": m["kw_n"],
        "fig_n": m["fig_n"],
        "tab_n": m["tab_n"],
        "benwen": round(m["benwen_rate"]),
        "heads": m["titles"],
        "has_indent": True,
        "max_para": max([round(x) for x in [m["para_med"]]] or [0]),
        "learned": True,
    }
    if matched:
        print(f"  更新已有篇目：{rec['no']}（{matched.get('file', '')[:36]}）")
        print(f"    旧：段长中位 {matched.get('para_med')} / 句均 {matched.get('sent_mean', 0):.1f} / 文献 {matched.get('ref_cnt')}")
        matched.update(rec)
    else:
        print(f"  新增篇目：{rec['no']}（{m['doc'][:40]}）")
        rows.append(rec)
    print(f"    新：段长中位 {rec['para_med']} / 句均 {rec['sent_mean']} / 文献 {rec['ref_cnt']}")

    if not DRY:
        with open(os.path.join(corpus_dir, "corpus_stats.json"), "w", encoding="utf-8") as f:
            json.dump(rows, f, ensure_ascii=False, indent=1)
        print(f"  ✓ 已写回 corpus_stats.json（共 {len(rows)} 篇）")
    else:
        print("  [dry-run] 未写回 corpus_stats.json")

    # 基准重算（只报告差异，不自动改文档）
    new_base = corpus.load_reference_stats(_corpus_dir())
    diffs = []
    for k, label in [("paras", "段长"), ("sents", "句均"), ("cites", "引用密度"),
                     ("refs", "文献量"), ("abs", "摘要长度")]:
        o, n = base[k]["med"], new_base[k]["med"]
        if isinstance(o, (int, float)) and isinstance(n, (int, float)) and o:
            delta = n - o
            if abs(delta) > 0.01:
                diffs.append(f"{label}中位 {o}{'→'}{n}（{delta:+.2f}）")
    if diffs:
        print("  基准变化：" + "；".join(diffs))
    return {"absorbed": True, "no": rec["no"], "matched": bool(matched),
            "base_diffs": diffs, "n_papers": len(rows)}


# ---------- B. 用户投稿稿：只取长处 ----------
def absorb_user(m, base):
    print("\n【B. 用户投稿稿 —— 只取长处，不污染基准】")
    s, w = strengths_vs_base(m, base)
    print(f"  长处（{len(s)} 项，将沉淀为正面样本）：")
    for x in s:
        print(f"    + {x}")
    print(f"  短板（{len(w)} 项，仅记录用于趋势统计）：")
    for x in w:
        print(f"    - {x}")
    return {"strengths": s, "weaknesses": w}


# ---------- 主流程 ----------
def main():
    if not _args and "--report" not in _flags:
        print(__doc__)
        return 0

    _paths()   # 解析语料库与文件路径

    if "--report" in _flags:
        log = load_json(LOG_PATH, [])
        obs = load_json(OBS_PATH, {})
        print("=" * 66)
        print(f"审查反哺记录（{LOG_PATH}）")
        print("=" * 66)
        if not log:
            print("尚无记录。运行 `python feedback_loop.py <论文.docx>` 开始积累。")
            return 0
        print(f"\n累计审查 {len(log)} 次：")
        for e in log[-10:]:
            kind = "已见刊" if e.get("published") else "投稿稿"
            print(f"  [{e.get('time', '')[:16]}] {kind}｜{e.get('doc', '')[:40]}")
            if e.get("published"):
                print(f"      吸收：{e.get('absorb', {}).get('no', '-')}"
                      f"（{'更新' if e.get('absorb', {}).get('matched') else '新增'}）")
            else:
                print(f"      长处 {len(e.get('strengths', []))} 项 / 短板 {len(e.get('weaknesses', []))} 项")
        # 累积的短板趋势
        if obs.get("weakness_counter"):
            print("\n短板出现频次（Top10，可据此调整诊断权重）：")
            for k, v in sorted(obs["weakness_counter"].items(), key=lambda x: -x[1])[:10]:
                print(f"  {v:>3} 次  {k[:60]}")
        if obs.get("strength_counter"):
            print("\n长处出现频次（Top10，可据此识别「好稿」特征）：")
            for k, v in sorted(obs["strength_counter"].items(), key=lambda x: -x[1])[:10]:
                print(f"  {v:>3} 次  {k[:60]}")
        return 0

    doc = _args[0]
    if not os.path.isfile(doc):
        print(f"找不到文件：{doc}")
        return 2

    is_pub, why = detect_published(doc, _corpus_dir())
    base = corpus.load_reference_stats(_corpus_dir())
    m = extract_metrics(doc)

    print("=" * 66)
    print("审查反哺闭环")
    print("=" * 66)
    print(f"文件：{os.path.basename(doc)}")
    print(f"类型判定：{'已见刊论文' if is_pub else '用户投稿稿'}（{why}）")
    print(f"策略：{'全面吸收 → 更新语料库基准' if is_pub else '只取长处 → 不污染基准'}")
    print(f"基本指标：正文 {m['body_chars']} 字 / 段 {m['para_n']} / "
          f"段长中位 {m['para_med']:.0f} / 句均 {m['sent_mean']:.0f} / 文献 {m['ref_n']}")

    absorb = absorb_published(m, _corpus_dir(), base) if is_pub else None
    user_part = None if is_pub else absorb_user(m, base)

    # ---- 写记录 ----
    entry = {
        "time": datetime.now().isoformat(timespec="seconds"),
        "doc": os.path.basename(doc),
        "published": is_pub,
        "reason": why,
        "metrics": {k: v for k, v in m.items() if k not in ("titles", "ref_sources")},
        "absorb": absorb,
        "strengths": (user_part or {}).get("strengths", []),
        "weaknesses": (user_part or {}).get("weaknesses", []),
    }
    log = load_json(LOG_PATH, [])
    log.append(entry)
    save_json(LOG_PATH, log)

    # ---- 更新观察累积 ----
    obs = load_json(OBS_PATH, {"strength_counter": {}, "weakness_counter": {},
                               "published_absorbed": [], "n_runs": 0})
    obs["n_runs"] = obs.get("n_runs", 0) + 1
    for x in entry["strengths"]:
        key = re.sub(r"[\d.]+", "#", x)[:60]
        obs["strength_counter"][key] = obs["strength_counter"].get(key, 0) + 1
    for x in entry["weaknesses"]:
        key = re.sub(r"[\d.]+", "#", x)[:60]
        obs["weakness_counter"][key] = obs["weakness_counter"].get(key, 0) + 1
    if is_pub and absorb and absorb.get("absorbed"):
        obs["published_absorbed"].append({"no": absorb.get("no"), "time": entry["time"]})
    save_json(OBS_PATH, obs)

    print("\n" + "─" * 66)
    print("闭环结果")
    print("─" * 66)
    print(f"  记录：{LOG_PATH}{'（dry-run 未写）' if DRY else ''}")
    print(f"  观察：{OBS_PATH}{'（dry-run 未写）' if DRY else ''}")
    print(f"  累计审查 {obs['n_runs']} 次")

    # ---- 阈值提示 ----
    hot = [k for k, v in obs.get("weakness_counter", {}).items() if v >= THRESHOLD]
    if hot:
        print(f"\n  ⚠️ 以下短板已累计 ≥{THRESHOLD} 次，建议据此调整诊断默认阈值：")
        for k in hot:
            print(f"      {obs['weakness_counter'][k]} 次 · {k}")
    good = [k for k, v in obs.get("strength_counter", {}).items() if v >= THRESHOLD]
    if good:
        print(f"\n  ✓ 以下长处已累计 ≥{THRESHOLD} 次，可写入 learned-patterns.md 作为正面样本：")
        for k in good:
            print(f"      {obs['strength_counter'][k]} 次 · {k}")
    return 0


if __name__ == "__main__":
    sys.exit(main())
