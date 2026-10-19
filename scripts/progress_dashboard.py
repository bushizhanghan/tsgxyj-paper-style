# -*- coding: utf-8 -*-
"""progress_dashboard.py —— 《图书馆学研究》投稿就绪度看板（方向八）

把投稿前必须过一遍的七项门槛汇成一张仪表盘：n/7 达标总览 + 逐项状态行
（含量化差距）+ 最优先修复建议 Top3（按对分数的边际影响排序）。

七项就绪指标（每项独立判定，任一 ✗ 即整体未就绪）：
    ① 红线 0 项            ← diagnose_paper 红线合计（官方合规/反模式超标/结构缺标配/摘要红线）
    ② 模拟审稿分 ≥85       ← mock_review 13 维百分制总分（差值 = 85 − 当前分）
    ③ 参考文献 ≥ 刊 P25     ← diagnose 分栏目基线的 refs.q25（差值 = P25 − 当前条数）
    ④ 无待补图占位符       ← table_fig_audit 占位计数（占位符/三线表缺陷）
    ⑤ 数字一致性无硬伤     ← number_consistency 硬伤数（待核对不算硬伤）
    ⑥ 术语无混用           ← term_consistency 混用组数（非偏好写法只提示、不计 ✗）
    ⑦ 官方投稿包要素齐备   ← diagnose.scan_official 的 package_notices
                              （作者简介 / 联系方式 / 定稿日期）

复用方式（不重复造轮子）：优先 import 各工具模块直接调其函数取数——看板要的是
「结构化数值」，不是各工具的完整报告，因此与 check_all.py 的 subprocess 串联互补。
mock_review.py 是脚本式模块（顶层即 sys.exit），故该项走 subprocess 解析 stdout
关键行（【总分】与修改路线图），与 check_all.py 的做法一致。
任一项依赖 import 失败时该项降级为「无法评估」，不计入 n/7 分母，并在输出中明示，
不让一个脚本的缺失拖垮整张看板。

用法:
    python progress_dashboard.py <论文.docx> [--corpus <dir>] [--top N] [--json]
    python progress_dashboard.py <论文.docx> --brief   # 只出看板与 Top3

退出码：0 = 七项全部就绪；1 = 存在未就绪项（或依赖缺失导致无法评估）。
说明：本看板只做形式就绪度汇总，指标口径沿用各工具；不替代细读，也不替代
check_all.py 的全量报告。
"""
import os
import re
import sys
import json
import io
import subprocess

# Windows 控制台 GBK 保护
if hasattr(sys.stdout, "reconfigure"):
    try:
        sys.stdout.reconfigure(encoding="utf-8", errors="replace")
    except Exception:
        pass

HERE = os.path.dirname(os.path.abspath(__file__))
if HERE not in sys.path:
    sys.path.insert(0, HERE)

REVIEW_TARGET = 85.0     # 指标② 就绪线（mock_review 判定线：≥85 形式层面优秀）
DASHBOARD_NAME = "投稿就绪度"

# 指标⑦ 投稿包三要素（与 dp.scan_official 的 package_notices 文案对应）
PACKAGE_ITEMS = ("作者简介", "联系方式", "定稿日期")

# 红线等效权重：仅用于红线内部排序（审稿失分通道一律排在红线之前，见 top_advice）。
# 取值依据该刊审稿关注度（mock_review 权重）：官方合规 7 > 净度 6 > 结构 4 ≈ 摘要 4。
RED_WEIGHT = {"official": 7.0, "anti": 6.0, "struct": 4.0, "abstract": 4.0}
RED_LABEL = {"official": "官方合规红线", "anti": "反模式超标红线",
             "struct": "结构缺标配红线", "abstract": "摘要红线"}
RED_HINT = {
    "official": "官方须知明文项（英文摘要/关键词、栏目对位、序号递增等）——见 diagnose_paper ⓪ 节。",
    "anti": "反模式超该刊基线 2 倍——见 diagnose_paper ③ 节与 anti-patterns.md。",
    "struct": "缺该刊收口标配章（对策/策略 67%、结语 52% 命中率）——见 diagnose_paper ① 节。",
    "abstract": "摘要含忌词/虚句或字数越界——见 diagnose_paper 摘要体检。",
}


# ---------------- 依赖加载（任一失败即降级，不崩溃） ----------------
def _load_tools():
    """按需 import 工具模块，返回 (modules, errors)。

    diagnose_paper / number_consistency / term_consistency 会拉 docx 与 corpus，
    缺 python-docx 或缺语料库时可能抛异常——统一捕获后记入 errors，
    由调用方把对应指标标为「无法评估」。
    """
    mods, errors = {}, {}
    for key, mod in [("dp", "diagnose_paper"), ("nc", "number_consistency"),
                     ("tc", "term_consistency"), ("tf", "table_fig_audit"),
                     ("vr", "verify_refs")]:
        try:
            mods[key] = __import__(mod)
        except Exception as e:      # ImportError / ModuleNotFoundError / 语法错误
            errors[mod] = f"{type(e).__name__}: {e}"
    return mods, errors


def _fmt_gap(v):
    """差距值格式化：整数不带小数点，小数保留 1 位。"""
    if v is None:
        return "?"
    try:
        fv = float(v)
    except (TypeError, ValueError):
        return str(v)
    return f"{fv:g}" if abs(fv - round(fv)) < 1e-9 else f"{fv:.1f}"


# ---------------- 指标②：mock_review（subprocess 解析 stdout） ----------------
def _mock_review(doc, corpus_dir):
    """跑 mock_review.py，解析【总分】与修改路线图。

    返回 {percent, total, max, roadmap:[(loss, dim, op)]}；失败返回 None。
    mock_review 是脚本式模块（顶层 sys.exit），不能 import，故走 subprocess。
    """
    cmd = [sys.executable, os.path.join(HERE, "mock_review.py"), doc]
    if corpus_dir:
        cmd += ["--corpus", corpus_dir]
    try:
        r = subprocess.run(cmd, capture_output=True, encoding="utf-8", errors="replace")
    except Exception as e:
        return {"error": f"{type(e).__name__}: {e}"}
    out = r.stdout or ""
    m = re.search(r"【总分】([\d.]+)\s*/\s*(\d+)", out)
    if not m:
        return {"error": (r.stderr or out or "无输出").strip()[-160:]}
    total, max_total = float(m.group(1)), int(m.group(2))
    roadmap = []
    for mm in re.finditer(r"^\s+\d+\.\s*\[失\s*([\d.]+)\s*分·([^\]]+)\]\s*(.+)$", out, re.M):
        roadmap.append((float(mm.group(1)), mm.group(2).strip(), mm.group(3).strip()))
    return {"total": total, "max": max_total,
            "percent": round(total / max_total * 100, 1) if max_total else 0.0,
            "roadmap": roadmap}


# ---------------- 指标采集 ----------------
def collect(doc, corpus_dir):
    """采集七项指标。返回 (rows, ctx)。

    rows: [{id, name, ok, mark, detail, gap, source, unavailable}]
    ctx:  看板附加上下文（审稿路线图、红线明细、投稿包提示、引用核验等）。
    """
    mods, load_errors = _load_tools()
    ctx = {"load_errors": load_errors, "roadmap": [], "red_detail": {},
           "package_notices": [], "nc_todo": 0, "nc_hard": [], "tc_hints": [],
           "tf_issues": 0, "vr": None, "review": None, "base": None}

    rows = []

    def add(idx, name, ok, detail, gap, source, unavailable=None):
        rows.append({"id": idx, "name": name, "ok": ok, "mark": "—" if ok is None else ("✓" if ok else "✗"),
                     "detail": detail, "gap": gap, "source": source,
                     "unavailable": unavailable})

    # ---------- 公共取数：诊断指标与官方合规（指标①③⑦ 共用，只算一次） ----------
    dp = mods.get("dp")
    m = official = anti = struct = abst = base = None
    if dp is not None:
        try:
            paras, refs, abs_text, kw_text, titles, body = dp.extract_paras(doc)
            m = dp.compute_metrics(body, refs, abs_text, kw_text, titles)
            abst = dp.scan_abstract(abs_text)
            anti = dp.scan_anti_patterns("".join(dp.clean(p) for p in body))
            struct = dp.scan_structure(m["titles"])
            official = dp.scan_official(kw_text, refs, m["titles"],
                                        "".join(dp.clean(p) for p in body), paras=paras)
            # 分栏目基线口径：与 diagnose_paper.main 一致——先按自动对位栏目取分栏目基准
            target_column = None
            try:
                target_column = official["columns"][0] if official["columns"] else None
            except Exception:
                target_column = None
            try:
                base = dp.corpus.load_reference_stats(corpus_dir, column=target_column)
            except Exception as e:
                ctx["base_error"] = f"{type(e).__name__}: {e}"
        except Exception as e:
            ctx["diag_error"] = f"{type(e).__name__}: {e}"
            dp = m = official = anti = struct = abst = base = None

    # ================= ① 红线 0 项 =================
    if dp is None:
        add(1, "红线 0 项", None, "该项无法评估（diagnose_paper 不可用或解析失败）", None,
            "diagnose_paper.py", unavailable=ctx.get("diag_error") or "diagnose_paper 不可用")
    else:
        n_anti = sum(1 for a in anti if a["flag"])
        n_struct = sum(1 for x in ("对策/策略/路径/启示（收口标配）", "结语")
                       if x in struct["missing"])
        n_abs = sum(1 for i in abst.get("issues", [])
                    if "忌词" in i or "虚句" in i or "偏长" in i)
        n_off = len(official.get("issues", []))
        n_red = n_anti + n_struct + n_abs + n_off
        ctx["red_detail"] = {"official": n_off, "anti": n_anti,
                             "struct": n_struct, "abstract": n_abs}
        add(1, "红线 0 项", n_red == 0,
            f"红线合计 {n_red} 项（官方合规 {n_off} / 反模式超标 {n_anti} / "
            f"结构缺标配 {n_struct} / 摘要红线 {n_abs}）",
            n_red, "diagnose_paper.py")

    # ================= ② 模拟审稿分 ≥85 =================
    mr = _mock_review(doc, corpus_dir)
    if mr is None or "error" in mr:
        add(2, "模拟审稿分 ≥85", None, "该项无法评估（mock_review.py 取数失败）", None,
            "mock_review.py", unavailable=(mr or {}).get("error") or "mock_review 未返回结果")
    else:
        ctx["review"] = mr
        ctx["roadmap"] = mr["roadmap"]
        gap = round(REVIEW_TARGET - mr["percent"], 1)
        add(2, "模拟审稿分 ≥85", mr["percent"] >= REVIEW_TARGET,
            f"模拟审稿 {mr['percent']:.1f}/100（目标 {REVIEW_TARGET:g}"
            + (f"，还差 {_fmt_gap(gap)} 分" if gap > 0 else "，已达标）"),
            gap if gap > 0 else 0, "mock_review.py")

    # ================= ③ 参考文献 ≥ 刊 P25 =================
    if dp is None or base is None or m is None:
        add(3, "参考文献 ≥ 刊 P25", None, "该项无法评估（诊断基线不可用）", None,
            "diagnose_paper.py", unavailable=ctx.get("base_error") or "diagnose_paper 不可用")
    else:
        q25 = (base.get("refs") or {}).get("q25") or 0
        n_refs = m["ref_n"]
        gap = q25 - n_refs
        add(3, "参考文献 ≥ 刊 P25", n_refs >= q25,
            f"参考文献 {n_refs} 条 vs 刊 P25 {q25} 条（基线：{base.get('source', '')}）"
            + (f"，还差 {_fmt_gap(gap)} 条" if gap > 0 else "，已达 P25"),
            gap if gap > 0 else 0, "diagnose_paper.py")

    # ================= ④ 无待补图占位符 =================
    tf = mods.get("tf")
    if tf is None:
        add(4, "无待补图占位符", None, "该项无法评估（table_fig_audit 不可用）", None,
            "table_fig_audit.py",
            unavailable=load_errors.get("table_fig_audit", "table_fig_audit 不可用"))
    else:
        try:
            res = tf.audit_document_tables_and_figures(doc)
            st = res.get("stats", {}) or {}
            ph = len(st.get("placeholders", []))
            ctx["tf_issues"] = len(res.get("issues", []))
            detail = f"待补图占位符 {ph} 处"
            if ctx["tf_issues"]:
                detail += f"；另有图表体例问题 {ctx['tf_issues']} 项"
            if ph == 0:
                detail += (f"（表 {st.get('word_tables', 0)} 个，三线表合规 "
                           f"{st.get('three_line_tables', 0)} 个）")
            add(4, "无待补图占位符", ph == 0, detail, ph, "table_fig_audit.py")
        except Exception as e:
            add(4, "无待补图占位符", None, "该项无法评估（table_fig_audit 取数失败）", None,
                "table_fig_audit.py", unavailable=f"{type(e).__name__}: {e}")

    # ================= ⑤ 数字一致性无硬伤 =================
    nc = mods.get("nc")
    if nc is None:
        add(5, "数字一致性无硬伤", None, "该项无法评估（number_consistency 不可用）", None,
            "number_consistency.py",
            unavailable=load_errors.get("number_consistency", "number_consistency 不可用"))
    else:
        try:
            reg = nc.split_regions(doc)
            abs_as = nc.extract_assertions([reg["abstract"]] if reg["abstract"] else [], "摘要")
            body_as = nc.extract_assertions(reg["body"], "正文")
            concl_as = nc.extract_assertions(reg["conclusion"], "结语")
            pool_as = body_as + concl_as
            hard, todo = [], []
            if abs_as:
                h1, t1, _o1 = nc.compare_regions(abs_as, pool_as, "摘要")
                hard += h1
                todo += t1
                ag, pg = nc.pct_groups(abs_as), nc.pct_groups(pool_as)
                todo += nc.check_enum_groups(ag, pg, "摘要")
                todo += nc.check_pct_sum(ag)
                h2, t2, _o2 = nc.compare_regions(concl_as, body_as, "结语")
                hard += h2
                todo += t2
                todo += nc.check_pct_sum(nc.pct_groups(concl_as))
            hard += nc.check_sample_logic(abs_as + body_as + concl_as)
            hard += nc.check_pct_bounds(abs_as + body_as + concl_as)
            todo += nc.check_rate_consistency(abs_as + body_as + concl_as)
            abs_claims = nc.extract_claims(reg["abstract"], "摘要")
            body_claims = []
            for i, p in enumerate(reg["body"]):
                body_claims += nc.extract_claims(p, "正文", i + 1)
            concl_claims = nc.extract_claims("\n".join(reg["conclusion"]), "结语")
            ch, ct = nc.compare_claims(abs_claims, body_claims + concl_claims, "摘要")
            hard += ch
            todo += ct
            todo += nc.check_claim_vs_enum(abs_claims, reg["body"], "正文", "摘要")
            ctx["nc_todo"] = len(todo)
            ctx["nc_hard"] = hard
            add(5, "数字一致性无硬伤", len(hard) == 0,
                f"数字硬伤 {len(hard)} 项"
                + (f"；另有待人工核对 {len(todo)} 项（非硬伤）" if todo else "；无待核对项"),
                len(hard), "number_consistency.py")
        except Exception as e:
            add(5, "数字一致性无硬伤", None, "该项无法评估（number_consistency 取数失败）", None,
                "number_consistency.py", unavailable=f"{type(e).__name__}: {e}")

    # ================= ⑥ 术语无混用 =================
    tc = mods.get("tc")
    if tc is None:
        add(6, "术语无混用", None, "该项无法评估（term_consistency 不可用）", None,
            "term_consistency.py",
            unavailable=load_errors.get("term_consistency", "term_consistency 不可用"))
    else:
        try:
            try:
                cdir = tc.corpus.find_corpus_dir(corpus_dir)
            except FileNotFoundError:
                cdir = None
            stats = tc.corpus_term_stats(cdir) if cdir else {}
            hits, _regions = tc.scan_doc(doc, stats)
            drifts, hints = [], []
            for g in tc.SYNONYM_GROUPS:
                present = [w for w in g if hits.get(w)]
                if not present:
                    continue
                # 限定形优先：已用限定形（如「信息可及性」）时裸词组不参与，避免重复报
                if g and g[0] in ("无障碍", "适老化", "视障", "老年人"):
                    qualified = [w for grp in tc.SYNONYM_GROUPS for w in grp
                                 if len(w) > len(g[0]) and g[0] in w and hits.get(w)]
                    if qualified:
                        continue
                pref = tc.preferred_term(g, stats) if stats else None
                real = list(present)
                for a in present:
                    for b in present:
                        if a != b and tc.is_not_synonym(a, b):
                            if a in real:
                                real.remove(a)
                            if b in real:
                                real.remove(b)
                if len(real) >= 2:
                    if not pref:
                        pref = max(real, key=lambda w: sum(len(v) for v in hits[w].values()))
                    for w in real:
                        if w != pref:
                            drifts.append((w, pref))
                elif len(present) == 1 and pref and present[0] != pref:
                    s = stats.get(present[0]) or {"papers": 0, "hits": 0}
                    ps = stats.get(pref) or {"papers": 0, "hits": 0}
                    if ps["papers"] >= 5 and s["papers"] <= 2:
                        hints.append((present[0], pref))
            ctx["tc_hints"] = hints
            ctx["tc_drifts"] = drifts
            add(6, "术语无混用", len(drifts) == 0,
                f"同义异形词混用 {len(drifts)} 组"
                + (f"；另有非偏好写法 {len(hints)} 处（提示级）" if hints else "；无非偏好写法"),
                len(drifts), "term_consistency.py")
        except Exception as e:
            add(6, "术语无混用", None, "该项无法评估（term_consistency 取数失败）", None,
                "term_consistency.py", unavailable=f"{type(e).__name__}: {e}")

    # ================= ⑦ 官方投稿包要素齐备 =================
    if dp is None or official is None:
        add(7, "官方投稿包要素齐备", None, "该项无法评估（diagnose_paper 不可用）", None,
            "diagnose_paper.py", unavailable=ctx.get("diag_error") or "diagnose_paper 不可用")
    else:
        pn = official.get("package_notices", []) or []
        ctx["package_notices"] = pn
        missing = [it for it in PACKAGE_ITEMS
                   if any(it in x for x in pn)]
        add(7, "官方投稿包要素齐备", len(pn) == 0,
            ("作者简介 / 联系方式 / 定稿日期 三项齐备" if not pn
             else "待补：" + "、".join(missing or [x[:12] for x in pn])
                  + f"（{len(pn)} 项）"),
            len(pn), "diagnose_paper.py")

    # ---------- 附注：引用核验（不计入七项，作看板附注） ----------
    vr = mods.get("vr")
    if vr is not None:
        try:
            body_txt, refs = vr.split_body_refs(doc)
            items = vr.parse_ref_items(refs)
            if items:
                nums = [n for n, _ in items]
                cited = vr.collect_cited(body_txt)
                ctx["vr"] = {"ref_n": len(items),
                             "missing": sorted(set(range(min(nums), max(nums) + 1)) - set(nums)),
                             "dup": [n for n in set(nums) if nums.count(n) > 1],
                             "ghost": sorted(cited - set(nums)),
                             "unused": sorted(set(nums) - cited)}
        except Exception as e:
            ctx["vr_error"] = f"{type(e).__name__}: {e}"

    return rows, ctx


# ---------------- 修复建议 Top N ----------------
def top_advice(rows, ctx, top_n=3):
    """按「对分数的边际影响」排序：审稿失分最大的维度优先，其次红线。

    通道一（审稿失分）：直接取 mock_review 修改路线图的失分绝对值，降序——
        该工具已按失分排序，看板不重复实现排序逻辑。
    通道二（红线）：按该刊审稿关注度给等效权重（见 RED_WEIGHT），仅用于红线内部排序。
    通道三（其余未就绪项）：不占分数但阻断就绪判定，按指标序号排。
    """
    advice = []

    for loss, dim, op in ctx.get("roadmap") or []:
        if loss <= 0:
            continue
        advice.append({"weight": float(loss), "channel": "审稿失分",
                       "title": f"{dim}：失 {_fmt_gap(loss)} 分",
                       "body": op})

    red_detail = ctx.get("red_detail") or {}
    for key, n in sorted(red_detail.items(),
                         key=lambda kv: (-RED_WEIGHT.get(kv[0], 4.0), -kv[1])):
        if not n:
            continue
        advice.append({"weight": RED_WEIGHT.get(key, 4.0) * min(int(n), 3),
                       "channel": "红线",
                       "title": f"{RED_LABEL.get(key, key)} {n} 项",
                       "body": RED_HINT.get(key, "")})

    EXTRA = {
        3: "按 diagnose 的分栏目 P25 补齐文献——「文献支撑不足」是常见退修理由。",
        4: "替换全部【图位/待补】占位标记为真实图表，或删除占位符（table_fig_audit）。",
        5: "摘要—正文同名指标必须一致，另核对稿内样本量逻辑（number_consistency）。",
        6: "同义异形词全文统一为该刊偏好词（term_consistency 给篇数证据）。",
        7: "补齐作者简介 / 联系方式 / 定稿日期（dp.scan_official 的 package_notices）。",
    }
    for r in sorted(rows, key=lambda x: x["id"]):
        if r["ok"] is False and r["id"] in EXTRA:
            advice.append({"weight": 1.0, "channel": "未就绪项",
                           "title": f"{r['name']}（指标{r['id']}）", "body": EXTRA[r["id"]]})

    advice.sort(key=lambda a: (-a["weight"], a["channel"]))
    out, seen = [], set()
    for a in advice:
        if a["title"] in seen:
            continue
        seen.add(a["title"])
        out.append(a)
        if len(out) >= max(1, top_n):
            break
    return out


# ---------------- 输出 ----------------
def render(doc, rows, ctx, top_n):
    lines = []
    lines.append("=" * 72)
    lines.append(f"{DASHBOARD_NAME}看板（投稿就绪度 · 七项指标）")
    lines.append(f"稿件：{os.path.basename(doc)}")
    lines.append("=" * 72)

    evaluable = [r for r in rows if r["ok"] is not None]
    n_pass = sum(1 for r in evaluable if r["ok"])
    n_all = len(evaluable)
    unavailable = [r for r in rows if r["ok"] is None]

    # 总览
    if n_all:
        bar_n = 24
        filled = int(round(n_pass / n_all * bar_n))
        lines.append(f"\n达标总览：{n_pass}/{n_all}  [{'█' * filled}{'░' * (bar_n - filled)}]")
    else:
        lines.append("\n达标总览：无法计算（全部指标依赖缺失）")
    if n_all and n_pass == n_all and not unavailable:
        verdict = "全部就绪，可进入投稿流程"
    elif n_pass < n_all:
        verdict = "存在未就绪项，处理后再投"
    else:
        verdict = "指标无法评估，先补齐依赖脚本"
    lines.append(f"判定：{verdict}")

    # 逐项状态行
    lines.append("\n逐项状态（✓ 达标 / ✗ 未达标 / — 无法评估）")
    lines.append("-" * 72)
    units = {1: " 项", 2: " 分", 3: " 条", 4: " 处", 5: " 项", 6: " 组", 7: " 项"}
    for r in rows:
        lines.append(f" {r['mark']} {r['id']}. {r['name']}")
        if r["ok"] is None:
            lines.append(f"      状态：该项无法评估——{r['unavailable']}")
            lines.append(f"      取数：{r['source']}（import/取数失败，已降级）")
            continue
        lines.append(f"      现状：{r['detail']}")
        lines.append(f"      差距：{'0（已达标）' if r['ok'] else _fmt_gap(r['gap']) + units.get(r['id'], '')}")
        lines.append(f"      取数：{r['source']}")

    # 最优先修复建议
    advice = top_advice(rows, ctx, top_n)
    lines.append("\n" + "-" * 72)
    lines.append(f"最优先修复建议 Top{len(advice)}（按对模拟审稿分的边际影响排序：")
    lines.append("审稿失分最大的维度优先，其次红线；红线内部按该刊关注度）")
    lines.append("-" * 72)
    if advice:
        for i, a in enumerate(advice, 1):
            lines.append(f" {i}. [{a['channel']}] {a['title']}")
            lines.append(f"      {a['body']}")
    else:
        lines.append(" 无需修复项——七项全部达标。")

    # 附注
    notes = []
    vr = ctx.get("vr")
    if vr and (vr["missing"] or vr["dup"] or vr["ghost"]):
        notes.append(f"引用核验：序号缺号 {vr['missing'] or '无'} / 重号 {vr['dup'] or '无'} / "
                     f"引用但列表无 {vr['ghost'] or '无'}（verify_refs，建议一并处理）")
    if ctx.get("nc_todo"):
        notes.append(f"数字待人工核对 {ctx['nc_todo']} 项（模糊量/多候选，非硬伤，见 number_consistency）")
    if ctx.get("tc_hints"):
        notes.append(f"非偏好写法 {len(ctx['tc_hints'])} 处（提示级、非混用，见 term_consistency）")
    if unavailable:
        notes.append(f"{len(unavailable)} 项指标因依赖缺失无法评估（见上），已从达标分母剔除")
    if notes:
        lines.append("\n附注")
        lines.append("-" * 72)
        for n in notes:
            lines.append(f" · {n}")

    lines.append("\n" + "=" * 72)
    lines.append("口径说明：七项指标分别取 diagnose_paper / mock_review / number_consistency /")
    lines.append("term_consistency / table_fig_audit / verify_refs 的既有判定，本看板只做汇总与排序，")
    lines.append("不重复实现判定逻辑；细读请用 check_all.py 全量总检。")
    lines.append("=" * 72)
    return "\n".join(lines)


def main():
    args = sys.argv[1:]
    positional = [a for a in args if not a.startswith("--")]
    flags = {a for a in args if a.startswith("--")}
    if not positional or positional[0] in ("-h", "--help"):
        print(__doc__)
        return 0
    doc = positional[0]
    if not os.path.isfile(doc):
        print(f"找不到文件：{doc}")
        return 2

    corpus_dir = None
    if "--corpus" in args:
        i = args.index("--corpus")
        corpus_dir = args[i + 1] if i + 1 < len(args) else None
        if corpus_dir and not os.path.isdir(corpus_dir):
            print(f"[提示] --corpus 指定的目录不存在：{corpus_dir}，改按自动定位语料库。")
            corpus_dir = None

    top_n = 3
    if "--top" in args:
        i = args.index("--top")
        try:
            top_n = max(1, int(args[i + 1]))
        except (IndexError, ValueError):
            print("[提示] --top 需要一个正整数参数，已用默认 3。")

    try:
        rows, ctx = collect(doc, corpus_dir)
    except Exception as e:
        print(f"[错误] 看板取数失败：{type(e).__name__}: {e}")
        print("       请确认文件为 .docx 且未损坏；依赖脚本缺失时对应指标会标「无法评估」。")
        return 1

    print(render(doc, rows, ctx, top_n))

    if "--json" in flags:
        evaluable = [r for r in rows if r["ok"] is not None]
        out = {
            "doc": os.path.basename(doc),
            "dashboard": DASHBOARD_NAME,
            "review_target": REVIEW_TARGET,
            "summary": {
                "passed": sum(1 for r in evaluable if r["ok"]),
                "total": len(evaluable),
                "unavailable": sum(1 for r in rows if r["ok"] is None),
            },
            "rows": rows,
            "context": {k: v for k, v in ctx.items() if k != "load_errors"},
            "load_errors": ctx.get("load_errors", {}),
            "advice": top_advice(rows, ctx, top_n),
            "rc": 0 if (evaluable and all(r["ok"] for r in evaluable)
                        and len(evaluable) == len(rows)) else 1,
        }
        jf = os.path.splitext(doc)[0] + ".progress.json"
        with io.open(jf, "w", encoding="utf-8") as f:
            f.write(json.dumps(out, ensure_ascii=False, indent=1))
        print(f"\nJSON 已写入：{jf}")

    evaluable = [r for r in rows if r["ok"] is not None]
    all_ready = bool(evaluable) and all(r["ok"] for r in evaluable) and len(evaluable) == len(rows)
    return 0 if all_ready else 1


if __name__ == "__main__":
    sys.exit(main())
