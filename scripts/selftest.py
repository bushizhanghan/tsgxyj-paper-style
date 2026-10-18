# -*- coding: utf-8 -*-
"""selftest.py —— scripts/ 工具链自检（修复即断言）

覆盖：corpus 定位/降级、diagnose 指标计算、结构/摘要体检、
verify_refs 核验逻辑、search 检索逻辑。不依赖真实投稿稿，
用合成 docx + 合成语料验证。

用法: python selftest.py            # 全量
      python selftest.py -v         # 打印每个用例
退出码 0 = 全过。
"""
import os
import sys
import re
import json
import tempfile
import io

HERE = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, HERE)

PASS, FAIL = 0, 0
VERBOSE = "-v" in sys.argv


def check(name, cond, detail=""):
    global PASS, FAIL
    if cond:
        PASS += 1
        if VERBOSE:
            print(f"  ✓ {name}")
    else:
        FAIL += 1
        print(f"  ✗ {name} {detail}")


# ============ 1. corpus.py ============
print("[1] corpus 定位与降级")
import corpus  # noqa: E402

with tempfile.TemporaryDirectory() as td:
    # 无标志文件的目录 → None
    check("空目录→None", corpus._is_corpus_dir(td) is False)
    # 只放 paragraphs.json → 识别为语料库
    open(os.path.join(td, "paragraphs.json"), "w", encoding="utf-8").write("[]")
    check("仅 paragraphs.json→True", corpus._is_corpus_dir(td) is True)
    check("load_paragraphs 空列表", corpus.load_paragraphs(td) == [])
    check("load_corpus_stats 缺文件→[]", corpus.load_corpus_stats(td) == [])
    st_ = corpus.load_reference_stats(td)
    check("无语料库降级静态基线", st_["static"] is True and st_["paras"]["med"] == 206)
    # 坏目录的 --corpus 应抛错
    try:
        corpus.find_corpus_dir(cli_corpus=os.path.join(td, "no", "such"))
        check("坏 --corpus 应抛 FileNotFoundError", False)
    except FileNotFoundError:
        check("坏 --corpus 应抛 FileNotFoundError", True)

# ============ 2. diagnose 指标与体检 ============
print("[2] diagnose_paper 指标与体检")
import diagnose_paper as dp  # noqa: E402

body = [
    "本研究旨在检验某机制的运行效果，采用混合方法开展两轮实验，获得了三百余份有效问卷数据。结果表明，该机制能够显著改善用户的可及性体验，并对不同类型障碍群体产生差异化影响，据此提出优化路径。",
    "效果很好。用户喜欢。数据支持。反应不错。结果满意。体验良好。",   # 句均极短
]
m = dp.compute_metrics(body, ["[1] 甲. 某文献[J]. 某刊, 2020."],
                       "针对某现象，本研究以某群体为对象，采用某方法，发现某结论，提供参考。",
                       "关键词：A；B；C", ["0 引言", "1 研究设计", "5 对策与启示", "6 结语"])
check("正文计字数", m["body_chars"] > 90, m["body_chars"])
check("关键词=3", m["kw_n"] == 3, m["kw_n"])
check("文献=1", m["ref_n"] == 1)
check("标题=4", len(m["titles"]) == 4)

# 反模式
anti = dp.scan_anti_patterns("这具有重要意义。这不仅体现了价值，而且标志着进步。综上所述，在此基础上。")
names = {a["name"]: a for a in anti}
check("万能句命中", names["万能句（具有重要意义/值得深入探讨/提供有益参考）"]["n"] >= 1,
      list(names))
check("不仅…而且命中", names["「不仅…而且」（该刊罕用，主流是更/还）"]["n"] >= 1, list(names))
check("空转连接命中", names["空转连接（综上/在此基础上/与此同时）"]["n"] >= 2, list(names))
# A 级红线
anti2 = dp.scan_anti_patterns("显而易见，该结论成立。众所周知这是共识。总而言之意义重大。")
n2 = {a["name"]: a for a in anti2}
check("A级红线句命中", n2["A级红线句（显而易见/众所周知/总而言之等）"]["n"] >= 3, list(n2))

# 段落级
issues = dp.scan_paragraphs(body)
check("碎段检出", any(i["chars"] < 110 and "碎段" in i["tags"] for i in issues), issues)
short_para = ("系统运行稳定。" * 20)   # 100 字、句均 6 字
issues2 = dp.scan_paragraphs([short_para])
check("句均偏短检出", any("偏短" in i["tags"] for i in issues2), issues2)
over_para = "这是一个很长的段落内容。" * 40   # 640 字超长段
check("超长段检出", any("超长段" in i["tags"] for i in dp.scan_paragraphs([over_para])))

# 结构完备度
s_ok = dp.scan_structure(["0 引言", "1 文献综述", "2 研究设计", "3 结果与分析",
                          "4 对策与启示", "5 结语"])
check("六标配全检出", len(s_ok["missing"]) == 0, s_ok["missing"])
s_bad = dp.scan_structure(["0 引言", "1 研究设计", "2 系统实现", "3 验证"])
check("缺对策/结语报出", "对策/策略/路径/启示（收口标配）" in s_bad["missing"]
      and "结语" in s_bad["missing"], s_bad["missing"])
s_h3 = dp.scan_structure(["1 绪论", "1.1 背景", "1.1.1 细节"])
check("三级标题计数", s_h3["h3_count"] == 1, s_h3["h3_count"])
# 三级标题常态（corpus 实测 85/108 篇含三级）：少量不得提示问题
s_h3_ok = dp.scan_structure(["0 引言", "1 文献综述", "2 研究设计", "2.1 数据",
                             "2.1.1 来源", "2.1.2 清洗", "3 对策与启示", "4 结语"])
check("三级标题为该刊常态(计数不报错)", s_h3_ok["h3_count"] == 2, s_h3_ok)
check("三级标题阈值 22", dp.H3_ALERT_THRESHOLD == 22, dp.H3_ALERT_THRESHOLD)

# 章节占比（scan_sections）
_fake_paras = [
    ("Heading", "0 引言"),
    ("", "引言内容。" * 30),
    ("Heading", "1 研究设计"),
    ("", "设计内容。" * 100),
    ("Heading", "2 结语"),
    ("", "结语内容。" * 10),
]
_secs = dp.scan_sections(_fake_paras, ["0 引言", "1 研究设计", "2 结语"])
check("章节切分可用", _secs is not None and _secs["total"] > 0, _secs)
if _secs:
    check("引言占比可算", "引言" in _secs["merged"], _secs["merged"])

# 摘要体检
ab_ok = dp.scan_abstract("针对某现象，本研究以某群体为对象，采用某方法构建了模型。"
                         "结果表明可分为三类。据此提出对策，为实践提供参考。")
check("摘要四要素全中", all(ab_ok["elements"].values()), ab_ok["elements"])
ab_bad = dp.scan_abstract("本文分析了某问题，提出了某建议。")
check("摘要忌词检出", any("忌词" in i for i in ab_bad["issues"]), ab_bad["issues"])
check("摘要偏短检出", any("偏短" in i for i in ab_bad["issues"]), ab_bad["issues"])
check("无摘要返回 present=False", dp.scan_abstract("")["present"] is False)

# ============ 3. verify_refs 核验逻辑 ============
print("[3] verify_refs 核验逻辑")
import verify_refs as vr  # noqa: E402

items = vr.parse_ref_items(["[1] 甲. 某文献[J]. 某刊, 2020.", "[2] 乙. 另一文献[M]. 北京: 某社, 2019.",
                            "[3] 丙. 第三条[J]. 某刊, 2021."])
check("解析3条", len(items) == 3 and items[0][0] == 1)
cited = vr.collect_cited("有研究指出该观点[1-3]。另有文献支持[2]。还有[5,6]。")
check("范围式展开1-3", {1, 2, 3} <= cited, cited)
check("列举式解析", {5, 6} <= cited, cited)
cited2 = vr.collect_cited("单独引用[7]")
check("单引解析", 7 in cited2, cited2)

# ============ 3b. ref_format GB/T 7714 ============
print("[3b] ref_format GB/T 7714 要素核验")
import ref_format as rf  # noqa: E402

good = "[6] 束漫, 孙蓓. 图书馆阅读障碍症群体服务的理论与实践[J]. 中国图书馆学报, 2014, 40(4): 92-99."
it = rf.parse_item(good)
check("解析编号6", it["num"] == 6, it["num"])
check("解析类型J", it["type"] == "J", it["type"])
check("解析年份2014", it["year"] == 2014, it["year"])
check("解析作者", "束漫" in it["authors"], it["authors"])
bad_items = [
    ("[7] 某作者. 某题名. 某刊, 2014.", "缺文献类型标识"),      # 无 [J]
    ("[8] 某作者. 某题名[J].", "缺年份"),                        # 无年份
    ("[9] 某题名[J]. 某刊, 2014, 40(4): 99-92.", "页码区间颠倒"),  # 页码颠倒
    ("[10] 某作者. 某题名[J]. 某刊, 2099, 40(4): 1-10.", "年份可疑"),
    ("[11] 某机构. 某报告[EB/OL]. https://example.com. (2020-01-01).", "缺年份"),
]
for raw, kind in bad_items:
    it = rf.parse_item(raw)
    got = []
    if not it["type"]:
        got.append("缺文献类型标识")
    if it["year"] is None:
        got.append("缺年份")
    elif it["year"] < 1900 or it["year"] > rf.CUR_YEAR + 1:
        got.append("年份可疑")
    pm = rf.find_page_inversion(it["rest"])
    if pm and int(pm[0]) > int(pm[1]):
        got.append("页码区间颠倒")
    check(f"检出「{kind}」", kind in got, (raw, got))
# 误报防护：日期/标准号/访问日期不得判为页码颠倒；URL 内年份不当年份
for raw, why in [
    ("[12] 某机构. 某法[Z/OL]. (2023-06-28)[2026-09-01]. https://x.com.", "日期与访问日期"),
    ("[13] 信息技术 某方法: GB/T 37668—2019[S]. 北京: 某社, 2019.", "标准号"),
]:
    it = rf.parse_item(raw)
    check(f"{why}不误报页码颠倒", rf.find_page_inversion(it["rest"]) is None,
          rf.find_page_inversion(it["rest"]))
it = rf.parse_item("[11] 某机构. 某报告[EB/OL]. https://example.com. (2020-01-01).")
check("访问日期不当作出版年", it["year"] != 2020 or "2020" not in it["rest"],
      it["year"])

# ============ 4. search 检索逻辑 ============
print("[4] search 检索逻辑")
import search as sr  # noqa: E402
check("AND 逻辑", sr.match(["无障碍", "服务"], False, "图书馆无障碍服务研究") is True)
check("AND 缺词", sr.match(["无障碍", "老年"], False, "图书馆无障碍服务研究") is False)
sr._or_mode = True
check("OR 逻辑", sr.match(["老年", "视障"], False, "图书馆视障服务研究") is True)
check("OR 缺词", sr.match(["老年", "儿童"], False, "图书馆视障服务研究") is False)
sr._or_mode = False
sr.keys = ["无障碍"]
sr.sys_re_args = [r"无障碍.{0,3}服务"]
loc, ctx = sr.build_hit_block({"no": "110", "issue": "2022-01", "file": "某标题_甲.pdf"},
                              3, "这是图书馆无障碍服务的一段文字。", 10, False)
check("锚点上下文含关键词", "无障碍" in ctx, ctx)
check("位置标注格式", loc.startswith("[110|2022-01|甲]"), loc)
check("作者提取", sr.author_of("某标题_张三.pdf") == "张三")

# ============ 4b. analyze_corpus 模式挖掘 ============
print("[4b] analyze_corpus 模式挖掘")
import analyze_corpus as ac  # noqa: E402

# 全角归一化
check("全角归一化", ac.norm("２０２６年【１】") == "2026年【1】", ac.norm("２０２６年【１】"))
# 段首分类
_d = ac.dim_openings([{"no": "1", "file": "a.pdf", "issue": "2026-01",
                       "paras": ["这是直接陈述起句的一段正文，字数足够长以便通过过滤条件，继续写一些内容。" * 2,
                                 "随着技术发展，图书馆服务发生根本性转变，传统服务模式已难以适应新时代的发展需求，亟须从理念到手段进行全面重构与系统革新。"]}])
check("段首分类计数", _d["n_paras"] == 2, _d["n_paras"])
check("直接陈述命中", any(k == "其他/直接陈述" for k, _ in _d["dist"]), _d["dist"])
check("现象切入命中", any(k == "现象/背景切入" for k, _ in _d["dist"]), _d["dist"])
# 引用统计口径（句末 vs 句中）
_c = ac.dim_citations([{"no": "1", "file": "a.pdf", "issue": "2026-01",
                        "paras": ["研究表明某观点成立[3]，另有学者支持该结论[4]。还有研究指出其他机制[5]。"
                                  "本段长度补充一些文字以确保通过最短段落过滤条件的限制要求，满足统计口径的基本门槛。"]}])
check("引用总数3", _c["tail"]["cite_total"] == 3, _c["tail"])
check("句末引用2处([4][5]句末、[3]句中)", _c["tail"]["sent_end"] == 2, _c["tail"])
# 数据句
_da = ac.dim_data([{"no": "1", "file": "a.pdf", "issue": "2026-01",
                    "paras": ["结果显示，412个候选词整合为148个主题词，其中第一类占比35%，为最高。"
                              "补充若干文字使本段长度达到统计过滤条件的最低门槛要求，确保模式匹配能够正常执行。"]}])
check("数据句识别", sum(v for _, v in _da["dist"]) >= 3, _da["dist"])
# 标题句式（语料库存在时）
_t = ac.dim_titles()
check("标题三级占比数值", isinstance(_t["h3_ratio"], float) and 20 < _t["h3_ratio"] < 50,
      _t["h3_ratio"])
# 摘要提取（语料库存在时）
_cd = corpus.find_corpus_dir()
if _cd:
    _abs = ac.load_abstracts(_cd)
    check("摘要提取>100篇", len(_abs) > 100, len(_abs))
    _ab = ac.dim_abstracts(_cd)
    check("摘要四要素率齐全", len(_ab["elem_rate"]) == 4, _ab["elem_rate"])
    check("摘要长度中位>200", _ab["len_med"] > 200, _ab["len_med"])

# ============ 4c. polish_lexicon 润色词典 ============
print("[4c] polish_lexicon 润色词典")
import polish_lexicon as pl  # noqa: E402

_v = pl.dim_vocab([{"no": "1", "file": "a.pdf", "issue": "2026-01",
                   "paras": ["大约有若干内容使得本研究得以开展，其中大部分工作基本完成，"
                             "大量样本被广泛应用于各类分析场景之中，能够持续提升整体效能水平。"]}])
_rows = {r["colloquial"]: r for r in _v["rows"]}
check('口语词表含「差不多」', "差不多" in _rows)
check('强断言词表含「显著」', "显著" in _v["strong"], list(_v["strong"])[:3])
check('审慎词表含「逐步」', "逐步" in _v["hedge"])
check("判定逻辑(刊内0次→替换)", _rows["差不多"]["verdict"] == "该刊几乎不用，改写应替换",
      _rows["差不多"]["verdict"])

# 段落级扫描：罕用词 + 「了」密度（润色词典接入）
_p1 = dp.scan_paragraphs(["这个系统搞得很好，特别给力。" * 6])
check("罕用词检出", any("罕用词" in i["tags"] for i in _p1), _p1)
_p2 = dp.scan_paragraphs([
    "该系统提升了资源检索的效率。平台完善了个性化推荐的功能模块。在线服务解决了远程访问的障碍问题。"
    "帮扶机制优化了工单流转的整体流程。智能模型改进了问答匹配的精度水平。系统扩展了多终端适配能力。"
    "整体运营改善了用户的可及性体验，各项指标均达到了预期目标，为后续迭代奠定了坚实基础。"])
check("「了」密度检出", any("「了」×" in i["tags"] for i in _p2), _p2)

# 句法特征统计口径
_syn = pl.dim_syntax([{"no": "1", "file": "a.pdf", "issue": "2026-01",
                       "paras": ["该理论被称为建构主义，即社会现实通过互动得以建构；"
                                 "研究旨在揭示其机制，样本由研究者来确定，共获得有效数据三百余条。"]}])
check("句法含被动统计", "被动「被」" in dict(_syn["counts"]), dict(_syn["counts"]))
check("句法含「即」解释句", any("即" in k for k, _ in _syn["counts"]), dict(_syn["counts"]))
check("句法含「旨在」", any("旨在" in k for k, _ in _syn["counts"]), dict(_syn["counts"]))
check("per_1k 数值化", all(isinstance(v, float) for v in _syn["per_1k"].values()))

# ============ 4d. citation_analysis 引文分析 ============
print("[4d] citation_analysis 引文分析")
import citation_analysis as ca  # noqa: E402

# 归一化（全角字母数字标点）
check("引文归一化(全角字母)", ca.norm("［Ｊ］２０２５（３）") == "[J]2025(3)",
      ca.norm("［Ｊ］２０２５（３）"))
# 条目解析
_items = ca.parse_refs("参考文献[1]甲. 某题名[J]. 某刊, 2020(3): 1-9.[2]乙. 另一本[M]. 北京: 某社, 2019.")
check("引文条目切分", len(_items) == 2 and _items[0][0] == 1, _items)
# 年份提取（含排除访问日期）
check("年份提取[J]", ca.ref_year("甲. 题名[J]. 某刊, 2020(3): 1-9.") == 2020)
check("年份提取[EB/OL]排除访问日期",
      ca.ref_year("某法[Z/OL]. (2023-06-28)[2026-09-01]. https://x.com.") == 2023,
      ca.ref_year("某法[Z/OL]. (2023-06-28)[2026-09-01]. https://x.com."))
check("年份提取无年份", ca.ref_year("某文献. 无年份著录.") is None)
# 类型识别
check("类型J", ca.ref_type("某题名[J]. 某刊, 2020.") == "J")
check("类型EB/OL", ca.ref_type("某报告[EB/OL]. https://x.com.") == "EB/OL")
check("类型无标识", ca.ref_type("某题名. 某刊, 2020.") == "无标识")
# 来源提取
check("来源J刊名", "图书馆学研究" in (ca.ref_source("甲. 题名[J]. 图书馆学研究, 2020(3): 1-9.", "J") or ""),
      ca.ref_source("甲. 题名[J]. 图书馆学研究, 2020(3): 1-9.", "J"))
check("来源M出版者", "商务印书馆" in (ca.ref_source("乙. 题名[M]. 北京: 商务印书馆, 2019.", "M") or ""),
      ca.ref_source("乙. 题名[M]. 北京: 商务印书馆, 2019.", "M"))
check("来源EB/OL域名去www",
      ca.ref_source("某网[EB/OL]. https://www.example.gov.cn/a/b.htm.", "EB/OL") == "example.gov.cn",
      ca.ref_source("某网[EB/OL]. https://www.example.gov.cn/a/b.htm.", "EB/OL"))
# 语料库基准（存在时）
_cd2 = corpus.find_corpus_dir()
if _cd2:
    _cs = ca.corpus_stats(_cd2)
    check("基准篇数>100", _cs and _cs["n_papers"] > 100, _cs and _cs["n_papers"])
    check("基准文献数>3000", _cs and _cs["n_refs"] > 3000, _cs and _cs["n_refs"])
    check("基准含J类占比", _cs and "J" in _cs["types"], _cs and list(_cs["types"])[:3])
    check("基准时效中位合理", _cs and 2000 < (_cs["year_med_med"] or 0) <= 2026,
          _cs and _cs["year_med_med"])
    check("高频来源含图情刊", _cs and any("图书馆" in s or "情报" in s
                                       for s, _ in _cs["top_sources"][:10]),
          _cs and [s for s, _ in _cs["top_sources"][:5]])

# ============ 4e. 回归保护（编码/跳过逻辑/空数据/私有路径） ============
print("[4e] 回归保护")
import glob as _glob

# 每个 subprocess.run 调用都应带 errors=replace（避免 UnicodeDecodeError 让 r.stdout 变 None）
_st = io.open(os.path.join(HERE, "selftest.py"), encoding="utf-8").read()
_calls = _st.count("subprocess.run(")
_safe = _st.count('encoding="utf-8", errors="replace"')
check("selftest subprocess 全带 errors=replace", _calls > 0 and _calls <= _safe,
      (_calls, _safe))

# check_all 跳过逻辑按脚本名而非显示序号
_ca = io.open(os.path.join(HERE, "check_all.py"), encoding="utf-8").read()
check("check_all 按脚本名跳过", "SKIP_SCRIPTS" in _ca and 'startswith("\u2464")' not in _ca)
check("check_all 有 quick 白名单", "QUICK_SCRIPTS" in _ca)

# benchmark 空段落保护
_bm = io.open(os.path.join(HERE, "benchmark_paper.py"), encoding="utf-8").read()
check("benchmark 防空段落崩溃", "if tl and lens:" in _bm and "\u65e0\u53ef\u7edf\u8ba1\u6bb5\u843d" in _bm)

# corpus 不含本机私有路径，支持用户配置
_cp = io.open(os.path.join(HERE, "corpus.py"), encoding="utf-8").read()
check("corpus 无本机私有路径", "\u5927\u521b\u9879\u76ee" not in _cp and "~/Desktop/" not in _cp)
check("corpus 支持用户配置文件", "_USER_CONFIG" in _cp and "tsgxyj_corpus" in _cp)

# 除 selftest 自身外，所有脚本应有 stdout UTF-8 保护
_miss = []
for _f in sorted(_glob.glob(os.path.join(HERE, "*.py"))):
    if os.path.basename(_f) == "selftest.py":
        continue
    _src = io.open(_f, encoding="utf-8").read()
    if 'reconfigure(encoding="utf-8"' not in _src:
        _miss.append(os.path.basename(_f))
check("所有脚本带 UTF-8 stdout 保护", not _miss, _miss)

# ============ 4f. 官方须知 + 反哺闭环 ============
print("[4f] 官方须知与反哺闭环")
# 官方摘要忌词（比此前更全：+我们/总结了/解决了）
_ab_bad = dp.scan_abstract("本文分析了某问题，我们总结了经验，提出了建议，解决了难题。")
_bad_txt = " ".join(_ab_bad["issues"])
check("官方忌词全六项", all(w in _bad_txt for w in ["本文", "我们", "分析了", "总结了", "提出了", "解决了"]),
      _ab_bad["issues"])
check("官方忌词计入issues", len(_ab_bad["issues"]) >= 6, len(_ab_bad["issues"]))
# 官方关键词语法
_off = dp.scan_official("关键词：研究；发展；视障服务", [], ["0 引言", "5 结语"], "视障用户的无障碍服务需求")
check("关键词泛词检出", any("泛词" in i for i in _off["issues"]), _off["issues"])
check("栏目对位非空", bool(_off["columns"]), _off["columns"])
check("读者工作优先匹配", _off["columns"][0] == "读者工作", _off["columns"][:3])
# 官方九栏齐全
check("官方九栏定义", len(dp.OFFICIAL_COLUMNS) == 9, len(dp.OFFICIAL_COLUMNS))
# 英文题名大写（ref_format）
_rf2 = rf.parse_item("[1] KOTIS K. Emerging Technologies in Smart Libraries[J]. Journal, 2025.")
_iss, _tc, _dp, _lat = rf.check_items([_rf2])
check("英文题名大写键存在", "en_title_issues" in _lat, list(_lat))
# 反哺闭环：已见刊/投稿稿判定
import feedback_loop as fb  # noqa: E402
_cd3 = corpus.find_corpus_dir()
if _cd3:
    _pub, _why = fb.detect_published("某论文_2026(1).docx", _cd3)
    check("年期次→判为已见刊", _pub is True, _why)
    _pub2, _why2 = fb.detect_published("我的投稿稿.docx", _cd3)
    check("无年期次→判为投稿稿", _pub2 is False, _why2)
# 闭环强弱项提取
_base2 = corpus.load_reference_stats(_cd3) if _cd3 else None
if _base2:
    _m = {"para_med": _base2["paras"]["med"], "sent_mean": _base2["sents"]["med"],
          "cite_rate": _base2["cites"]["med"], "abs_len": _base2["abs"]["med"],
          "ref_n": _base2["refs"]["med"], "le_per_1k": 1.0, "bei_per_1k": 0.3,
          "anti_flagged": [], "abs_elements": {"目的": True, "方法": True,
                                               "结果": True, "价值": True},
          "abs_issues": []}
    _s, _w = fb.strengths_vs_base(_m, _base2)
    check("达标稿→多长处少短板", len(_s) >= 4 and len(_w) <= 1, (_s, _w))
    _m2 = dict(_m)
    _m2.update({"para_med": 50, "ref_n": 5, "anti_flagged": ["三连×化"]})
    _s2, _w2 = fb.strengths_vs_base(_m2, _base2)
    check("差稿→检出短板", len(_w2) >= 2, _w2)

# ============ 5. 合成 docx 端到端 ============
print("[5] 合成 docx 端到端（可选，缺 python-docx 则跳过）")
print("[5] 合成 docx 端到端（可选，缺 python-docx 则跳过）")
try:
    from docx import Document
    HAVE_DOCX = True
except ImportError:
    HAVE_DOCX = False

if HAVE_DOCX:
    with tempfile.TemporaryDirectory() as td:
        p = os.path.join(td, "t.docx")
        doc = Document()
        doc.add_paragraph("摘要：针对某现象，本研究以某群体为对象，采用某方法构建了模型。结果表明可分为三类。据此提出对策，为实践提供参考。")
        doc.add_paragraph("关键词：A；B；C；D")
        doc.add_paragraph("0 引言")
        doc.add_paragraph("这是引言的正文段落，围绕研究背景与缺口展开，字数足够长以通过最短段落过滤条件，确保被计入统计口径。")
        doc.add_paragraph("1 文献综述")
        doc.add_paragraph("既有研究为本研究奠定了基础，但在特定群体与场景方面仍有拓展空间，本章梳理相关线索并指出缺口所在。")
        doc.add_paragraph("2 研究设计")
        doc.add_paragraph("本章说明研究对象与方法，交代样本规模与时间范围，并给出研究思路与指标设计，内容长度满足统计要求。")
        doc.add_paragraph("3 结果与分析")
        doc.add_paragraph("结果显示样本可分为三类，第一类占比百分之三十，第二类占比百分之四十五，其余为第三类，差异具有统计显著性。")
        doc.add_paragraph("4 对策与启示")
        doc.add_paragraph("针对上述发现，应完善机制设计，优化服务流程，加强馆员培训，并建立反馈闭环，从而提升整体可及性水平。")
        doc.add_paragraph("5 结语")
        doc.add_paragraph("本文构建了某模型并提出对策，但仍存在样本范围有限等局限，未来将扩大研究场景进一步验证。")
        doc.add_paragraph("参考文献")
        doc.add_paragraph("[1] 甲. 某文献[J]. 某刊, 2020.")
        doc.add_paragraph("[2] 乙. 另一文献[M]. 北京: 某社, 2019.")
        doc.save(p)

        paras, refs, abs_t, kw_t, titles, body2 = dp.extract_paras(p)
        check("端到端：标题6个", len(titles) == 6, titles)
        check("端到端：文献2条", len(refs) == 2, refs)
        check("端到端：摘要提取", len(abs_t) > 20, abs_t[:30])
        m2 = dp.compute_metrics(body2, refs, abs_t, kw_t, titles)
        check("端到端：正文>200字", m2["body_chars"] > 200, m2["body_chars"])
        s2 = dp.scan_structure(titles)
        check("端到端：结构无缺失", len(s2["missing"]) == 0, s2["missing"])
        ab2 = dp.scan_abstract(abs_t)
        check("端到端：摘要无红线", not [i for i in ab2["issues"] if "忌词" in i], ab2["issues"])

    # ---- 6. xref 域化全链路 ----
    print("[6] xref 域化全链路（build → verify）")
    saved_argv = sys.argv[:]
    with tempfile.TemporaryDirectory() as td:
        px = os.path.join(td, "x.docx")
        doc = Document()
        doc.add_paragraph("既有研究表明某观点[1]。另有学者支持[2]。还有研究[3]指出。范围式[2-3]在此。")
        doc.add_paragraph("参考文献")
        doc.add_paragraph("[1] 甲. 文献一[J]. 某刊, 2020.")
        doc.add_paragraph("[2] 乙. 文献二[M]. 北京: 某社, 2019.")
        doc.add_paragraph("[3] 丙. 文献三[J]. 某刊, 2021.")
        doc.save(px)

        before = "\n".join(pp.text for pp in Document(px).paragraphs)
        # build_xref / xref_verify 是脚本型模块（导入即执行），用 subprocess 测
        import subprocess
        r = subprocess.run([sys.executable, os.path.join(HERE, "build_xref.py"), px, "--dry-run"],
                           capture_output=True, encoding="utf-8", errors="replace")
        out = r.stdout
        check("dry-run 退出码0", r.returncode == 0, r.stderr[-200:])
        check("dry-run 报告范围式跳过", "范围式标注" in out, out[:200])
        check("dry-run 不写文件", "未写文件" in out)
        after_dry = "\n".join(pp.text for pp in Document(px).paragraphs)
        check("dry-run 文本未变", before == after_dry)

        # 正式 build
        r = subprocess.run([sys.executable, os.path.join(HERE, "build_xref.py"), px, "--backup"],
                           capture_output=True, encoding="utf-8", errors="replace")
        out = r.stdout
        check("build 退出码0", r.returncode == 0, r.stderr[-300:])
        check("build 生成3个书签", "文末书签: [1, 2, 3]" in out, out[:200])
        check("build 替换3个域", "正文替换域数: 3" in out, out[:200])
        check("--backup 生成 .bak.docx", os.path.isfile(os.path.splitext(px)[0] + ".bak.docx"))

        # verify
        bak = os.path.splitext(px)[0] + ".bak.docx"
        r = subprocess.run([sys.executable, os.path.join(HERE, "xref_verify.py"), px, bak],
                           capture_output=True, encoding="utf-8", errors="replace")
        out = r.stdout
        check("verify 退出码0", r.returncode == 0, r.stderr[-300:])
        check("verify 文本零变化", "全文文本是否一致: True" in out, out[:200])
        check("verify 域配平", "配平: True" in out)
        check("verify 引用均有书签", "引用编号是否都有书签: True" in out)
        check("verify 正文上标3处", "正文上标 [n] 计数: 3" in out, out[:400])
        check("verify 文末非上标3处", "文末编号保持非上标: 3" in out)

        # ---- 7. 退域 → 重建 往返（编辑前退域工作流） ----
        print("[7] unfield → rebuild 往返")
        # 当前 px 已域化；先退域
        r = subprocess.run([sys.executable, os.path.join(HERE, "unfield_refs.py"), px],
                           capture_output=True, encoding="utf-8", errors="replace")
        out = r.stdout
        check("unfield 退出码0", r.returncode == 0, r.stderr[-300:])
        check("unfield 折叠3个域", "折叠域数: 3" in out, out[:200])
        check("unfield 移除书签start", "移除书签元素: 3" in out, out[:200])
        xml_after = Document(px).element.xml
        check("退域后无 REF 域", "REF _Ref" not in xml_after)
        check("退域后无书签", "_Ref1" not in xml_after)
        txt_after = "\n".join(pp.text for pp in Document(px).paragraphs)
        check("退域后文本仍一致", txt_after == before, txt_after[:200])
        # 重建域
        r = subprocess.run([sys.executable, os.path.join(HERE, "build_xref.py"), px],
                           capture_output=True, encoding="utf-8", errors="replace")
        out = r.stdout
        check("rebuild 退出码0", r.returncode == 0, r.stderr[-300:])
        check("rebuild 恢复3个域", "正文替换域数: 3" in out, out[:200])
        xml_re = Document(px).element.xml
        check("rebuild 后 REF 域回来", "REF _Ref1" in xml_re)
        txt_re = "\n".join(pp.text for pp in Document(px).paragraphs)
        check("rebuild 后文本仍一致", txt_re == before)
        # 重建后再 verify 一次
        r = subprocess.run([sys.executable, os.path.join(HERE, "xref_verify.py"), px],
                           capture_output=True, encoding="utf-8", errors="replace")
        check("rebuild 后 verify 合格", "引用编号是否都有书签: True" in r.stdout, r.stdout[:300])
        sys.argv = saved_argv

    # ---- 8. 模拟审稿 + 整稿语言体检 ----
    print("[8] mock_review + language_audit（端到端）")
    with tempfile.TemporaryDirectory() as td:
        pv = os.path.join(td, "rv.docx")
        doc = Document()
        doc.add_paragraph("摘要：针对某现象，本研究以某群体为对象，采用某方法构建了模型。"
                          "结果表明可分为三类。据此提出对策，为实践提供参考。")
        doc.add_paragraph("关键词：A；B；C；D")
        doc.add_paragraph("0 引言")
        doc.add_paragraph("随着技术发展，图书馆服务发生根本转变。既有研究多聚焦单一群体，"
                          "尚缺乏对多元障碍群体的系统探讨，本研究旨在填补这一缺口。")
        doc.add_paragraph("1 研究设计")
        doc.add_paragraph("本研究以某群体为对象，采用问卷调查法，共发放问卷300份，"
                          "回收有效问卷286份，有效率为95.3%，调查时间为2023年3月至5月。")
        doc.add_paragraph("2 结果与分析")
        doc.add_paragraph("结果显示可分为三类。一是信息获取类，占比42.6%。"
                          "二是操作适配类，占比35.1%。三是情境支持类，占比22.3%。")
        doc.add_paragraph("3 对策与启示")
        doc.add_paragraph("针对上述发现，应完善适配机制。针对操作差异，应优化交互设计。"
                          "针对情境需求，应健全支持体系。")
        doc.add_paragraph("4 结语")
        doc.add_paragraph("本研究构建了三类需求框架并提出对策，但仍存在样本范围有限等局限。")
        doc.add_paragraph("参考文献")
        doc.add_paragraph("[1] 甲. 某文献[J]. 某刊, 2020, 40(3): 12-18.")
        doc.add_paragraph("[2] 乙. 另一文献[M]. 北京: 某社, 2019.")
        doc.save(pv)

        r = subprocess.run([sys.executable, os.path.join(HERE, "mock_review.py"), pv],
                           capture_output=True, encoding="utf-8", errors="replace")
        out = r.stdout
        check("mock_review 退出码0", r.returncode == 0, r.stderr[-400:])
        check("mock_review 有总分", "【总分】" in out and "/ 100" in out, out[:300])
        check("mock_review 十维以上齐全",
              all(d in out for d in ["选题与定位", "结构与篇幅", "方法与数据", "论证链",
                                     "文献质量", "引文规范", "体例规范", "语言与文风",
                                     "润色与语感", "净度", "图表质量", "结构创新度"]), out[:300])
        check("mock_review 有判定", "【判定】" in out)
        check("mock_review 有免责", "免责" in out and "模拟" in out)
        check("mock_review 有路线图", "修改路线图" in out)
        import re as _re
        m = _re.search(r"【总分】([\d.]+) / 100", out)
        check("mock_review 分数在合理区间", m and 0 <= float(m.group(1)) <= 100,
              m.group(1) if m else "无")
        check("mock_review 权重合计100", "100 分制" in out, out[:120])

        r = subprocess.run([sys.executable, os.path.join(HERE, "language_audit.py"), pv],
                           capture_output=True, encoding="utf-8", errors="replace")
        out = r.stdout
        check("language_audit 退出码0(A类无)", r.returncode == 0, r.stderr[-400:])
        check("language_audit 七类齐全",
              all(c in out for c in ["【A】", "【B】", "【C】", "【D】", "【E】", "【F】", "【G】"]),
              out[:200])
        check("language_audit 报告A类结论", "无该刊 0 次口语词" in out, out[-300:])

        # 植入问题句，验证能检出
        pv2 = os.path.join(td, "rv2.docx")
        doc2 = Document()
        doc2.add_paragraph("摘要：本项目做了些工作，取得了良好效果。")
        doc2.add_paragraph("关键词：A")
        doc2.add_paragraph("0 引言")
        doc2.add_paragraph("这个系统搞得挺不错的，差不多能解决大家的问题，咱们老百姓用得也靠谱。"
                           "该系统明显提升了效率，完全解决了所有问题，被广泛应用于各类场景。")
        doc2.add_paragraph("参考文献")
        doc2.add_paragraph("[1] 甲. 某文献. 某刊, 2020.")
        doc2.save(pv2)
        r = subprocess.run([sys.executable, os.path.join(HERE, "language_audit.py"), pv2],
                           capture_output=True, encoding="utf-8", errors="replace")
        out = r.stdout
        check("language_audit 检出A类口语词", r.returncode == 1 and "【A】该刊 0 次口语词" in out,
              out[:400])
        check("language_audit 检出「差不多/搞得/咱们/靠谱」",
              all(w in out for w in ["差不多", "搞得", "咱们", "靠谱"]), out[:600])
        check("language_audit 检出「被」被动", "【E】" in out and "被广泛" in out, out[:600])
else:
    print("  （跳过：未安装 python-docx）")

# ============ 汇总 ============
print()
print("=" * 46)
print(f"  selftest: 通过 {PASS} | 失败 {FAIL}")
print("=" * 46)
sys.exit(1 if FAIL else 0)
