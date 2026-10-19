# -*- coding: utf-8 -*-
"""meta_scale.py —— 摘要字数 / 最大段长 / 图表数量对标（三项合一）

三项都是 diagnose_paper 已采集但**未与该刊分布对照判定**的指标，
合并为一个工具一次跑完（避免重复读 docx）：

  · 摘要字数（该刊 107 篇实测：P10=201、P25=238、P50=262、P75=303、P90=339）：
    官方要求「200 字左右」；<150 字（仅 0.9% 篇目）显单薄，>350 字（4.7%）
    偏离主流。diagnose_paper 只对比中位，不判区间；
  · 篇内最大段长（该刊实测：P50=826、P75=1054、P90=1336，95.4% 篇目含
    >500 字的段）——**该刊并不排斥长段**，综述段/论证段可达千字。
    故 diagnose_paper 的「>450 字考虑切分」只作提示，此处给该刊真实分位
    作参照，避免作者按通用写作教条把该刊正常长段切碎；
  · 图表数量（该刊实测：图均值 2.4、表均值 4.3；无图篇目 13.9%、无表 16.7%）——
    实证论文有图有表是常态，但纯思辨型稿无图表亦属正常（该刊 1/7 篇目无图），
    故只在对标评估时报形态，不作硬性要求。

用法:
    python meta_scale.py <论文.docx> [--corpus <dir>] [--json]
退出码：0 = 三项均在常规区间；1 = 摘要字数越界（<150 或 >400）。
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

# 摘要字数（该刊 107 篇实测分位）
ABS_P10, ABS_P25, ABS_P50, ABS_P75, ABS_P90 = 201, 238, 262, 303, 339
ABS_MIN, ABS_MAX = 150, 400          # 硬区间（<150 仅 0.9% 篇目；>400 超 P90 较多）
# 篇内最大段长（该刊 108 篇实测分位）
MAXPARA_P50, MAXPARA_P75, MAXPARA_P90 = 826, 1054, 1336
# 图表（该刊 108 篇实测）
FIG_MEAN, TAB_MEAN = 2.4, 4.3
NO_FIG_RATE, NO_TAB_RATE = 0.139, 0.167


def analyze(doc):
    paras, refs, abs_text, kw_text, titles, body = dp.extract_paras(doc)
    m = dp.compute_metrics(body, refs, abs_text, kw_text, titles)

    abs_len = m["abs_len"]
    # compute_metrics 不返回 max_para（该键只在 corpus_stats.json 有），自行计算
    max_para = max((len(dp.clean(p)) for p in body), default=0)
    fig_n, tab_n = m["fig_n"], m["tab_n"]

    issues, notices = [], []

    # ① 摘要字数
    if abs_len and abs_len < ABS_MIN:
        issues.append(f"摘要 {abs_len} 字，低于该刊单篇常规下限（{ABS_MIN}；"
                      f"该刊 P10={ABS_P10}、P25={ABS_P25}）——报道性摘要需完整"
                      f"呈现目的/方法/结果/结论，过短难以承载")
    elif abs_len > ABS_MAX:
        issues.append(f"摘要 {abs_len} 字，超过该刊单篇常规上限（{ABS_MAX}；"
                      f"该刊 P90={ABS_P90}）——官方要求「200 字左右」，压缩冗余")

    # ② 最大段长（只作该刊分位参照，不作硬判）
    if max_para > MAXPARA_P90:
        notices.append(f"篇内最大段 {max_para} 字，超过该刊 P90（{MAXPARA_P90}）——"
                       f"该刊并不排斥长段（P50={MAXPARA_P50}），但超 P90 建议确认"
                       f"是否为多主题未切分（用 paragraph_structure 查）")
    elif max_para > MAXPARA_P50:
        notices.append(f"篇内最大段 {max_para} 字，高于该刊中位（{MAXPARA_P50}）——"
                       f"属该刊常见长段（综述/论证段），无需强行切分")

    # ③ 图表数量
    if fig_n == 0 and tab_n == 0:
        notices.append(f"全文无图表——该刊无图篇目仅 {NO_FIG_RATE:.0%}、无表 "
                       f"{NO_TAB_RATE:.0%}；若属实证研究建议补图表，纯思辨型稿可接受")
    else:
        if fig_n > 8:
            notices.append(f"图 {fig_n} 张偏多（该刊均值 {FIG_MEAN}）——确认是否必要")
        if tab_n > 10:
            notices.append(f"表 {tab_n} 张偏多（该刊均值 {TAB_MEAN}）——确认是否可合并")

    return {"abs_len": abs_len, "max_para": max_para, "fig_n": fig_n, "tab_n": tab_n,
            "issues": issues, "notices": notices}


def main():
    args = [a for a in sys.argv[1:] if not a.startswith("--")]
    flags = set(a for a in sys.argv[1:] if a.startswith("--"))
    if not args or args[0] in ("-h", "--help"):
        print(__doc__)
        sys.exit(0)
    DOC = args[0]
    if not os.path.isfile(DOC):
        sys.exit(f"找不到文件：{DOC}")

    res = analyze(DOC)
    print("=" * 74)
    print("摘要字数 / 最大段长 / 图表数量对标（该刊实测分位）")
    print(f"稿件：{os.path.basename(DOC)}")
    print(f"摘要 {res['abs_len']} 字（刊 P25–P90：{ABS_P25}–{ABS_P90}，中位 {ABS_P50}）")
    print(f"篇内最大段 {res['max_para']} 字（刊 P50–P90：{MAXPARA_P50}–{MAXPARA_P90}）")
    print(f"图 {res['fig_n']} 张（刊均值 {FIG_MEAN}）｜ 表 {res['tab_n']} 张"
          f"（刊均值 {TAB_MEAN}）")
    print("=" * 74)

    if res["issues"]:
        print(f"🔴 发现 {len(res['issues'])} 项问题：")
        for it in res["issues"]:
            print(f"   · {it}")
    else:
        print("✅ 摘要字数与图表数量在该刊常规区间。")
    if res["notices"]:
        print(f"\n🟡 建议核对 {len(res['notices'])} 项：")
        for it in res["notices"]:
            print(f"   · {it}")

    if "--json" in flags:
        jf = os.path.splitext(DOC)[0] + ".meta_scale.json"
        with open(jf, "w", encoding="utf-8") as f:
            json.dump(res, f, ensure_ascii=False, indent=1)
        print(f"\nJSON 已写入：{jf}")

    sys.exit(1 if res["issues"] else 0)


if __name__ == "__main__":
    main()
