# -*- coding: utf-8 -*-
"""ji_density.py —— 「即」解释句式密度核验

anti-patterns.md 第 24 条：「解释不用『即』——该刊 418 处『即』句是标准
解释手段，比破折号/括号更贴刊风」。language_audit 的 F 类从**缺失**角度查
（长解释未用「即」），本脚本从**密度分布**角度补：

该刊 118 篇 121 万字实测：「即」共 458 处 = **3.78/万字**。
「即」句密度显著低于此，说明作者的解释习惯偏离该刊（多用破折号/括号/
「也就是」等替代）；显著高于此，则要核对是否过度使用。

  · 密度 <1.5/万字 → 🟡 明显低于该刊（3.78），检查长解释是否都用了
    破折号/括号而非「××，即××」；
  · 密度 >8/万字 → 🟡 偏高，核对是否堆砌；
  · 同时报破折号解释与括号解释的数量作对照（该刊「即」458 处 vs
    长破折号/括号解释的量级），让作者看清自己用的是哪一路。

用法:
    python ji_density.py <论文.docx> [--json]
退出码：恒 0（提示级工具，解释手段选择属风格）。
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
import diagnose_paper as dp  # noqa: E402

# 该刊 121 万字实测
CORPUS_JI_RATE = 3.78        # 「即」/万字
CORPUS_JI_TOTAL = 458
LOW_LIMIT, HIGH_LIMIT = 1.5, 8.0


def analyze(doc):
    paras, refs, abs_text, kw_text, titles, body = dp.extract_paras(doc)
    abs_c = dp.clean(abs_text or "")
    body_c = dp.clean("\n".join(body))
    n_chars = len(abs_c) + len(body_c)
    if n_chars < 200:
        return {"rate": 0.0, "issues": [], "notices": ["正文过短，无法统计"]}

    n_ji = (abs_c + body_c).count("即")
    # 破折号解释与括号解释数量（作对照）
    n_dash = len(re.findall(r"—[^—]{10,80}", abs_c + body_c))
    n_par = len(re.findall(r"（[^）]{12,80}）", abs_c + body_c))

    rate = n_ji / n_chars * 10000
    issues, notices = [], []
    if rate < LOW_LIMIT:
        notices.append(f"「即」密度 {rate:.2f}/万字，明显低于该刊（{CORPUS_JI_RATE}，"
                       f"121 万字 {CORPUS_JI_TOTAL} 处）——检查长解释是否用了破折号/括号"
                       f"（本文破折号解释 {n_dash} 处、括号解释 {n_par} 处）；"
                       f"该刊标准解释手段是「××，即××」")
    elif rate > HIGH_LIMIT:
        notices.append(f"「即」密度 {rate:.2f}/万字，高于该刊（{CORPUS_JI_RATE}）"
                       f"较多——核对是否堆砌")

    return {"rate": round(rate, 2), "n_ji": n_ji, "n_dash": n_dash,
            "n_par": n_par, "n_chars": n_chars,
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
    print("「即」解释句式密度核验（该刊 121 万字实测基线）")
    print(f"稿件：{os.path.basename(DOC)}")
    print(f"「即」{res.get('n_ji', 0)} 处 = {res.get('rate', 0)}/万字"
          f"（该刊 {CORPUS_JI_RATE}/万字，共 {CORPUS_JI_TOTAL} 处）")
    print(f"对照：破折号解释 {res.get('n_dash', 0)} 处 ｜ 括号解释 {res.get('n_par', 0)} 处")
    print("=" * 74)

    if res["issues"]:
        print(f"🔴 发现 {len(res['issues'])} 项问题：")
        for it in res["issues"]:
            print(f"   · {it}")
    if res["notices"]:
        print(f"🟡 建议核对 {len(res['notices'])} 项：")
        for it in res["notices"]:
            print(f"   · {it}")
    if not res["issues"] and not res["notices"]:
        print("✅ 「即」密度在该刊常规区间。")

    if "--json" in flags:
        jf = os.path.splitext(DOC)[0] + ".ji_density.json"
        with open(jf, "w", encoding="utf-8") as f:
            json.dump(res, f, ensure_ascii=False, indent=1)
        print(f"\nJSON 已写入：{jf}")

    sys.exit(1 if res["issues"] else 0)


if __name__ == "__main__":
    main()
