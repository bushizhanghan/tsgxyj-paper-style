# -*- coding: utf-8 -*-
"""author_meta.py —— 作者信息形态与一致性核验（官方 §8）

官方投稿须知 §8 明文：

    「作者简介（包括姓名、单位、职称、学位，如为硕士生导师、博士生导师
      请注明，以及研究方向）
      示例：王悦欣 ＊＊＊大学副教授，硕士生导师，博士。研究方向：知识
      产权研究、专利分析等。通讯作者。
      如有通讯作者，请在该作者信息之后标明「通讯作者」。
      邮编及详细通信地址，电话，电子邮箱，定稿日期。」

submission_pack.py 已按 §8 做**要素存在性**清点（姓名/单位/职称/学位/
研究方向/导师/通讯作者逐项 ✓✗）。本脚本补另一半——**形态与一致性**，
这些正是编辑在初审里逐条挑的问题：

  · 姓名一致性：文首署名与作者简介块的姓名是否一致（含「张　瀚」与
    「张瀚」的全角空格差异、以及简介里写成英文名/缺姓等）；
  · 职称学位表述是否规范：官方示例为「副教授，硕士生导师，博士」式
    ——职称、导师身份、学位三者以逗号串接；本脚本查「教授博士」式
    粘连漏逗号、「博士，副教授」式语序颠倒、「副研究员，硕士」等
    学位与职称用词是否在该刊常见集合内；
  · 研究方向是否 ≥2 个方向：官方示例列 2 个（「知识产权研究、专利分析」），
    只写 1 个或「等」前为空的提示补全；
  · 单位是否为全称：示例用「＊＊＊大学」全称，检出「××大学」「某大学」
    等占位写法，以及「吉大」「北大」式简称；
  · 导师身份与通讯作者标注的形态提示。

用法:
    python author_meta.py <论文.docx> [--json]
退出码：0 = 形态与一致性无问题；1 = 检出姓名不一致、单位占位/简称、
        或研究方向仅 1 个。

说明：复用 submission_pack.bio_block 截取简介块（与其同一口径），
       不重复做存在性检查。
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
import diagnose_paper as dp  # noqa: E402
import submission_pack as spk  # noqa: E402

# 官方示例：王悦欣 ＊＊＊大学副教授，硕士生导师，博士。研究方向：…
# 职称（该刊常见集合）
TITLES = ["教授", "副教授", "讲师", "助理教授", "研究馆员", "副研究馆员",
          "馆员", "助理馆员", "研究员", "副研究员", "助理研究员",
          "高级工程师", "工程师", "副教授（青年）", "青年教授"]
# 学位
DEGREES = ["博士", "硕士", "学士", "博士研究生", "硕士研究生", "在读博士"]
# 导师身份
MENTORS = ["硕士生导师", "博士生导师", "硕导", "博导", "硕士导师", "博士导师"]

# 单位占位/简称写法（需改全称）
UNIT_PLACEHOLDER = [
    (r"\*{2,}\s*大学|\*{2,}", "星号占位（＊＊＊大学）"),
    (r"(?<![A-Za-z])××|某某|某大|某馆", "「××/某某」占位"),
    (r"^.{0,4}(?:吉大|北大|清华|人大|浙大|南大|武大|中大|川大|山大|"
     r"复旦|交大|同济|南开)$", "高校简称（应写全称）"),
]
# 单位全称特征词（出现即视为「像全称」）
UNIT_FULL_HINT = re.compile(r"大学|学院|图书馆|研究院|研究所|研究中心|"
                            r"委员会|办公厅|学会|出版社|有限公司|集团")
# 简介块内单位行（示例：单位名紧随姓名之后）
UNIT_IN_BLOCK = re.compile(r"([\u4e00-\u9fffA-Za-z]{2,25}?"
                           r"(?:大学|学院|图书馆|研究院|研究所|研究中心))")
# 研究方向行
RESEARCH_DIR_RE = re.compile(r"研究方向[：:]\s*(.+)$")
# 研究方向分隔符
DIR_SPLIT_RE = re.compile(r"[;；,，、/]\s*|\s{2,}")


def norm_name(s):
    """姓名归一化：去空白（含全角空格）、去常见称谓后缀。"""
    c = re.sub(r"[\s\u3000]+", "", s or "")
    c = re.sub(r"(教授|副教授|讲师|研究员|馆员|博士|硕士)$", "", c)
    return c


def cn_names_in(text):
    """文本中的中文姓名候选（2–4 字连续中文，排除常见职称学位词）。"""
    stop = {"大学", "学院", "图书馆", "研究", "通讯作者", "作者简介", "作者信息"}
    out = []
    for m in re.finditer(r"[\u4e00-\u9fff]{2,4}", text or ""):
        w = m.group(0)
        if w in stop:
            continue
        out.append(w)
    return out


def research_dirs(block):
    """研究方向列表（≥2 个为官方示例形态）。

    研究方向可能写在「研究方向：」标签后同句，也可能另起一段续写；
    这里在整块内取标签后的全部内容，再切到句末/下一标签。
    """
    m = RESEARCH_DIR_RE.search(block or "")
    if not m:
        return []
    seg = m.group(1)
    # 截到句末（或「通讯作者」标签前），避免吞掉后续段落
    seg = re.split(r"[。！？!?]", seg)[0]
    seg = re.split(r"(?:通讯作者|通信作者)", seg)[0]
    seg = seg.strip().rstrip("。.").strip()
    parts = [p.strip() for p in DIR_SPLIT_RE.split(seg) if p.strip()]
    parts = [re.sub(r"等$", "", p) for p in parts]
    return [p for p in parts if p]


def title_degree_shape(block):
    """职称—导师—学位串联形态检查，返回 (issues, notices, info)。"""
    issues, notices = [], []
    info = {"titles": [], "degrees": [], "mentors": []}
    for t in TITLES:
        if re.search(re.escape(t), block):
            info["titles"].append(t)
    for d in DEGREES:
        if re.search(re.escape(d), block):
            info["degrees"].append(d)
    for mt in MENTORS:
        if re.search(mt, block):
            info["mentors"].append(mt)

    # ① 粘连漏逗号：「副教授硕士生导师」式
    for m in re.finditer(r"([\u4e00-\u9fff]{2,8}?)(硕士生导师|博士生导师|硕导|博导)",
                         block):
        pre = m.group(1)
        if pre and not pre.endswith("，") and not pre.endswith(",") \
                and not pre.endswith("："):
            # 前面是职称词才算问题（如「副教授硕士生导师」）
            if any(pre.endswith(t) for t in TITLES):
                issues.append(f"「{pre}{m.group(2)}」职称与导师身份之间缺逗号——"
                              f"官方示例为「副教授，硕士生导师，博士」式串联")

    # ② 学位与职称粘连（「副教授博士」式）
    for m in re.finditer(r"(教授|副教授|讲师|研究馆员|副研究馆员|研究员|副研究员)"
                         r"(博士|硕士|学士)", block):
        issues.append(f"「{m.group(0)}」职称与学位之间缺逗号——官方示例为"
                      f"「副教授，硕士生导师，博士」式串联")

    # ③ 学位前置（「博士，副教授」式语序颠倒）
    if re.search(r"(博士|硕士)[,，]\s*(教授|副教授|讲师|研究员|副研究员|"
                 r"研究馆员|副研究馆员)", block):
        notices.append("学位写在职称之前（「博士，副教授」式）——官方示例"
                       "顺序为「职称，导师身份，学位」")

    # ④ 学位与职称用词不在该刊常见集合
    if block and not info["titles"] and not info["degrees"]:
        notices.append("简介块内未检出职称/学位词——请按官方示例写明"
                       "「职称，硕士生导师，学位」")

    # ⑤ 硕博导注明（官方：如为硕导/博导请注明）
    if info["degrees"] and "博士" in info["degrees"] and not info["mentors"]:
        notices.append("具博士学位但未注明是否硕士生/博士生导师——官方要求"
                       "「如为硕士生导师、博士生导师请注明」")

    return issues, notices, info


def unit_shape(block):
    """单位写法检查，返回 (issues, notices, unit)。"""
    issues, notices = [], []
    m = UNIT_IN_BLOCK.search(block or "")
    unit = m.group(1) if m else ""
    for pat, desc in UNIT_PLACEHOLDER:
        if re.search(pat, block or ""):
            issues.append(f"单位含{desc}——官方示例用全称「＊＊＊大学」，"
                          f"投稿须写真实全称")
    if unit and not UNIT_FULL_HINT.search(unit):
        notices.append(f"单位「{unit}」未见「大学/学院/图书馆/研究院」等全称"
                       f"特征词——确认是否为全称（官方示例：＊＊＊大学）")
    if unit and len(unit) <= 3 and UNIT_FULL_HINT.search(unit):
        notices.append(f"单位「{unit}」偏短——确认是否为规范全称")
    return issues, notices, unit


def analyze(doc):
    paras, refs, abs_text, kw_text, titles, body = dp.extract_paras(doc)
    # 复用 submission_pack 的简介块口径（文末 + 「作者简介」起 15 行）
    tail = spk.tail_text(paras)
    bblock = spk.bio_block(tail)

    issues, notices = [], []
    stat = {"bio_block": bblock[:80] if bblock else ""}

    if not bblock:
        # 存在性缺失由 submission_pack 报告，此处只说明形态核验无法进行
        return {"issues": issues,
                "notices": ["文末未见「作者简介」块——形态与一致性核验"
                            "无法进行（存在性见 submission_pack.py）"],
                "stat": stat}

    # 简介块只取到第一个「。」为止的句子（官方示例为一句话），
    # 避免把后续通信地址段并入简介块，造成单位/职称误配
    first_sent = re.split(r"(?<=。)", bblock)[0]
    stat["bio_sentence"] = first_sent[:90]

    # 简介块首句（官方示例形态：姓名＋单位＋职称…＋研究方向…，一句话）
    bio = first_sent

    # ---------- ① 姓名一致性 ----------
    sig = spk.signature_name(paras)      # 文首署名（与 submission_pack 同口径）
    names_in_block = cn_names_in(bio)
    stat.update({"signature": sig, "names_in_block": names_in_block[:6]})
    if sig:
        sig_n = norm_name(sig)
        hit = any(norm_name(n) == sig_n for n in names_in_block)
        if not hit:
            issues.append(f"文首署名「{sig}」与作者简介块的姓名"
                          f"（{'、'.join(names_in_block[:4]) or '未检出'}）"
                          f"不一致——按官方 §8「作者简介（包括姓名…）」核对")

    # ---------- ② 职称/学位/导师 串联形态 ----------
    t_iss, t_not, info = title_degree_shape(bio)
    issues += t_iss
    notices += t_not
    stat.update({"titles": info["titles"], "degrees": info["degrees"],
                 "mentors": info["mentors"]})

    # ---------- ③ 研究方向 ----------
    dirs = research_dirs(bio)
    stat["dirs"] = dirs
    if not dirs:
        # 研究方向可能跨句/跨段（示例中同句，实操里常见另起一行）
        dirs = research_dirs(bblock)
        stat["dirs"] = dirs
    if not dirs:
        issues.append("作者简介未检出「研究方向：」——官方 §8 要求写明研究方向")
    elif len(dirs) < 2:
        issues.append(f"研究方向仅 {len(dirs)} 个（「{dirs[0]}」）——官方示例"
                      f"列 2 个（「知识产权研究、专利分析」），建议补足")
    elif len(dirs) == 2:
        notices.append(f"研究方向 {len(dirs)} 个（{'、'.join(dirs)}），与官方"
                       f"示例形态一致 ✓")

    # ---------- ④ 单位 ----------
    u_iss, u_not, unit = unit_shape(bio)
    issues += u_iss
    notices += u_not
    stat["unit"] = unit

    # ---------- ⑤ 通讯作者标注形态 ----------
    if re.search(r"通讯作者|通信作者", bblock):
        notices.append("简介块内含「通讯作者」标注 ✓（官方要求标注在该作者"
                       "信息之后）")

    return {"issues": issues, "notices": notices, "stat": stat}


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
    st = res["stat"]
    print("=" * 72)
    print("作者信息形态与一致性核验（官方须知 §8）")
    print(f"稿件：{os.path.basename(DOC)}")
    print("=" * 72)

    print(f"\n文首署名：{st.get('signature') or '（未检出）'}")
    if st.get("bio_sentence"):
        print(f"简介块首句：{st['bio_sentence'][:70]}")
    if st.get("unit"):
        print(f"单位（简介块内）：{st['unit']}")
    if st.get("titles") or st.get("degrees") or st.get("mentors"):
        print(f"职称：{'、'.join(st['titles']) or '（未检出）'} ｜ "
              f"学位：{'、'.join(st['degrees']) or '（未检出）'} ｜ "
              f"导师身份：{'、'.join(st['mentors']) or '（未检出）'}")
    if st.get("dirs"):
        print(f"研究方向 {len(st['dirs'])} 个：{'、'.join(st['dirs'])}")

    print("\n" + "-" * 72)
    if res["issues"]:
        print(f"🔴 发现 {len(res['issues'])} 项形态/一致性问题：")
        for it in res["issues"]:
            print(f"   · {it}")
    else:
        print("✅ 作者信息形态与一致性：未发现问题。")
    if res["notices"]:
        print(f"\n🟡 建议核对 {len(res['notices'])} 项：")
        for it in res["notices"]:
            print(f"   · {it}")
    print("\n【官方示例】（official-guidelines.md §8 原文）：")
    print(f"   {spk.BIO_EXAMPLE}")

    if "--json" in flags:
        jf = os.path.splitext(DOC)[0] + ".author_meta.json"
        with open(jf, "w", encoding="utf-8") as f:
            json.dump(res, f, ensure_ascii=False, indent=1)
        print(f"\nJSON 已写入：{jf}")

    sys.exit(1 if res["issues"] else 0)


if __name__ == "__main__":
    main()
