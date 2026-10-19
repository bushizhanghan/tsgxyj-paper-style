# -*- coding: utf-8 -*-
"""submission_pack.py —— 投稿材料包完整性核验（投稿前最后一道清单）

官方须知（references/official-guidelines.md §8）明文要求文末列：
    作者简介（姓名、单位、职称、学位；硕导/博导注明；研究方向）
    邮编及详细通信地址、电话、电子邮箱、定稿日期
另外该刊实操要求随邮件附：论文稿件、查重报告（<7%）、AIGC 检测、作者简介。

diagnose_paper 的官方合规维只报「文末未见作者简介/联系方式/定稿日期」三条，
不核**要素是否齐全**。本脚本做细粒度清点：

  · 作者简介要素：姓名 / 单位 / 职称 / 学位 / 研究方向 /（硕导、博导）/
    通讯作者标注，逐项 ✓✗，并给出该刊的真实示例格式；
  · 联系方式要素：通信地址 / 邮编 / 电话 / 电子邮箱 / 定稿日期，逐项 ✓✗；
  · 姓名-单位一致性：正文署名与作者简介里的姓名是否对应；
  · 全部要素齐全且格式合规时给「可投」结论，否则列出**待补清单**。

用法:
    python submission_pack.py <论文.docx> [--json]
退出码：0 = 材料包要素齐全；1 = 存在待补要素。

说明：本工具核「稿件内是否写明」，不核外部文件（查重报告等）是否存在——
      那部分见 references/checklists.md 的「投稿材料」清点单。
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

# 作者简介必须要素：(名称, 正则, 是否必需)
BIO_ITEMS = [
    # 姓名：块首 2–4 字中文姓名（可带全角空格分隔，如「张　瀚」）；
    # 兼容未剥离标签的形态（「作者简介：张三」），两种都认
    ("姓名", r"^\s*(?:作者简介|作者信息)?(?:】)?[：:]?\s*[\u4e00-\u9fff](?:[\s\u3000]*[\u4e00-\u9fff]){1,3}", True),
    ("单位", r"[\u4e00-\u9fff]{2,20}(?:大学|学院|图书馆|研究院|研究所|中心)", True),
    ("职称", r"(?:教授|副教授|讲师|助理教授|研究馆员|副研究馆员|馆员|助理馆员|"
             r"研究员|副研究员|高级工程师|工程师|编辑)", True),
    ("学位", r"(?:博士|硕士|学士|博士研究生|硕士研究生)", True),
    ("研究方向", r"研究方向[：:]", True),
    ("导师身份", r"(?:硕士生导师|博士生导师|硕导|博导|硕士导师|博士导师)", False),
    ("通讯作者", r"通讯作者|通信作者", False),
]
# 联系信息必须要素
CONTACT_ITEMS = [
    ("通信地址", r"(?:通信地址|通讯地址|地址)[：:]|省.{0,12}市.{0,20}(?:路|街|号)", True),
    ("邮编", r"邮(?:政编码|编)[：:]?\s*\d{6}", True),
    ("电话", r"(?:电话|手机|联系方式)[：:]?\s*[\d\-\s]{7,}", True),
    ("电子邮箱", r"[A-Za-z0-9._%+\-]+@[A-Za-z0-9.\-]+\.[A-Za-z]{2,}", True),
    ("定稿日期", r"(?:定稿日期|收稿日期)[：:]?\s*\d{4}", True),
]

# 该刊作者简介示例（official-guidelines.md §8 原文）
BIO_EXAMPLE = ("王悦欣 ＊＊＊大学副教授，硕士生导师，博士。"
               "研究方向：知识产权研究、专利分析等。通讯作者。")


def gather_text(paras):
    """全文文本（含文末所有段落，不限正文区域）。"""
    return "\n".join(t for _, t in paras if t)


def tail_text(paras):
    """文末区域文本（最后 40 段或参考文献之后的部分）。"""
    texts = [t for _, t in paras if t]
    return "\n".join(texts[-40:])


def signature_name(paras):
    """正文署名：出现「摘要」之前的最后一行短中文（多为作者名）。"""
    for _, t in paras[:12]:
        s = (t or "").strip()
        if re.search(r"^【?摘\s*要】?[：:]?", s[:12]):
            return None
    names = []
    for _, t in paras[:10]:
        s = (t or "").strip()
        if not s:
            continue
        if re.search(r"^【?摘\s*要】?[：:]?", s[:12]):
            break
        # 署名行：2–5 字中文，可能带空格分隔（张　瀚）
        c = re.sub(r"\s+", "", s)
        if 2 <= len(c) <= 5 and re.match(r"^[\u4e00-\u9fff]+$", c):
            names.append(c)
    return names[-1] if names else None


def check_items(text, items):
    """逐项检查要素，返回 [(名称, 是否找到, 是否必需)]。"""
    out = []
    for name, pat, required in items:
        # 用 MULTILINE：块首锚点 ^ 在多行文本的首行生效
        found = bool(re.search(pat, text or "", re.M))
        out.append((name, found, required))
    return out


def bio_block(tail):
    """截取「作者简介」块文本（从「作者简介」到文末，最多 15 行）。

    要素必须在**简介块内**才算命中——否则正文里「××大学副教授」这类
    描述性文字会让「单位/职称」误判为已具备。
    """
    m = re.search(r"(?:【)?(?:作者简介|作者信息)(?:】)?[：:]?", tail)
    if not m:
        return ""
    seg = tail[m.end():]
    return "\n".join(seg.splitlines()[:15])


def analyze(doc):
    paras, refs, abs_text, kw_text, titles, body = dp.extract_paras(doc)
    tail = tail_text(paras)
    bblock = bio_block(tail)

    # 作者简介要素：在简介块内核（无简介块时全部判缺失，交给整块判定）
    bio = check_items(bblock, BIO_ITEMS) if bblock else \
        [(n, False, req) for n, _, req in BIO_ITEMS]
    # 联系方式要素：在文末核（联系方式常独立成行，不必在简介块内）
    contact = check_items(tail, CONTACT_ITEMS)

    # 作者简介是否整体存在（先用 diagnose 的 package_notices 判定口径对齐）
    official = dp.scan_official(kw_text, refs, titles,
                                "".join(dp.clean(p) for p in body), paras=paras)
    has_bio_block = bool(re.search(r"作者简介|作者信息|【作者简介】", tail))
    has_contact_block = bool(re.search(r"通信地址|通讯地址|邮(?:编|政编码)|电话|"
                                       r"电子邮箱|[A-Za-z0-9._%+\-]+@", tail))
    has_date = bool(re.search(r"(?:定稿日期|收稿日期)", tail))

    missing_bio = [n for n, ok, req in bio if req and not ok]
    missing_contact = [n for n, ok, req in contact if req and not ok]

    sig = signature_name(paras)
    # 姓名一致性：署名是否出现在作者简介块里
    name_consistent = None
    if sig and bblock:
        name_consistent = sig in bblock

    return {
        "bio": bio, "contact": contact,
        "has_bio_block": has_bio_block, "has_contact_block": has_contact_block,
        "has_date": has_date,
        "missing_bio": missing_bio, "missing_contact": missing_contact,
        "signature": sig, "name_consistent": name_consistent,
        "package_notices": official.get("package_notices", []),
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

    res = analyze(DOC)
    print("=" * 72)
    print("投稿材料包完整性核验（官方须知 §8 明文要求）")
    print(f"稿件：{os.path.basename(DOC)}")
    if res["signature"]:
        print(f"正文署名：{res['signature']}")
    print("=" * 72)

    def show(title, items, missing):
        print(f"\n【{title}】")
        for name, ok, required in items:
            mark = "✓" if ok else ("✗" if required else "—")
            tag = "" if required else "（如有）"
            print(f"  {mark} {name}{tag}")
        if missing:
            print(f"  → 待补：{'、'.join(missing)}")

    show("作者简介要素", res["bio"], res["missing_bio"])
    show("联系方式要素", res["contact"], res["missing_contact"])

    issues = []
    if not res["has_bio_block"]:
        issues.append("文末缺「作者简介」整块")
    elif res["missing_bio"]:
        issues.append(f"作者简介缺要素：{'、'.join(res['missing_bio'])}")
    if not res["has_contact_block"]:
        issues.append("文末缺「联系方式」整块")
    elif res["missing_contact"]:
        issues.append(f"联系方式缺要素：{'、'.join(res['missing_contact'])}")
    if not res["has_date"]:
        issues.append("文末缺「定稿日期」")
    if res["name_consistent"] is False:
        issues.append(f"正文署名「{res['signature']}」未出现在作者简介中（核对姓名一致）")

    print("\n" + "-" * 72)
    if issues:
        print(f"🔴 材料包待补 {len(issues)} 项：")
        for it in issues:
            print(f"   · {it}")
        print("\n【该刊作者简介示例】（official-guidelines.md §8 原文）：")
        print(f"   {BIO_EXAMPLE}")
        print("   邮编及详细通信地址、电话、电子邮箱、定稿日期。")
    else:
        print("✅ 投稿材料包要素齐全（稿件内部分）。")
        print("\n提醒：外部件仍需备齐——查重报告（<7%）、AIGC 检测（<10%）、"
              "署名确认（录用后不可改）。")

    if "--json" in flags:
        jf = os.path.splitext(DOC)[0] + ".submission_pack.json"
        with open(jf, "w", encoding="utf-8") as f:
            json.dump(res, f, ensure_ascii=False, indent=1)
        print(f"\nJSON 已写入：{jf}")

    sys.exit(1 if issues else 0)


if __name__ == "__main__":
    main()
