# tsgxyj-paper-style

《图书馆学研究》（CSSCI）期刊论文体例与文风 Skill——基于 108 篇全刊样本量化统计与逐段功能标注的论文写作 / 改写 / 诊断 / 引用核验 / 模拟审稿全流程辅助。

## 这是什么

一个 Agent Skill（Claude Code 等可加载的 SKILL.md 格式），把《图书馆学研究》的期刊风格变成可执行、可核验的规则集与工作流。

## 与通用学术写作 skill 的差异

同类高星项目的「期刊风格学习」通常基于 8–10 篇样本的定性归纳（如 lishix520/academic-paper-skills、ARS 的风格校准）。本 skill 的做法：

| 差异点 | 说明 |
|---|---|
| 全刊样本量化 | 108 篇全刊样本（2026 年 1—9 期）逐篇行级几何提取，统计带四分位区间与样本稳定性验证（分割检验差异 3%） |
| 规则可核验 | 每条规则标注数据出处或「分析判断」置信度；不凭印象出规则 |
| 真实反馈校准 | 模拟审稿的关注点与口吻按投稿者社群反馈校准（如查重实操口径 <7%） |
| 正反示范 | 改写含「过度改写」反例——量化证明 +16% 字数、句均 70 超线的失败案例与修正纪律 |
| 损伤排查 | 提取管线自带受损篇目检查（段数 <12 或字数 <5000 自动标记），防止静默数据污染 |

## 七种模式

| 模式 | 用途 | 一句话 |
|---|---|---|
| G 对标评估 | 投稿前横向定位 | 与语料库逐维对比，输出百分位定位（`benchmark_paper.py`） |
| A 诊断 | 投稿前自查 | 对照实测基准体检，输出「位置 → 问题 → 依据 → 建议」清单 |
| B 改写 | 保真改稿 | 先测段再动手；词层优先于句层；幅度上限 ±10%；不增饰 |
| C 新写 | 新章节 | 写作前必答四问；骨架与「对策—发现」对仗表确认后动笔 |
| D 引用核验 | 投稿前 | 序号连续 / 正文—列表对应 / 标注形态 / 语料库溯源四项核验 |
| E 模拟审稿 | 投稿前 | 13 维 × 100 分制打分卡 + 按失分排序的修改路线图（声明：特征模拟，非真实评审） |
| F 回复审稿 | 返修阶段 | 逐条回应制 + 投稿邮件与材料模板 |

> 投稿前一条命令跑完二十项（GB/T 7714 + D + 引文结构 + A + 润色体检 + E + G + 论证链对仗 + 图表三线表 + 摘要—正文数字一致性 + 术语漂移）：`python scripts/check_all.py 论文.docx`（自动汇总红线/待办与模拟审稿分数，并输出与上次总检的指标对比，退出码 0 = 无红线）。术语漂移已在总检内（待办级），数字硬伤为红线级。返修阶段另用 `python scripts/response_check.py --reviews 审稿意见.docx --response 回复信.docx` 核对「意见条数 vs 回应区块数 / 三件套 / 位置精确度」。

## 文件结构

```
SKILL.md                        # 入口：七模式 + 核心画像速查 + 使用纪律
references/
  style-rules.md                # 九维量化规则（每条带数据出处）
  section-guides.md             # 分节写作配方（引言到结语 + 写作前必答问题）
  rewrite-examples.md           # 改写操作纪律 + 三版正反示范
  reviewer-simulation.md        # 模拟审稿 13 维 100 分制打分卡 + 真实审稿意见特征
  response-template.md          # 回复审稿意见 + 投稿邮件模板
  checklists.md                 # 投稿前体例自检清单
  anti-patterns.md              # 反模式清单（原文实测分级：A 级红线/B 级控密度 + 分条偏好）
  polish-lexicon.md             # 润色语言词典（口语→书面次数实证、程度词、句法特征）
  exemplars.md                  # 样章摘录（注明篇目出处）
  close-reads.md                # 对标论文逐段功能标注
  phrase-library.md             # 句式骨架库（含防模板腔提示）
  patterns.md                   # 原文写作模式库（analyze_corpus.py 生成，七类句式实测）
scripts/
  corpus.py                     # 语料库自动定位与加载（--corpus / 环境变量 / 逐级探测）
  diagnose_paper.py             # 模式 A：论文诊断（结构/摘要体检 + 动态基线 + 反模式 + 段落级）
  search.py                     # 语料库段落检索（118 篇，AND/OR/正则/主题筛选）
  analyze_corpus.py             # 原文模式挖掘（7 维句式统计，产出 patterns.md）
  polish_lexicon.py             # 润色词典（口语→书面次数实证/程度词/句法，产出 polish-lexicon.md）
  language_audit.py             # 整稿语言体检（七类逐句定位：口语词/强断言/了/被/即/反模式）
  mock_review.py                # 模式 E：模拟审稿 13 维 100 分制自动打分 + 修改路线图
  verify_refs.py                # 模式 D：参考文献四项核验
  ref_format.py                 # GB/T 7714 文献表要素/形态核验（误报防护）
  citation_analysis.py          # 引文结构与时效分析（对照 5115 条文献基准）
  benchmark_paper.py            # 模式 G：百分位对标评估
  argument_check.py             # 论证链对仗：发现句↔对策句逐条匹配（主题词重叠，收口章优先）
  table_fig_audit.py            # 图表体例与三线表规范深度审查（无竖线三线表 + 占位符检测）
  number_consistency.py         # 摘要—正文数字一致性核验器（样本量/百分比/题次交叉比对，红线级，check_all 第⑩项）
  term_consistency.py           # 术语漂移检测器（同义异形词混用，附该刊 118 篇原文偏好篇数与共现率白名单，check_all 第⑪项）
  response_check.py             # 审稿回复逐条回应核对器（模式 F：条数/三件套/位置精确度，退出码 1 = 有硬伤）
  draft_scaffold.py             # 句式脚手架生成器（五大核心章节地道骨架填空）
  check_all.py                  # 投稿前一键总检（二十项工具串联 + 增量快照对比与刊中位可视化）
  fetch_reviews_direct.py       # 投稿者反馈抓取与解析（维护 submission-reality.md）
  feedback_loop.py              # 审查反哺闭环（已见刊全面吸收/用户稿只取长处）
  citation_quality.py           # 引文—论点关联核验（装饰性引用专查）
  progress_dashboard.py         # 投稿就绪度看板（七项指标 + Top3 修复建议）
  review_triage.py              # 审稿意见分类助手（必改/疑问/建议/表扬）
  self_repeat.py                # 稿内自我重复检测（查重前哨，跨区域重复）
  english_audit.py              # 英文摘要与关键词对称性核验
  heading_audit.py              # 标题编号体例（跳级/重号/断号）
  citation_distribution.py      # 引用密度分布（结构性失引，论证空窗）
  figure_data_check.py          # 图表与正文数字一致性（读真实表格）
  submission_pack.py            # 投稿材料包完整性（官方 §8 要素）
  corpus_similarity.py         # 与已见刊论文句面相合度（外部查重前哨）
  conclusion_echo.py           # 摘要与结语呼应核验（贡献承诺一致性）
  selftest.py                   # 工具链自检（409 项断言，含 xref 往返、13 维守卫、数字一致性、术语漂移与审稿回复核对，改脚本后必跑）
  build_xref.py / unfield_refs.py / xref_verify.py   # Word 交叉引用域
```

## 数据方法

- 语料：该刊 2026 年 1—9 期 108 篇（全刊统计基准）+ 2021—2025 年「无障碍／老年／特殊群体」主题 10 篇（扩展对标区）
- 提取管线：PyMuPDF 行级几何提取——双栏自适应分栏、首行缩进分段、行内按 x 坐标排序、全角字符归一化、受损篇目自动标记
- 稳定性：按期次分割两半检验，结构指标差异 3%（饱和）；引用密度为高方差指标，按区间使用
- 原文级证据优先：`analyze_corpus.py` 直接读 118 篇原文文本挖掘句式模式，与统计层交叉验证；冲突时以原文级证据为准（例：语料库实测 85/108 篇含三级标题、每篇中位 6.5 个，据此修正了早期「不堆三级标题」的归纳结论）
- 局限：单刊、单年份为主；统计口径与「分析判断」条目在各文件内均有标注

## 安装与使用

将本目录复制到 agent 的 skills 目录（如 `~/.claude/skills/tsgxyj-paper-style`），SKILL.md 的 description 用于触发加载。

### 配套工具（scripts/）

`corpus.py` 提供语料库自动定位（`--corpus <dir>` > 环境变量 `TSGXYJ_CORPUS` > 当前目录 > 技能目录向上逐级 > 常见论文工作区路径）。诊断/检索/核验需要 python-docx，检索与溯源需要语料库（`corpus_stats.json` / `paragraphs.json`，**语料库本身不随本仓库分发**——见版权声明；缺语料库时 `diagnose_paper.py` 自动降级为内置静态基线）。

```bash
# 模式 A：论文诊断（基线动态取自语料库 108 篇实测）
python scripts/diagnose_paper.py 论文.docx            # 四段完整报告
python scripts/diagnose_paper.py 论文.docx --brief     # 只出优先问题清单（红线在前）
python scripts/diagnose_paper.py 论文.docx --json      # 附 JSON 报告

# 语料库段落检索（118 篇，写某主题前先搜该刊怎么写）
python scripts/search.py 无障碍                        # AND 多词默认
python scripts/search.py --or 助听器 手语              # OR
python scripts/search.py --re "适老.{0,4}服务"         # 正则
python scripts/search.py --count 无障碍                # 按篇统计

# 原文模式挖掘（写某类段前查该刊真实句式分布）
python scripts/analyze_corpus.py                       # 七维全量
python scripts/analyze_corpus.py 对策段句式             # 单维度
python scripts/analyze_corpus.py --json                # 写语料库 patterns.json

# 润色词典（模式 B 改写的词层依据）
python scripts/polish_lexicon.py                       # 全量：口语→书面/程度词/句法
python scripts/polish_lexicon.py 词汇替换               # 单维度
python scripts/polish_lexicon.py 赋能                   # 查某个词在该刊的真实用法与次数

# 整稿语言体检（润色的落地点：逐句定位七类问题）
python scripts/language_audit.py 论文.docx             # A 类口语词非空则退出码 1
python scripts/language_audit.py 论文.docx --only A,D  # 只查口语词与「了」
python scripts/language_audit.py 论文.docx --limit 20  # 每类显示 20 条

# 模式 E 模拟审稿（13 维 100 分制）
python scripts/mock_review.py 论文.docx               # <60 分退出码 1
python scripts/mock_review.py 论文.docx --strict      # 扣分 ×1.5

# 论证链对仗核对（模式 E 论证链维度深挖：发现↔对策逐条）
python scripts/argument_check.py 论文.docx            # 存在疑似缺回应退出码 1

# 模式 D：参考文献核验
python scripts/verify_refs.py 论文.docx

# 模式 G：对标评估
python scripts/benchmark_paper.py 论文.docx --corpus <语料库目录>

# GB/T 7714 文献表要素核验
python scripts/ref_format.py 论文.docx

# 句式脚手架生成器（写稿/改稿提速）
python scripts/draft_scaffold.py --section intro --topic "数字图书馆"
python scripts/draft_scaffold.py --all

# 图表三线表规范与占位符审查
python scripts/table_fig_audit.py 论文.docx

# 摘要—正文数字一致性核验（样本量/百分比/题次交叉比对，红线级）
python scripts/number_consistency.py 论文.docx

# 术语漂移检测（同义异形词混用，附该刊 118 篇原文偏好词与篇数；已接入总检第⑪项）
python scripts/term_consistency.py 论文.docx
python scripts/term_consistency.py 论文.docx --corpus <语料库目录>

# 审稿回复逐条回应核对（模式 F：意见条数 vs 回应区块数 / 三件套 / 位置精确度）
python scripts/response_check.py --reviews 审稿意见.docx --response 回复信.docx
python scripts/response_check.py --reviews 审稿意见.md --response 回复信.md --top 20 --json

# 投稿前一键总检（串联 7714 + D + 引文结构 + A + 润色 + E + G + 论证链 + 图表体例 + 数字一致性 + 术语漂移，
#                   自动汇总红线/待办，并输出与上次总检的指标对比）
python scripts/check_all.py 论文.docx            # 退出码 0 = 无红线
python scripts/check_all.py 论文.docx --quick    # 快速模式（只跑前两项）

# 工具链自检（改 scripts/ 后必跑，409 项断言）
python scripts/selftest.py -v

# Word 交叉引用域：编辑前退域 → 改文字 → 单次重建 → 核验
python scripts/unfield_refs.py 论文.docx --backup     # 退域为纯文本 [n]
# …在 Word 中编辑…
python scripts/build_xref.py 论文.docx --backup        # 重建域（只跑一次）
python scripts/xref_verify.py 论文.docx                # 核验（文本零变化/配平/上标）
python scripts/build_xref.py 论文.docx --dry-run       # 或先干跑，看会替换/跳过哪些标注
```

## 边界声明

- 本 skill 是**形式层工具**：体例、结构、文风、规范。研究本身的创新性与学术价值不在其能力范围，也不应有
- 所有统计基于公开刊文；期刊论文版权归原刊与作者，**本仓库不含任何期刊全文**
- 标注「分析判断」的规则为样本归纳，非期刊编辑部口径
- 模拟审稿不替代真实同行评审

## License

MIT
