# Changelog

本项目遵循 [Keep a Changelog](https://keepachangelog.com/zh-CN/1.1.0/) 与
[语义化版本](https://semver.org/lang/zh-CN/)。

## [Unreleased]

### Changed

- **文风检测改为完整采用 [academic-humanizer](https://github.com/AIScientists-Dev/academic-humanizer)
  （MIT）的逻辑**，删除本仓自研的一切去 AI 味规则：
  - 移除：S/A/B「AI 味」12 类规则、R「语体」6 条、D「防御性写作」6 条、三轴框架
  - 移除：`--max-register` / `--max-defensive` 参数与 `register_score` / `defensive_score`
  - 新增：Layer 1 通用痕迹（T1.1–T1.12）、Layer 2 学术痕迹（T2.1–T2.11）、
    Layer 4 主张-证据（C1–C4），以及 Layer 3 保留清单的代码化豁免
    （句中带数字则不报；证据型模糊限制不报）
  - 新增：`--audit` 输出审计清单（层/编号/类型/位置/改法），对齐上游 Read→Audit→Rewrite→Report
  - 文档：`references/anti-ai-academic-zh.md` 整篇重写为六层中文适配 + 许可与来源声明；
    SKILL.md §12 与 §0.0、README「文风」节同步；测试全部按新规则重写（19 条）

### Added

- **文风第三条轴：防御性写作（D 级规则）**——「别为想象中的批评者写作」
  - 6 条规则：D1 免责声明/自我辩护、D2 铺垫式让步、D3 模糊限定叠加、
    D4 预设质疑的"不是…而是…"、D5 结尾强行平衡、D6 弱化到看不出结论
  - `ai_tone_check.py` 新增 `defensive_score` 与 `--max-defensive`；
    报告改为三轴并列（AI 味 / 语体 / 防御性），JSON 输出同步
  - `references/anti-ai-academic-zh.md` 新增 §9（六条规则、必须保留的边界条件、
    逐句自检、三轴对照表）；文档开头改为"一条主线：措辞强度 = 证据强度"
  - README「文风」节改为三轴；SKILL.md 新增 §0.0「一条主线」说明所有规则的三类来源
    （国标原文 / 可测量证据 / 读者优先），并在 §12 给出三轴说明
  - 新增 5 个防御性用例（含"有证据的边界条件不算防御"的反向保护）
- **模板自动判定**：`--profile` 省略时按任务类型选模板（`resolve_profile`）——
  有 `journal` 块 → 双栏 csae；`meta.doc_type` 含毕业设计/学位论文/课程报告/研究报告
  → 单栏 thesis；含期刊/学报/投稿/学术论文 → 双栏；只写"论文" → 双栏；其它 → 单栏。
  构建时打印"模板：csae（依据）"
- **qa 自动识别版式**：省略 `--profile` 时看文档是否双栏，自动选用期刊/学位论文检查项
- README 新增「三种用法：怎么说、怎么切单栏/双栏」；SKILL.md 顶部新增 §0 模板选择规则
- **期刊模板 `--profile csae`（农业工程学报）**：参数实测自 4 篇该刊论文
  （A4 双栏、栏宽 8.66 cm、标题黑体 18 pt、作者仿宋 14 pt、摘要 9 pt、
  一级标题黑体 11 pt 无装饰线、图题楷体 9 pt、表题黑体 9 pt、三线表 0.75/0.5 pt、
  奇偶页页眉含刊名卷期、页码在页眉）
  - 新增 `JournalBuilder` + `CSAE_SPEC` + `build_journal_from_model`
  - 内容模型新增 `journal` / `title_en` / `authors` / `affiliations` /
    `abstract_en` 等字段；`examples/journal_csae/` 为完整样例
  - `build` / `qa` 均支持 `--profile`
  - qa 期刊分支检查：双栏、页眉刊名与卷期、图表题中英对照、三线表 0.75 pt、正文字号
- **正文引文标注（GB/T 7714—2015 顺序编码制）**：段落里写 `[[1]]`、`[[2,3]]`、`[[4-6]]`，
  渲染为**上标方括号**；qa 新增三项检查——标注必须为上标、编号不超出文后条目数、
  每条文献都被引用、条目含文献类型标识（[J]/[M]/[C]/[EB/OL]…）
- **语体检查（反"大白话"）**：`ai_tone_check.py` 新增 R 级规则与 `register_score`，
  与 AI 味分数并列报告；等级 学术语体 / 略松 / 偏口语 / 大白话；
  `--max-register` 可设阈值；`references/anti-ai-academic-zh.md` 新增 §8
  （AI 腔 vs 大白话 vs 学术语体三方对照、六组改写示例、语体分读法）
- **目录缩进按 GB/T 7713.1—2025 附录 E**：显式定义 TOC1/TOC2/TOC3 样式，
  二级缩进 **1 个字**、三级 **2 个字**（原来用 Word 默认，二级缩进约 2 字，偏大）；
  页码右对齐 + 点线引导
- **提示段不再生成彩色底框**：`callout` 默认无底色无边框，只用加粗引导词；
  确需方框时显式 `"box": true`；示例文档已移除该提示框
- qa 新增"正文段落无彩色底纹"检查（WARN）
- **三级标题支持**：`{"level": 3, "text": "1.1.1 …"}` → `Heading 3`（12 pt 加粗、大纲级别 2）；
  模型写法见 SKILL.md 内容模型表
- **目录默认列到二级**：`TOC \o "1-2"`（原来只列一级，几乎无法导航）；
  `toc_levels` 可改 `1-3`，小型报告仍可显式用 `1-1`
- qa 新增：目录层级检查（1-1 判 WARN）、标题层级 ≤4 级与"章节标题 ≤15 字"
  （GB/T 7713.2—2022 5.2.2）、标题层级分布统计
- **前置部分顺序按 GB/T 7713.1—2025 表 A.1 校正**：默认 封面 → 摘要 → Abstract → 目录；
  支持 `"front_order": "toc-first"` 以适配要求"目录在前"的校规；qa 新增该顺序检查（WARN）
- 修正：**英文摘要原先与中文摘要同页**（用二级标题接排），不符合 5.2.9"另起一页"；
  现改为独立成页、一级标题外观、不进目录
- `references/front-matter-order.md`：摘要/目录顺序的规范依据（国标原文 + 四校规范 + 常见误解）
- **按 GB/T 7713.2—2022 逐条加约束**（原文 5.1 / 5.2.4 / 5.3.5 / 5.4.1–5.4.5 / 5.5.1–5.5.5）：
  - 插图尺寸：宽 ≤ 版心（通栏 155 mm、半栏 75 mm），超出版心出图时直接报错；
    图高 ≤ 22 cm（图+题+注须容于一页，5.4.5）
  - **图内文字插入后有效字号 ≥6 pt**：`figures.py` 把"设计宽度/基准字号"写进 PNG 元数据，
    插入时换算，缩小过狠会警告（5.4.1"图中文字清晰可见"）
  - 图题/表题：`图N` + **1 个全角空格** + 题名（5.4.3 / 5.5.3）
  - 图注/表注：自动补末尾句号（5.4.3 / 5.5.3）
  - 表身空缺语义按 5.5.4 修正：`None`/`""` → 空白（无此项或未测量），
    `"—"` → 测量过而未发现，`0` → 实为零（原先一律写"—"是错的）
  - qa 新增 4 项：插图宽/高/字号、表头斜线、注末尾句号、先文后图（宜条款 → WARN）
- `references/figure-table-standard.md` 重写为**逐条对照版**：引用标准原文条款，
  并给出"条款 → 代码 → 校验"对照表与尺寸推导依据
- `scripts/palette.py`：**唯一配色源**（图与 Word 共用）
  - Okabe-Ito 8 色定性板、Paul Tol bright、Viridis/cividis 顺序板、RdBu/vik 发散板
  - 感知明度 `lstar()/delta_l()`（CIE L*）、Viénot 色觉模拟、红绿配对判定
  - `check_series()` 返回 (硬错, 提示)：色板外颜色与红绿色盲下同色为硬错，
    灰度接近/蓝黄接近/红绿组合为提示（靠线型+标记兜底）
  - `resolve_cmap()` 直接拒绝 jet/rainbow 等彩虹色谱
- `scripts/figures.py`：matplotlib 插图后端（取代 Pillow 版）
  - 毫米版面（半栏 75 / 通栏 160 / 自定义），按插入后不缩放设计字号线宽
  - 坐标轴 `量符号/单位`（符号斜体、单位正体）、只留左/下轴线、柱图零基线
  - 自动冗余编码（颜色 + 线型 + 标记）、误差线、显著性括号、(a)(b) 分图号
  - 默认 1000 dpi + PDF/SVG 矢量；`--grayscale` 输出灰度版自检
  - 配色不合规直接报错（非警告）
- `references/figure-table-standard.md`：图表规范核对版（GB/T 7713.2—2022 已核实编号/名称/
  发布与实施日期；明确区分"国标规定"与"期刊惯例"）
- `references/palette.md`：配色规范（数据类型→色板、色盲底线、灰度校验、学科语义）
- `scripts/apd.py figure` 子命令（`chart` 保留为同义旧名）
- 表格增强：`columns` 列元信息（自动"量/单位"表头、按列精度、数值右对齐）、
  `caption_en` 中英对照题注、`^{a}` 上标（显著性/脚注）、空缺自动写 `—`
- `qa_check.py` 新增 4 项：同列小数位一致、文档用色在批准色板内、彩色 ≤5、
  插图分辨率 ≥600 dpi、插图宽度 ≤160 mm

### Changed

- **开箱可用性**：`apd doctor` 分级报告（必需/推荐/可选），缺 matplotlib 只跳过出图；
  `apd figure` / `render_preview` 缺依赖时给一行安装提示并返回 2，不再抛栈
- `install.ps1` 默认自动 `pip install -r requirements.txt`（`-SkipDeps` 可跳过）
- 中文字体支持 `APD_CJK_FONT` 环境变量覆盖；找不到中文字体时提示一次
- `requirements.txt` 分层注释（必需 / 出图 / 可选 PyMuPDF）
- `apd figure` 增加 `--palette`（与 `figures.py --palette` 等价）
- README 新增「三步跑起来」与「环境要求」表；SKILL.md 顶部补一句先跑 doctor

- 三线表**栏目线由 1 pt 改为 0.75 pt**（顶/底线保持 1.5 pt），对齐期刊惯例
- 图内字号调整为六号（8.5 pt）、轴线 0.7 pt，按 140 mm 插入时不再缩放
- `make_chart.py` 保留为兼容入口（旧 `width_px`/`y_min` 字段自动映射），内部改走 figures.py
- 文档版式色改为从 `palette.DOC` 读取（色值不变，来源唯一）

- `references/anti-ai-academic-zh.md`：学术中文的「AI 味」规则
  - 借 blader/humanizer v3.1.0 的根因分类（用姿态/规则/权威/模板代替信息）与 S/A/B 强度分级
  - 借 op7418/Humanizer-zh 的中文分组与"保事实优先"约束，补 `B2.13 装饰性标点`
  - **学术豁免清单**：方法段被动句、术语化名词化、必要限定、引用格式、图表题注不乱洗
  - 12 类典型句的 before/after、改写协议、交付前六问
- `scripts/ai_tone_check.py`：启发式检测器（只报警，不改稿）
  - 输入 `--docx` / `--json`（内容模型）/ `--text`
  - 词典 + 正则 + 结构指标（句长变异、连接词密度、数字密度、段首复读标题）
  - 输出规则编号、段落号、原句、改法；`--max-score` 作 CI 门禁，`--json-out` 机器可读
- `scripts/apd.py tone` 子命令
- `scripts/qa_check.py`：新增「文风」提示项（36 项），**只报 WARN 不判 FAIL**
- `tests/test_ai_tone.py`：20 个用例（判定方向、数据/引语/被动/数值范围豁免、
  规则与文档一一对应、CLI）
- `SKILL.md`：第 4 条硬规矩 + §11「正文的 AI 味」

## [1.0.0] - 2026-10-24

首个可用版本：把《学术论文 / 研究报告 Word 排版与生成规范》变成可执行的格式系统。

### Added

- `scripts/docx_builder.py`：格式库与 `PaperBuilder` API
  - A4 页面、页边距、首页不同、页眉页脚距离
  - 宋体/黑体 + Times New Roman/Arial 双字体（清除主题字体）
  - 正文 12 pt / 1.5 倍 / 首行 2 字符 / 两端对齐
  - 一级标题 16 pt 深蓝 + 0.75 pt 蓝色下边框，二级标题 13 pt 学术蓝
  - 封面（标题层级、蓝色横线、无边框信息表、底部单位与日期）
  - 摘要/Abstract（`Front Heading`：外观同一级标题但不进目录）
  - 三线表（顶/底线 1.5 pt、表头线 1 pt、无竖线、表头重复、行不跨页）
  - 图片等比缩放 + 图题/图注 + keep_with_next
  - 公式：1×3 无边框表格，居中公式 + 右对齐编号
  - 参考文献：9 pt + 0.74 cm 悬挂缩进 + 自动补编号
  - 页眉 `标题 · 类型`、页脚 `第 X 页 / 共 Y 页`（PAGE / NUMPAGES 域）
  - 目录域 `TOC \o "1-1" \h \z \u`，只显示一级标题
  - `w:view=print` + `w:zoom=100%`，默认页面视图
- `scripts/omml_latex.py`：LaTeX 子集 → Word 原生 OMML（分式、根式、求和/积分、
  上下标、重音、定界符、正体文字），不依赖外部数学库
- `scripts/build_docx.py`：JSON 内容模型 → docx，支持 `--spec` 整体换肤
- `scripts/qa_check.py`：35 项规范自动质检（页面/视图/样式/三线表/图表题/图片/
  公式/目录/页眉页脚/参考文献/字体/编号连续性/占位符），支持 `--spec` 校准
- `scripts/render_preview.py`：PDF 逐页转 PNG，检出空白页、未定义书签、原始域代码
- `scripts/postprocess.ps1`：Word COM 更新域与目录、导出 PDF、重设页面视图
- `scripts/postprocess.sh`：无 Word 时用 LibreOffice 兜底
- `scripts/set_view.py`：只重写 `word/settings.xml`，强制页面视图 100%
- `scripts/make_chart.py`：Okabe-Ito 配色图表 PNG（柱状/分组柱状/折线），
  按 300 dpi 反算字号，无需 matplotlib
- `scripts/apd.py`：统一命令行（`doctor` / `build` / `post` / `qa` / `preview` /
  `chart` / `demo`）
- `install.ps1` / `install.sh`：一键安装到 DSH / Codex / Claude 等 agent 的 skills 目录
- `references/`：规范全文、交付前核对清单、已知坑与修复
- `examples/`：最小示例、完整样例、Python API 示例
- 测试：84 个 pytest 用例覆盖页面、样式、三线表、图片、公式、质检、图表与 CLI
- CI：Windows / Linux / macOS × Python 3.9 / 3.12

### Fixed

- `w:hanging` 单位错误：python-docx 会把 EMU 直接写进 twips，导致参考文献整段被推到
  页面外。改为直接写 XML，并提供 `hanging_cm` 参数
- `-Cm(0.74)` 取负后 `Length` 退化为 `int`，被当作磅值再 ×20；新增 `_to_length()` 兜底
- 样式中的 `w:asciiTheme` 会覆盖显式字体，导致"宋体/黑体"在 Word 里不生效
- `.ps1` 在 Windows PowerShell 5.1 下按 ANSI 解析导致中文乱码与语法错误，改为 UTF-8 BOM
- **CLI 在非 UTF-8 控制台（Windows 管道 = cp1252）打印中文时抛 `UnicodeEncodeError`
  并退出码 1**，导致 CI 的 Windows 作业全红。新增 `scripts/_console.py`，各入口
  `main()` 开头把 stdout/stderr 重配为 UTF-8，并加回归测试
- Pages 工作流在仓库未开启 Pages 时于 "Configure Pages" 失败；`configure-pages` 增加
  `enablement: true` 自动开启

[Unreleased]: https://github.com/Znk1314/academic-paper-docx/compare/v1.0.0...HEAD
[1.0.0]: https://github.com/Znk1314/academic-paper-docx/releases/tag/v1.0.0
