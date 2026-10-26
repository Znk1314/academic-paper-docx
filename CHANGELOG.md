# Changelog

本项目遵循 [Keep a Changelog](https://keepachangelog.com/zh-CN/1.1.0/) 与
[语义化版本](https://semver.org/lang/zh-CN/)。

## [Unreleased]

### Added

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
