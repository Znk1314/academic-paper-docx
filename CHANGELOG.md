# Changelog

本项目遵循 [Keep a Changelog](https://keepachangelog.com/zh-CN/1.1.0/) 与
[语义化版本](https://semver.org/lang/zh-CN/)。

## [Unreleased]

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
- 测试：81 个 pytest 用例覆盖页面、样式、三线表、图片、公式、质检、图表与 CLI
- CI：Windows / Linux / macOS × Python 3.9 / 3.12

### Fixed

- `w:hanging` 单位错误：python-docx 会把 EMU 直接写进 twips，导致参考文献整段被推到
  页面外。改为直接写 XML，并提供 `hanging_cm` 参数
- `-Cm(0.74)` 取负后 `Length` 退化为 `int`，被当作磅值再 ×20；新增 `_to_length()` 兜底
- 样式中的 `w:asciiTheme` 会覆盖显式字体，导致"宋体/黑体"在 Word 里不生效
- `.ps1` 在 Windows PowerShell 5.1 下按 ANSI 解析导致中文乱码与语法错误，改为 UTF-8 BOM

[Unreleased]: https://github.com/OWNER/academic-paper-docx/compare/v1.0.0...HEAD
[1.0.0]: https://github.com/OWNER/academic-paper-docx/releases/tag/v1.0.0
