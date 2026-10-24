# academic-paper-docx

> 把「学术论文 / 研究报告 Word 排版与生成规范」变成一套**能跑、能验、能交付**的格式系统。
> 你只提供内容，脚本保证页面、字体、字号、配色、标题层级、三线表、图题表题、OMML 公式、
> 自动目录、页眉页脚、参考文献格式、默认页面视图全部正确。

[![CI](https://github.com/OWNER/academic-paper-docx/actions/workflows/ci.yml/badge.svg)](https://github.com/OWNER/academic-paper-docx/actions/workflows/ci.yml)
[![License: MIT](https://img.shields.io/badge/license-MIT-blue.svg)](LICENSE)
[![Python](https://img.shields.io/badge/python-3.9%2B-blue.svg)](pyproject.toml)
[![QA](https://img.shields.io/badge/QA-35%20checks-009E73.svg)](scripts/qa_check.py)
[![Agent](https://img.shields.io/badge/agent-DSH%20%7C%20Codex%20%7C%20Claude-1F3A5F.svg)](#安装)

**它同时是一个 agent skill**：仓库根目录就是 skill 根目录（`SKILL.md` 在根上），
`git clone` 到 `~/.dsh/skills/` 或 `~/.codex/skills/` 就能被 agent 直接使用。

[English](README.en.md) · [格式规范全文](references/format-spec.md) · [在线落地页](docs/index.html) · [示例 PDF](docs/demo/demo_report.pdf)

---

## 为什么需要它

让模型「写个 Word」的十种翻车方式，这里逐条固化了修复：

| 常见翻车 | 根因 | 本项目的处理 |
|---|---|---|
| 宋体/黑体不生效 | 只设了 `w:ascii`，样式里还有 `w:asciiTheme` 主题字体覆盖 | 写入前清除 theme 属性 |
| 缩进靠敲空格 | 手动空格换字体就错位 | `firstLineChars=200`（真·2 字符） |
| 参考文献整段跑出页面 | `w:hanging` 单位坑：python-docx 把 EMU 当 twips | 直接写 XML，`hanging_cm` 参数 |
| 三线表变全框线 | `tblBorders` 只改了部分边 | 6 条边全显式声明 + 表头 `tcBorders` |
| 表头跨页丢失 | 缺 `w:tblHeader`；行可拆分 | 表头重复 + 全行 `cantSplit` |
| 公式是图片 | 不支持 OMML | 自写 LaTeX 子集 → OMML，双击可编辑 |
| 目录是手打文本 | 没用域 | 真 `TOC \o "1-1"` 域 + `dirty` 标记 |
| 页码手打「第 3 页」 | 没用域 | `PAGE` / `NUMPAGES` |
| 打开是 Web 版式 | Word 保存会把视图改回 web | 最后一次保存后再写 `w:view=print` |
| 图表编号一改就乱 | 编号写死在正文 | 编号由脚本接管，重排自动连续 |

## 成品预览

全部由 [`examples/demo/content_model.json`](examples/demo/content_model.json) 自动生成，**未做人工排版**：

| 封面 | 摘要 | 图与公式 | 三线表 |
|---|---|---|---|
| ![封面](docs/screenshots/01-cover.png) | ![摘要](docs/screenshots/02-abstract.png) | ![图与公式](docs/screenshots/05-figure-equation.png) | ![三线表](docs/screenshots/06-table.png) |

完整 7 页见 [docs/demo/demo_report.pdf](docs/demo/demo_report.pdf)。

## 安装

### 方式一：装成 agent skill（推荐）

```powershell
# Windows
git clone https://github.com/OWNER/academic-paper-docx "$env:USERPROFILE\.dsh\skills\academic-paper-docx"
python -m pip install -r "$env:USERPROFILE\.dsh\skills\academic-paper-docx\requirements.txt"
```

```bash
# macOS / Linux
git clone https://github.com/OWNER/academic-paper-docx ~/.dsh/skills/academic-paper-docx
python3 -m pip install -r ~/.dsh/skills/academic-paper-docx/requirements.txt
```

也可以用自带脚本（支持 dsh / codex / claude / 项目级 / 自定义目录，可选符号链接）：

```powershell
pwsh -File .\install.ps1 -Target codex -Force
```

```bash
LINK=1 ./install.sh codex
```

装好后重启 agent（或新开会话），技能 `academic-paper-docx` 就会被发现。

### 方式二：当普通工具用

```bash
git clone https://github.com/OWNER/academic-paper-docx
cd academic-paper-docx
python -m pip install -r requirements.txt
python scripts/apd.py doctor     # 环境自检
```

不装依赖也能读 `references/format-spec.md` 当**格式规范手册**用。

## 三步上手

```bash
# 1) 内容模型 → docx
python scripts/apd.py build work/content.json -o work/report.docx

# 2) 更新域与目录、导出 PDF、把视图固定为页面视图（Windows + Word）
python scripts/apd.py post work/report.docx

# 3) 质检 + 逐页预览
python scripts/apd.py qa work/report.docx --pdf work/report.pdf
python scripts/apd.py preview work/report.pdf -o work/preview
```

一条命令跑完整示例（图表 → docx → PDF → 质检 → 预览）：

```bash
python scripts/apd.py demo -o build/demo
```

没有 Word（macOS / Linux）：`bash scripts/postprocess.sh work/report.docx`（LibreOffice 兜底）。
文档已写入 `w:updateFields`，用 Word/WPS 打开会自动更新目录；如未展开按 `Ctrl+A` → `F9`。

## 格式指标

这些值都写在 [`scripts/docx_builder.py`](scripts/docx_builder.py) 顶部的 `SPEC` 里，并被质检脚本逐条断言：

| 项目 | 值 |
|---|---|
| 纸张 / 页边距 | A4 21×29.7 cm；上 2.4 / 下 2.2 / 左 2.6 / 右 2.4 cm；页眉 1.2 / 页脚 1.0 cm；首页不同 |
| 正文 | 宋体 + Times New Roman，12 pt，1.5 倍行距，首行 2 字符，段后 6 pt，两端对齐 |
| 一级标题 | 黑体 + Arial，16 pt 粗体 `#1F3A5F`，段前 16 / 段后 8 pt，下边框 0.75 pt `#0072B2`，keep-with-next |
| 二级标题 | 黑体 + Arial，13 pt 粗体 `#0072B2`，段前 10 / 段后 4 pt |
| 封面 | 主标题 24 pt、副标题 20 pt、类型行 15 pt、1.5 pt 学术蓝横线、无边框信息表 |
| 三线表 | 顶/底线 1.5 pt，表头线 1 pt，无竖线，表头重复，行不拆分，表内 9 pt |
| 图 / 表题 | 9.5 pt 粗体 `#1F3A5F` 居中；图题在下、表题在上；图注/表注 8.5 pt `#595959` |
| 公式 | OMML 居中，编号 `（1）` 右对齐，1×3 无边框表格定位 |
| 列表 | 不用圆点，统一 `（1）（2）（3）` |
| 目录 | `TOC \o "1-1" \h \z \u`，只显示一级标题；摘要/Abstract 用 `Front Heading`（外观同 H1 但不进目录） |
| 页眉 / 页脚 | 8.5 pt 灰右对齐 `标题 · 类型`；8 pt 灰居中 `第 X 页 / 共 Y 页` |
| 参考文献 | 9 pt，悬挂缩进 0.74 cm，GB/T 7714，自动补 `[n]` |
| 视图 | `w:view=print` + `w:zoom=100%` |

配色：深海军蓝 `#1F3A5F`、学术蓝 `#0072B2`、暖橙 `#E69F00`、学术绿 `#009E73`、
正文 `#262626`、次级 `#595959`、浅底 `#EAF2FA`、淡边框 `#B4C6E7`。

## 架构

```
内容模型 (JSON)                格式库 (SPEC)                   产物
┌──────────────┐    build     ┌──────────────────┐   docx   ┌────────────┐
│ content.json │ ───────────► │ docx_builder.py  │ ───────► │ report.docx │
│  meta        │              │ 页面/样式/封面    │          └─────┬──────┘
│  abstract    │              │ 三线表/图/公式    │                │ Word COM
│  sections[]  │   equation   │ 目录域/页眉页脚   │                ▼
│  references  │ ───────────► │ omml_latex.py    │          ┌────────────┐
└──────────────┘              └──────────────────┘          │ report.pdf │
                                                             └─────┬──────┘
                            ┌──────────────────┐   QA 35 项        │ 逐页渲染
                            │ qa_check.py      │ ◄─────────────────┤
                            └──────────────────┘                   ▼
                                                        preview/page-*.png
```

```
academic-paper-docx/
├── SKILL.md                  agent 入口（DSH / Codex / Claude 通用）
├── scripts/
│   ├── apd.py                统一 CLI：doctor / build / post / qa / preview / chart / demo
│   ├── docx_builder.py       格式库 + PaperBuilder API（SPEC 在文件顶部）
│   ├── omml_latex.py         LaTeX 子集 → OMML（零外部依赖）
│   ├── build_docx.py         内容模型 JSON → docx
│   ├── qa_check.py           35 项规范质检
│   ├── make_chart.py         Okabe-Ito 配色图表（Pillow，无需 matplotlib）
│   ├── render_preview.py     PDF → 逐页 PNG + 空白页/书签体检
│   ├── postprocess.ps1       Word COM：更新域/目录 + PDF + 固定 print 视图
│   ├── postprocess.sh        LibreOffice 兜底
│   └── set_view.py           只改 settings.xml，强制页面视图
├── references/
│   ├── format-spec.md        规范全文（权威）
│   ├── checklist.md          交付前人工核对清单
│   └── troubleshooting.md    已知坑与修复
├── examples/
│   ├── minimal/content.json  最小示例（一页，无图）
│   ├── demo/                 完整样例：内容模型 + 图表规格 + 配图
│   └── python_api.py         不用 JSON、直接写 Python 的示例
├── tests/                    81 个 pytest 用例
├── docs/                     落地页 + 截图 + 示例 PDF
└── .github/                  CI / Release / Pages / Issue 模板
```

## 内容模型

```jsonc
{
  "meta": {
    "org": "示例大学", "doc_type": "本科毕业设计（论文）",
    "title": "标题", "subtitle": "副标题",
    "info": [["论文编号", "…"], ["指导教师", "…"]],
    "college": "学院", "date": "2026 年 3 月", "header_type": "本科毕业设计（论文）"
  },
  "abstract":    { "heading": "摘　要", "paragraphs": ["…"], "keywords": ["…"] },
  "abstract_en": { "heading": "Abstract", "paragraphs": ["…"], "keywords": ["…"] },
  "toc": true,
  "sections": [
    { "level": 1, "text": "1 引言" },
    { "type": "p", "text": "正文段落…" },
    { "type": "list", "items": ["要点一", "要点二"] },
    { "type": "table", "caption": "示例表", "headers": ["列1", "列2"],
      "rows": [["1", "2"]], "widths": [7, 8.5], "note": "注：示例数据。" },
    { "type": "figure", "path": "figs/fig1.png", "caption": "示例图", "width_cm": 14 },
    { "type": "equation", "latex": "\\hat{x} = \\frac{1}{N}\\sum_{i=1}^{N} x_i" },
    { "type": "callout", "title": "说明", "text": "…" }
  ],
  "references": ["作者. 题名[J]. 刊名, 2020, 1(2): 3-4."]
}
```

块类型：`h1`/`h2`（也可用 `level`）、`p`、`list`、`table`、`figure`、`equation`、
`callout`、`pagebreak`。字段细节见 [SKILL.md](SKILL.md#2-先写内容模型核心步骤)
与 [examples/minimal/content.json](examples/minimal/content.json)。

换配色/字号用 `--spec`，质检要带同一份（否则颜色断言会 FAIL）：

```bash
python scripts/apd.py build work/content.json -o work/report.docx --spec work/spec.json
python scripts/apd.py qa work/report.docx --spec work/spec.json
```

## 公式

`equation.latex` 支持 LaTeX 子集，转换结果是真的 Word 公式（`m:oMath`），双击即可编辑：

```
上下标 x_i  x^2  x_i^2       分式 \frac{a}{b}        根式 \sqrt{x}  \sqrt[n]{x}
求和/积分 \sum_{i=1}^{n}  \int_0^T      括号 \left( ... \right)
重音 \hat{x} \bar{x} \tilde{x} \vec{x}   正体 \mathrm{RMSE}  \text{中文单位}
希腊字母 \alpha \beta \theta \Delta \omega     运算符 \times \le \approx \pm \infty \to
```

## 质检

```text
[表格]
    OK   数据表为三线表（顶/底线 1.5pt、表头线 1pt、无竖线）
    OK   表格统计：三线表 1 个 / 无边框表 3 个
[图表]
    OK   图题 / 表题：9.5 pt 加粗 #1F3A5F
    OK   图编号从 1 连续不跳号
[图片]
    OK   图片等比缩放且不超版心宽度
    OK   图片与图题同页（keep_with_next）
[公式]
    OK   公式为 Word 原生 OMML（可编辑）
    OK   公式编号连续、位于右侧单元格
...
==============================================================================
通过 35 项，警告 1 项，失败 0 项
结论：符合规范。
```

35 项覆盖：页面几何、默认视图、三个核心样式、正文缩进/行距/对齐、列表形态、
三线表边框与表头、图表题字号/颜色/编号连续性、图片等比与宽度上限、OMML 与编号、
TOC 域、页眉页脚域、参考文献缩进、交叉引用、字体白名单、空白页、占位符残留。
`--strict` 把 WARN 也当失败，适合 CI 门禁。

## 跨平台

| 环境 | docx | 更新目录/域 | 导出 PDF | 说明 |
|---|---|---|---|---|
| Windows + Microsoft Word | ✅ | ✅ `postprocess.ps1` | ✅ | 推荐路径，能力最完整 |
| Windows + WPS | ✅ | 打开时更新 | 手动 | TOC 域兼容 |
| macOS / Linux + LibreOffice | ✅ | 打开时更新 | ✅ `postprocess.sh` | 脚本不更新域 |
| 纯 Python | ✅ | — | — | 只生成 docx 也能用 |

## 测试

```bash
python -m pip install -r requirements-dev.txt
python -m ruff check .
python -m pytest -q      # 81 passed
```

用例覆盖：页面几何、三个核心样式、三线表边框、表头重复、图片等比与宽度限制、
OMML 结构、公式编号连续性、悬挂缩进单位（历史 bug 回归）、默认视图、
`set_view.py` 只改 `settings.xml`、质检的 FAIL/WARN 判定、图表尺寸与配色、
CLI 子命令、`SKILL.md` frontmatter、`.ps1` 必须带 UTF-8 BOM。

CI：Windows / Linux / macOS × Python 3.9 / 3.12，跑 lint + 单测 + 端到端构建与质检。

## 已知限制

- **文档真实性靠人**：脚本不会替你判断数据是否虚构，只用 `[待补充/待核实]` 与"模拟数据"标注来提醒。
- **Word 是最终裁判**：LibreOffice 对 OMML 与域的支持不完整，交付前建议用 Word 打开核对。
- **换配色需要一起换图**：图表 PNG 由 `make_chart.py` 按同一套色板生成，改色板要重出图。
- **不自动生成论文内容**：这是排版工具，内容仍然要你（或你的 agent）写。
- 质检里颜色/字号断言写死默认调色板，换肤后必须带 `--spec`。

## 贡献

见 [CONTRIBUTING.md](CONTRIBUTING.md)。最重要的两条：**格式参数集中在 `SPEC`**、
**改格式必须同步改 `qa_check.py` 与 `tests/`**。

## 致谢与许可

- 代码：[MIT](LICENSE)。
- `references/format-spec.md` 是排版规范的原始文档，著作权归原作者，仓库内仅作实现参照。
- 配色采用 [Okabe–Ito](https://jfly.uni-koeln.de/color/) 色盲友好方案。
- 依赖：[python-docx](https://python-docx.readthedocs.io/)、[Pillow](https://python-pillow.org/)、[PyMuPDF](https://pymupdf.readthedocs.io/)。

> 发布到自己的账号时，请把文中 `OWNER` 替换成 GitHub 用户名，步骤见
> [docs/publishing.md](docs/publishing.md)。
