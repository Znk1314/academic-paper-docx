---
name: academic-paper-docx
description: 生成符合《学术论文 / 研究报告 Word 排版与生成规范》的 .docx（蓝绿学术配色、黑体标题+宋体正文、三线表、OMML 可编辑公式、自动目录、页码页眉、GB/T 7714 参考文献、默认页面视图）。当用户要写/生成/排版学术论文、毕业设计、课程报告、研究报告、课题组汇报的 Word 文档，或要求"按某套格式规范出 Word"时使用。
whenToUse: 用户要求产出 .docx 论文/报告/毕业设计，或要求按学术排版规范（三线表、图题表题、自动目录、OMML 公式、GB/T 7714）生成 Word 文档时。
---

# 学术论文 Word 生成（academic-paper-docx）

## 0.0 一条主线：这个 skill 的所有规则都回到三件事

看起来规则很多（版式、图表、配色、文风），其实只有三条来源，每条都能指回去：

| 来源 | 怎么落地 | 举例 |
|---|---|---|
| **国标原文**（可查条款号） | `references/` 里逐条对照，代码里注释条款号 | 图题空 1 个汉字（GB/T 7713.2 5.4.3）、引文上标方括号（GB/T 7714） |
| **可测量的证据**（数字、不靠感觉） | 出图/质检把它变成断言 | 图宽 ≤ 版心、图内字 ≥6 pt、灰度 ΔL*≥20、色盲距离、文风两分数 |
| **读者优先**（去装饰、措辞与证据相符） | 默认无底纹、无装饰线、无彩虹色、无过度断言 | 一级标题不画彩条、提示段不加蓝框、禁用 jet |

**遇到新要求时先问：它属于哪一条？** 能指到国标就写成硬校验；能测量就写成断言；
都不属于，就只写进文档当建议，不塞进代码。

## 0. 先选版式（两套模板，别混用）

| 任务 | 模板 | 样子 |
|---|---|---|
| 期刊论文 / 学术论文 / 投稿（如农业工程学报） | **csae** | 双栏 8.66 cm、黑体 18 pt 题名、正文 10 pt、页码在页眉 |
| 毕业设计 / 学位论文 / 课程报告 / 研究报告 | **thesis**（默认） | 单栏、封面+摘要+目录、正文 12 pt、页脚页码 |

**判定顺序**（`scripts/docx_builder.resolve_profile`）：`--profile` > 模型 `profile` 字段 >
有 `journal` 块 → csae > `meta.doc_type` 含"毕业设计/学位论文/课程报告/研究报告" → thesis >
含"期刊/学报/投稿/学术论文" → csae > 只写"论文" → csae > 其它 → thesis。

- 用户只说"写一篇论文" → 用 **csae（双栏）**
- 用户说"毕业设计""学位论文""课程报告""研究报告" → 用 **thesis（单栏）**
- 用户明确说某个刊名 → 用 **csae**，并在 `journal` 字段填刊名/卷期/doi
- 建完打印的"模板：xxx（依据）"一行就是判定结果，如实告诉用户

判定拿不准时：**先问一句**"这是投期刊（双栏）还是毕业设计/课程报告（单栏）？"，
不要默认套错模板——两者的字号、栏数、三线表线宽、页眉都不同。


> **先跑通再动笔**：`python <SKILL_DIR>/scripts/apd.py doctor` 会告诉你这台机器
> 能做什么、缺什么。只装 python-docx + lxml 就能出完整 docx；matplotlib 只在出图时用。
> 缺依赖时命令会给一行安装提示，不会抛栈。

把"内容模型"交给固定格式库，产出**格式合规、可编辑、可直接打印**的 Word 论文/报告。
内容可以任意变化，**格式参数不要改**——这是本 skill 的全部价值。

权威依据：`references/format-spec.md`（完整规范，任何时候与本文冲突以它为准）。

> **安装 / 更新**：本目录整体就是技能本体，放到 `<agent>/skills/academic-paper-docx/`（DSH 为
> `~/.dsh/skills/`，Codex 为 `~/.codex/skills/`）即可，文件夹名不要改；然后
> `python -m pip install -r <本目录>/requirements.txt`，重启 agent 生效。
> 详细步骤与排障见 `README.md` 的「安装」章节。

## 0. 四条硬规矩

1. **不要手写 python-docx 代码去拼格式**。格式已在 `scripts/docx_builder.py` 里实现，
   你只需要写**内容模型 JSON**（或调用 Python API）。
2. **不要复用示例内容**。`examples/demo/` 里的题目、姓名、单位、数据、文献、图都是演示用的，
   必须全部替换成用户真实内容。
3. **不要虚构数据/文献**。无法确认的数值、指标、文献一律标注 `[待补充/待核实]`，
   模拟数据必须在表注/图注里写明"模拟数据"。
4. **正文按「人话」写，不按模板写**。写作时就要避开 AI 味（空转意义句、数据缺席的拔高、
   假对比、段末复述、元话语铺垫、术语漂移），规则见
   `references/anti-ai-academic-zh.md`，交付前跑 `scripts/apd.py tone` 自检。
   注意：学术文体里的**被动句、名词化术语、必要限定不是 AI 味**，别乱洗。

## 1. 判断环境（先做，10 秒）

```bash
python <SKILL_DIR>/scripts/apd.py doctor
```

`doctor` 会打印依赖、Word/LibreOffice、中文字体与大致的可用能力。缺依赖就装：

```bash
python -m pip install -r <SKILL_DIR>/requirements.txt
```

能力差异（决定第 4 步怎么做）：

| 平台 | 更新目录/域 + 导出 PDF | 说明 |
|---|---|---|
| Windows + 已装 Word | `scripts/postprocess.ps1`（推荐） | Word COM，最完整 |
| macOS / Linux / 无 Word | `scripts/postprocess.sh` | LibreOffice 导 PDF；目录靠打开时自动更新 |
| 只要 docx，不要 PDF | 跳过第 4 步 | 仍必须做第 5 步质检 |

> 下文命令里的 `python scripts/xxx.py` 都可以换成统一入口
> `python <SKILL_DIR>/scripts/apd.py {build|post|qa|preview|chart|demo}`。

## 2. 先写内容模型（核心步骤）

最快的起点：把 `examples/minimal/content.json`（一页、无图）复制成工作文件，跑通后再加内容；
需要完整形态时参考 `examples/demo/content_model.json`。**内容全部换成本文用户自己的**。

```jsonc
{
  "meta": {
    "org": "学校/单位", "doc_type": "论文类型（如 本科毕业设计（论文））",
    "title": "标题", "subtitle": "副标题",
    "note": "封面说明行（可省）",
    "info": [["论文编号", "…"], ["学生姓名", "…"], ["指导教师", "…"]],  // 封面无边框信息表，按需增删
    "college": "学院/部门", "date": "2026 年 3 月",
    "header_type": "页眉右侧的类型后缀"       // 页眉 = title · header_type
  },
  "abstract":   { "heading": "摘　要", "paragraphs": ["…"], "keywords": ["…", "…"] },
  "abstract_en":{ "heading": "Abstract", "paragraphs": ["…"], "keywords": ["…"] },
  "toc": true, "toc_title": "目　录", "toc_levels": "1-2",   // 目录层级，可改 1-3
  "sections": [ /* 见下表 */ ],
  "references": ["[1] …", "…"]            // 不写 [n] 时自动补号
}
```

`sections` 里可用的块：

| 块 | 写法 | 说明 |
|---|---|---|
| 一级标题 | `{"level":1,"text":"1 引言"}` | 编号写进 text，全文统一（`1` 或 `一、`）；首个一级标题自动另起一页 |
| 二级标题 | `{"level":2,"text":"1.1 研究背景"}` | 不自动加编号 |
| 三级标题 | `{"level":3,"text":"1.1.1 更小的节"}` | 12 pt 加粗；是否进目录取决于 `toc_levels`（默认 `1-2`，即三级不进目录；要进目录改 `1-3`） |
| 正文段 | `{"type":"p","text":"…"}` | 12 pt / 1.5 倍 / 首行 2 字符 / 两端对齐 |
| 编号列表 | `{"type":"list","items":["…","…"]}` | 生成（1）（2）（3），不用圆点 |
| 三线表 | `{"type":"table","caption":"…","headers":[…],"rows":[[…]],"widths":[3,3.8,8.7],"aligns":["center","left","left"],"note":"注：…"}` | 宽度自动归一到 15.5 cm；编号自动 |
| 图 | `{"type":"figure","path":"figs/x.png","caption":"…","note":"注：…","width_cm":14}` | 只给宽度 → 等比缩放；图题在下 |
| 公式 | `{"type":"equation","latex":"\\hat{x}_k = \\frac{1}{N}\\sum_{i=0}^{N-1} z_{k-i}","number":"（1）"}` | 转成 Word 原生 OMML；编号右对齐 |
| 正文引文 | 段落里写 `[[1]]`、`[[2,3]]`、`[[4-6]]` | 渲染为**上标方括号**（GB/T 7714—2015 顺序编码制） |
| 提示段 | `{"type":"callout","title":"说明","text":"…"}` | **默认无底色无边框**，只用加粗引导词
（论文正文不惯用彩色框）；确需方框时加 `"box": true` |
| 分页 | `{"type":"pagebreak"}` | 少用，优先靠标题分页 |

**表题/图题只写标题文字即可**（写 `表1 xxx` 也行，会被自动重排为正确连续编号）。
但正文里必须出现对应引用：`见图1`、`见表2`、`见式（1）`——质检会检查这一点。

### 公式写法（LaTeX 子集 → OMML）

支持 `_` `^`、`\frac{}{}`、`\sqrt{}`、`\sqrt[n]{}`、`\sum_{}^{}`、`\prod`、`\int`、`\left( \right)`、
`\hat{}` `\bar{}` `\tilde{}` `\vec{}`、`\mathrm{}` `\text{}`（正体，中文也用它）、
希腊字母 `\alpha \beta \theta \Delta \omega …`、运算符 `\times \le \approx \pm \infty \to …`。

```jsonc
{"type":"equation","latex":"\\mathrm{RMSE} = \\sqrt{\\frac{1}{M}\\sum_{i=1}^{M}\\left( \\hat{x}_i - x_i \\right)^2}"}
```

不要用图片或纯文本冒充公式，也不要在正文里写 `x_i` 这种下划线（会被当成普通字符）。

## 3. 生成 docx

```bash
python <SKILL_DIR>/scripts/build_docx.py work/content.json -o work/report.docx
```

需要配色/字号变体时，复制一份 spec 覆盖（只改下面这些"允许改"的项）：

```bash
python <SKILL_DIR>/scripts/build_docx.py work/content.json -o work/report.docx --spec work/spec.json
# spec.json 例：{"colors":{"navy":"1F3A5F","blue":"0072B2"},"sizes":{"body":12}}
```

## 4. 更新目录/域 + 导出 PDF（关键，别省）

Word COM（Windows，推荐）：

```powershell
pwsh -File <SKILL_DIR>/scripts/postprocess.ps1 -Docx work/report.docx -Pdf work/report.pdf
```

这一步做三件事：`Fields.Update()` → `TablesOfContents.Item(1).Update()` → 保存+导 PDF，
最后再跑一次 `set_view.py` 把视图固定为**页面视图 100%**（否则 Word 可能把视图存成 Web 版式）。

没有 Word（macOS/Linux）：

```bash
bash <SKILL_DIR>/scripts/postprocess.sh work/report.docx
```

LibreOffice 不会更新 TOC 域，文档已设 `w:updateFields`，用户用 Word 打开时会自动更新；
交付时要在说明里写清"打开后如目录未展开，Ctrl+A → F9"。

## 5. 质检（必做，交付前跑）

```bash
python <SKILL_DIR>/scripts/qa_check.py work/report.docx --pdf work/report.pdf
python <SKILL_DIR>/scripts/render_preview.py work/report.pdf -o work/preview
python <SKILL_DIR>/scripts/apd.py tone --docx work/report.docx     # 文风自检
```

> 如果第 3 步用了 `--spec` 整体换配色/字号，质检也要带同一份：
> `qa_check.py work/report.docx --spec work/spec.json`，否则颜色/字号检查项会报 FAIL。

- `qa_check.py` 把规范变成 30+ 项检查（页面/视图/样式/三线表/图表题/图片等比/OMML/目录/
  页眉页脚/参考文献/字体/占位符/文风…），有 FAIL 必须修到通过。文风一项**只报 WARN 不判失败**
  （分数是启发式的，见 §11）。
- `render_preview.py` 把 PDF 逐页转 PNG，并自动查空白页、未定义书签、原始域代码。
- **必须真的打开预览图逐页看**：孤行、图题与图不同页、表格只剩表头、算式错位这类问题机器查不出来。
- 人工核对还有 `references/checklist.md`。

## 6. 交付

```text
report.docx   可直接打开：页面视图 100%、目录可更新、公式可编辑、表格可编辑、图片可替换
report.pdf    预览/送审用
图表源文件或来源说明 + 参考文献列表 + 数据来源说明
```

## 7. Python API（不想用 JSON 时）

```python
import sys; sys.path.insert(0, r"<SKILL_DIR>/scripts")
from docx_builder import PaperBuilder

b = PaperBuilder()                 # 可传 spec 覆盖
b.setup()                          # 页面 + 样式 + docDefaults
b.add_cover({"org": "…", "title": "…", "info": [["作者", "…"]]})
b.add_abstract_zh("摘　要", ["…"], ["关键词1", "关键词2"], page_break=True)
b.add_toc()                        # 目录域
b.add_heading1("1 引言", page_break=True)
b.add_body("…")
b.add_table(headers=["列1", "列2"], rows=[["1", "2"]], caption="示例表",
            widths=[6, 9.5], note="注：…")
b.add_figure("figs/fig1.png", caption="示例图", width_cm=14.0)
b.add_equation(r"y = ax + b")
b.add_references(["作者. 题名[J]. 刊名, 2020, 1(2): 3-4."])
b.add_header_footer({"title": "…", "header_type": "…"})
b.doc.save("report.docx")
```

完整可运行示例见 `examples/python_api.py`；可跑通的成品样例见 `examples/demo/`。

## 8. 允许改 / 禁止改

| 允许改（属于内容或整体换肤） | 禁止改（格式系统） |
|---|---|
| 标题、章节、正文、图表数据、图片、参考文献 | 页面尺寸与页边距 |
| 封面字段（增删 info 行） | 字体族（黑体/宋体 + Arial/Times） |
| 整体配色（≤4 个主色，色盲友好，黑白可辨） | 字号层级 |
| 有无附录/致谢等章节 | 三线表规则、图题在下/表题在上 |
| 图表数量与编号（由脚本自动连续） | 目录域、页眉页脚、OMML 公式、页面视图 |

换配色时务必**整组替换**（深色/强调/辅助/浅底），不要只改标题颜色——图表 PNG 也要一起重做
（`scripts/make_chart.py` 读同一套色板）。

## 9. 常见故障

| 现象 | 原因 / 修复 |
|---|---|
| 目录显示成 `TOC \o "1-2" ...` 原始域 | 域结构被破坏；删掉整个目录区域重插，或直接重跑 `postprocess.ps1` |
| 目录后多一张空白页 | 分页符段落 + 标题 `page_break_before` 同时生效；只留一种 |
| 表格跨页只剩表头 | 表头 `w:tblHeader` 缺失 / 行未 `w:cantSplit`；用 `add_table` 会自动处理 |
| 图片变形 | 只设宽度即可（脚本已等比）；不要手动拉伸 |
| 参考文献整段跑出页面 | `w:hanging` 单位写错（把 EMU 当 twips）；用 `hanging_cm=` 参数 |
| Word 里字体不是你设的 | 样式里残留主题字体（`asciiTheme`）；`set_run_font` 已清除，手写代码别漏 |
| `.ps1` 报语法错误/中文乱码 | 文件必须存成 **UTF-8 with BOM**，Windows PowerShell 5.1 才认 |
| 视图变成 Web 版式 | 最后一次保存后又让 Word 打开过；重跑 `set_view.py` |

更多见 `references/troubleshooting.md`（含规范 §19 全部条目）。

## 10. 目录结构

```text
academic-paper-docx/
├── SKILL.md                    本文件（agent 入口）
├── README.md                   项目说明（GitHub 首页）
├── requirements.txt
├── scripts/
│   ├── apd.py                  统一 CLI：doctor/build/post/qa/preview/chart/tone/demo
│   ├── build_docx.py           内容模型 JSON → docx（主入口）
│   ├── docx_builder.py         格式库 / Python API（SPEC 在文件顶部）
│   ├── palette.py              唯一配色源（Okabe-Ito / Tol / Viridis 家族）
│   ├── figures.py              插图渲染（毫米版面 / 1000 dpi / 矢量 / 灰度自检）
│   ├── omml_latex.py           LaTeX 子集 → OMML
│   ├── ai_tone_check.py        「AI 味」启发式检测（只报警，不改稿）
│   ├── make_chart.py           规范配色图表 PNG
│   ├── postprocess.ps1         Word COM：更新域/目录 + PDF + 固定视图
│   ├── postprocess.sh          LibreOffice 兜底
│   ├── set_view.py             只改 settings.xml 强制 print 视图
│   ├── render_preview.py       PDF → PNG 逐页预览 + 体检
│   └── qa_check.py             规范自动质检（40 项，含配色/精度/分辨率/文风）
├── references/
│   ├── format-spec.md          完整规范（权威，35 KB）
│   ├── figure-table-standard.md 图表规范（核对 GB/T 7713.2 与期刊惯例）
│   ├── palette.md              配色规范（色板、色盲、灰度、语义）
│   ├── anti-ai-academic-zh.md  学术文风规则（academic-humanizer 六层，MIT）
│   ├── checklist.md            交付前人工核对清单
│   └── troubleshooting.md      已知坑与修复
├── examples/
│   ├── minimal/content.json    最小示例（一页，无图）—— 建议从这里开始
│   ├── demo/                   完整样例（content_model.json + charts + figs）
│   └── python_api.py           Python API 示例
├── tests/                      133 个 pytest 用例（改格式后跑一遍）
└── docs/                       落地页 + 截图 + 示例 PDF
```

> 改了 `scripts/` 里的格式逻辑后，务必跑 `python -m pytest -q` 与
> `python scripts/apd.py qa` 验证，不要只靠肉眼。

## 10.5 前置部分顺序：摘要在前、目录在后

**默认顺序：封面 → 摘要（中文）→ Abstract → 目录 → 正文。**

依据 **GB/T 7713.1—2025《信息与文献 编写规则 第 1 部分：学位论文》附录 A 表 A.1**：
致谢 → **摘要页** → 序言/前言 → **目次页** → 插图和附表清单；
5.2.9 还要求**英文摘要另起一页**置于中文摘要之后。
中科院、武汉科技大学（本科毕业设计）、华东师范大学的规范与此一致。

少数学校模板要求"目录在前"，在内容模型里加一行即可切换：

```jsonc
{ "front_order": "toc-first" }
```

`apd qa` 会检查该顺序（判 WARN，因为校规可能不同）。
完整证据与原文摘录见 `references/front-matter-order.md`。

## 10.6 期刊投稿模板（`--profile csae`）

除学位论文，还可按《农业工程学报》实测版式出稿（双栏、黑体标题、9 pt 摘要、
楷体图题、0.75/0.5 pt 三线表、奇偶页页眉）：

```bash
python <SKILL_DIR>/scripts/apd.py build work/paper.json -o work/paper.docx --profile csae
python <SKILL_DIR>/scripts/apd.py qa    work/paper.docx --profile csae
```

内容模型需补 `title_en` / `authors` / `affiliations` / `abstract_en` / `journal`
等字段，完整样例见 `examples/journal_csae/content.json`；
参数表与实测依据见 `references/journal-csae.md`。

> 学位论文用默认 `--profile thesis`（或不写），不要混用：两者的页边距、字号、
> 栏数、三线表线宽都不同。

## 11. 图表与配色（GB/T 7713.2 + 期刊惯例）

完整规范：`references/figure-table-standard.md`、`references/palette.md`。

**不要自己配颜色、自己调像素**，按下面的流程走：

### 11.1 出图

```bash
python <SKILL_DIR>/scripts/figures.py work/fig1.json -o work/figs/fig1.png --grayscale
```

图的 JSON 关键字段：

```jsonc
{
  "kind": "line|bar|grouped_bar|scatter|errorbar|box|heatmap",
  "xlabel": {"label": "采样周期", "symbol": "t", "unit": "s"},
  "ylabel": {"label": "平均功耗", "symbol": "P", "unit": "mW"},
  "width_mm": "half | full | 120",      // 半栏 75 / 通栏 155（= 版心宽）
  "style": "cn | nature | ieee",
  "dpi": 1000, "formats": ["png", "pdf"],
  "series": [{"name": "本文", "values": [15.9, 10.4], "error": [0.5, 0.4]}],
  "brackets": [{"x1": 0, "x2": 1, "label": "**"}],   // 显著性
  "cmap": "viridis"                     // 热图；jet/rainbow 会直接报错
}
```

配色可选"主题色卡"（9 套设计稿配色，自动取安全子集并打印实测判据）：

```bash
python <SKILL_DIR>/scripts/apd.py palette            # 列出全部色卡与 3/4/5 类判定
python <SKILL_DIR>/scripts/figures.py spec.json -o f.png --palette rose6
```

> 色卡是"整体协调"的设计稿，不等于"N 类分类色板"。同色系深浅当分类色会糊，
> 本 skill 会用 `safe_subset()` 自动挑可区分的子集，并给出 CVD/灰度实测值。

出图时的硬约束（违反会报错，不是建议）：
- 颜色只能取自 `palette.SERIES`（Okabe-Ito）；红绿色盲下同色的组合直接拒绝。
- 定量数据禁用 `jet`/`rainbow`。
- 每个系列自动配"颜色 + 线型 + 标记"三重编码，黑白打印与色觉障碍都能读。

### 11.2 配表

表用 `columns` 给列元信息，表头自动写成"量名称/单位"，数值自动按位数格式化：

```jsonc
{"type": "table",
 "caption": "不同采样周期下的功耗", "caption_en": "Power versus sampling period",
 "columns": [
   {"label": "采样周期", "unit": "s", "digits": 0},
   {"label": "平均功耗", "unit": "mW", "digits": 1}
 ],
 "rows": [[30, 15.9], [60, 10.4]],
 "note": "注：a 为模拟数据。标 * 者 P<0.05。"}
```

自动做到：顶/底线 1.5 pt、栏目线 **0.75 pt**、无竖线、表头无斜线、表头重复、
行不跨页、同列小数位一致、`^{a}` 渲染为上标、中英对照题注。

**空缺值按 GB/T 7713.2—2022 5.5.4 区分**：`None`/`""` → 空白（无此项或未测量）；
要表示"测量过而未发现"就在数据里显式写 `"—"`；实测为零写 `0`。三者语义不同，不要混用。

### 11.3 尺寸与版面约束（GB/T 7713.2—2022 5.4.1）

标准原文只要求插图"**大小适当，图中文字清晰可见**"（5.4.1），毫米数由版心与 CY/T 171 推。
本 skill 把可算的都做成硬约束：

| 约束 | 取值 | 越界后果 |
|---|---|---|
| 图宽 | 半栏 75 mm / 通栏 **≤155 mm**（= 版心 16.0 cm 减余量） | `figures.py` 直接报错 |
| 图高 | **≤22 cm** | 生成时警告（图+题+注会跨页） |
| 图内文字插入后 | **≥6 pt** | 生成时警告，并指出原设计字号与缩放比 |
| 图题/表题 | `图N` + **1 个全角空格** + 题名，居中 | 自动 |
| 图注/表注 | 置于图下/表下，末尾自动补"。" | 自动 |
| 先文后图/表 | 正文提及在前（5.4.5 / 5.5.5） | `apd qa` 检查 |

出图时按**最终插入尺寸**设计：半栏图就写 `"width_mm": "half"`。
拿通栏图缩到半栏，155 mm 的 8.5 pt 会变成 4.1 pt——这正是 5.4.1 要避免的。
具体算法与实测表见 `references/figure-table-standard.md` §1.1。

### 11.4 交付前自检

`qa_check.py` 里与图表/配色相关的检查项（任一 FAIL 必须修）：

| 组 | 检查 |
|---|---|
| 表格 | 三线表 1.5/0.75 pt 无竖线；同列小数位一致 |
| 图片 | 分辨率 ≥600 dpi（线条图目标 1000）；宽 ≤版心、高 ≤22 cm；插入后图内文字 ≥6 pt；等比不变形 |
| 配色 | 全稿用色来自 `palette.py`；彩色 ≤5 种 |

## 12. 正文文风：按 academic-humanizer 的逻辑改稿

改稿规则整套采用
[AIScientists-Dev/academic-humanizer](https://github.com/AIScientists-Dev/academic-humanizer)（MIT），
中文适配见 `references/anti-ai-academic-zh.md`。**本 skill 不再自研去 AI 味规则。**

一条原则：**没有动词可以强过它的证据。** 每一条主张都要挣得自己的数字、图或引用。

四步（与上游一致）：**Read**（认文体与场合）→ **Audit**（列命中与证据状态，不动笔）→
**Rewrite**（结构内容不变，去痕迹、把过强的话降到证据水平）→ **Report**（改动清单）。

六层结构：

| 层 | 内容 |
|---|---|
| Layer 1 | 通用 AI 痕迹（意义拔高、宣传词、无源权威、长句、破折号…） |
| Layer 2 | 学术 AI 痕迹 2.1–2.11（过度断言动词、显著性炒作、新颖性填充、贡献清单套话…） |
| Layer 3 | **必须保留**（有证据的模糊限制、被动语态、"我们"、术语、公式；数字与引用一字不改） |
| Layer 4 | 主张-证据纪律（缺证据→补指针或降调；动词强于证据→降级；量级模糊→区间+口径） |
| Layer 5 | 语体与场合匹配（期刊风格、作者本人节奏） |
| Layer 6 | 基金本子模式（NSF/NIH：保留愿景，查"主张 ↔ 可行性"） |

脚本只做可枚举部分的 **Audit**，输出两个分数：

| 分数 | 含义 | 对应层 |
|---|---|---|
| `tells_score` | AI 痕迹密度 | Layer 1 + Layer 2 |
| `claims_score` | 主张强于证据的程度 | Layer 4 |

```bash
python <SKILL_DIR>/scripts/apd.py tone --docx work/report.docx
python <SKILL_DIR>/scripts/apd.py tone --docx work/report.docx --audit    # 审计清单
python <SKILL_DIR>/scripts/apd.py tone --json work/content.json --max-score 40   # CI 门禁

# 基金本子（Layer 6）：结构 + aims 质量 + 主张↔可行性
python <SKILL_DIR>/scripts/apd.py tone --text aims.md --mode proposal --agency nih
```

**Layer 6 只查结构性问题**（P1 结构缺失、P2 aims 数量、P3 方法当目标、P4 多米诺依赖、
P5 雄心缺可行性、P6 中心假设含糊、P7 Broader Impacts 模板化）；
"这个 aim 值不值得做"仍需人判断。

改稿时守三条：**数字/公式/引用不许改**；**有证据的模糊限制必须保留**（把 suggest 改成
prove 是制造过度断言）；**主张必须落到数字/图/表/引用**。
`qa_check` 的文风项只提示（≥40 分告警），不判 FAIL。

