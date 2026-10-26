---
name: academic-paper-docx
description: 生成符合《学术论文 / 研究报告 Word 排版与生成规范》的 .docx（蓝绿学术配色、黑体标题+宋体正文、三线表、OMML 可编辑公式、自动目录、页码页眉、GB/T 7714 参考文献、默认页面视图）。当用户要写/生成/排版学术论文、毕业设计、课程报告、研究报告、课题组汇报的 Word 文档，或要求"按某套格式规范出 Word"时使用。
whenToUse: 用户要求产出 .docx 论文/报告/毕业设计，或要求按学术排版规范（三线表、图题表题、自动目录、OMML 公式、GB/T 7714）生成 Word 文档时。
---

# 学术论文 Word 生成（academic-paper-docx）

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
  "toc": true, "toc_title": "目　录", "toc_levels": "1-1",
  "sections": [ /* 见下表 */ ],
  "references": ["[1] …", "…"]            // 不写 [n] 时自动补号
}
```

`sections` 里可用的块：

| 块 | 写法 | 说明 |
|---|---|---|
| 一级标题 | `{"level":1,"text":"1 引言"}` | 编号写进 text，全文统一（`1` 或 `一、`）；首个一级标题自动另起一页 |
| 二级标题 | `{"level":2,"text":"1.1 研究背景"}` | 不自动加编号 |
| 正文段 | `{"type":"p","text":"…"}` | 12 pt / 1.5 倍 / 首行 2 字符 / 两端对齐 |
| 编号列表 | `{"type":"list","items":["…","…"]}` | 生成（1）（2）（3），不用圆点 |
| 三线表 | `{"type":"table","caption":"…","headers":[…],"rows":[[…]],"widths":[3,3.8,8.7],"aligns":["center","left","left"],"note":"注：…"}` | 宽度自动归一到 15.5 cm；编号自动 |
| 图 | `{"type":"figure","path":"figs/x.png","caption":"…","note":"注：…","width_cm":14}` | 只给宽度 → 等比缩放；图题在下 |
| 公式 | `{"type":"equation","latex":"\\hat{x}_k = \\frac{1}{N}\\sum_{i=0}^{N-1} z_{k-i}","number":"（1）"}` | 转成 Word 原生 OMML；编号右对齐 |
| 提示框 | `{"type":"callout","title":"说明","text":"…"}` | 浅蓝底 + 淡蓝边框 |
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
| 目录显示成 `TOC \o "1-1" ...` 原始域 | 域结构被破坏；删掉整个目录区域重插，或直接重跑 `postprocess.ps1` |
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
│   ├── omml_latex.py           LaTeX 子集 → OMML
│   ├── ai_tone_check.py        「AI 味」启发式检测（只报警，不改稿）
│   ├── make_chart.py           规范配色图表 PNG
│   ├── postprocess.ps1         Word COM：更新域/目录 + PDF + 固定视图
│   ├── postprocess.sh          LibreOffice 兜底
│   ├── set_view.py             只改 settings.xml 强制 print 视图
│   ├── render_preview.py       PDF → PNG 逐页预览 + 体检
│   └── qa_check.py             规范自动质检（36 项，含文风提示项）
├── references/
│   ├── format-spec.md          完整规范（权威，35 KB）
│   ├── anti-ai-academic-zh.md  「AI 味」规则、学术豁免清单、改写协议
│   ├── checklist.md            交付前人工核对清单
│   └── troubleshooting.md      已知坑与修复
├── examples/
│   ├── minimal/content.json    最小示例（一页，无图）—— 建议从这里开始
│   ├── demo/                   完整样例（content_model.json + charts + figs）
│   └── python_api.py           Python API 示例
├── tests/                      104 个 pytest 用例（改格式后跑一遍）
└── docs/                       落地页 + 截图 + 示例 PDF
```

> 改了 `scripts/` 里的格式逻辑后，务必跑 `python -m pytest -q` 与
> `python scripts/apd.py qa` 验证，不要只靠肉眼。

## 11. 正文的「AI 味」：怎么认、怎么改、哪里不能改

完整规则在 `references/anti-ai-academic-zh.md`（含 before/after 与学术豁免清单）。
这里只留执行要点。

**写作时（不是写完再洗）就避开这 6 类**：

| 类别 | 典型句 | 改成 |
|---|---|---|
| 空转意义句 | 具有重要的理论意义和现实价值 | 这个结论约束了什么 |
| 数据缺席的拔高 | 显著提升了精度 | 指标 + 数值 + 对比 + 条件 |
| 假对比/假递进 | 不仅是…更是…；不是…而是… | 直接说肯定项 |
| 段末复述收尾 | 由此可见…的重要性 | 删掉，或写带条件的推论 |
| 元话语铺垫 | 本节将…；随着…的快速发展 | 直接给内容 |
| 术语漂移 | 节点/采集端/终端设备混用 | 首次定义后固定不变 |

**不要洗掉的（学术语体，不是 AI 味）**：方法段的被动句、术语化名词化（收敛性、有限元分析）、
必要限定（可能/尚未验证）、每章一次的"综上所述"、引用格式、公式符号、图表题注、参考文献。

**改写协议**：先标记（S 一次即改 / A 重复即改 / B 需凑证据）→ 再重写（不新增事实、不删独立信息、
不改确定程度、术语固定、句长交替）→ 后检（朗读 + 查 S 级残留）→ 直接交成稿。

**自检**：

```bash
python <SKILL_DIR>/scripts/apd.py tone --docx work/report.docx
python <SKILL_DIR>/scripts/apd.py tone --json work/content.json --max-score 40   # CI 门禁
```

分数只是**定位工具**：高不等于 AI 写的（模板公文、翻译腔也会高），低也不等于合格
（真问题常在论证与数据）。它列出规则编号、段落号和原句，改不改由人判断。


