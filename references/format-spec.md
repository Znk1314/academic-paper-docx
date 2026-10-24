# 学术论文 / 研究报告 Word 文档排版与生成规范

> 本规范用于指导另一个 Agent 或排版人员，生成与目标示例风格一致、可直接用于论文初稿、课程报告或课题组汇报的 Word 文档。
>
> 本规范描述的是**格式系统**，不是固定内容。论文主题、章节标题、图表数量、参考文献都可以变化；页面、字体、字号、配色、标题层级、三线表、图题、公式、页眉页脚、目录等格式必须保持一致。

> **硬性要求：生成的 `.docx` 必须默认以页面视图 / 打印布局打开，缩放 100%，不能是 Web 版式。**
>
> 具体设置见第 18 节。Word COM 保存后可能把视图改回 Web，所以必须在最后一次保存之后重新设置 `w:view=print`。

> **使用边界：本文档只提供格式模板，不提供论文内容模板。**
>
> - 文中的章节标题、表题、图题、参考文献、封面信息都是**占位示例**，必须全部替换。
> - 不要复用示例中的项目名称、学校、姓名、数据、图表、文献。
> - 第 9 节的结构只是“一般论文可以长什么样”的参考，不是必须遵循的目录。
> - 生成新文档时，唯一必须保留的是格式参数和排版规则，不是示例内容。

---

## 1. 成品定位

目标成品应具备以下特征：

- A4 纵向论文 / 报告版式；
- 蓝绿色学术配色，低饱和、清爽、简洁；
- 中文为主，英文摘要、术语、参考文献保留；
- 标题黑体、正文宋体 / Times New Roman；
- 正文小四、1.5 倍行距、首行缩进 2 字符；
- 标题层级清晰，使用 Word 内置标题样式；
- 自动目录域，可在 Word 中更新；
- 页眉页脚、页码完整；
- 数据表统一为三线表，不使用竖线；
- 图题在下、表题在上；
- 公式使用 Word 内置 OMML 可编辑公式，编号右对齐；
- 图表编号连续，正文有“见图 X / 见表 X / 见式（X）”的引用；
- 参考文献按 GB/T 7714 格式；
- 不虚构数据，无法确认处标注 `[待补充/待核实]`；
- **打开 Word 默认进入页面视图 / 打印布局，缩放 100%。**

---

## 2. 页面设置

| 项目 | 值 |
|---|---|
| 纸张 | A4，21.0 cm × 29.7 cm |
| 上边距 | 2.4 cm |
| 下边距 | 2.2 cm |
| 左边距 | 2.6 cm |
| 右边距 | 2.4 cm |
| 页眉距离 | 1.2 cm |
| 页脚距离 | 1.0 cm |
| 装订线 | 0 cm |
| 首页不同 | 是（封面不显示页眉页脚） |

推荐使用 `python-docx` 设置：

```python
from docx.shared import Cm

sec = doc.sections[0]
sec.page_width = Cm(21.0)
sec.page_height = Cm(29.7)
sec.top_margin = Cm(2.4)
sec.bottom_margin = Cm(2.2)
sec.left_margin = Cm(2.6)
sec.right_margin = Cm(2.4)
sec.header_distance = Cm(1.2)
sec.footer_distance = Cm(1.0)
sec.different_first_page_header_footer = True
```

正文可用文字宽度约为 **15.8 cm**。图片最大宽度不要超过 15.5 cm。

---

## 3. 色彩系统

全文只使用以下 3—4 个主色系，避免高饱和荧光色和过度装饰。

| 用途 | 颜色名 | HEX |
|---|---|---|
| 主标题、深色强调 | 深海军蓝 | `#1F3A5F` |
| 二级标题、强调线 | 学术蓝 | `#0072B2` |
| 辅助强调 | 暖橙 | `#E69F00` |
| 正向/完成 | 学术绿 | `#009E73` |
| 正文 | 深灰 | `#262626` |
| 次级说明 | 中灰 | `#595959` |
| 浅色提示框 | 浅蓝 | `#EAF2FA` |
| 浅色边框 | 淡蓝灰 | `#B4C6E7` |
| 白色 | 白 | `#FFFFFF` |

配色原则：

- 全文同一变量、同一类别、同一实验组保持同色；
- 图表配色优先使用 Okabe-Ito / ColorBrewer 等色盲友好方案；
- 同一张图不超过 4 个主色；
- 黑白打印时仍能通过线型、标记和灰度区分；
- 不使用渐变阴影、3D 柱状图、花哨背景。

---

## 4. 字体与字号

### 4.1 中英文字体

| 元素 | 中文字体 | 英文字体 |
|---|---|---|
| 标题 | 黑体 | Arial |
| 正文 | 宋体 | Times New Roman |
| 表格 | 宋体 | Times New Roman |
| 图题/表题 | 黑体 | Arial |
| 图注/表注 | 宋体 | Times New Roman |
| 页眉页脚 | 宋体 | Times New Roman |
| 参考文献 | 宋体 | Times New Roman |

### 4.2 字号

| 元素 | 字号 |
|---|---|
| 封面主标题 | 24 pt |
| 封面副标题 | 20 pt |
| 封面顶部项目类型 | 15 pt |
| 一级标题 | 16 pt |
| 二级标题 | 13 pt |
| 正文 | 12 pt（小四） |
| 英文摘要 | 11 pt |
| 关键词 | 10.5 pt |
| 表题/图题 | 9.5 pt |
| 表格正文 | 9 pt（小五；密集表可 8.8 或 8） |
| 表格注释/图注 | 8.5 pt |
| 页眉 | 8.5 pt |
| 页脚 | 8 pt |
| 参考文献 | 9 pt |

### 4.3 字体设置代码

```python
from docx.oxml.ns import qn
from docx.shared import Pt

def set_run_font(run, cn="宋体", en="Times New Roman", size=12,
                 bold=False, color=None, italic=False):
    run.font.name = en
    run.font.size = Pt(size)
    run.bold = bold
    run.italic = italic
    if color:
        run.font.color.rgb = color
    rPr = run._element.get_or_add_rPr()
    rFonts = rPr.find(qn("w:rFonts"))
    if rFonts is None:
        rFonts = OxmlElement("w:rFonts")
        rPr.append(rFonts)
    rFonts.set(qn("w:ascii"), en)
    rFonts.set(qn("w:hAnsi"), en)
    rFonts.set(qn("w:eastAsia"), cn)
```

---

## 5. 段落与标题样式

### 5.1 正文

- 字体：宋体 / Times New Roman；
- 字号：12 pt；
- 行距：1.5 倍；
- 首行缩进：2 字符，约 24 pt；
- 段后：6 pt；
- 对齐：两端对齐；
- 不手动敲空格缩进，使用段落格式。

```python
p.paragraph_format.line_spacing = 1.5
p.paragraph_format.first_line_indent = Pt(24)
p.paragraph_format.space_after = Pt(6)
p.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
```

### 5.2 一级标题

- 字体：黑体 / Arial；
- 字号：16 pt；
- 加粗；
- 颜色：`#1F3A5F`；
- 段前 16 pt，段后 8 pt；
- 左对齐；
- 段落下边框：0.75 pt，颜色 `#0072B2`；
- 与下段同页（keep with next）；
- 使用 Word `Heading 1` 样式，便于自动目录。

### 5.3 二级标题

- 字体：黑体 / Arial；
- 字号：13 pt；
- 加粗；
- 颜色：`#0072B2`；
- 段前 10 pt，段后 4 pt；
- 左对齐；
- 与下段同页；
- 使用 Word `Heading 2` 样式。

### 5.4 标题编号

- 一级标题：`一、二、三、……` 或 `1 2 3 ……`，全文统一；
- 二级标题：`3.1 3.2 3.3 ……`；
- 不使用 Word 自动多级列表，避免跨平台编号错乱；
- 标题编号与正文引用、目录保持一致。

---

## 6. 封面

封面结构（以下全部是占位符，必须替换）：

```text
[学校 / 单位名称] [项目类型 / 论文类型]
[论文 / 报告标题]
[副标题]
[说明行，可选]

[无边框信息表]
[编号标签]      [编号]
[作者 / 负责人] [姓名]
[学院 / 部门]   [名称]
[专业 / 班级]   [名称]
[指导教师]      [姓名 职称]
[项目 / 研究期限] [起止时间]
[日期标签]      [日期]

[学校 / 单位名称]  [学院 / 部门]
[日期]
```

格式：

- 顶部项目类型 / 论文类型：15 pt，黑体 / Arial，蓝色；
- 主标题：24 pt，黑体 / Arial，深海军蓝；
- 副标题：20 pt，黑体 / Arial，学术蓝；
- 说明行：11 pt，宋体 / Times New Roman，灰色；
- 一条 1.5 pt 学术蓝横线；
- 信息表：2 列，无边框，左列加粗深蓝，右列正文色；
- 底部单位与日期：12 pt，灰色，居中。

注意：

- 不要复用示例中的学校、项目名称、姓名、编号、日期；
- 封面信息表只保留实际需要的字段，不必强行凑满；
- 封面不显示页眉页脚，设置 `different_first_page_header_footer = True`。

---

## 7. 摘要与关键词

### 7.1 中文摘要

- 标题“摘 要”：一级标题样式；
- 正文：12 pt，1.5 倍行距，首行缩进 2 字符；
- 内容：研究背景、方法、阶段结果、下一步，200—350 字；
- 不写未完成内容为已完成；
- 不虚构指标。

### 7.2 关键词

```text
关键词：关键词1；关键词2；关键词3；关键词4；关键词5
```

- “关键词：”用黑体加粗；
- 关键词用宋体/Times New Roman；
- 10.5 pt；
- 分号分隔，末尾不加句号。

### 7.3 英文摘要

- 标题“Abstract”：二级标题样式；
- 正文：11 pt，1.5 倍行距；
- 与中文摘要内容对应；
- 英文关键词：

```text
Keywords: keyword1; keyword2; keyword3; keyword4; keyword5
```

---

## 8. 目录

### 8.1 目录形式

- 使用 Word 自动目录域；
- 建议只显示一级标题，保证目录一页：
  `TOC \o "1-1" \h \z \u`
- 如果一级标题超过 15 个，可显示二级，但需要把目录样式字体缩小到 10 pt 左右，并减少段间距。

### 8.2 目录标题

- “目 录”：18 pt，黑体 / Arial，深海军蓝，居中；
- 不使用 Heading 1 样式，避免目录把自己列进去。

### 8.3 目录域代码

```python
def add_toc_field(paragraph):
    run = paragraph.add_run()
    begin = OxmlElement("w:fldChar")
    begin.set(qn("w:fldCharType"), "begin")
    instr = OxmlElement("w:instrText")
    instr.set(qn("xml:space"), "preserve")
    instr.text = 'TOC \\o "1-1" \\h \\z \\u'
    sep = OxmlElement("w:fldChar")
    sep.set(qn("w:fldCharType"), "separate")
    text = OxmlElement("w:t")
    text.text = "右键此处选择“更新域”生成目录"
    end = OxmlElement("w:fldChar")
    end.set(qn("w:fldCharType"), "end")
    for el in (begin, instr, sep, text, end):
        run._r.append(el)
```

### 8.4 更新方式

- 生成 docx 后，用 Word COM 打开：
  1. `Fields.Update()`；
  2. `TablesOfContents.Item(1).Update()`；
  3. 保存并导出 PDF。
- 不要在 Word 更新后手动删除/移动目录段落，否则容易破坏 TOC 域。
- 如果 TOC 显示为 `TOC \o ...` 原始代码，说明域结构被破坏，应删除整个目录区域后重新插入 TOC 域。

---

## 9. 结构示例（仅参考，不是格式要求）

> 下面是一般论文 / 研究报告的常见结构，只用于说明“标题层级如何组织”。
>
> 它不是必须遵循的目录，也不携带任何项目内容。生成新文档时，请完全替换标题、章节数量、顺序和内容；唯一必须保留的是前文的字体、字号、标题样式、段落、目录、表格、图片、公式等格式。
>
> 不要复用示例中的主题、项目名、数据、图表、参考文献。

| 层级 | 示例标题 | 说明 |
|---|---|---|
| 摘要 | 摘要 / Abstract | 中英文摘要、关键词 |
| 一级 | 1 引言 | 背景、问题、目标 |
| 一级 | 2 相关工作 | 文献综述、研究现状 |
| 一级 | 3 方法 | 研究设计、模型、算法 |
| 二级 | 3.1 数据 | 数据来源、预处理 |
| 二级 | 3.2 模型 | 方法细节、公式 |
| 一级 | 4 实验设置 | 数据、指标、对比方法 |
| 一级 | 5 结果与分析 | 主结果、消融、可视化 |
| 一级 | 6 讨论 | 机制、局限性 |
| 一级 | 7 结论与展望 | 总结与未来工作 |
| 一级 | 参考文献 | GB/T 7714 |
| 可选 | 附录 | 补充材料 |

说明：

- 阶段汇报、课程报告、项目申报书可以用完全不同的标题，例如“项目背景”“研究内容”“阶段成果”“存在问题”“后期计划”等；这些属于内容层，不属于格式规范。
- 标题编号可以用“一、二、三”或“1 2 3”，全文统一即可。
- 不要因为示例里是“1 引言”，就强制每篇文档都叫“1 引言”。
- 结构可以比示例多、少、合并或拆分；只要标题层级、样式、目录和编号规则一致，格式就是合格的。

---

## 10. 三线表规范

### 10.1 总体要求

- 全文数据表统一三线表；
- 仅保留顶线、表头下线和底线；
- 不使用竖线；
- 表格居中；
- 表格宽度适配 A4 正文宽度，约 15.5 cm；
- 表题在表格上方，格式“表X 表题”，居中；
- 表注在表格下方，说明数据来源、缩写、显著性、缺失值；
- 表头重复，长表跨页时重复表头；
- 表格行禁止跨页拆分；
- 表内文字 9 pt，表头加粗；
- 数据不得虚构，模拟/示意数据明确标注。

### 10.2 表题

- 9.5 pt，黑体 / Arial，加粗，深海军蓝；
- 居中；
- 段前 6 pt，段后 3 pt；
- 与表格同页（keep with next）。

### 10.3 表注

- 8.5 pt，宋体 / Times New Roman，灰色 `#595959`；
- 左对齐；
- 段前 2 pt，段后 6 pt；
- 以“注：”开头。

### 10.4 三线表边框代码

```python
def set_three_line_table(table):
    tblPr = table._tbl.tblPr
    for el in tblPr.findall(qn("w:tblBorders")):
        tblPr.remove(el)

    borders = OxmlElement("w:tblBorders")
    for edge in ["top", "left", "bottom", "right", "insideH", "insideV"]:
        el = OxmlElement(f"w:{edge}")
        if edge in ("top", "bottom"):
            el.set(qn("w:val"), "single")
            el.set(qn("w:sz"), "12")
            el.set(qn("w:space"), "0")
            el.set(qn("w:color"), "000000")
        else:
            el.set(qn("w:val"), "nil")
        borders.append(el)
    tblPr.append(borders)

    for cell in table.rows[0].cells:
        tcPr = cell._tc.get_or_add_tcPr()
        tcBorders = OxmlElement("w:tcBorders")
        bottom = OxmlElement("w:bottom")
        bottom.set(qn("w:val"), "single")
        bottom.set(qn("w:sz"), "8")
        bottom.set(qn("w:space"), "0")
        bottom.set(qn("w:color"), "000000")
        tcBorders.append(bottom)
        tcPr.append(tcBorders)
```

### 10.5 表头重复与行禁止拆分

```python
# 表头重复
trPr = table.rows[0]._tr.get_or_add_trPr()
tblHeader = OxmlElement("w:tblHeader")
tblHeader.set(qn("w:val"), "true")
trPr.append(tblHeader)

# 每一行禁止跨页
for row in table.rows:
    trPr = row._tr.get_or_add_trPr()
    trPr.append(OxmlElement("w:cantSplit"))
```

### 10.6 列宽

- 表格总宽约 15.5 cm；
- 列宽按内容分配，避免频繁换行；
- 示例：

```python
widths = [3.0, 3.8, 8.7]  # cm，总计约 15.5 cm
for i, w in enumerate(widths):
    for row in table.rows:
        row.cells[i].width = Cm(w)
```

### 10.7 数据格式

- 数值按小数点或有效位数对齐；
- 单位写在表头括号内，如“频率/Hz”“误差/m”；
- 百分比、显著性标记、缺失值在表注说明；
- 同一列保持同一量级和有效位数；
- 表格与图不重复表达同一组数据。

---

## 11. 图片规范

### 11.1 尺寸与分辨率

| 场景 | 推荐宽度 |
|---|---|
| 普通 16:9 示意图 | 14.0 cm |
| 内容密集的流程图/系统图 | 15.0—15.5 cm |
| 竖版图片 | 高度不超过 9.5 cm，宽度按比例 |
| 最大宽度 | 15.5 cm |

分辨率换算参考：

- 1672 × 941 px 的 16:9 图片，插入宽度 14.0 cm 时约 **303 dpi**；
- 同一张图插入宽度 15.5 cm 时约 **274 dpi**；
- 因此，普通图优先 14.0 cm，只有内容密集、文字偏小的流程图才用 15.0—15.5 cm；
- 如果严格坚持 300 dpi，保持 14.0 cm，或先用高质量算法把源图放大到约 1900—2000 px 宽再插入。

- 图片必须保持原始宽高比，不得只改宽或只改高；
- 源图建议 ≥ 1600 px 宽，300 dpi；
- 插入 Word 后保持清晰，不得拉伸变形；
- 实验截图可作为证据图，但需高分辨率并注明来源。

### 11.2 图题与图注

- 图题在图下方，格式“图X 图题”，居中；
- 图题 9.5 pt，黑体 / Arial，加粗，深海军蓝；
- 图注 8.5 pt，宋体 / Times New Roman，灰色；
- 图注以“注：”开头，说明数据来源、缩写、误差线、显著性；
- 图题与图片同页，图片段落设置 `keep_with_next = True`。

### 11.3 图片来源

- 实验截图、实物照片、AI 生图均可；
- AI 生图建议 16:9、300 dpi、白色/浅色背景、蓝绿学术配色；
- AI 生图如要求中文标签，提示词中列出全部中文文字，并要求“无错别字”；
- 生成后必须检查中文错字；如有错字，用同色文本框覆盖；
- 不得用 AI 图伪造实验数据或实验结果；
- 数据图优先使用可编辑图表，不用截图代替。

如需让 AI 直接生成中文标签，可使用下面的提示词骨架：

```text
科研风格[图类型]，[视角/构图]；中心/主体为[主体内容]；
图中中文标签清晰准确，标注“[标签1]”“[标签2]”“[标签3]”……；
[主色]与浅灰配色，白色/浅色背景，扁平化/写实科技风格，
层次清晰，适合学术汇报，16:9，300 dpi，无错别字，无水印。
```

注意：

- 中文标签必须逐项列出，不要只写“加上中文标签”；
- 生成后逐字检查，AI 仍可能把“树莓派”写成“树莓派”以外的错字；
- 错字用同色文本框覆盖，不要用拉伸图片的方式掩盖。

### 11.4 图片插入代码

```python
from docx.shared import Cm
from docx.enum.text import WD_ALIGN_PARAGRAPH

p = doc.add_paragraph()
p.alignment = WD_ALIGN_PARAGRAPH.CENTER
p.paragraph_format.keep_with_next = True
p.paragraph_format.space_before = Pt(6)
p.paragraph_format.space_after = Pt(2)

run = p.add_run()
run.add_picture("figure.png", width=Cm(14.0))
```

---

## 12. 公式规范

### 12.1 总体要求

- 必须使用 Word 内置 OMML 公式，不得用图片或纯文本代替；
- 公式居中；
- 编号右对齐，格式“（1）（2）（3）”；
- 公式编号连续；
- 正文用“见式（1）”引用；
- 变量、参数斜体；函数名、单位、数字、运算符正体；
- 公式后按语法加标点；
- 符号首次出现时解释。

### 12.2 公式布局

推荐使用 1×3 无边框表格：

```text
[空]    [公式 OMML]    [（1）]
```

- 中间单元格居中放公式；
- 右侧单元格右对齐放编号；
- 表格无边框；
- 公式单元格宽度约 11.5 cm，编号单元格约 2 cm。

### 12.3 OMML 示例代码

```python
def mr(text):
    r = OxmlElement("m:r")
    t = OxmlElement("m:t")
    t.text = text
    r.append(t)
    return r

def msub(base, sub):
    s = OxmlElement("m:sSub")
    e = OxmlElement("m:e")
    e.append(mr(base))
    sb = OxmlElement("m:sub")
    sb.append(mr(sub))
    s.append(e)
    s.append(sb)
    return s

omath = OxmlElement("m:oMath")
omath.append(msub("θ", "k"))
omath.append(mr("="))
omath.append(msub("θ", "k-1"))
omath.append(mr("+"))
omath.append(msub("ω", "k"))
omath.append(mr("Δt"))
cell.paragraphs[0]._p.append(omath)
```

---

## 13. 列表与编号

- 全文不使用圆形项目符号 `• ● ○ ▪`；
- 需要列举时统一使用：

```text
（1）……
（2）……
（3）……
```

- 编号段落使用正文样式，不使用 `List Bullet`；
- 11 pt 或 12 pt，1.5 倍行距；
- 段前 2 pt，段后 2 pt；
- 左缩进 0，首行缩进 0；
- 编号与正文之间不加空格。

如果需要：

```python
para.style = doc.styles["Normal"]
r = para.add_run(f"（{i}）{text}")
```

---

## 14. 页眉页脚

### 14.1 页眉

- 首页不同；
- 页眉右对齐；
- 文字：`[文档标题] + “· [文档类型]”`；
- 8.5 pt，宋体 / Times New Roman，灰色；
- 不加横线或只加极细灰线。

### 14.2 页脚

- 居中；
- 格式：`第 X 页 / 共 Y 页`；
- 8 pt，宋体 / Times New Roman，灰色；
- 使用 `PAGE` 和 `NUMPAGES` 域，不手写页码。

### 14.3 页码域代码

```python
def add_field(paragraph, instr):
    run = paragraph.add_run()
    begin = OxmlElement("w:fldChar")
    begin.set(qn("w:fldCharType"), "begin")
    instrText = OxmlElement("w:instrText")
    instrText.set(qn("xml:space"), "preserve")
    instrText.text = instr
    end = OxmlElement("w:fldChar")
    end.set(qn("w:fldCharType"), "end")
    run._r.append(begin)
    run._r.append(instrText)
    run._r.append(end)
```

---

## 15. 参考文献

- 按 GB/T 7714 格式；
- 9 pt，宋体 / Times New Roman；
- 建议悬挂缩进 0.74 cm；
- 编号 `[1] [2] [3]`；
- 正文引用与文献编号对应；
- 不引用与正文无关的文献；
- 不虚构文献；
- 无法确认的文献标注 `[待补充/待核实]`。

示例（仅格式，必须替换成真实文献）：

```text
[1] 作者. 题名[J]. 刊名, 年, 卷(期): 起止页码.
[2] 作者. 书名[M]. 版次. 出版地: 出版社, 年: 页码.
[3] 作者. 题名[C]//会议名称. 出版地: 出版者, 年: 页码.
[4] 作者. 题名[D]. 城市: 学校, 年.
```

注意：

- 以上不是真实文献，只是 GB/T 7714 的格式占位；
- 不要把示例文献直接写进新文档；
- 正文引用编号必须与文献表一一对应。

---

## 16. 图表编号与交叉引用

- 图：`图1 图2 图3 ……`；
- 表：`表1 表2 表3 ……`；
- 公式：`（1）（2）（3）……`；
- 正文引用：

```text
见图1
见表2
见式（1）
```

- 图表编号在全文连续，不跳号；
- 图表标题与正文引用必须一致；
- 图表与正文、参考文献数据一致；
- 如使用 Word 交叉引用域，需在最终 PDF 中检查是否显示为“错误！未定义书签”。

---

## 17. 推荐生成流程

### 17.1 总体流程

```text
1. 整理内容模型
2. 生成 docx 骨架
3. 设置页面、样式、颜色
4. 写封面、摘要、目录
5. 写正文、表格、公式、图片
6. 写参考文献
7. 设置页眉页脚、页码
8. 保存 docx
9. Word COM 更新目录与域
10. 导出 PDF
11. 渲染 PDF 页面逐页检查
12. 修复版式问题
13. 最终 QA
```

### 17.2 内容模型

在写代码前，先确定：

```text
文档标题
副标题
项目/作者/单位/日期
一级标题列表
二级标题列表
每个表格的表头、行数据、列宽、表注
每张图的路径、图题、图注、宽度
公式及编号
参考文献列表
```

不要一边写代码一边临时决定结构，否则容易出现编号、目录、图表引用不一致。

### 17.3 python-docx 生成

推荐使用：

```bash
python -m pip install python-docx
```

基本流程：

```python
from docx import Document
from docx.shared import Pt, Cm, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml import OxmlElement
from docx.oxml.ns import qn

doc = Document()
# 页面设置
# 样式设置
# 封面
# 摘要
# 目录域
# 正文
# 表格
# 图片
# 公式
# 页眉页脚
# 参考文献
doc.save("output.docx")
```

### 17.4 Word COM 更新目录与导出 PDF

```powershell
$word = New-Object -ComObject Word.Application
$word.Visible = $false
$doc = $word.Documents.Open($docxPath)
$doc.Fields.Update() | Out-Null
if ($doc.TablesOfContents.Count -gt 0) {
    $doc.TablesOfContents.Item(1).Update() | Out-Null
}
$doc.Save()
$doc.ExportAsFixedFormat($pdfPath, 17)
$doc.Close($false)
$word.Quit()
```

### 17.5 PDF 逐页检查

```bash
pdftoppm -png -f 1 -l 15 -r 100 output.pdf preview/page
```

逐页检查：

- 封面是否居中；
- 目录是否一页；
- 标题是否孤行；
- 表格是否跨页拆分；
- 图片是否变形；
- 图题是否与图片同页；
- 页脚是否过大；
- 公式是否可编辑；
- 参考文献是否溢出。

### 17.6 最小可复用主流程

把生成过程拆成固定函数，按顺序调用。不要把所有代码写成一个巨型脚本，否则后期很难改版式。

```python
def build_report():
    doc = Document()

    # 1. 页面与全局样式
    setup_page(doc)                 # A4、页边距、首页不同
    setup_normal_style(doc)         # 宋体/Times 12 pt、1.5 倍行距
    setup_heading_styles(doc)       # H1 16 pt、H2 13 pt

    # 2. 前置部分
    add_cover(doc)
    add_abstract(doc)               # 中文摘要 + 关键词 + 英文摘要
    add_toc(doc)                    # TOC 域，只显示一级标题

    # 3. 正文（以下标题只是示例，必须替换）
    add_h1(doc, "1 引言")
    add_body(doc, "...")

    add_h1(doc, "2 相关工作")
    add_body(doc, "...")

    add_h1(doc, "3 方法")
    add_h2(doc, "3.1 数据")
    add_body(doc, "...")
    add_data_table(
        doc,
        caption="表1  示例数据表",
        headers=["列1", "列2", "列3"],
        rows=[...],
        widths=[3.0, 3.8, 8.7],
        note="注：此处为示例表注，必须替换。",
    )

    add_h2(doc, "3.2 模型")
    add_body(doc, "...")
    add_equation(doc)               # OMML 公式 + 右侧编号
    add_figure(
        doc,
        path="figure.png",
        caption="图1  示例图",
        note="注：此处为示例图注，必须替换。",
        width_cm=14.0,
    )

    # 4. 实验、结果、讨论、结论
    add_h1(doc, "4 实验设置")
    add_h1(doc, "5 结果与分析")
    add_h1(doc, "6 讨论")
    add_h1(doc, "7 结论与展望")

    # 5. 参考文献
    add_references(doc, refs)

    # 6. 页眉页脚与字段
    add_header_footer(doc)
    enable_update_fields(doc)

    doc.save("report.docx")


def postprocess(docx_path, pdf_path):
    # Word COM：更新目录、更新域、保存、导出 PDF
    update_toc_and_export_pdf(docx_path, pdf_path)
    # 最后再把视图设为 print，避免 Word 保存时改回 web
    set_print_view(docx_path)
    # 渲染 PDF 逐页检查
    render_pdf_preview(pdf_path)
```

调用顺序的原则：

```text
页面 → 样式 → 封面/摘要/目录 → 正文 → 表格/图片/公式
→ 参考文献 → 页眉页脚 → 保存 → Word 更新目录 → PDF → 逐页检查
```

不要在最后一步大改正文结构，否则目录页码会全部变化。

### 17.7 表格函数骨架

```python
def add_data_table(doc, caption, headers, rows, widths, note):
    add_table_caption(doc, caption)

    table = doc.add_table(rows=1, cols=len(headers))
    table.alignment = WD_TABLE_ALIGNMENT.CENTER
    table.autofit = False

    for i, w in enumerate(widths):
        for row in table.rows:
            row.cells[i].width = Cm(w)

    # 表头
    for i, h in enumerate(headers):
        cell = table.rows[0].cells[i]
        cell.text = ""
        p = cell.paragraphs[0]
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        r = p.add_run(h)
        set_run_font(r, "黑体", "Arial", 9, True, DARK)

    # 表头重复
    trPr = table.rows[0]._tr.get_or_add_trPr()
    tblHeader = OxmlElement("w:tblHeader")
    tblHeader.set(qn("w:val"), "true")
    trPr.append(tblHeader)
    trPr.append(OxmlElement("w:cantSplit"))

    # 数据行
    for row_data in rows:
        cells = table.add_row().cells
        trPr = table.rows[-1]._tr.get_or_add_trPr()
        trPr.append(OxmlElement("w:cantSplit"))
        for i, value in enumerate(row_data):
            cell = cells[i]
            cell.text = ""
            p = cell.paragraphs[0]
            p.alignment = (
                WD_ALIGN_PARAGRAPH.CENTER if i == 0
                else WD_ALIGN_PARAGRAPH.LEFT
            )
            r = p.add_run(str(value))
            set_run_font(r, "宋体", "Times New Roman", 9, False, DARK)

    set_three_line_table(table)
    add_table_note(doc, note)
```

表格函数的参数必须包含：

```text
caption   表题
headers   表头
rows      行数据
widths    列宽
note      表注
```

这样换主题时只改数据，不用重写格式代码。

---

## 18. 默认页面视图

Word 默认打开时必须是页面视图，不是 Web 版式。

在 `settings.xml` 中设置：

```xml
<w:view w:val="print"/>
<w:zoom w:percent="100"/>
```

`python-docx` 写法：

```python
settings = doc.settings.element
for el in settings.findall(qn("w:view")):
    settings.remove(el)
view = OxmlElement("w:view")
view.set(qn("w:val"), "print")
settings.append(view)

for el in settings.findall(qn("w:zoom")):
    settings.remove(el)
zoom = OxmlElement("w:zoom")
zoom.set(qn("w:percent"), "100")
settings.append(zoom)
```

注意：

- Word COM 保存后可能把视图改回 Web；
- 必须在最后一次 Word 保存之后，再用 python-docx 设置 `print`；
- 设置后不要再让 Word 打开并保存该文件。

---

## 19. 常见坑与修复

### 19.1 目录变成原始 TOC 代码

现象：

```text
TOC \o "1-1" \h \z \u [标题] ...
```

原因：

- 目录域结构被破坏；
- 手动删除/移动了目录结果段落；
- 字段 `begin/separate/end` 不完整。

修复：

- 删除整个目录区域；
- 重新插入完整 TOC 域；
- 用 Word 更新整个目录；
- 更新后不要手动改目录段落。

### 19.2 目录后出现空白页

原因：

- 标题已有 `page_break_before = True`，前面又插入了分页符段落；
- 目录刚好占满一页，分页符又生成一页。

修复：

- 删除目录后的空分页段落；
- 只保留标题的 `page_break_before = True`；
- 不要同时使用“分页符段落 + 标题前分页”。

### 19.3 表格跨页只剩表头

原因：

- 表头与第一行没有绑定；
- 行允许跨页拆分。

修复：

- 表头行设置 `w:tblHeader`；
- 所有行设置 `w:cantSplit`；
- 表题设置 `keep_with_next`；
- 必要时在表题前加分页。

### 19.4 图片变形

原因：

- 只设置了宽度，没有按比例设置高度；
- 或在 Word 中手动拉伸。

修复：

```python
ratio = shape.height / shape.width
shape.width = Cm(14.0)
shape.height = int(Cm(14.0) * ratio)
```

### 19.5 图片太小/太大

- 普通 16:9 图：14.0 cm；
- 密集流程图：15.0—15.5 cm；
- 不要超过正文宽度 15.8 cm；
- 图注、图注与图片同页。

### 19.6 圆形项目符号

- 不要使用 `List Bullet`；
- 不要输入 `• ● ○ ▪`；
- 统一改成 `（1）（2）（3）`；
- 设置 `paragraph.style = Normal`，清除 `w:numPr`。

### 19.7 Word 文件被锁定

现象：

```text
PermissionError: [Errno 13] Permission denied
```

原因：

- Word 仍在后台打开文件；
- 预览器/同步盘锁定。

修复：

- 关闭 Word；
- 另存为新文件名；
- 不要强行覆盖被锁文件。

### 19.8 删除附录时误删正文

原因：

- 匹配到了目录中的“附录A”条目，而不是正文的 Heading 1。

修复：

```python
if para.style.name == "Heading 1" and para.text.strip().startswith("附录A"):
    ...
```

### 19.9 页脚数字太大

- 页脚字号 8 pt；
- 颜色灰色 `#595959`；
- 不要和正文一样 12 pt；
- 不要加粗。

### 19.10 图片中文标签错字

- AI 生图对中文仍可能出错；
- 生成后逐字检查；
- 错字用同色文本框覆盖；
- 不要直接拉伸图片来“修复”文字。

---

## 20. 质检清单

### 20.1 页面

- [ ] A4，页边距正确；
- [ ] 首页不显示页眉页脚；
- [ ] 默认页面视图 / 打印布局；
- [ ] 缩放 100%。

### 20.2 标题与目录

- [ ] 一级标题 16 pt 黑体，深蓝；
- [ ] 二级标题 13 pt 黑体，蓝色；
- [ ] 目录为一页；
- [ ] 目录条目与正文标题一致；
- [ ] 目录页码正确；
- [ ] 目录中没有多余条目。

### 20.3 正文

- [ ] 宋体 / Times New Roman 12 pt；
- [ ] 1.5 倍行距；
- [ ] 首行缩进 2 字符；
- [ ] 段后 6 pt；
- [ ] 没有手动空格缩进；
- [ ] 没有大段无分段文字。

### 20.4 表格

- [ ] 三线表，无竖线；
- [ ] 表题在上；
- [ ] 表注在下；
- [ ] 表头加粗；
- [ ] 表头跨页重复；
- [ ] 行不跨页拆分；
- [ ] 表格宽度不超版面；
- [ ] 单位、百分比、缺失值说明完整。

### 20.5 图片

- [ ] 图片不变形；
- [ ] 图题在下；
- [ ] 图注在下；
- [ ] 图题与图片同页；
- [ ] 图片清晰，无模糊；
- [ ] 中文标签无错字；
- [ ] 图片编号连续；
- [ ] 正文有“见图X”。

### 20.6 公式

- [ ] OMML 可编辑；
- [ ] 公式居中；
- [ ] 编号右对齐；
- [ ] 编号连续；
- [ ] 正文有“见式（X）”；
- [ ] 符号有解释。

### 20.7 列表

- [ ] 无圆形符号；
- [ ] 使用（1）（2）（3）；
- [ ] 编号连续；
- [ ] 缩进统一。

### 20.8 页眉页脚

- [ ] 页眉右对齐，8.5 pt；
- [ ] 页脚居中，8 pt；
- [ ] 页码为 PAGE / NUMPAGES 域；
- [ ] 页码不遮挡正文。

### 20.9 参考文献

- [ ] GB/T 7714；
- [ ] 编号连续；
- [ ] 正文引用对应；
- [ ] 不引用无关文献；
- [ ] 不虚构文献。

### 20.10 内容真实性

- [ ] 未完成内容不写成已完成；
- [ ] 未测试指标标注 `[待补充/待核实]`；
- [ ] 模拟/示意数据明确标注；
- [ ] 图表数据与正文一致。

---

## 21. 主题替换方法

主题可以变化，但格式系统不变。

替换主题时只改：

```text
文档标题
章节标题
正文内容
图表数据
图片
参考文献
```

不要改：

```text
页面尺寸
页边距
字体族
字号层级
标题颜色体系
三线表规则
图题/表题位置
页眉页脚
目录域
公式编号规则
```

如果必须换配色，建议整体替换：

```text
深色主色：原 #1F3A5F → 新主色
强调色：原 #0072B2 → 新强调色
辅助色：原 #E69F00 / #009E73 → 新辅助色
浅底：原 #EAF2FA → 新浅底
```

并确保：

- 新配色不超过 4 个主色；
- 黑白打印仍可区分；
- 图表、标题、提示框使用同一套颜色；
- 不要只改标题颜色，不改图表颜色。

---

## 22. 交付标准

最终交付应包含：

```text
1. Word 文档（.docx）
2. PDF 预览（.pdf）
3. 配图源文件或图片来源说明
4. 图表数据来源说明
5. 参考文献列表
```

Word 文档必须满足：

- 可直接打开；
- 默认以页面视图 / 打印布局打开，缩放 100%；
- 目录可更新；
- 公式可编辑；
- 表格可编辑；
- 图片可替换；
- 页眉页脚正常；
- 无空白页；
- 无乱码；
- 无圆形项目符号；
- 无图片变形；
- 无“错误！未定义书签”；
- 无虚构数据。

---

## 23. 参数速查表

| 类别 | 参数 |
|---|---|
| 纸张 | A4，21.0 × 29.7 cm |
| 页边距 | 上 2.4 / 下 2.2 / 左 2.6 / 右 2.4 cm |
| 页眉页脚距离 | 1.2 / 1.0 cm |
| 正文 | 宋体 + Times New Roman，12 pt，1.5 倍行距，首行缩进 24 pt |
| 一级标题 | 黑体 + Arial，16 pt，粗体，`#1F3A5F`，下边框 `#0072B2` |
| 二级标题 | 黑体 + Arial，13 pt，粗体，`#0072B2` |
| 表题/图题 | 黑体 + Arial，9.5 pt，粗体，`#1F3A5F`，居中 |
| 表格正文 | 宋体 + Times New Roman，9 pt；密集表 8.8 或 8 pt |
| 表注/图注 | 宋体 + Times New Roman，8.5 pt，`#595959` |
| 页眉 | 宋体 + Times New Roman，8.5 pt，灰色，右对齐 |
| 页脚 | 宋体 + Times New Roman，8 pt，灰色，居中 |
| 参考文献 | 宋体 + Times New Roman，9 pt，GB/T 7714 |
| 三线表 | 顶线/底线 1.5 pt，表头下线 1 pt，无竖线 |
| 图片 | 16:9，普通 14.0 cm，密集 15.0—15.5 cm，保持比例 |
| 公式 | OMML，居中，编号右对齐 `（1）` |
| 列表 | 不用圆点，统一 `（1）（2）（3）` |
| 目录 | `TOC \o "1-1" \h \z \u`，只显示一级标题 |
| 页面视图 | `<w:view w:val="print"/>`，缩放 100% |
| 配色 | 深蓝 `#1F3A5F`、蓝 `#0072B2`、橙 `#E69F00`、绿 `#009E73` |
| 浅底/边框 | `#EAF2FA` / `#B4C6E7` |

---

## 24. 一句话总结

这是一套“**蓝绿色学术配色 + 黑体标题 + 宋体正文 + 三线表 + 16:9 图 + OMML 公式 + 自动目录 + 小号页脚 + （1）（2）（3）编号**”的 Word 论文/报告格式系统。

生成时先搭格式骨架，再填内容；生成后用 Word 更新目录、导出 PDF、逐页检查。只要保持这套格式规则，主题和内容可以变化，成品风格仍然统一、清爽、简洁、可用于正式汇报。
