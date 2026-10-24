# 已知坑与修复（含规范 §19 全部条目）

> 标注「本 skill 已自动处理」的项，用 `docx_builder.py` 就不会踩到；手写 python-docx 时才需要自己处理。

## 目录

1. [目录变成原始 TOC 代码](#1-目录变成原始-toc-代码)
2. [目录后出现空白页](#2-目录后出现空白页)
3. [表格跨页只剩表头](#3-表格跨页只剩表头)
4. [图片变形 / 太小太大](#4-图片变形--太小太大)
5. [圆形项目符号](#5-圆形项目符号)
6. [Word 文件被锁定](#6-word-文件被锁定)
7. [删除附录时误删正文](#7-删除附录时误删正文)
8. [页脚数字太大](#8-页脚数字太大)
9. [图片中文标签错字](#9-图片中文标签错字)
10. [python-docx 专属坑](#10-python-docx-专属坑)
11. [跨平台与编码](#11-跨平台与编码)

---

## 1. 目录变成原始 TOC 代码

**现象**：目录区域显示 `TOC \o "1-1" \h \z \u …`

**原因**：域结构不完整——`begin / instrText / separate / end` 缺一截，或手工删改过域结果段落。

**修复**：删掉整个目录区域，重新插入完整 TOC 域；用 Word 更新整个目录；更新后不要再手工改目录段落。
（本 skill 已自动处理：每个域片段独立成 run，`w:fldChar begin` 带 `w:dirty="true"`。）

## 2. 目录后出现空白页

**原因**：标题已有 `page_break_before=True`，前面又插了分页符段落 → 多出一页。

**修复**：只保留一种分页方式。本书 `add_heading1(..., page_break=True)` 与首个一级标题自动分页，
**不要再额外 `add_pagebreak()`**。

## 3. 表格跨页只剩表头

**原因**：表头与第一行没绑定；行允许跨页拆分；表题没 `keep_with_next`。

**修复**（本 skill 已自动处理）：表头行 `w:tblHeader`，所有行 `w:cantSplit`，表题 `keep_with_next`。

## 4. 图片变形 / 太小太大

**原因**：只改宽或只改高；或手拉过。

**修复**：`run.add_picture(path, width=Cm(14.0))` 只给宽度即等比。
常用宽度：普通 16:9 图 `14.0 cm`；密集流程图 `15.0–15.5 cm`；**上限 15.5 cm**（版心 15.8 cm）。
1672×941 px 的图在 14.0 cm 时约 303 dpi，够用。

## 5. 圆形项目符号

**原因**：用了 `List Bullet` 样式，或直接敲了 `• ● ○ ▪`。

**修复**：改成 `（1）（2）（3）`；`paragraph.style = Normal` 并去掉 `w:numPr`。
（本 skill 的 `add_list()` 直接生成带全角括号的编号段落，不用列表样式。）

## 6. Word 文件被锁定

**现象**：`PermissionError: [Errno 13] Permission denied`

**原因**：Word 后台还开着该文件；预览器/同步盘占用。

**修复**：关闭 Word（含后台残留 `WINWORD.EXE`），或另存新文件名；不要强写被锁文件。

## 7. 删除附录时误删正文

**原因**：匹配到目录里的"附录A"条目，而不是正文的 Heading 1。

**修复**：

```python
if para.style.name == "Heading 1" and para.text.strip().startswith("附录A"):
    ...
```

## 8. 页脚数字太大

**修复**：页脚 8 pt、灰色 `#595959`、不加粗、居中。（本 skill 已处理：页脚 run 显式 8 pt。）

## 9. 图片中文标签错字

AI 生图对中文仍会出错：生成后**逐字**检查；错字用同色文本框覆盖；不要靠拉伸图片掩盖。
提示词里把每个中文标签逐项列出（不要只写"加上中文标签"）。

## 10. python-docx 专属坑

| 坑 | 说明 | 处理 |
|---|---|---|
| 主题字体覆盖 | `w:rFonts` 里若有 `w:asciiTheme`/`w:eastAsiaTheme`，显式 `w:ascii` 不生效 | 写入前删除 theme 属性（`_force_rfonts`） |
| `w:hanging` 单位错 | 某些版本把 EMU 直接写进 `w:hanging`，参考文献被推到页外 | 自己写 `w:hanging`，或用 `hanging_cm=` 参数 |
| `-Cm(0.74)` 退化 | `Length` 是 int 子类，取负后变普通 int，类型信息丢失 | 用 `Cm(-0.74)` 或 `hanging_cm=`；`_to_length()` 会按量级兜底 |
| 首行缩进与悬挂缩进打架 | 样式里有 `w:firstLineChars`，段落的 `w:hanging` 会被压掉 | 样式只设 `w:firstLine`；字符缩进逐段设 |
| `w:numPr` 残留 | 段落用了列表样式，改样式后编号还在 | 显式清除 `w:numPr`，或用 `Normal` 样式 |
| 元素顺序非法 | `w:tblBorders`/`w:cantSplit`/`w:pBdr` 插错位置，Word 可能报"内容有问题" | 按 OOXML 顺序插入（`insert_ordered`） |
| `table.autofit=False` 后列宽仍不准 | 只设了 `tcW`，没设 `w:tblGrid` | 同时改 `w:gridCol`（`set_col_widths`） |
| 单元格文字继承 12 pt | 单元格段落走 `Normal` 样式 | 每个单元格显式设字号与缩进 0 |

## 11. 跨平台与编码

| 问题 | 修复 |
|---|---|
| `.ps1` 在 Windows PowerShell 5.1 报语法错误 / 中文乱码 | 存成 **UTF-8 with BOM** |
| 控制台中文输出乱码 | 设置 `PYTHONIOENCODING=utf-8`（不影响文件内容） |
| macOS/Linux 没有 Word COM | 用 `postprocess.sh`（LibreOffice）；TOC 依赖打开时自动更新 |
| Word 保存后视图变回 Web 版式 | 最后一次保存后再跑 `set_view.py`，之后别再让 Word 打开保存 |
| LibreOffice 打开 docx 后另存 | 可能丢 OMML 的部分排版；交付以 Word/WPS 为准，PDF 用 Word 导出最稳 |
| WPS | 基本兼容；TOC 域同样支持。若目录不展开，Ctrl+A → F9 |
