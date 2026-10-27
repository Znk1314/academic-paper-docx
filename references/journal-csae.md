# 期刊模板：农业工程学报（Transactions of the CSAE）

> 参数不是猜的：逐项量自 4 篇《农业工程学报》论文（2025 年第 41 卷）
> —— 土壤分离能力（机器学习）、棉花叶面积指数（机器学习）、YOLO-EDW 害虫检测、
> YOLO-HSI 西梅无损检测。用 `pymupdf` 读 PDF 的字体/字号/坐标得到。

## 1. 实测参数

| 项目 | 实测值 | 说明 |
|---|---|---|
| 页面 | A4 595.3×841.9 pt（21×29.7 cm） | 与国标一致 |
| 版式 | **双栏**，栏宽 8.66 cm，栏间距 0.73 cm | 正文双栏，题名/摘要/英文块通栏 |
| 页边距 | 左 1.46 / 右 1.28 / 上 2.70 / 下 2.00 cm | 页眉 1.90 cm |
| **中文标题** | **黑体 18 pt**，居中 | |
| 作者 | **仿宋 14 pt**，居中，角标 1,2 上标 | |
| 单位 | 宋体 9 pt，居中，全角括号 | |
| 摘要 / 关键词 | **宋体 9 pt**；标签"摘　要："黑体 9 pt | 顶格，不分段缩进 |
| doi / 中图分类号 / 文献标志码 / 文章编号 | 9 pt | 黑体标签 + Times 数字 |
| 引用格式块 | 9 pt | 中文引用式样 |
| **英文标题** | **Times New Roman Bold 16 pt** | 居中 |
| 英文作者 | Times 10.5 pt | 姓全大写 |
| 英文单位 | Times 9 pt | 括号内分号分隔 |
| **英文摘要 / Keywords** | **Times 10 pt** | Keywords 标签加粗 |
| **正文** | **宋体 10 pt / Times 10 pt** | 行距约 1.15，首行缩进 2 字符 |
| **一级标题** | **黑体 11 pt**（编号 Times Bold 11） | 形如"0　引　言"，**无装饰线** |
| **二级标题** | **黑体 10 pt**（编号 Times Bold 10） | 形如"1.1　研究区域" |
| **图题** | **楷体 9 pt** "图1　…"；英文 **Times 9 pt** "Fig.1　…" | **英文另起一行** |
| **表题** | **黑体 9 pt** "表 1　…"；英文 **Times 9 pt** "Table 1　…" | **英文另起一行** |
| 表内文字 | 宋体 8 pt | |
| 表注 / 图注 | 8 pt | 表注在表下，末尾加"。" |
| **三线表** | 顶/底线 **0.75 pt**，栏目线 **0.5 pt** | 比学位论文（1.5/0.75）更细 |
| 参考文献 | 8 pt | GB/T 7714—2015 顺序编码制 |
| **页眉** | 奇偶页不同，9 pt | 见下 |
| 页脚 | 无（页码在页眉） | |
| 正文引文 | **上标方括号** [1]、[2,3]、[4-6] | 与国标一致 |

### 页眉（实测）

奇数页（首页）两行三栏：

```
第 41 卷第 15 期                                          农业工程学报                                          Vol.41 No.15
2025 年 8 月          Transactions of the Chinese Society of Agricultural Engineering          Aug. 2025
```

偶数页一行三栏：

```
第 15 期                              张三等：面向实验室环境监测的低功耗边缘融合方法                              99
```

## 2. 用法

```bash
python scripts/apd.py build work/paper.json -o work/paper.docx --profile csae
python scripts/apd.py post  work/paper.docx --pdf work/paper.pdf
python scripts/apd.py qa    work/paper.docx --pdf work/paper.pdf --profile csae
```

内容模型里多出的字段（`examples/journal_csae/content.json` 是完整样例）：

```jsonc
{
  "profile": "csae",
  "title_en": "…", "authors": ["张三","李四"], "author_sup": {"张三":"1"},
  "affiliations": ["1. …，北京 100083"],
  "authors_en": ["ZHANG San"], "affiliations_en": ["1. …"],
  "abstract_en": {"paragraphs": ["…"], "keywords": ["…"]},
  "journal": {"volume":"41","issue":"15","year":"2025","month":"8",
              "month_en":"Aug.","doi":"…","clc":"S126","doc_code":"A",
              "article_no":"…","cite_zh":"…"},
  "header_short": "张三等：题名",
  "sections": [ {"level":1,"text":"0　引　言"}, … ],
  "references": ["…"]
}
```

## 3. 与学位论文模板的差别

| 项目 | 学位论文（thesis） | 期刊稿（csae） |
|---|---|---|
| 版式 | 单栏 15.8 cm | **双栏 8.66 cm** |
| 字体 | 标题黑体 16/13 pt、正文宋体 12 pt | 标题黑体 18/11/10 pt、正文宋体 10 pt |
| 一级标题 | 深蓝 + 蓝色下边框 | 纯黑，无边框 |
| 图表题 | 中文 + 英文同行，9.5 pt | 中文一行、英文一行，各 9 pt |
| 三线表 | 1.5 / 0.75 pt | 0.75 / 0.5 pt |
| 前置部分 | 封面 → 摘要 → Abstract → 目录 | 题名/作者/单位/摘要/关键词/doi/分类号/引用格式/英文块 |
| 页眉 | 论文题目（单行） | 期刊名 + 卷期 + 英文刊名 + 年月（两行三栏） |
| 页码 | 页脚"第 X 页 / 共 Y 页" | 页眉（偶数页右侧） |

## 4. 已知取舍

- **通栏图**：期刊允许整栏宽（18.25 cm）图，需要单独插入"单栏分节"。当前实现统一
  按单栏 8.6 cm 处理；确需通栏时可在该图前后手工插分节，或用 `width_cm` 覆盖后调版。
- **公式编号**：沿用"右对齐 (1)"，与期刊一致；期刊中公式多用 Times 斜体变量，
  本实现走 OMML 公式域，Word 内可编辑。
- **收稿日期/基金项目/作者简介**：期刊首页脚注区。当前未生成，如需要可加
  `journal.fund` / `journal.author_bio`（字段已在 `JournalBuilder` 预留读取位）。

## 5. 换其他期刊

`CSAE_SPEC` 与 `JournalBuilder` 是参数化的：改 `SPEC` 里的字号/页边距/栏数、
改 `add_journal_header_footer()` 的页眉文字，即得到另一家期刊的模板。
`apd qa --profile csae` 的检查项也可按刊物增删。
