# 配方出图：选配方 → 拷源码 → 适配 → 核对

本 skill **不再有通用渲染器**（原 `figures.py` / `make_chart.py` / JSON 规格已删除）。
出图改为"**配方 + 源码适配**"：每个图型是一份**完整、可独立运行**的 matplotlib 脚本，
你把整份源码拷到工作区，在真实文件上改数据、标签与布局，改完再核对。

> 思路参考了 [vivid-illustration-skill](https://github.com/yjz211/vivid-illustration-skill)
> 的"模板 + 适配 + 保真检查"工作流。**但该仓库采用 Personal Non-Commercial License，
> 明确禁止改编与再分发**，所以本仓库的配方、配色与代码全部是**自行编写**的，
> 没有复制或改写其任何模板与源码。

## 为什么这样做

| 通用渲染器 | 配方 + 源码适配 |
|---|---|
| 图型由 JSON 字段决定，复杂设计写不进去 | 每张图有独立脚本，渐变、注释、面板、拼图都能写 |
| 改一次样式要改渲染器 | 改的是**你自己项目里的脚本**，改完就留着 |
| 交付物只有 PNG | 交付 **PNG/PDF + 可复现源码**，换数据重跑即可 |

## 四步

```bash
# 1) 选：看配方说明、数据契约、保真要点
python scripts/apd.py recipe list
python scripts/apd.py recipe info line-multiseries

# 2) 拷：完整源码 + 公共件拷到工作区
python scripts/apd.py recipe new line-multiseries --out work/figs
#    → work/figs/gen_fig_line_multiseries.py + work/figs/_kit.py

# 3) 适配并在真实脚本上跑
python scripts/apd.py recipe run work/figs/gen_fig_line_multiseries.py \
       work/data/power.csv work/figs/fig2.png 140

# 4) 核对：数据契约 + 相对原配方改了哪些行
python scripts/apd.py recipe check work/figs/gen_fig_line_multiseries.py --data work/data/power.csv
```

## 配方目录（当前 6 个）

| 配方 | 图型 | 数据契约（CSV） |
|---|---|---|
| `line-multiseries` | 多系列折线 + SD 误差线 + 显著性 | `x,<系列...>[,err_<系列>]` |
| `bar-single` | 单系列分类柱状（零基线） | `category,value[,err]` |
| `bar-grouped` | 分组柱状 + 误差线 | `category,<系列...>[,err_<系列>]` |
| `scatter-fit` | 散点 + 1:1 参考线 | `x,<系列...>` |
| `boxgroup` | 分组箱线（含离群点） | `group,value`（长表） |
| `heatmap-seq` | 顺序色板热图（cividis） | `row,<列...>` |
| `nature-multipanel` | Nature 双栏四面板 + a/b/c/d | `panel,group,x,y` |
| `nature-dotplot` | 小组样本点图（替代均值柱） | `group,value`（每组 3—12 点） |

每个配方目录下有 `recipe.py`（完整源码）、`meta.json`（适用/契约/保真/自查）、`preview.png`（预览）。

## 适配时必须守住的（保真清单）

1. **颜色 + 线型/标记双重编码**——不能只靠颜色区分组别（色盲与黑白打印）。
2. **轴标目是"量符号/单位"**，符号斜体、单位正体（GB/T 7713.2 5.4.4）。
3. **字号与线宽按插入宽度设计**：`width_mm` 取 75（半栏）/155（通栏）；
   插入后图内文字必须 ≥6 pt，docx 端会用 PNG 元数据再核算一次。
4. **不要为了适配数据整段重写绘图主体**：优先改位置、间距、尺寸；
   新增组别可扩画布/面板。
5. **不加装饰**：无 3D、无渐变背景、无彩虹色谱（`heatmap-seq` 会在代码里直接拒绝）。
6. **数据不补造**：某层没有证据就按统计含义移除，并在图注说明。

## 期刊尺寸与 Nature 体例

按刊取成品宽度用 `K.width_for(journal, span)`：

| 刊 | 单栏 | 双栏 |
|---|---|---|
| nature | 89 mm | 183 mm |
| science | 55 mm | 183 mm |
| ieee | 88.9 mm | 181 mm |
| elsevier | 90 mm | 190 mm |
| pnas | 87 mm | 178 mm |
| csae | 86 mm | 182 mm |
| thesis | 75 mm | 155 mm |

体例规则（按最终尺寸出图、矢量优先、色盲冗余编码、最小 6 pt、误差必交代、
主动拦截 9 类问题图型）见 `references/nature-style.md`（规则借鉴 MIT 许可的
scipilot-figure-skill 与 Icarus-Figures，代码自写）。

## 与文档的衔接

- 图注必须说明误差线含义（SD/SEM/95% CI）与显著性水平——这是 `qa` 会检查的。
- `docx` 的 `add_figure()` 会读 PNG 里的 `Software: academic-paper-docx width_mm=… font_pt=…`
  元数据，核算"插入后图内文字还剩几 pt"，低于 6 pt 会警告。
- 尺寸、分辨率、三线表、配色等规范见 `references/figure-table-standard.md` 与
  `references/palette.md`；出图配色沿用同一套 Okabe-Ito（配方 `_kit.py` 内常量）。
