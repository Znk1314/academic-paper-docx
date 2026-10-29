# Nature / Science 体例：期刊成品图的硬规则

> **规则来源（均为 MIT 许可，可借鉴并以本文档形式再表述）**
> - [Haojae/scipilot-figure-skill](https://github.com/Haojae/scipilot-figure-skill)（MIT，2.7k★）：
>   最终尺寸出图、矢量优先、色盲友好、最小 6 pt、误差必交代、主动拦截 18 条画图禁忌
> - [TAO-QKV/Icarus-Figures](https://github.com/TAO-QKV/Icarus-Figures)（MIT）：
>   构图范式与期刊级工艺
> - [jeffreyjing/AgentFigureGallery](https://github.com/Dsadd4/AgentFigureGallery)（MIT）：nature-style 参考图库
>
> 本仓库的配方、`_kit.py` 与检查代码是自己写的；上面这些是**规则**层面的采用。

## 1. 成品尺寸（按刊取宽度，插入文档后不再缩放）

| 刊 | 单栏 | 双栏 | 字号 | 备注 |
|---|---|---|---|---|
| Nature | **89 mm** | **183 mm** | 7 pt（最小 6 pt） | 面板标签用加粗小写 a/b/c |
| Science | 55 mm | 183 mm | 6.5—7 pt | 字号更小、线更细 |
| IEEE | 88.9 mm | 181 mm | 8 pt | 面板标签 (a)(b)(c) |
| Elsevier | 90 mm | 190 mm | 8 pt | 单栏与双栏间距大 |
| PNAS | 87 mm | 178 mm | 7—8 pt | |
| 农业工程学报 | 86 mm | 182 mm | 9 pt | 中文标题楷体/黑体见期刊模板 |
| 学位论文 | 75 mm | 155 mm | 9.5 pt | 版心 16 cm |

`recipes/_kit.py` 里是 `JOURNAL_WIDTHS`，取宽度用 `K.width_for("nature", "double")`。

## 2. 五条硬原则

1. **按最终尺寸出图**：`figsize` 直接设成品毫米数。插入 Word 后缩放会让 7 pt 变 4 pt，
   投稿自检直接打回。
2. **矢量优先**：数据图出 PDF/SVG；位图才用 PNG（300—600 dpi 起，本仓库默认 1000）。
   **绝不用 JPEG** 存数据图。
3. **色盲友好**：Okabe-Ito 色板 + **冗余编码**（线型 / 标记），并出灰度版自检。
4. **字号可读**：正文与刻度 7—9 pt，**最小 6 pt**。脚本会把设计字号写进 PNG 元数据，
   docx 端据此核算"插入后还剩几 pt"。
5. **误差必有交代**：误差棒/阴影/箱线必须在图注写清 **误差类型（SD/SEM/95% CI/IQR）、
   n、检验方法与校正、显著性符号含义**。SD 与 SEM 差一个 √n，混淆会反转结论。

## 3. 主动拦截：这些图型要先劝阻

| 触发条件 | 后果 | 替代 |
|---|---|---|
| 每组 n<10 还想画均值柱 | 柱高+一根误差棒掩盖分布与 n | **小组样本点图**（`nature-dotplot`）：点全画出来 + 中位线 |
| 双 Y 轴展示无关变量 | 视觉相关是作图者造出来的 | 拆上下子图共享 x，或标准化共轴 |
| 饼图显示占比 | 人眼判角度差长度约 3 倍 | 横向柱状（按值排序） |
| 3D 柱 / 3D 饼 | 视角扭曲所有数值 | 2D 柱或热图 |
| 比例图 Y 轴不从 0 起 | 小差异被放大 | 从 0 起，或 log 轴，或明确断裂标记 |
| 连续色映射无 colorbar | 读者不知深浅对应何值 | 必加 colorbar + 标目单位 |
| 分类 x 用折线连均值 | 暗示不存在的连续关系 | 散点 / 点图 / 柱状 |
| 一图塞 5 个论点 | 没有论点 | 拆图：一图一结论 |
| jet / rainbow | 感知不均匀，凭空造峰 | cividis / viridis / RdBu_r |

`apd recipe check` 会对适配后的脚本做静态扫描，命中 `plt.pie`、`twinx`、`jet`、
"柱状 + 误差棒但无散点" 会**明确提示**（不阻断，但留记录）。

## 4. 面板标签与组合图

- 面板标签用**轴分数坐标 + 统一偏移**（`K.panel_label`），不要手摆 `ax.text(0.02, 0.95)`：
  constrained layout 下每块位置不同，手摆必错位。
- Nature = 加粗小写 `a`；IEEE = `(a)`。
- 组合图统一：同一变量在所有面板**同色同标记**、字号一致、坐标尺度可比时才共轴。
- 只左列标 y、只底行标 x；图例用 `loc="outside lower center"` 占独立空间，避免压数据。

## 5. 出图后的三层自检

| 层 | 查什么 | 本仓库怎么落地 |
|---|---|---|
| 语义 | 图型是否适合数据、误差口径是否正确 | 本文件 §3 + 图注必须写误差类型 |
| 形式 | 尺寸/字号/分辨率/矢量 | `apd qa`：图宽 ≤ 版心、图内字 ≥6 pt、分辨率 ≥600 dpi |
| 视觉 | 图例压数据、标签裁切、面板错位、灰度可分 | `apd recipe check` + 打开 `preview.png` 人工看 |

## 6. 用法

```bash
apd recipe info nature-multipanel        # 183 mm 双栏四面板
apd recipe new nature-multipanel --out work/figs
apd recipe run work/figs/gen_fig_nature_multipanel.py data.csv fig1.png 183

apd recipe info nature-dotplot           # 89 mm 单栏，n=3—12 的点图
```
