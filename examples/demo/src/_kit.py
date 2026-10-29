"""配方公共件：字体、配色、坐标标签、保存（自带 PNG 元数据）。

每个配方脚本都只依赖本文件，复制到工作区后即可独立运行，不依赖 skill 内部模块。
保持"可交付、可复现、可继续改"：字体与配色常量集中在这里，改一处全图生效。
"""

from __future__ import annotations

import os
import sys

import matplotlib

matplotlib.use("Agg")
import matplotlib.pyplot as plt  # noqa: E402
from matplotlib import font_manager  # noqa: E402

MM_PER_INCH = 25.4
MIN_FONT_PT = 6.0

# 定性色板（Okabe-Ito，色盲安全；与 skill 的 palette.py 同源）
SERIES = ["#000000", "#E69F00", "#56B4E9", "#009E73",
          "#F0E442", "#0072B2", "#CC79A7", "#999999"]
LINESTYLES = ["-", "--", ":", "-."]
MARKERS = ["o", "s", "^", "D", "v", "P", "X", "*"]
GREY = ["#000000", "#404040", "#808080", "#C0C0C0", "#F2F2F2"]
AXIS = "#262626"
GRID = "#D9D9D9"

# 字号/线宽（按"插入后不缩放"设计；厘米宽 × 25.4 = 毫米）
STYLE = {
    "cn": dict(font=8.5, label=9.0, legend=8.0, lw=1.1, ms=4.0,
               tick=0.7, spine=0.7, grid_alpha=0.25),
    "nature": dict(font=7.0, label=7.0, legend=6.0, lw=0.75, ms=2.8,
                   tick=0.5, spine=0.5, grid_alpha=0.0),
    "ieee": dict(font=8.0, label=8.0, legend=7.0, lw=1.2, ms=4.0,
                 tick=0.8, spine=0.8, grid_alpha=0.2),
}
CJK_FONTS = ["Microsoft YaHei", "SimHei", "Noto Sans CJK SC",
             "Source Han Sans SC", "PingFang SC", "WenQuanYi Zen Hei"]


def setup(width_mm=140.0, height_mm=None, aspect=0.62, style="cn"):
    """建好画布与坐标轴：字体、线宽、单侧轴线、rcParams 一次到位。"""
    st = STYLE.get(style, STYLE["cn"])
    available = {f.name for f in font_manager.fontManager.ttflist}
    cjk = os.environ.get("APD_CJK_FONT", "").strip() or next(
        (n for n in CJK_FONTS if n in available), None)
    family = ([cjk] if cjk else []) + ["Arial", "Helvetica", "DejaVu Sans"]
    plt.rcParams.update({
        "font.family": "sans-serif",
        "font.sans-serif": family,
        "axes.unicode_minus": False,
        "pdf.fonttype": 42,
        "ps.fonttype": 42,
        "svg.fonttype": "none",
        "font.size": st["font"],
        "axes.labelsize": st["label"],
        "legend.fontsize": st["legend"],
        "xtick.labelsize": st["font"],
        "ytick.labelsize": st["font"],
        "axes.linewidth": st["spine"],
        "xtick.major.width": st["tick"],
        "ytick.major.width": st["tick"],
        "xtick.major.size": 3.0,
        "ytick.major.size": 3.0,
    })
    h = height_mm if height_mm is not None else width_mm * aspect
    fig, ax = plt.subplots(figsize=(width_mm / MM_PER_INCH, h / MM_PER_INCH))
    fig.set_layout_engine("constrained")
    for side in ("top", "right"):
        ax.spines[side].set_visible(False)
    for side in ("left", "bottom"):
        ax.spines[side].set_linewidth(st["spine"])
        ax.spines[side].set_color(AXIS)
    ax.tick_params(direction="out", colors=AXIS)
    return fig, ax, st


def label(name, unit="", symbol=""):
    """轴的"量符号/单位"：符号斜体、单位正体（GB/T 7713.2 5.4.4）。"""
    if symbol:
        body = "$%s$" % symbol
    else:
        body = name
    return "%s/%s" % (body, unit) if unit else body


def style_of(i):
    """第 i 个系列的颜色 + 线型 + 标记（冗余编码，黑白打印也能区分）。"""
    return dict(color=SERIES[i % len(SERIES)],
                linestyle=LINESTYLES[i % len(LINESTYLES)],
                marker=MARKERS[i % len(MARKERS)])


def grid(ax, st, axis="y"):
    if st["grid_alpha"] <= 0:
        return
    ax.grid(True, axis=axis, linestyle=":", linewidth=0.5,
            color=GRID, alpha=st["grid_alpha"])
    ax.set_axisbelow(True)


def bracket(ax, x1, x2, label_text, y=None, st=None):
    """显著性括号：横线 + 上标记号。"""
    lo, hi = ax.get_ylim()
    span = hi - lo
    y = hi - span * 0.08 if y is None else y
    h = span * 0.02
    ax.plot([x1, x1, x2, x2], [y, y + h, y + h, y],
            color=AXIS, linewidth=0.7, clip_on=False)
    ax.text((x1 + x2) / 2, y + h, label_text, ha="center", va="bottom",
            fontsize=(st or STYLE["cn"])["font"])
    ax.set_ylim(lo, hi + span * 0.06)


def save(fig, out, width_mm=140.0, base_font_pt=8.5, style="cn",
         dpi=1000, formats=("png", "pdf")):
    """保存并写元数据：docx 端据此核算"插入后图内文字还剩几 pt"。"""
    out = os.path.abspath(out)
    os.makedirs(os.path.dirname(out) or ".", exist_ok=True)
    stem, ext = os.path.splitext(out)
    ext = (ext or ".png").lower().lstrip(".")
    software = ("academic-paper-docx width_mm=%.1f font_pt=%.1f style=%s"
                % (width_mm, base_font_pt, style))
    saved = []
    for fmt in dict.fromkeys([ext] + [f.lower() for f in formats]):
        target = out if fmt == ext else "%s.%s" % (stem, fmt)
        # PNG 支持任意 tEXt（docx 端读 Software 核算有效字号）；PDF 只认标准字段
        meta = ({"Software": software} if fmt == "png"
                else {"Creator": software, "Producer": software}
                if fmt == "pdf" else None)
        fig.savefig(target, dpi=dpi, facecolor=fig.get_facecolor(),
                    metadata=meta)
        saved.append(target)
    plt.close(fig)
    return saved[0]


def read_csv(path):
    """读 CSV：第一行表头，其余为数值/字符串；跳过空行与 # 注释。"""
    import csv

    with open(path, encoding="utf-8-sig", newline="") as fh:
        rows = [r for r in csv.reader(fh)
                if r and not (r[0] or "").startswith("#")]
    if not rows:
        raise ValueError("数据文件为空：%s" % path)
    header = [h.strip() for h in rows[0]]
    body = []
    for r in rows[1:]:
        cells = []
        for cell in r:
            cell = (cell or "").strip()
            try:
                cells.append(float(cell))
            except ValueError:
                cells.append(cell)
        body.append(cells)
    return header, body


def column(header, body, name):
    if name not in header:
        raise KeyError("数据缺少列 %r；现有列：%s" % (name, "、".join(header)))
    i = header.index(name)
    return [row[i] for row in body]


def require(cond, message):
    if not cond:
        raise SystemExit("数据契约不满足：%s" % message)


def main_guard():
    if len(sys.argv) < 3:
        raise SystemExit("用法：python %s <data.csv> <out.png> [width_mm]"
                         % os.path.basename(sys.argv[0]))
