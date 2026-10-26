"""论文插图渲染（matplotlib 后端）：按 GB/T 7713.2 与期刊通行惯例出图。

核心约束（全部可测）：

* **版面对齐**：`width_mm` 取 75（半栏）/160（通栏），按插入后不缩放设计字号与线宽。
* **坐标轴规范**：只留左/下轴线；标目为"量符号/单位"，符号斜体、单位正体；
  轴线结束于刻度范围；柱图 y 轴默认从 0 起。
* **冗余编码**：颜色之外自动配线型与标记（色觉障碍与黑白打印双保险）。
* **色板纪律**：颜色只取 `palette.SERIES`；定量数据禁用 jet/rainbow；顺序/发散自动选板。
* **分辨率**：线条图默认 1000 dpi PNG + PDF/SVG 矢量；照片类另走下采样。
* **自检**：`grayscale=True` 额外输出灰度版，用于"黑白打印还能不能读"的肉眼与数值校验。

用法：

    python scripts/figures.py spec.json -o figs/fig1.png
    python scripts/figures.py spec.json -o figs/fig1.png --grayscale
"""

from __future__ import annotations

import argparse
import json
import os
import sys

import matplotlib

matplotlib.use("Agg")  # 必须在 pyplot 之前；保证无显示环境可用

import matplotlib.pyplot as plt  # noqa: E402
from matplotlib import font_manager  # noqa: E402
from matplotlib.lines import Line2D  # noqa: E402

HERE = os.path.dirname(os.path.abspath(__file__))
if HERE not in sys.path:
    sys.path.insert(0, HERE)

import palette  # noqa: E402

MM_PER_INCH = 25.4
ASPECT = {"line": 0.62, "bar": 0.66, "grouped_bar": 0.70,
          "scatter": 0.62, "errorbar": 0.62, "box": 0.55, "heatmap": 0.78}


def _setup_fonts() -> str:
    """挑选可用的中文字体；找不到就退回 DejaVu（数字仍正常，中文会缺字）。"""
    available = {f.name for f in font_manager.fontManager.ttflist}
    chosen = next((n for n in palette.FONTS["cjk"] if n in available), None)
    family = [chosen] if chosen else []
    family += [n for n in palette.FONTS["latin"] if n in available] or ["DejaVu Sans"]
    plt.rcParams.update({
        "font.family": "sans-serif",
        "font.sans-serif": family,
        "axes.unicode_minus": False,   # 负号用 ASCII，避免方框
        "pdf.fonttype": 42,            # 矢量图内嵌 TrueType，可在 Illustrator 里编辑
        "ps.fonttype": 42,
        "svg.fonttype": "none",
    })
    return chosen or "DejaVu Sans"


def parse_axis(value, fallback: str = "") -> dict:
    """把 '平均功耗 / mW' 或 {label,symbol,unit} 解析成标目结构。"""
    if isinstance(value, dict):
        return {
            "label": str(value.get("label", "")).strip(),
            "symbol": str(value.get("symbol", "")).strip(),
            "unit": str(value.get("unit", "")).strip(),
        }
    text = str(value or fallback).strip()
    unit = ""
    if "/" in text:
        label, _, maybe_unit = text.rpartition("/")
        if maybe_unit and len(maybe_unit) <= 12:
            text, unit = label.strip(), maybe_unit.strip()
    return {"label": text, "symbol": "", "unit": unit}


def axis_text(axis: dict) -> str:
    """量符号/单位：符号斜体、单位正体；没有符号时用名称（中文保持正体）。"""
    label, symbol, unit = axis["label"], axis["symbol"], axis["unit"]
    if symbol:
        body = r"$%s$" % symbol if not symbol.startswith("$") else symbol
        if label and label != symbol:
            body = (r"$\it{%s}$" % symbol)
    else:
        body = label
    return "%s/%s" % (body, unit) if unit else body


def _dpi_for(spec: dict) -> int:
    kind = str(spec.get("kind", "line"))
    default = 300 if kind == "photo" else 1000
    return int(spec.get("dpi") or default)


def _width_mm(spec: dict) -> float:
    raw = spec.get("width_mm")
    if isinstance(raw, str) and raw in palette.WIDTH_MM:
        return palette.WIDTH_MM[raw]
    if isinstance(raw, (int, float)):
        return float(raw)
    if spec.get("layout") in palette.WIDTH_MM:
        return palette.WIDTH_MM[spec["layout"]]
    return palette.WIDTH_MM["half"]


def _series(spec: dict) -> list[dict]:
    out = []
    for i, s in enumerate(spec.get("series") or []):
        style = palette.serie_style(i)
        out.append({
            "name": str(s.get("name", "系列%d" % (i + 1))),
            "values": list(s.get("values") or []),
            "error": list(s.get("error") or []),
            "color": palette.normalize(s.get("color") or style["color"]),
            "linestyle": s.get("linestyle") or style["linestyle"],
            "marker": s.get("marker") or style["marker"],
        })
    return out


def _check_colors(series: list[dict], spec: dict) -> None:
    """硬错抛异常；灰度/蓝黄提示只告警（靠线型与标记兜底）。"""
    colors = [s["color"] for s in series]
    if len(colors) < 2:
        return
    errors, warnings = palette.check_series(colors)
    for msg in warnings:
        print("提示：%s（已用线型/标记做冗余编码）" % msg, file=sys.stderr)
    if errors:
        if spec.get("strict_palette", True):
            raise ValueError("配色不合规：\n  - " + "\n  - ".join(dict.fromkeys(errors)))
        for msg in errors:
            print("警告：%s" % msg, file=sys.stderr)


def render(spec: dict, out_path: str, grayscale: bool = False) -> str:
    """渲染一张图，返回主输出路径。"""
    kind = str(spec.get("kind") or spec.get("type") or "line")
    if kind not in ASPECT:
        raise ValueError("未知图型：%r（支持 %s）" % (kind, "/".join(ASPECT)))
    style_name = str(spec.get("style") or "cn")
    style = dict(palette.STYLE.get(style_name, palette.STYLE["cn"]))
    series = _series(spec)
    if not series:
        raise ValueError("series 为空：至少需要一个数据系列")

    _check_colors(series, spec)

    width_mm = _width_mm(spec)
    height_mm = float(spec.get("height_mm") or width_mm * ASPECT[kind])
    plt.rcParams.update({
        "font.size": style["font_size"],
        "axes.labelsize": style["label_size"],
        "legend.fontsize": style["legend_size"],
        "xtick.labelsize": style["font_size"],
        "ytick.labelsize": style["font_size"],
        "axes.linewidth": style["spine_width"],
        "xtick.major.width": style["tick_width"],
        "ytick.major.width": style["tick_width"],
        "xtick.major.size": 3.0,
        "ytick.major.size": 3.0,
    })
    font_used = _setup_fonts()
    if font_used == "DejaVu Sans" and spec.get("strict_font", False):
        raise RuntimeError("未找到中文字体，图内中文会缺字；请安装" +
                           palette.FONTS["cjk"][0])

    fig, ax = plt.subplots(figsize=(width_mm / MM_PER_INCH, height_mm / MM_PER_INCH))
    fig.set_layout_engine("constrained")
    fig.patch.set_facecolor(palette.FIGURE_NEUTRAL["background"])
    ax.set_facecolor(palette.FIGURE_NEUTRAL["background"])

    categories = [str(c) for c in (spec.get("categories") or [])]
    xs = list(spec.get("x") or range(len(series[0]["values"])))

    if kind in ("bar", "grouped_bar"):
        _plot_bars(ax, series, categories, style, spec)
    elif kind == "box":
        _plot_box(ax, series, categories, style)
    elif kind == "heatmap":
        _plot_heatmap(fig, ax, series, spec, style)
    else:  # line / scatter / errorbar
        _plot_lines(ax, series, xs, kind, style, spec)

    ax.set_xlabel(axis_text(parse_axis(spec.get("xlabel"))))
    ax.set_ylabel(axis_text(parse_axis(spec.get("ylabel"))))
    if spec.get("x_range"):
        ax.set_xlim(*spec["x_range"])
    if spec.get("y_range"):
        ax.set_ylim(*spec["y_range"])

    _style_axes(ax, style, style_name)
    if spec.get("grid", style["grid_alpha"] > 0):
        ax.grid(True, axis="y", linestyle=":", linewidth=0.5,
                color=palette.FIGURE_NEUTRAL["grid"], alpha=style["grid_alpha"])
        ax.set_axisbelow(True)

    if spec.get("title"):
        ax.set_title(str(spec["title"]), fontsize=style["label_size"], pad=4)
    if spec.get("panel"):
        ax.text(0.01, 0.98, str(spec["panel"]), transform=ax.transAxes,
                ha="left", va="top", fontsize=style["label_size"], fontweight="bold")

    show_legend = spec.get("legend")
    if show_legend is None:
        show_legend = kind in ("line", "scatter", "errorbar", "box") or len(series) > 1
    if show_legend:
        _legend(ax, series, kind, style)

    _annotate_significance(ax, spec, style, kind)

    out_path = os.path.abspath(out_path)
    os.makedirs(os.path.dirname(out_path) or ".", exist_ok=True)
    dpi = _dpi_for(spec)
    stem, ext = os.path.splitext(out_path)
    ext = (ext or ".png").lower()
    formats = [str(f).lower() for f in (spec.get("formats") or ["png"])]
    if ext.lstrip(".") not in formats:
        formats.insert(0, ext.lstrip("."))

    saved = []
    for fmt in formats:
        target = out_path if fmt == ext.lstrip(".") else "%s.%s" % (stem, fmt)
        # 不用 bbox_inches="tight"：tight 会把溢出标签算进来，导致实际宽度偏离 mm 设计值
        fig.savefig(target, dpi=dpi, facecolor=fig.get_facecolor())
        saved.append(target)

    if grayscale:
        grey_spec = dict(spec)
        grey_spec["series"] = [dict(s, color=palette.normalize("#000000")
                                   if i == 0 else palette.normalize("#7F7F7F")
                                   if i == 1 else palette.normalize("#BFBFBF"))
                               for i, s in enumerate(spec.get("series") or [])]
        grey_spec["strict_palette"] = False
        grey_path = "%s-grey%s" % (stem, ext or ".png")
        render(grey_spec, grey_path, grayscale=False)
        saved.append(grey_path)

    plt.close(fig)
    return saved[0] if saved else out_path


def _style_axes(ax, style: dict, style_name: str) -> None:
    axis_color = palette.FIGURE_NEUTRAL["axis"]
    for side in ("top", "right"):
        ax.spines[side].set_visible(style_name == "ieee")
        ax.spines[side].set_linewidth(style["spine_width"])
        ax.spines[side].set_color(axis_color)
    for side in ("left", "bottom"):
        ax.spines[side].set_linewidth(style["spine_width"])
        ax.spines[side].set_color(axis_color)
    ax.tick_params(direction="out", colors=axis_color)
    for label in ax.get_xticklabels() + ax.get_yticklabels():
        label.set_color(palette.FIGURE_NEUTRAL["text"])


def _plot_lines(ax, series, xs, kind, style, spec) -> None:
    for s in series:
        values = s["values"]
        if kind == "scatter":
            ax.plot(xs[:len(values)], values, linestyle="none", marker=s["marker"],
                    markersize=style["marker_size"], color=s["color"],
                    label=s["name"], zorder=3)
        elif kind == "errorbar" or s["error"]:
            ax.errorbar(xs[:len(values)], values, yerr=s["error"] or None,
                        color=s["color"], marker=s["marker"],
                        linestyle=s["linestyle"], linewidth=style["line_width"],
                        markersize=style["marker_size"], capsize=2.5,
                        elinewidth=style["line_width"] * 0.8, label=s["name"])
        else:
            ax.plot(xs[:len(values)], values, color=s["color"],
                    linestyle=s["linestyle"], marker=s["marker"],
                    linewidth=style["line_width"],
                    markersize=style["marker_size"], label=s["name"])


def _plot_bars(ax, series, categories, style, spec) -> None:
    n = max(len(categories), 1)
    groups = max(len(series), 1)
    slot = 1.0 / n
    bar_w = slot * 0.66 / groups
    import numpy as np
    idx = np.arange(n)
    for si, s in enumerate(series):
        offsets = idx + (si - (groups - 1) / 2) * bar_w
        values = s["values"][:n]
        ax.bar(offsets, values, width=bar_w * 0.9, color=s["color"],
               edgecolor=palette.FIGURE_NEUTRAL["axis"], linewidth=0.4,
               label=s["name"], zorder=2)
        if s["error"]:
            ax.errorbar(offsets, values, yerr=s["error"][:len(values)],
                        fmt="none", ecolor=palette.FIGURE_NEUTRAL["axis"],
                        elinewidth=style["line_width"] * 0.8, capsize=2.0)
    ax.set_xticks(idx)
    ax.set_xticklabels(categories or ["%d" % (i + 1) for i in idx])
    if spec.get("zero_base", True):
        ax.set_ylim(bottom=0)


def _plot_box(ax, series, categories, style) -> None:
    data = [s["values"] for s in series]
    bp = ax.boxplot(data, patch_artist=True, widths=0.55,
                    medianprops={"color": palette.FIGURE_NEUTRAL["axis"], "linewidth": 1.0},
                    boxprops={"linewidth": 0.7}, whiskerprops={"linewidth": 0.7},
                    capprops={"linewidth": 0.7}, flierprops={"markersize": 2})
    for i, patch in enumerate(bp["boxes"]):
        patch.set_facecolor(series[i]["color"])
        patch.set_alpha(0.85)
        patch.set_edgecolor(palette.FIGURE_NEUTRAL["axis"])
    ax.set_xticks(range(1, len(series) + 1))
    ax.set_xticklabels(categories or [s["name"] for s in series])


def _plot_heatmap(fig, ax, series, spec, style) -> None:
    import numpy as np
    data = np.array(series[0]["values"], dtype=float)
    rows = spec.get("rows") or [str(i + 1) for i in range(data.shape[0])]
    data = data.reshape(len(rows), -1) if data.ndim == 1 else data
    kind = "divergent" if (spec.get("cmap_kind") == "divergent" or spec.get("diverging")) \
        else "sequential"
    cmap = palette.resolve_cmap(spec.get("cmap") or "viridis", kind)
    image = ax.imshow(data, cmap=cmap, aspect="auto",
                      vmin=spec.get("vmin"), vmax=spec.get("vmax"))
    ax.set_yticks(range(len(rows)))
    ax.set_yticklabels(rows)
    if spec.get("categories"):
        ax.set_xticks(range(len(spec["categories"])))
        ax.set_xticklabels([str(c) for c in spec["categories"]])
    bar = fig.colorbar(image, ax=ax, pad=0.02, fraction=0.05)
    bar.ax.tick_params(labelsize=style["font_size"] - 0.5)
    if spec.get("colorbar_label"):
        bar.set_label(axis_text(parse_axis(spec["colorbar_label"])),
                      fontsize=style["font_size"])


def _legend(ax, series, kind, style) -> None:
    pos = "upper right"
    if kind in ("bar", "grouped_bar"):
        pos = "best"
    if kind == "box":
        handles = [Line2D([], [], color=s["color"], linewidth=6, label=s["name"])
                   for s in series]
        ax.legend(handles=handles, loc=pos, frameon=False)
        return
    ax.legend(loc=pos, frameon=False, handlelength=2.2,
              borderaxespad=0.3, labelspacing=0.3)


def _annotate_significance(ax, spec, style, kind) -> None:
    """显著性标注：brackets=[{x1,x2,y,label}]，标签如 '**'。"""
    brackets = spec.get("brackets") or []
    if not brackets:
        return
    lo, hi = ax.get_ylim()
    span = hi - lo
    for item in brackets:
        x1, x2 = float(item.get("x1", 0)), float(item.get("x2", 1))
        y = float(item.get("y", hi - span * 0.08))
        h = span * 0.02
        ax.plot([x1, x1, x2, x2], [y, y + h, y + h, y], color=palette.FIGURE_NEUTRAL["axis"],
                linewidth=0.7, clip_on=False)
        ax.text((x1 + x2) / 2, y + h, str(item.get("label", "*")),
                ha="center", va="bottom", fontsize=style["font_size"])
    ax.set_ylim(lo, hi + span * 0.06)


def main(argv=None):
    ap = argparse.ArgumentParser(description="按国标/期刊惯例渲染论文插图")
    ap.add_argument("spec", help="图的 JSON 规格")
    ap.add_argument("-o", "--output", required=True)
    ap.add_argument("--grayscale", action="store_true",
                    help="同时输出灰度版，用于黑白打印自检")
    args = ap.parse_args(argv)

    with open(os.path.abspath(args.spec), encoding="utf-8-sig") as fh:
        spec = json.load(fh)
    path = render(spec, args.output, grayscale=args.grayscale)
    print("已生成：%s" % path)
    return 0


if __name__ == "__main__":
    from _console import setup as _setup_console
    _setup_console()
    sys.exit(main())
