#!/usr/bin/env python
"""生成符合规范配色的学术图表 PNG（16:9 / 1672x941，插入 14 cm 时约 300 dpi）。

不需要 matplotlib，只用 Pillow。配色默认 Okabe-Ito 色盲友好方案。

用法：
    python make_chart.py chart.json -o figs/fig1.png

chart.json:
{
  "type": "bar" | "grouped_bar" | "line",
  "title": "三种方案的耗时对比",
  "xlabel": "方案", "ylabel": "耗时 / s",
  "categories": ["基线", "改进", "本文"],
  "series": [{"name": "耗时", "values": [12.4, 9.8, 7.2], "color": "0072B2"}],
  "value_labels": true,
  "legend": false,
  "y_min": 0, "y_max": null,
  "width_px": 1672, "height_px": 941
}
"""

from __future__ import annotations

import argparse
import json
import os
import sys

import _console
from PIL import Image, ImageDraw, ImageFont

PALETTE = {
    "navy": "#1F3A5F", "blue": "#0072B2", "orange": "#E69F00",
    "green": "#009E73", "sky": "#56B4E9", "purple": "#CC79A7",
    "yellow": "#F0E442", "ink": "#262626", "gray": "#595959",
    "light": "#D9D9D9", "grid": "#E6E6E6", "white": "#FFFFFF",
}
DEFAULT_CYCLE = ["blue", "orange", "green", "sky", "purple", "navy"]

FONT_CANDIDATES = [
    r"C:\Windows\Fonts\msyh.ttc",
    r"C:\Windows\Fonts\msyhl.ttc",
    r"C:\Windows\Fonts\simhei.ttf",
    r"C:\Windows\Fonts\simsun.ttc",
    "/System/Library/Fonts/PingFang.ttc",
    "/usr/share/fonts/truetype/wqy/wqy-zenhei.ttc",
    "/usr/share/fonts/opentype/noto/NotoSansCJK-Regular.ttc",
    "/usr/share/fonts/truetype/dejavu/DejaVuSans.ttf",
]


def _font(size, bold=False):
    paths = FONT_CANDIDATES if not bold else (
        [r"C:\Windows\Fonts\msyhbd.ttc", r"C:\Windows\Fonts\simhei.ttf",
         r"C:\Windows\Fonts\msyh.ttc"] + FONT_CANDIDATES)
    for path in paths:
        if os.path.exists(path):
            try:
                return ImageFont.truetype(path, size)
            except OSError:
                continue
    return ImageFont.load_default()


def _color(name):
    if not name:
        return PALETTE["blue"]
    if name.startswith("#"):
        return name
    return PALETTE.get(name, "#" + name)


def _nice_ticks(lo, hi, count=5):
    import math
    if hi <= lo:
        hi = lo + 1
    span = hi - lo
    raw = span / max(count - 1, 1)
    mag = 10 ** math.floor(math.log10(raw)) if raw > 0 else 1
    step = mag
    for mult in (1, 2, 2.5, 5, 10):
        step = mag * mult
        if step >= raw:
            break
    lo_t = math.floor(lo / step) * step
    ticks = []
    hi_t = lo_t
    while hi_t < hi - step * 1e-9:
        ticks.append(hi_t)
        hi_t += step
    ticks.append(hi_t)          # 保证刻度覆盖最大值，避免数据点被裁掉
    if len(ticks) < 2:
        ticks = [lo, hi]
    return ticks


def _value_formatter(values):
    dec = 0
    for v in values:
        if v is None:
            continue
        s = repr(float(v))
        if "." in s and "e" not in s:
            dec = max(dec, len(s.split(".")[1]))
    dec = min(dec, 2)
    return lambda v: format(float(v), ".%df" % dec)


def _fmt(value, step):
    if step >= 1:
        return "%g" % round(value, 2)
    import math
    dec = max(0, min(3, int(math.ceil(-math.log10(step))) if step > 0 else 3))
    return format(value, ".%df" % dec)


def render(spec: dict, out_path: str) -> str:
    W = int(spec.get("width_px", 1672))
    H = int(spec.get("height_px", 941))
    kind = spec.get("type", "bar")
    series = spec.get("series") or []
    categories = spec.get("categories") or []
    show_values = spec.get("value_labels", True)
    show_legend = spec.get("legend", kind == "line" or len(series) > 1)

    for i, s in enumerate(series):
        s["_color"] = _color(s.get("color") or DEFAULT_CYCLE[i % len(DEFAULT_CYCLE)])

    if kind == "line" and spec.get("x") is not None:
        xs = list(spec["x"])
    else:
        xs = list(range(len(categories)))

    values_all = [v for s in series for v in s.get("values", []) if v is not None]
    lo = spec.get("y_min")
    lo = min(values_all) if lo is None else lo
    hi = spec.get("y_max")
    hi = max(values_all) if hi is None else hi
    if hi == lo:
        hi = lo + 1
    pad = (hi - lo) * 0.12
    if spec.get("y_min") is None:
        lo = 0 if lo >= 0 else lo - pad
    if spec.get("y_max") is None:
        hi = hi + pad
    ticks = _nice_ticks(lo, hi, 5)
    lo, hi = ticks[0], ticks[-1]
    vfmt = _value_formatter(values_all)

    img = Image.new("RGB", (W, H), PALETTE["white"])
    d = ImageDraw.Draw(img)

    f_title = _font(46, True)
    f_axis = _font(34)
    f_tick = _font(32)
    f_legend = _font(32)
    f_value = _font(30)

    left, right, top, bottom = 168, W - 64, 96, H - 172
    if spec.get("title"):
        d.text(((left + right) // 2, 44), spec["title"], font=f_title,
               fill=PALETTE["navy"], anchor="mm")
    else:
        top = 60

    plot_w, plot_h = right - left, bottom - top

    def y_pos(v):
        return bottom - (v - lo) / (hi - lo) * plot_h

    # 网格
    step = ticks[1] - ticks[0] if len(ticks) > 1 else 1
    for t in ticks:
        y = y_pos(t)
        d.line([(left, y), (right, y)], fill=PALETTE["grid"], width=2)
        d.text((left - 18, y), _fmt(t, step), font=f_tick,
               fill=PALETTE["gray"], anchor="rm")

    # 坐标轴
    d.line([(left, top - 10), (left, bottom)], fill=PALETTE["gray"], width=3)
    d.line([(left, bottom), (right + 10, bottom)], fill=PALETTE["gray"], width=3)

    # x 轴刻度文字
    if categories and kind != "line":
        slot = plot_w / len(categories)
        for i, cat in enumerate(categories):
            d.text((left + slot * (i + 0.5), bottom + 30), str(cat),
                   font=f_tick, fill=PALETTE["ink"], anchor="ma")
    elif kind == "line":
        n = max(len(xs), 1)
        for i, xv in enumerate(xs):
            px = left + (plot_w * (i / max(n - 1, 1)))
            d.text((px, bottom + 30), "%g" % float(xv), font=f_tick,
                   fill=PALETTE["ink"], anchor="ma")

    # 数据
    if kind in ("bar", "grouped_bar"):
        n_series = max(len(series), 1)
        slot = plot_w / max(len(categories), 1)
        group_w = slot * 0.62
        bar_w = group_w / n_series * 0.82
        for si, s in enumerate(series):
            vals = s.get("values", [])
            for i, v in enumerate(vals):
                if v is None:
                    continue
                cx = left + slot * (i + 0.5) - group_w / 2 + bar_w * (si + 0.5)
                x0, x1 = cx - bar_w / 2, cx + bar_w / 2
                y0, y1 = y_pos(v), y_pos(0 if lo <= 0 <= hi else lo)
                top_y, bot_y = min(y0, y1), max(y0, y1)
                d.rectangle([x0, top_y, x1, bot_y], fill=s["_color"])
                if show_values:
                    d.text((cx, top_y - 14), vfmt(v), font=f_value,
                           fill=PALETTE["ink"], anchor="mb")
    else:  # line
        n = max(len(xs), 1)
        for s in series:
            vals = s.get("values", [])
            pts = []
            for i, v in enumerate(vals):
                if v is None:
                    continue
                px = left + (plot_w * (i / max(n - 1, 1)))
                pts.append((px, y_pos(v)))
            if len(pts) > 1:
                d.line(pts, fill=s["_color"], width=6, joint="curve")
            for px, py in pts:
                r = 10
                d.ellipse([px - r, py - r, px + r, py + r], fill=PALETTE["white"],
                          outline=s["_color"], width=5)
            if show_values:
                for i, (px, py) in enumerate(pts):
                    d.text((px, py - 22), vfmt(vals[i]), font=f_value,
                           fill=PALETTE["ink"], anchor="mb")

    # 轴标题
    if spec.get("xlabel"):
        d.text(((left + right) // 2, H - 46), spec["xlabel"], font=f_axis,
               fill=PALETTE["ink"], anchor="mm")
    if spec.get("ylabel"):
        tmp = Image.new("RGBA", (int(plot_h * 0.9), 60), (255, 255, 255, 0))
        td = ImageDraw.Draw(tmp)
        td.text((tmp.width // 2, 30), spec["ylabel"], font=f_axis,
                fill=PALETTE["ink"], anchor="mm")
        tmp = tmp.rotate(90, expand=True)
        img.paste(tmp, (12, top + plot_h // 2 - tmp.height // 2), tmp)

    # 图例
    if show_legend and series:
        x = right - 24
        y = top + 6
        items = [(s.get("name", ""), s["_color"]) for s in series]
        widths = [d.textlength(name, font=f_legend) for name, _ in items]
        box_w = max(widths) + 74
        box_h = 46 * len(items) + 18
        d.rounded_rectangle([x - box_w, y, x, y + box_h], radius=10,
                            fill=(255, 255, 255), outline=PALETTE["light"], width=2)
        for i, (name, color) in enumerate(items):
            cy = y + 14 + 46 * i + 16
            d.rectangle([x - box_w + 18, cy - 12, x - box_w + 50, cy + 12],
                        fill=color)
            d.text((x - box_w + 62, cy), name, font=f_legend,
                   fill=PALETTE["ink"], anchor="lm")

    out_path = os.path.abspath(out_path)
    os.makedirs(os.path.dirname(out_path) or ".", exist_ok=True)
    img.save(out_path, "PNG", dpi=(300, 300))
    return out_path


def main(argv=None):
    _console.setup()
    ap = argparse.ArgumentParser(description="学术图表 PNG 生成（Pillow）")
    ap.add_argument("spec", help="图表 JSON 路径")
    ap.add_argument("-o", "--output", required=True)
    args = ap.parse_args(argv)
    with open(os.path.abspath(args.spec), encoding="utf-8") as fh:
        spec = json.load(fh)
    path = render(spec, args.output)
    print("已生成：%s" % path)
    return 0


if __name__ == "__main__":
    sys.exit(main())
