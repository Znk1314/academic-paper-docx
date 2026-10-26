#!/usr/bin/env python
"""兼容入口：旧脚本名保留，内部改用 `figures.py`（matplotlib 后端）。

为什么换后端：Pillow 画的图没有真正的矢量输出、字号无法按毫米版面精确控制，
色板与色觉/灰度校验也无法复用。现在统一走 figures.py：

    python scripts/figures.py spec.json -o figs/fig1.png      # 推荐
    python scripts/make_chart.py spec.json -o figs/fig1.png    # 等价（自动转字段）

旧字段 `type` / `width_px` / `y_min` / `y_max` 会被自动映射；
`value_labels` 不再支持（论文图里数值用图注或正文说明更规范）。
"""

from __future__ import annotations

import argparse
import json
import os
import sys

HERE = os.path.dirname(os.path.abspath(__file__))
if HERE not in sys.path:
    sys.path.insert(0, HERE)

import figures  # noqa: E402


def convert(spec: dict) -> dict:
    """旧 spec → 新 spec；已是新格式时基本原样返回。"""
    out = dict(spec)
    out["kind"] = spec.get("kind") or spec.get("type") or "line"

    # 旧版按像素给尺寸（1672 px @300 dpi ≈ 141 mm），转成毫米版面
    if "width_mm" not in out and "width_px" in spec:
        px = float(spec.get("width_px") or 1672)
        out["width_mm"] = round(px / 300 * 25.4, 1)
        out.setdefault("dpi", 300)
    out.setdefault("width_mm", 140)
    out.setdefault("style", "cn")
    out.setdefault("dpi", 300)

    if "y_range" not in out:
        lo, hi = spec.get("y_min"), spec.get("y_max")
        if lo is not None and hi is not None:
            out["y_range"] = [lo, hi]
        elif lo is not None:
            out["y_range"] = None
            out["zero_base"] = lo == 0
    if spec.get("value_labels"):
        print("提示：value_labels 已停用，数值请写进图注或正文。", file=sys.stderr)
    return out


def render(spec: dict, out_path: str, grayscale: bool = False) -> str:
    return figures.render(convert(spec), out_path, grayscale=grayscale)


def main(argv=None):
    from _console import setup as _setup_console
    _setup_console()
    ap = argparse.ArgumentParser(description="兼容入口：旧参数 → figures.py")
    ap.add_argument("spec")
    ap.add_argument("-o", "--output", required=True)
    ap.add_argument("--grayscale", action="store_true")
    args = ap.parse_args(argv)
    with open(os.path.abspath(args.spec), encoding="utf-8-sig") as fh:
        spec = json.load(fh)
    path = render(spec, args.output, grayscale=args.grayscale)
    print("已生成：%s" % path)
    return 0


if __name__ == "__main__":
    sys.exit(main())
