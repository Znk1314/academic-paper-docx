import os
import sys

# 两种布局都要能找到 _kit：仓库内（recipes/）与工作区（与脚本同目录）
_here = os.path.dirname(os.path.abspath(__file__))
for _p in (_here, os.path.dirname(_here)):
    if _p not in sys.path:
        sys.path.insert(0, _p)
import _kit as K  # noqa: E402

"""配方 scatter-fit：散点 + 参考线（1:1 或线性拟合）。

数据契约（CSV）：x,<系列1>[,<系列2>]；可用 ref=true 画 y=x 参考线
设计要点：标记与颜色双重编码；拟合线用虚线；轴范围按数据 5% 留白。
"""


def config():
    return dict(width_mm=140.0, style="cn", x_name="参考值", x_unit="",
                x_symbol="x", y_name="实测值", y_unit="", y_symbol="y",
                ref_line=True)


def main():
    K.main_guard()
    data_path, out = sys.argv[1], sys.argv[2]
    width_mm = float(sys.argv[3]) if len(sys.argv) > 3 else config()["width_mm"]
    cfg = config()
    header, body = K.read_csv(data_path)
    xs = K.column(header, body, "x")
    series = [h for h in header if h != "x"]

    fig, ax, st = K.setup(width_mm=width_mm, style=cfg["style"])
    lo = min(xs) * 0.95
    hi = max(xs) * 1.05
    if cfg["ref_line"]:
        ax.plot([lo, hi], [lo, hi], linestyle="--", linewidth=st["lw"] * 0.8,
                color=K.SERIES[7], zorder=1, label="1:1 参考线")
    for i, name in enumerate(series):
        stl = K.style_of(i + 5)
        ys = K.column(header, body, name)
        ax.plot(xs[:len(ys)], ys, linestyle="none", marker=stl["marker"],
                markersize=st["ms"], color=stl["color"], label=name, zorder=3)
    ax.set_xlim(lo, hi)
    ax.set_ylim(lo, hi)
    ax.set_xlabel(K.label(cfg["x_name"], cfg["x_unit"], cfg["x_symbol"]))
    ax.set_ylabel(K.label(cfg["y_name"], cfg["y_unit"], cfg["y_symbol"]))
    ax.legend(loc="upper left", frameon=False)
    K.grid(ax, st)
    K.save(fig, out, width_mm=width_mm, base_font_pt=st["font"], style=cfg["style"])
    K.say("已生成：%s" % out)


if __name__ == "__main__":
    main()
