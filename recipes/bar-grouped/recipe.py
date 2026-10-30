import os
import sys

# 两种布局都要能找到 _kit：仓库内（recipes/）与工作区（与脚本同目录）
_here = os.path.dirname(os.path.abspath(__file__))
for _p in (_here, os.path.dirname(_here)):
    if _p not in sys.path:
        sys.path.insert(0, _p)
import _kit as K  # noqa: E402

"""配方 bar-grouped：分组柱状（多系列 + 误差线）。

数据契约（CSV）：category,<系列1>,<系列2>,...[,err_<系列>]
设计要点：组宽占类别位 0.8，组内均分；y 轴从 0 起；图例在右上空白处。
"""


def config():
    return dict(width_mm=140.0, style="cn", x_name="类别", y_name="数值",
                y_unit="", y_symbol="", legend_loc="upper right")


def main():
    K.main_guard()
    data_path, out = sys.argv[1], sys.argv[2]
    width_mm = float(sys.argv[3]) if len(sys.argv) > 3 else config()["width_mm"]
    cfg = config()
    header, body = K.read_csv(data_path)
    cats = [str(c) for c in K.column(header, body, "category")]
    series = [h for h in header if h != "category" and not h.startswith("err_")]
    K.require(series, "至少一个数值系列列")
    n, groups = len(cats), len(series)
    bar_w = 0.8 / groups
    idx = list(range(n))

    fig, ax, st = K.setup(width_mm=width_mm, style=cfg["style"])
    for si, name in enumerate(series):
        vals = K.column(header, body, name)
        offs = [i + (si - (groups - 1) / 2) * bar_w for i in idx]
        ax.bar(offs, vals, width=bar_w * 0.86, color=K.SERIES[si % len(K.SERIES)],
               edgecolor="white", linewidth=0.8, label=name, zorder=2)
        if "err_" + name in header:
            errs = K.column(header, body, "err_" + name)
            ax.errorbar(offs, vals, yerr=errs, fmt="none", ecolor=K.AXIS,
                        elinewidth=st["lw"] * 0.8, capsize=2.0)
    ax.set_xticks(idx)
    ax.set_xticklabels(cats)
    ax.set_xlabel(cfg["x_name"])
    ax.set_ylabel(K.label(cfg["y_name"], cfg["y_unit"], cfg["y_symbol"]))
    ax.set_ylim(bottom=0)
    ax.legend(loc="lower center", bbox_to_anchor=(0.5, 1.02), ncol=len(series),
              frameon=False, handlelength=1.6, columnspacing=1.2)
    K.grid(ax, st)
    K.save(fig, out, width_mm=width_mm, base_font_pt=st["font"], style=cfg["style"])
    K.say("已生成：%s" % out)


if __name__ == "__main__":
    main()
