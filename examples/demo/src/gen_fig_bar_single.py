import os
import sys

# 两种布局都要能找到 _kit：仓库内（recipes/）与工作区（与脚本同目录）
_here = os.path.dirname(os.path.abspath(__file__))
for _p in (_here, os.path.dirname(_here)):
    if _p not in sys.path:
        sys.path.insert(0, _p)
import _kit as K  # noqa: E402

"""配方 bar-single：单系列分类柱状（零基线 + 可选误差线）。

数据契约（CSV）：category,value[,err]
设计要点：y 轴从 0 起；柱子宽度占类别位 0.55；有 err 时画 SD 误差线。
"""


def config():
    return dict(width_mm=140.0, style="cn", x_name="类别", y_name="数值",
                y_unit="", y_symbol="", grid=True)


def main():
    K.main_guard()
    data_path, out = sys.argv[1], sys.argv[2]
    width_mm = float(sys.argv[3]) if len(sys.argv) > 3 else config()["width_mm"]
    cfg = config()
    header, body = K.read_csv(data_path)
    cats = [str(c) for c in K.column(header, body, "category")]
    vals = K.column(header, body, "value")
    errs = K.column(header, body, "err") if "err" in header else None
    K.require(len(cats) == len(vals), "category 与 value 行数不一致")

    fig, ax, st = K.setup(width_mm=width_mm, style=cfg["style"])
    idx = range(len(vals))
    ax.bar(idx, vals, width=0.55, color=K.SERIES[5], edgecolor=K.AXIS,
           linewidth=0.4, zorder=2)
    if errs:
        ax.errorbar(idx, vals, yerr=errs, fmt="none", ecolor=K.AXIS,
                    elinewidth=st["lw"] * 0.8, capsize=2.0)
    ax.set_xticks(list(idx))
    ax.set_xticklabels(cats)
    ax.set_xlabel(cfg["x_name"])
    ax.set_ylabel(K.label(cfg["y_name"], cfg["y_unit"], cfg["y_symbol"]))
    ax.set_ylim(bottom=0)
    if cfg["grid"]:
        K.grid(ax, st)
    K.save(fig, out, width_mm=width_mm, base_font_pt=st["font"], style=cfg["style"])
    print("已生成：%s" % out)


if __name__ == "__main__":
    main()
