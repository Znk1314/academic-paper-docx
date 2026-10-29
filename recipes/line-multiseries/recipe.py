"""配方 line-multiseries：多系列折线（SD 误差线 + 显著性括号）。

数据契约（CSV，第一行表头）：
    x,<系列1>,<系列2>,...          必填；x 为数值
    err_<系列1>,err_<系列2>,...    可选；同长度，缺省则不画误差线
设计要点（适配时保留）：
    · 颜色 + 线型 + 标记三重编码（色盲与黑白打印都能区分）
    · 误差线上限/下限等长，cap 2.5 pt；图例置于图内空白处
    · 轴标目"量符号/单位"，符号斜体、单位正体（GB/T 7713.2 5.4.4）
"""

import os
import sys

# 两种布局都要能找到 _kit：仓库内（recipes/）与工作区（与脚本同目录）
_here = os.path.dirname(os.path.abspath(__file__))
for _p in (_here, os.path.dirname(_here)):
    if _p not in sys.path:
        sys.path.insert(0, _p)
import _kit as K  # noqa: E402


def config():
    return dict(width_mm=140.0, style="cn", x_symbol="t", x_name="采样周期",
                x_unit="s", y_symbol="P", y_name="平均功耗", y_unit="mW",
                brackets=[], legend_loc="upper right")


def main():
    K.main_guard()
    data_path, out = sys.argv[1], sys.argv[2]
    width_mm = float(sys.argv[3]) if len(sys.argv) > 3 else config()["width_mm"]
    cfg = config()

    header, body = K.read_csv(data_path)
    xs = K.column(header, body, "x")
    series = [h for h in header if h != "x" and not h.startswith("err_")]
    K.require(len(series) >= 1, "至少要有一个数值系列列（x 之外的列）")
    K.require(len(xs) >= 2, "x 至少 2 个点")

    fig, ax, st = K.setup(width_mm=width_mm, style=cfg["style"])
    for i, name in enumerate(series):
        ys = K.column(header, body, name)
        stl = K.style_of(i)
        err = None
        if "err_" + name in header:
            err = K.column(header, body, "err_" + name)
            K.require(len(err) == len(ys), "err_%s 与 %s 长度不一致" % (name, name))
        ax.errorbar(xs[:len(ys)], ys, yerr=err, label=name,
                    color=stl["color"], marker=stl["marker"],
                    linestyle=stl["linestyle"], linewidth=st["lw"],
                    markersize=st["ms"], capsize=2.5,
                    elinewidth=st["lw"] * 0.8)

    ax.set_xlabel(K.label(cfg["x_name"], cfg["x_unit"], cfg["x_symbol"]))
    ax.set_ylabel(K.label(cfg["y_name"], cfg["y_unit"], cfg["y_symbol"]))
    K.grid(ax, st)
    for item in cfg["brackets"]:
        K.bracket(ax, item["x1"], item["x2"], item["label"], y=item.get("y"))
    ax.legend(loc=cfg["legend_loc"], frameon=False, handlelength=2.2,
              borderaxespad=0.3, labelspacing=0.3)
    K.save(fig, out, width_mm=width_mm, base_font_pt=st["font"],
           style=cfg["style"])
    K.say("已生成：%s" % out)


if __name__ == "__main__":
    main()
