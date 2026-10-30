import os
import sys

_here = os.path.dirname(os.path.abspath(__file__))
for _p in (_here, os.path.dirname(_here)):
    if _p not in sys.path:
        sys.path.insert(0, _p)
import _kit as K  # noqa: E402

"""配方 dual-split：双变量拆图（替代双 Y 轴）。

数据契约（CSV）：x,left,right
    x      横轴（连续或有序）
    left   上图变量（如左轴量纲）
    right  下图变量（如右轴量纲）
设计要点：
    · 双 Y 轴会把"视觉相关"造成出来，这里拆成上下两个子图、共享 x
    · 两个子图各自标轴（量符号/单位），不共用 y 刻度
    · 上下两图 xy 对齐：左边界与绘图区宽度一致
"""


def config():
    return dict(width_mm=140.0, style="cn", x_name="时间", x_unit="d",
                x_symbol="t", left_name="浓度", left_unit="mg/L", left_symbol="C",
                right_name="去除率", right_unit="%", right_symbol="R")


def main():
    K.main_guard()
    data_path, out = sys.argv[1], sys.argv[2]
    width_mm = float(sys.argv[3]) if len(sys.argv) > 3 else config()["width_mm"]
    cfg = config()

    header, body = K.read_csv(data_path)
    xs = K.column(header, body, "x")
    left = K.column(header, body, "left")
    right = K.column(header, body, "right")
    K.require(len(xs) == len(left) == len(right), "三列长度需一致")

    K.apply_style(cfg["style"])
    fig, (ax1, ax2) = K.plt.subplots(
        2, 1, sharex=True, figsize=(width_mm / K.MM_PER_INCH,
                                    width_mm * 0.72 / K.MM_PER_INCH))
    fig.set_layout_engine("constrained")
    st = K.STYLE[cfg["style"]]

    ax1.plot(xs, left, color=K.SERIES[0], marker="o", markersize=st["ms"],
             linewidth=st["lw"], markeredgecolor="white", markeredgewidth=0.4)
    ax1.set_ylabel(K.label(cfg["left_name"], cfg["left_unit"], cfg["left_symbol"]))
    ax1.set_ylim(0, max(left) * 1.15)

    ax2.plot(xs, right, color=K.SERIES[1], marker="^", markersize=st["ms"],
             linewidth=st["lw"], markeredgecolor="white", markeredgewidth=0.4)
    ax2.set_ylabel(K.label(cfg["right_name"], cfg["right_unit"], cfg["right_symbol"]))
    ax2.set_xlabel(K.label(cfg["x_name"], cfg["x_unit"], cfg["x_symbol"]))
    ax2.set_ylim(0, 100)

    for ax in (ax1, ax2):
        for side in ("top", "right"):
            ax.spines[side].set_visible(False)
        K.grid(ax, st)
    ax1.set_title("a", loc="left", fontweight="bold", fontsize=st["label"] + 1)
    ax2.set_title("b", loc="left", fontweight="bold", fontsize=st["label"] + 1)
    K.say("· 已拆成上下两图、共享 x：不制造虚假的视觉相关（替代双 Y 轴）")
    K.save(fig, out, width_mm=width_mm, base_font_pt=st["font"], style=cfg["style"])
    K.say("已生成：%s" % out)


if __name__ == "__main__":
    main()
