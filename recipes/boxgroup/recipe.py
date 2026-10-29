import os
import sys

# 两种布局都要能找到 _kit：仓库内（recipes/）与工作区（与脚本同目录）
_here = os.path.dirname(os.path.abspath(__file__))
for _p in (_here, os.path.dirname(_here)):
    if _p not in sys.path:
        sys.path.insert(0, _p)
import _kit as K  # noqa: E402

"""配方 boxgroup：分组箱线（含离群点）。

数据契约（CSV）：group,value（长表，每行一个观测）
设计要点：箱体按组着色并半透明；中位线加粗；须为 1.5 倍 IQR。
"""


def config():
    return dict(width_mm=140.0, style="cn", x_name="组别", y_name="数值",
                y_unit="", y_symbol="", rotate=0)


def main():
    K.main_guard()
    data_path, out = sys.argv[1], sys.argv[2]
    width_mm = float(sys.argv[3]) if len(sys.argv) > 3 else config()["width_mm"]
    cfg = config()
    header, body = K.read_csv(data_path)
    groups = [str(g) for g in K.column(header, body, "group")]
    values = K.column(header, body, "value")
    names, data = [], []
    for g, v in zip(groups, values):
        if g not in names:
            names.append(g)
            data.append([])
        data[names.index(g)].append(float(v))
    K.require(all(len(d) >= 3 for d in data), "每组至少 3 个观测")

    fig, ax, st = K.setup(width_mm=width_mm, style=cfg["style"])
    bp = ax.boxplot(data, patch_artist=True, widths=0.55,
                    medianprops=dict(color=K.AXIS, linewidth=1.0),
                    boxprops=dict(linewidth=0.7), whiskerprops=dict(linewidth=0.7),
                    capprops=dict(linewidth=0.7), flierprops=dict(markersize=2))
    for i, patch in enumerate(bp["boxes"]):
        patch.set_facecolor(K.SERIES[5 if i % 3 == 0 else (1 if i % 3 == 1 else 3)])
        patch.set_alpha(0.85)
        patch.set_edgecolor(K.AXIS)
    ax.set_xticks(range(1, len(names) + 1))
    ax.set_xticklabels(names, rotation=cfg["rotate"])
    ax.set_xlabel(cfg["x_name"])
    ax.set_ylabel(K.label(cfg["y_name"], cfg["y_unit"], cfg["y_symbol"]))
    K.grid(ax, st)
    K.save(fig, out, width_mm=width_mm, base_font_pt=st["font"], style=cfg["style"])
    print("已生成：%s" % out)


if __name__ == "__main__":
    main()
