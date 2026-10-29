import os
import sys

# 两种布局都要能找到 _kit：仓库内（recipes/）与工作区（与脚本同目录）
_here = os.path.dirname(os.path.abspath(__file__))
for _p in (_here, os.path.dirname(_here)):
    if _p not in sys.path:
        sys.path.insert(0, _p)
import _kit as K  # noqa: E402

"""配方 nature-dotplot：小组样本点图（替代"均值柱 + 误差棒"）。

数据契约（CSV）：group,value
    每组 3—12 个观测（n<10 时**禁止**用均值柱：柱高+一根误差棒会掩盖分布）
设计要点（适配时保留）：
    · 每个观测点都画出来（stripplot 式抖动），叠加中位数横线
    · 组数 ≤6；宽度取 Nature 单栏 89 mm
    · 字号 7 pt；轴线 0.5 pt；无网格
"""


def config():
    return dict(journal="nature", style="nature", span="single",
                x_name="组别", y_name="数值", y_unit="", y_symbol="",
                jitter=0.06)


def main():
    K.main_guard()
    data_path, out = sys.argv[1], sys.argv[2]
    cfg = config()
    width_mm = float(sys.argv[3]) if len(sys.argv) > 3 else K.width_for(cfg["journal"], cfg["span"])

    header, body = K.read_csv(data_path)
    groups_raw = [str(g) for g in K.column(header, body, "group")]
    values = [float(v) for v in K.column(header, body, "value")]
    names, data = [], []
    for g, v in zip(groups_raw, values):
        if g not in names:
            names.append(g)
            data.append([])
        data[names.index(g)].append(v)
    K.require(all(3 <= len(d) <= 12 for d in data), "每组 3—12 个观测；更少请合并组，更多请用箱线")
    import statistics

    fig, ax, st = K.setup(width_mm=width_mm, aspect=0.72, style=cfg["style"])
    for i, (name, vals) in enumerate(zip(names, data)):
        stl = K.style_of(i * 2)
        jit = [((j % 5) - 2) * cfg["jitter"] / 2 for j in range(len(vals))]
        ax.plot([i + 1 + j for j in jit], vals, linestyle="none",
                marker=stl["marker"], markersize=st["ms"],
                color=stl["color"], label=name)
        med = statistics.median(vals)
        ax.plot([i + 1 - 0.22, i + 1 + 0.22], [med, med],
                color=K.AXIS, linewidth=1.0, zorder=3)
    ax.set_xticks(range(1, len(names) + 1))
    ax.set_xticklabels(names)
    ax.set_xlabel(cfg["x_name"])
    ax.set_ylabel(K.label(cfg["y_name"], cfg["y_unit"], cfg["y_symbol"]))
    K.say("· 中位线为粗横线；n<10 时这是替代均值柱的正确画法")
    K.save(fig, out, width_mm=width_mm, base_font_pt=st["font"], style=cfg["style"])
    K.say("已生成：%s" % out)


if __name__ == "__main__":
    main()
