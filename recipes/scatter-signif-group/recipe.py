import math
import os
import sys

_here = os.path.dirname(os.path.abspath(__file__))
for _p in (_here, os.path.dirname(_here)):
    if _p not in sys.path:
        sys.path.insert(0, _p)
import _kit as K  # noqa: E402

"""配方 scatter-signif-group：分组散点 + 显著性桥。

数据契约（CSV 长表）：group,value[,pair,label]
    group  组名（2—5 组）
    value  观测值（每组 ≥5 个；n<10 时必须露点，这是本配方的存在理由）
    pair   可选：要画显著性桥的组对，如 "A|B"；每行配对重复即可
    label  可选：该桥的标记，如 "*"、"**"、"n.s."
设计要点：
    · 每个观测点都画出来（抖动），叠加均值横线与 ±SE 误差棒
    · 显著性桥只画给定组对，标记必须与图注里的检验方法、校正方式一致
    · 不画"只有柱高和一根误差棒"的均值柱（n<10 时会掩盖分布）
"""


def config():
    return dict(width_mm=140.0, style="cn", x_name="组别", y_name="数值",
                y_unit="", y_symbol="", jitter=0.075, pairs=[])


def main():
    K.main_guard()
    data_path, out = sys.argv[1], sys.argv[2]
    width_mm = float(sys.argv[3]) if len(sys.argv) > 3 else config()["width_mm"]
    cfg = config()

    header, body = K.read_csv(data_path)
    rows = [dict(group=str(r[header.index("group")]),
                 value=float(r[header.index("value")]),
                 pair=str(r[header.index("pair")]) if "pair" in header else "",
                 label=str(r[header.index("label")]) if "label" in header else "")
            for r in body]
    names, data = [], []
    for r in rows:
        if r["group"] not in names:
            names.append(r["group"])
            data.append([])
        data[names.index(r["group"])].append(r["value"])
    K.require(2 <= len(names) <= 5, "组数 2—5；更多请拆图")
    K.require(all(len(d) >= 5 for d in data), "每组至少 5 个观测（本配方就是给小组样本用的）")

    fig, ax, st = K.setup(width_mm=width_mm, aspect=0.66, style=cfg["style"])
    for i, vals in enumerate(data):
        color = K.SERIES[i % len(K.SERIES)]
        jit = [((j % 5) - 2) * cfg["jitter"] for j in range(len(vals))]
        ax.plot([i + 1 + d for d in jit], vals, linestyle="none", marker="o",
                markersize=st["ms"], color=color, alpha=0.75,
                markeredgecolor="white", markeredgewidth=0.4, zorder=2)
        mean = sum(vals) / len(vals)
        se = (sum((v - mean) ** 2 for v in vals) / (len(vals) - 1)) ** 0.5 / math.sqrt(len(vals))
        ax.plot([i + 0.78, i + 1.22], [mean, mean], color=color, linewidth=2.0, zorder=3)
        ax.errorbar(i + 1, mean, yerr=se, fmt="none", ecolor=color,
                    elinewidth=0.9, capsize=2.2, zorder=3)

    ax.set_xticks(range(1, len(names) + 1))
    ax.set_xticklabels(names)
    ax.set_xlabel(cfg["x_name"])
    ax.set_ylabel(K.label(cfg["y_name"], cfg["y_unit"], cfg["y_symbol"]))
    K.grid(ax, st)

    # 显著性桥：按数据里出现的 pair 顺序画，标记用数据里给的 label
    bridges = []
    for r in rows:
        if r["pair"] and r["pair"] not in [b[0] for b in bridges]:
            bridges.append((r["pair"], r["label"]))
    if bridges:
        lo, hi = ax.get_ylim()
        span = hi - lo
        ax.set_ylim(lo, hi + span * 0.10 * len(bridges))
        for k, (pair, label) in enumerate(bridges):
            left, right = [p.strip() for p in pair.split("|")][:2]
            if left not in names or right not in names:
                continue
            x1, x2 = names.index(left) + 1, names.index(right) + 1
            y = hi + span * (0.03 + 0.10 * k)
            ax.plot([x1, x1, x2, x2], [y - span * 0.012, y, y, y - span * 0.012],
                    color=K.AXIS, linewidth=0.9, clip_on=False)
            ax.text((x1 + x2) / 2, y, label, ha="center", va="bottom",
                    fontsize=st["font"], color=K.AXIS)

    K.say("· 显著性标记含义（检验方法与校正）必须写进图注，本配方不猜")
    K.save(fig, out, width_mm=width_mm, base_font_pt=st["font"], style=cfg["style"])
    K.say("已生成：%s" % out)


if __name__ == "__main__":
    main()
