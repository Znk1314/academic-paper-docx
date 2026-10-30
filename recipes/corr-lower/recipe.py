import math
import os
import sys

_here = os.path.dirname(os.path.abspath(__file__))
for _p in (_here, os.path.dirname(_here)):
    if _p not in sys.path:
        sys.path.insert(0, _p)
import _kit as K  # noqa: E402

"""配方 corr-lower：相关性矩阵（下三角 + 系数）。

数据契约（CSV 宽表）：第 1 列为样本标签，其余每列一个变量（数值）
设计要点：
    · 只画下三角，对角为 1.00；上三角留白，避免同一系数出现两次
    · 发散色板 RdBu_r，以 0 为中心（正负对称），色条标 r
    · 单元格写系数，字号按格子大小自动缩小；|r|≥0.7 的格子加粗
    · 图注必须写清是 Pearson 还是 Spearman，以及样本量 n
"""


def config():
    return dict(width_mm=140.0, style="cn", method="pearson", vmin=-1.0, vmax=1.0)


def _pearson(xs, ys):
    n = len(xs)
    mx, my = sum(xs) / n, sum(ys) / n
    sxx = sum((x - mx) ** 2 for x in xs)
    syy = sum((y - my) ** 2 for y in ys)
    sxy = sum((x - mx) * (y - my) for x, y in zip(xs, ys))
    if sxx <= 0 or syy <= 0:
        return 0.0
    return sxy / math.sqrt(sxx * syy)


def _rank(vals):
    order = sorted(range(len(vals)), key=lambda i: vals[i])
    ranks = [0.0] * len(vals)
    i = 0
    while i < len(order):
        j = i
        while j + 1 < len(order) and vals[order[j + 1]] == vals[order[i]]:
            j += 1
        avg = (i + j) / 2.0 + 1
        for k in range(i, j + 1):
            ranks[order[k]] = avg
        i = j + 1
    return ranks


def main():
    K.main_guard()
    data_path, out = sys.argv[1], sys.argv[2]
    width_mm = float(sys.argv[3]) if len(sys.argv) > 3 else config()["width_mm"]
    cfg = config()

    header, body = K.read_csv(data_path)
    labels = [str(r[0]) for r in body]
    cols = [K.column(header, body, col) for col in header[1:]]
    names = header[1:]
    n = len(cols)
    K.require(n >= 2, "至少 2 个变量")
    K.require(len(labels) == len(cols[0]), "每列行数需一致")

    if cfg["method"].lower() == "spearman":
        cols = [_rank([float(v) for v in c]) for c in cols]

    mat = [[0.0] * n for _ in range(n)]
    for i in range(n):
        for j in range(n):
            mat[i][j] = _pearson([float(v) for v in cols[i]],
                                 [float(v) for v in cols[j]])

    fig, ax, st = K.setup(width_mm=width_mm, aspect=0.82, style=cfg["style"])
    masked = [[(mat[i][j] if j <= i else float("nan")) for j in range(n)]
              for i in range(n)]
    im = ax.imshow(masked, cmap="RdBu_r", vmin=cfg["vmin"], vmax=cfg["vmax"])
    ax.set_xticks(range(n))
    ax.set_yticks(range(n))
    ax.set_xticklabels(names, rotation=35, ha="right")
    ax.set_yticklabels(names)
    for i in range(n):
        for j in range(n):
            if j > i:
                continue
            r = mat[i][j]
            dark = abs(r) > 0.55
            ax.text(j, i, "%.2f" % r, ha="center", va="center",
                    fontsize=st["font"] - 0.5,
                    fontweight="bold" if abs(r) >= 0.7 else "normal",
                    color="white" if dark else K.AXIS)
    bar = fig.colorbar(im, ax=ax, pad=0.02, fraction=0.05)
    bar.set_label("$r$", fontsize=st["font"])
    K.say("· 图注需写明：%s 相关，n=%d；下三角显示，上三角留白" % (cfg["method"], len(labels)))
    K.save(fig, out, width_mm=width_mm, base_font_pt=st["font"], style=cfg["style"])
    K.say("已生成：%s" % out)


if __name__ == "__main__":
    main()
