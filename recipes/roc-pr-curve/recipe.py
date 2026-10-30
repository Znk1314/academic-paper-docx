import os
import sys

_here = os.path.dirname(os.path.abspath(__file__))
for _p in (_here, os.path.dirname(_here)):
    if _p not in sys.path:
        sys.path.insert(0, _p)
import _kit as K  # noqa: E402

"""配方 roc-pr-curve：ROC 与 PR 曲线（含 AUC）。

数据契约（CSV 长表）：model,kind,x,y
    model  方法名
    kind   "roc"（x=FPR, y=TPR）或 "pr"（x=Recall, y=Precision）
    x, y   已算好的曲线点（升序）；两种 kind 都有时并排两个面板
设计要点：
    · ROC 画对角虚线作随机基线；PR 画阳性率水平线作基线
    · 图例给 AUC（梯形积分），保留 3 位小数
    · 曲线点多时只画线不画点；点少（≤8）时加标记，避免"假平滑"
"""


def config():
    return dict(width_mm=140.0, style="cn")


def _auc(xs, ys):
    pairs = sorted(zip(xs, ys))
    area = 0.0
    for (x1, y1), (x2, y2) in zip(pairs, pairs[1:]):
        area += (x2 - x1) * (y1 + y2) / 2.0
    return area


def main():
    K.main_guard()
    data_path, out = sys.argv[1], sys.argv[2]
    width_mm = float(sys.argv[3]) if len(sys.argv) > 3 else config()["width_mm"]
    cfg = config()

    header, body = K.read_csv(data_path)
    curves = {}
    for r in body:
        model = str(r[header.index("model")])
        kind = str(r[header.index("kind")]).lower()
        curves.setdefault(kind, {}).setdefault(model, [[], []])
        curves[kind][model][0].append(float(r[header.index("x")]))
        curves[kind][model][1].append(float(r[header.index("y")]))
    K.require(curves, "数据为空")
    kinds = [k for k in ("roc", "pr") if k in curves]

    K.apply_style(cfg["style"])            # 多子图必须先配字体，否则中文变方框
    fig, axes = K.plt.subplots(1, len(kinds), squeeze=False,
                               figsize=(width_mm / K.MM_PER_INCH,
                                        width_mm * (0.46 if len(kinds) > 1 else 0.62)
                                        / K.MM_PER_INCH))
    fig.set_layout_engine("constrained")
    for ax, kind in zip(axes[0], kinds):
        for i, (model, (xs, ys)) in enumerate(curves[kind].items()):
            stl = K.style_of(i)
            ax.plot(xs, ys, color=stl["color"], linewidth=K.STYLE["cn"]["lw"],
                    marker=stl["marker"] if len(xs) <= 8 else None,
                    markersize=K.STYLE["cn"]["ms"] * 0.9,
                    label="%s（%s %.3f）" % (
                        model, "AUC" if kind == "roc" else "AP", _auc(xs, ys)))
        if kind == "roc":
            ax.plot([0, 1], [0, 1], linestyle="--", linewidth=0.8, color=K.GREY[2])
            ax.set_xlabel("假阳性率（1 − 特异度）")
            ax.set_ylabel("真阳性率（灵敏度）")
        else:
            ax.set_xlabel("召回率（Recall）")
            ax.set_ylabel("精确率（Precision）")
        ax.set_xlim(0, 1.02)
        ax.set_ylim(0, 1.02)
        ax.legend(loc="lower right" if kind == "roc" else "lower left",
                  frameon=False, fontsize=K.STYLE["cn"]["legend"])
        for side in ("top", "right"):
            ax.spines[side].set_visible(False)
    K.say("· 图注需写明：阈值选取方式、正负样本数与置信区间（若有）")
    K.save(fig, out, width_mm=width_mm, base_font_pt=K.STYLE["cn"]["font"],
           style=cfg["style"])
    K.say("已生成：%s" % out)


if __name__ == "__main__":
    main()
