import os
import sys

# 两种布局都要能找到 _kit：仓库内（recipes/）与工作区（与脚本同目录）
_here = os.path.dirname(os.path.abspath(__file__))
for _p in (_here, os.path.dirname(_here)):
    if _p not in sys.path:
        sys.path.insert(0, _p)
import _kit as K  # noqa: E402

"""配方 nature-multipanel：Nature 双栏多面板（a/b/c/d 面板标签）。

数据契约（CSV，长表）：panel,group,x,y
    panel  面板标识（a/b/c/d 或 P1..P4），最多 4 个
    group  组别（同一 group 在四个面板里同色——跨面板语义一致）
    每个 panel 内同一 group 至少 3 个点
设计要点（适配时保留）：
    · 整图 183 mm（Nature 双栏），插入后不再缩放
    · 字号 7 pt，最小 6 pt；轴线 0.5 pt；无上/右轴线、无网格
    · 面板标签用 figure 坐标加粗小写字母，不手摆 ax.text
    · 四个面板共用同一坐标范围与同一套颜色，避免"看起来不一样"
"""


def config():
    return dict(journal="nature", style="nature", span="double",
                x_symbol="x", x_name="变量", x_unit="",
                y_symbol="y", y_name="响应", y_unit="",
                layout=["a", "b", "c", "d"])


def main():
    K.main_guard()
    data_path, out = sys.argv[1], sys.argv[2]
    cfg = config()
    width_mm = float(sys.argv[3]) if len(sys.argv) > 3 else K.width_for(cfg["journal"], cfg["span"])

    header, body = K.read_csv(data_path)
    rows = []
    for r in body:
        rows.append(dict(panel=str(r[header.index("panel")]),
                         group=str(r[header.index("group")]),
                         x=float(r[header.index("x")]),
                         y=float(r[header.index("y")])))
    K.require(len(rows) >= 12, "至少 12 行（4 面板 × 3 组 × 1 点）")
    panels = cfg["layout"][:4]
    groups = sorted({r["group"] for r in rows})
    K.require(len(groups) <= len(K.SERIES) // 2, "组别过多，建议拆图")

    xs = [r["x"] for r in rows]
    ys = [r["y"] for r in rows]
    xpad = (max(xs) - min(xs)) * 0.06 or 1
    ypad = (max(ys) - min(ys)) * 0.06 or 1

    K.apply_style(cfg["style"])            # 字体/字号必须先配，否则中文变方框
    fig, axes = K.plt.subplots(2, 2, figsize=(width_mm / K.MM_PER_INCH,
                                              width_mm * 0.60 / K.MM_PER_INCH))
    fig.set_layout_engine("constrained")
    fig.subplots_adjust(wspace=0.28, hspace=0.34)
    flat = axes.ravel()
    for idx, label in enumerate(panels):
        ax = flat[idx]
        for g_i, g in enumerate(groups):
            pts = [r for r in rows if r["panel"] == label and r["group"] == g]
            if not pts:
                continue
            stl = K.style_of(g_i * 2)
            ax.plot([p["x"] for p in pts], [p["y"] for p in pts], linestyle="none",
                    marker=stl["marker"], markersize=K.STYLE["nature"]["ms"],
                    color=stl["color"], label=g if idx == 0 else None)
        ax.set_xlim(min(xs) - xpad, max(xs) + xpad)
        ax.set_ylim(min(ys) - ypad, max(ys) + ypad)
        for side in ("top", "right"):
            ax.spines[side].set_visible(False)
        if idx >= 2:                        # 只底行标 x
            ax.set_xlabel(K.label(cfg["x_name"], cfg["x_unit"], cfg["x_symbol"]))
        else:
            ax.set_xticklabels([])
        if idx % 2 == 0:                    # 只左列标 y
            ax.set_ylabel(K.label(cfg["y_name"], cfg["y_unit"], cfg["y_symbol"]))
        K.panel_label(fig, ax, label, style="nature")
    # "outside" 让图例占独立空间，constrained_layout 会为它让位，不会被裁掉
    fig.legend(loc="outside lower center", frameon=False, ncol=len(groups))
    K.save(fig, out, width_mm=width_mm, base_font_pt=K.STYLE["nature"]["font"],
           style="nature")
    K.say("已生成：%s" % out)


if __name__ == "__main__":
    main()
