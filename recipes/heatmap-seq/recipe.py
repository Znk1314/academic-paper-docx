import os
import sys

# 两种布局都要能找到 _kit：仓库内（recipes/）与工作区（与脚本同目录）
_here = os.path.dirname(os.path.abspath(__file__))
for _p in (_here, os.path.dirname(_here)):
    if _p not in sys.path:
        sys.path.insert(0, _p)
import _kit as K  # noqa: E402

"""配方 heatmap-seq：顺序色板热图（cividis）。

数据契约（CSV）：row,col1,col2,...（第一列为行标签，其余为数值）
设计要点：只用感知均匀顺序色板（禁 jet/rainbow）；色条标目带单位；行标签不省略。
"""


def config():
    return dict(width_mm=140.0, style="cn", cmap="cividis",
                row_name="节点", col_name="时段", value_name="温度",
                value_unit="℃", bar_symbol="T")


def main():
    K.main_guard()
    data_path, out = sys.argv[1], sys.argv[2]
    width_mm = float(sys.argv[3]) if len(sys.argv) > 3 else config()["width_mm"]
    cfg = config()
    if str(cfg["cmap"]).lower() in ("jet", "rainbow", "nipy_spectral", "hsv"):
        raise SystemExit("禁止使用彩虹色谱：亮度不单调，会伪造数据边界")
    header, body = K.read_csv(data_path)
    rows = [str(r[0]) for r in body]
    cols = header[1:]
    matrix = [[float(v) for v in r[1:]] for r in body]
    K.require(all(len(r) == len(cols) for r in matrix), "矩阵每行长度需等于列数")

    fig, ax, st = K.setup(width_mm=width_mm, aspect=0.78, style=cfg["style"])
    im = ax.imshow(matrix, cmap=cfg["cmap"], aspect="auto")
    ax.set_yticks(range(len(rows)))
    ax.set_yticklabels(rows)
    ax.set_xticks(range(len(cols)))
    ax.set_xticklabels(cols)
    ax.set_xlabel(cfg["col_name"])
    ax.set_ylabel(cfg["row_name"])
    bar = fig.colorbar(im, ax=ax, pad=0.02, fraction=0.05)
    bar.set_label(K.label(cfg["value_name"], cfg["value_unit"], cfg["bar_symbol"]),
                  fontsize=st["font"])
    K.save(fig, out, width_mm=width_mm, base_font_pt=st["font"], style=cfg["style"])
    K.say("已生成：%s" % out)


if __name__ == "__main__":
    main()
