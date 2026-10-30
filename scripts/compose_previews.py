#!/usr/bin/env python
"""把预览页拼成"两页一行"的图片，供 README / 落地页使用。

为什么不手工截图：手工图会过期（改了示例、图变了，README 里还是旧图）。
这里每次都从**当前 PDF 渲染出的页面图**重新拼，保证文档里看到的永远是刚生成的那一版。

用法：
    python scripts/compose_previews.py
    python scripts/compose_previews.py --thesis-pages 1,2,4,5,7,8 --csae-pages 1,2,3,4

约定：
    * 输入：examples/demo/build/preview/page-NN.png
            examples/journal_csae/build/preview/page-NN.png（由 render_preview.py 生成）
    * 输出：docs/screenshots/p1-*.png … 每张 = 左右两页
    * 一张图里必须是两页：README 里不允许"一页占一行"
"""

from __future__ import annotations

import argparse
import os

from PIL import Image, ImageDraw

REPO = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
THESIS = os.path.join(REPO, "examples", "demo", "build", "preview")
CSAE = os.path.join(REPO, "examples", "journal_csae", "build", "preview")
OUT = os.path.join(REPO, "docs", "screenshots")

TARGET_W = 660          # 单页宽度（拼完约 1342 px，GitHub 上清晰又不至于过大）
GAP = 22
BORDER = (214, 214, 214)

# (输出名, 页面目录, 左页码, 右页码, 说明)
PAIRS = [
    ("p1-cover-abstract", THESIS, 1, 2, "单栏：封面 + 摘要"),
    ("p2-toc-body", THESIS, 4, 5, "单栏：目录 + 正文首页"),
    ("p3-table-line", THESIS, 7, 8, "单栏：三线表 + 折线图 + 分组柱"),
    ("p4-scatter-box", THESIS, 9, 10, "单栏：散点 + 箱线 + 热图 + 结论"),
    ("p5-csae-cover", CSAE, 1, 2, "双栏：题名摘要 + 图 1—3 与表 2"),
    ("p6-csae-tail", CSAE, 3, 4, "双栏：图 4—5 与表 3 + 参考文献"),
]


# 图型库：两图一行（同一行放两个配方预览）
RECIPE_PAIRS = [
    ("g1-signif-corr", "scatter-signif-group", "corr-lower",
     "分组散点 + 显著性桥 ｜ 相关性矩阵（下三角 + 系数）"),
    ("g2-roc-surv", "roc-pr-curve", "survival-km",
     "ROC / PR 曲线（AUC） ｜ Kaplan-Meier 生存曲线"),
    ("g3-dual-heat", "dual-split", "heatmap-seq",
     "双变量拆图（替代双 Y 轴） ｜ 顺序色板热图"),
]


def _load(folder, page):
    path = os.path.join(folder, "page-%02d.png" % page)
    if not os.path.exists(path):
        raise SystemExit("缺页面图：%s（先跑 render_preview.py）" % path)
    return Image.open(path).convert("RGB")


def _fit(img, width=TARGET_W):
    w, h = img.size
    return img.resize((width, int(h * width / w)), Image.LANCZOS)


def compose(left_path, right_path, out_path):
    left, right = _fit(_load(*left_path)), _fit(_load(*right_path))
    height = max(left.size[1], right.size[1])
    canvas = Image.new("RGB", (TARGET_W * 2 + GAP + 2, height + 2), "white")
    canvas.paste(left, (1, 1))
    canvas.paste(right, (TARGET_W + GAP + 1, 1))
    ImageDraw.Draw(canvas).rectangle(
        [0, 0, canvas.size[0] - 1, canvas.size[1] - 1], outline=BORDER)
    os.makedirs(os.path.dirname(out_path), exist_ok=True)
    canvas.save(out_path, "PNG", optimize=True)
    return canvas.size


def _load_recipe(rid):
    path = os.path.join(REPO, "recipes", rid, "preview.png")
    if not os.path.exists(path):
        raise SystemExit("缺配方预览：%s（先跑该配方的 recipe.py）" % path)
    return Image.open(path).convert("RGB")


def _fit_h(img, height):
    w, h = img.size
    return img.resize((int(w * height / h), height), Image.LANCZOS)


def compose_figures(height=430):
    """把两个配方的预览上下拼成一行（各自缩到等高）。"""
    made = []
    for name, left_rid, right_rid, caption in RECIPE_PAIRS:
        left = _fit_h(_load_recipe(left_rid), height)
        right = _fit_h(_load_recipe(right_rid), height)
        canvas = Image.new("RGB",
                           (left.size[0] + right.size[0] + GAP + 2, height + 2),
                           "white")
        canvas.paste(left, (1, 1))
        canvas.paste(right, (left.size[0] + GAP + 1, 1))
        ImageDraw.Draw(canvas).rectangle(
            [0, 0, canvas.size[0] - 1, canvas.size[1] - 1], outline=BORDER)
        out = os.path.join(OUT, name + ".png")
        canvas.save(out, "PNG", optimize=True)
        made.append(name)
        print("  %-18s %sx%s  %s" % (name, canvas.size[0], canvas.size[1], caption))
    return made


def main(argv=None):
    ap = argparse.ArgumentParser(description="拼两页一行的预览图")
    ap.add_argument("--out", default=OUT)
    ap.add_argument("--figures", action="store_true",
                    help="同时拼图型库的两图一行图（需要各配方 preview.png）")
    ap.add_argument("--keep", action="store_true",
                    help="保留输出目录里不在本次清单中的旧图（默认不删，只覆盖）")
    args = ap.parse_args(argv)

    made = []
    for name, folder, left, right, caption in PAIRS:
        size = compose((folder, left), (folder, right),
                       os.path.join(args.out, name + ".png"))
        made.append(name)
        print("  %-20s %sx%s  %s" % (name, size[0], size[1], caption))
    print("共 %d 张（每张两页）→ %s" % (len(made), args.out))
    if args.figures:
        prints = compose_figures()
        print("图型库 %d 张（每张两个配方）→ %s" % (len(prints), args.out))
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
