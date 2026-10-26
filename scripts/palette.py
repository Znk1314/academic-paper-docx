"""科研配色：图与 Word 文档共用的**唯一色值来源**。

设计原则（对应 references/palette.md）：

1. **数据色只用经过验证的色盲安全色板**（Okabe-Ito / Paul Tol / Viridis 家族），
   不做"凭感觉调色"。图上任何系列色都必须来自 `SERIES`。
2. **色板类型跟着数据类型走**：定性（分类）→ `SERIES`；顺序（连续）→ `SEQUENTIAL`；
   发散（围绕中心值）→ `DIVERGENT`。定量数据**禁用** jet/rainbow，本模块直接拒绝。
3. **冗余编码**：颜色之外必须同时用线型/标记区分（`serie_style()` 提供循环）。
4. **灰度与色觉校验可执行**：`grayscale_hex()` / `delta_l()` / `simulate_cvd()`
   让"转灰度还能区分吗"变成可断言的数字，而不是主观判断。
5. **文档版式色不参与数据编码**，但仍受数量（≤4 彩色）与一致性约束；
   `DOC` 里的色值代表示例文档的外观，改这里即可全稿换肤。

所有色值统一大写 `#RRGGBB`。
"""

from __future__ import annotations

import colorsys
import math

# ---------------------------------------------------------------- 定性色板

# Okabe-Ito（Okabe & Ito 色觉通用设计；Bang Wong, Nature Methods 2011 推广）。
# 顺序即推荐使用顺序：黑 → 橙 → 天蓝 → 蓝绿 → 黄 → 蓝 → 红紫 → 灰。
OKABE_ITO: dict[str, str] = {
    "black": "#000000",
    "orange": "#E69F00",
    "sky": "#56B4E9",
    "green": "#009E73",
    "yellow": "#F0E442",
    "blue": "#0072B2",
    "purple": "#CC79A7",
    "grey": "#999999",
}
SERIES: list[str] = list(OKABE_ITO.values())
SERIES_NAMES: list[str] = list(OKABE_ITO.keys())

# Paul Tol bright（色盲安全，饱和度更高，用于需要 >5 类或屏幕展示）
TOL_BRIGHT: list[str] = [
    "#4477AA", "#EE6677", "#228833", "#CCBB44",
    "#66CCEE", "#AA3377", "#BBBBBB",
]

# 灰阶安全色：全彩不可用时的兜底（≥5 级，阶差 ≥20% 亮度）
GREY_RAMP: list[str] = ["#000000", "#404040", "#808080", "#C0C0C0", "#F2F2F2"]

# ---------------------------------------------------------------- 连续色板

# 顺序（低→高单调）：viridis 感知均匀；cividis 针对红绿色盲优化
SEQUENTIAL: dict[str, str] = {
    "viridis": "viridis",
    "cividis": "cividis",
    "batlow": "batlow",      # 需 cmcrameri；未安装时回退 viridis
    "grey": "Greys",
}

# 发散（围绕 0/对照）：中心中性、两端对比色相
DIVERGENT: dict[str, str] = {
    "RdBu": "RdBu_r",
    "vik": "vik",            # 需 cmcrameri；未安装时回退 RdBu_r
    "coolwarm": "coolwarm",
}

# 明令禁止：亮度不单调，会在数据中伪造边界（规范原话：不得用于定量数据）
FORBIDDEN_CMAPS = {"jet", "rainbow", "gist_rainbow", "nipy_spectral", "hsv", "gist_ncar"}

# ---------------------------------------------------------------- 文档版式色

# 彩色数 = 4（深蓝/蓝/橙/绿），落在"主色 ≤4—5"的通行区间内。
# 其中蓝 #0072B2、橙 #E69F00、绿 #009E73 直接取自 Okabe-Ito；
# heading 是中性深蓝，只用于标题层级，不参与任何数据编码。
DOC: dict[str, str] = {
    "heading": "#1F3A5F",     # 一级标题（中性深蓝，非数据色）
    "accent": "#0072B2",      # 二级标题 / 强调线（Okabe-Ito blue）
    "warn": "#E69F00",        # 辅助强调（Okabe-Ito orange）
    "ok": "#009E73",          # 正向 / 完成（Okabe-Ito green）
    "ink": "#262626",         # 正文（近黑，非纯黑以减轻屏幕眩光）
    "muted": "#595959",       # 次级说明 / 图注
    "pale": "#EAF2FA",        # 浅底提示框（蓝系 8% 明度）
    "rule": "#B4C6E7",        # 浅边框
    "paper": "#FFFFFF",
}

# 图中允许出现的"非数据"元素色（背景、轴线、刻度、网格）
FIGURE_NEUTRAL: dict[str, str] = {
    "background": "#FFFFFF",
    "axis": "#262626",
    "grid": "#D9D9D9",
    "text": "#1A1A1A",
}

# 字体：中文/西文/数学
FONTS: dict[str, list[str]] = {
    "cjk": ["Microsoft YaHei", "SimHei", "Noto Sans CJK SC",
            "Source Han Sans SC", "PingFang SC", "WenQuanYi Zen Hei"],
    "latin": ["Arial", "Helvetica", "DejaVu Sans"],
}

# 版面宽度（mm）。GB/T 7713.2—2022 5.4.1 只要求"大小适当、图中文字清晰可见"，
# 具体数值按版心推导：本模板 A4 21 cm − 左右边距 2.6/2.4 cm = 版心 16.0 cm，
# 通栏取 155 mm（与三线表同宽，留 5 mm 安全余量），半栏取 75 mm。
WIDTH_MM: dict[str, float] = {
    "half": 75.0,      # 半栏
    "full": 155.0,     # 通栏（= 版心宽 15.5 cm）
    "column-half": 75.0,
    "column-full": 155.0,
    "nature-half": 89.0,
    "nature-full": 183.0,
}

# 图内文字的硬下限（pt）：六号 7.5 pt 是常用最小字号，低于 6 pt 印刷后基本不可读
MIN_FIGURE_FONT_PT = 6.0

# 线宽/字号（pt）——按"插入后不再缩放"设计
STYLE: dict[str, dict] = {
    "cn": {           # 中文期刊：正文小四(12)，图内文字六号(8.5)，图题小五(9)
        "font_size": 8.5, "label_size": 9.0, "legend_size": 8.0,
        "line_width": 1.1, "marker_size": 4.0, "tick_width": 0.7,
        "spine_width": 0.7, "grid_alpha": 0.25,
    },
    "nature": {       # Nature 风：更细的线、更小的字
        "font_size": 7.0, "label_size": 7.0, "legend_size": 6.0,
        "line_width": 0.75, "marker_size": 2.8, "tick_width": 0.5,
        "spine_width": 0.5, "grid_alpha": 0.0,
    },
    "ieee": {         # IEEE 风：稍粗的线，全框可选
        "font_size": 8.0, "label_size": 8.0, "legend_size": 7.0,
        "line_width": 1.2, "marker_size": 4.0, "tick_width": 0.8,
        "spine_width": 0.8, "grid_alpha": 0.2,
    },
}

# 冗余编码循环：颜色之外，线型与标记也必须区分（规范 2.2）
LINESTYLES: list[str] = ["-", "--", ":", "-."]
MARKERS: list[str] = ["o", "s", "^", "D", "v", "P", "X", "*"]


# ---------------------------------------------------------------- 基础换算

def normalize(value: str) -> str:
    """'#000000' / '000000' / '#000' → '#RRGGBB'。"""
    v = (value or "").strip().lstrip("#")
    if len(v) == 3:
        v = "".join(ch * 2 for ch in v)
    if len(v) != 6:
        raise ValueError("非法颜色值：%r" % value)
    return "#" + v.upper()


def hex_to_rgb(value: str) -> tuple[float, float, float]:
    v = normalize(value).lstrip("#")
    return tuple(int(v[i:i + 2], 16) / 255.0 for i in (0, 2, 4))  # type: ignore[return-value]


def rgb_to_hex(rgb) -> str:
    return "#" + "".join("%02X" % max(0, min(255, round(c * 255))) for c in rgb)


def _srgb_to_linear(c: float) -> float:
    return c / 12.92 if c <= 0.04045 else ((c + 0.055) / 1.055) ** 2.4


def _linear_to_srgb(c: float) -> float:
    c = max(0.0, min(1.0, c))
    return 12.92 * c if c <= 0.0031308 else 1.055 * (c ** (1 / 2.4)) - 0.055


def relative_luminance(value: str) -> float:
    """WCAG 相对亮度（0=黑, 1=白）。"""
    r, g, b = (_srgb_to_linear(c) for c in hex_to_rgb(value))
    return 0.2126 * r + 0.7152 * g + 0.0722 * b


def grayscale_hex(value: str) -> str:
    """按亮度转灰度（模拟黑白打印）。"""
    lum = relative_luminance(value)
    v = _linear_to_srgb(lum)
    return rgb_to_hex((v, v, v))


def lstar(value: str) -> float:
    """CIE L*（0—100，感知均匀的明度）。"相邻灰阶差 ≥20%" 按 L* 判定。"""
    lum = relative_luminance(value)
    return 116.0 * (lum ** (1 / 3.0)) - 16.0 if lum > 0.008856 else 903.3 * lum


def delta_l(a: str, b: str) -> float:
    """两色的感知明度差 ΔL*（0—100），用于"转灰度后是否还能区分"。"""
    return abs(lstar(a) - lstar(b))


def contrast_ratio(a: str, b: str = "#FFFFFF") -> float:
    """WCAG 对比度（正文文字建议 ≥4.5，大字/图形 ≥3）。"""
    la, lb = relative_luminance(a), relative_luminance(b)
    hi, lo = max(la, lb), min(la, lb)
    return (hi + 0.05) / (lo + 0.05)


# ---------------------------------------------------------------- 色觉模拟

# Viénot, Brettel & Mollon (1999) 的线性 RGB 近似矩阵
_CVD_MATRIX = {
    "protanopia": ((0.152286, 1.052583, -0.204868),
                   (0.114503, 0.786281, 0.099216),
                   (-0.003882, -0.048116, 1.051998)),
    "deuteranopia": ((0.367322, 0.860646, -0.227968),
                     (0.280085, 0.672501, 0.047413),
                     (-0.011820, 0.042940, 0.968881)),
    "tritanopia": ((1.255528, -0.076749, -0.178779),
                   (-0.078411, 0.930809, 0.147602),
                   (0.004733, 0.691367, 0.303900)),
}


def simulate_cvd(value: str, kind: str = "deuteranopia") -> str:
    """模拟色觉障碍（protanopia / deuteranopia / tritanopia）下的显色。"""
    key = kind.lower()
    if key not in _CVD_MATRIX:
        raise ValueError("未知色觉类型：%s" % kind)
    r, g, b = (_srgb_to_linear(c) for c in hex_to_rgb(value))
    m = _CVD_MATRIX[key]
    out = tuple(
        sum(m[i][j] * (r, g, b)[j] for j in range(3)) for i in range(3)
    )
    return rgb_to_hex(tuple(_linear_to_srgb(c) for c in out))


def cvd_distance(a: str, b: str, kind: str = "deuteranopia") -> float:
    """两种颜色在指定色觉条件下的"可区分度"（0—100，越大越可分）。"""
    sa, sb = simulate_cvd(a, kind), simulate_cvd(b, kind)
    ra, rb = hex_to_rgb(sa), hex_to_rgb(sb)
    return math.dist(ra, rb) * 100


# ---------------------------------------------------------------- 组合校验

def serie_style(index: int) -> dict:
    """第 index 个数据系列的颜色 + 线型 + 标记（冗余编码）。"""
    return {
        "color": SERIES[index % len(SERIES)],
        "linestyle": LINESTYLES[index % len(LINESTYLES)],
        "marker": MARKERS[index % len(MARKERS)],
    }


def theme_cmap(name: str):
    """把主题色卡变成 matplotlib 连续色图（用于热图、密度图、渐变）。"""
    from matplotlib.colors import LinearSegmentedColormap
    return LinearSegmentedColormap.from_list(str(name), theme(name)["colors"])


def resolve_cmap(name: str, kind: str = "sequential") -> str:
    """把逻辑名解析成 matplotlib colormap；拒绝彩虹色谱。"""
    key = (name or "").strip()
    if key.lower() in FORBIDDEN_CMAPS:
        raise ValueError(
            "禁止使用 %r：彩虹色谱亮度不单调，会在定量数据中伪造边界"
            "（规范要求改用 viridis / cividis / 顺序色板）" % name)
    if key in THEMES:          # 主题色卡直接当色图
        return theme_cmap(key)
    table = DIVERGENT if kind == "divergent" else SEQUENTIAL
    cmap = table.get(key, key)
    try:  # cmcrameri 未安装时回退，保证跨平台可用
        import matplotlib  # noqa: F401
        from matplotlib import colormaps
        if cmap in colormaps:
            return cmap
    except Exception:  # pragma: no cover
        pass
    if cmap in {"batlow", "oleron"}:
        return SEQUENTIAL["viridis"]
    if cmap in {"vik", "roma", "broc"}:
        return DIVERGENT["RdBu"]
    return cmap


def check_series(colors: list[str], *, min_delta_l: float = 20.0):
    """校验一组"同时出现在一张图里"的系列色。

    返回 ``(errors, warnings)``：

    * **errors（硬错）**：不在批准色板内；在红绿色盲（deuteranopia/protanopia，
      约占男性 8%）下几乎同色 —— 这两类必须改。
    * **warnings（提示）**：灰度亮度差 < 20%（黑白打印会糊）；蓝黄色盲下接近。
      这两类靠"颜色 + 线型 + 标记"的冗余编码兜住，不强行阻断。
    """
    errors: list[str] = []
    warnings: list[str] = []
    normalized = [normalize(c) for c in colors]
    allowed = ({normalize(c) for c in SERIES} | {normalize(c) for c in TOL_BRIGHT}
               | {normalize(c) for c in ALL_THEME_COLORS})
    for c in normalized:
        if c not in allowed:
            errors.append("%s 不在批准色板内（应取自 Okabe-Ito / Tol bright）" % c)

    for i in range(len(normalized)):
        for j in range(i + 1, len(normalized)):
            a, b = normalized[i], normalized[j]
            if delta_l(a, b) < min_delta_l:
                warnings.append("%s 与 %s 灰度亮度差 %.0f%% < %.0f%%，"
                                "黑白打印需依赖线型/标记区分"
                                % (a, b, delta_l(a, b), min_delta_l))
            for kind in ("deuteranopia", "protanopia"):
                dist = cvd_distance(a, b, kind)
                if dist < 8:
                    errors.append("%s 与 %s 在 %s 下几乎同色（距离 %.0f）"
                                  % (a, b, kind, dist))
                elif dist < 12:
                    warnings.append("%s 与 %s 在 %s 下接近（距离 %.0f），"
                                    "靠线型/标记区分" % (a, b, kind, dist))
            if cvd_distance(a, b, "tritanopia") < 12:
                warnings.append("%s 与 %s 在蓝黄色盲下接近（蓝黄色盲罕见）" % (a, b))
            if is_red_green_pair(a, b):
                warnings.append("%s 与 %s 构成红—绿组合，建议改用蓝—橙等组合" % (a, b))
    return errors, warnings


def is_red_green_pair(a: str, b: str) -> bool:
    """是否为红—绿配对（规范明令避免的对比组合）。"""
    ha, sa, va = colorsys.rgb_to_hsv(*hex_to_rgb(a))
    hb, sb, vb = colorsys.rgb_to_hsv(*hex_to_rgb(b))
    ha, hb = ha * 360, hb * 360
    green_a = 75 <= ha <= 165
    green_b = 75 <= hb <= 165
    red_a = ha <= 45 or ha >= 340
    red_b = hb <= 45 or hb >= 340
    return (red_a and green_b) or (green_a and red_b)


def doc_colors() -> dict[str, str]:
    """文档版式色（大写规范化后的副本）。"""
    return {k: normalize(v) for k, v in DOC.items()}


def chromatic_count(colors) -> int:
    """统计"彩色"数量（排除纯灰/黑白）。"""
    n = 0
    for c in colors:
        h, s, v = colorsys.rgb_to_hsv(*hex_to_rgb(c))
        if s > 0.08 and 0.05 < v < 0.98:
            n += 1
    return n

# ---------------------------------------------------------------- 主题色卡

# 用户提供的科研配色设计稿（9 套）。注意：**这些是"主题色卡"，不是 N 类分类色板**——
# 多数含同色系深浅（如 rose6 有三个红色阶），直接当 N 个类别用会在灰度/色盲下糊掉。
# 所以额外提供 `safe_subset(name, n)`：按可区分度自动挑出能当 n 类用的子集。
THEMES: dict[str, dict] = {
    "rose6": {
        "label": "玫瑰豆沙（低饱和红+雾蓝）", "kind": "qualitative",
        "colors": ["#B87B7B", "#D49A9A", "#F2DADB", "#D2DDE0", "#7B97A1", "#144A58"],
        "roles": {"主强调": "#B87B7B", "次强调": "#D49A9A", "浅底": "#F2DADB",
                  "过渡": "#D2DDE0", "分类": "#7B97A1", "深色文字": "#144A58"},
    },
    "spectral8": {
        "label": "暖冷渐变（ColorBrewer Spectral 8 级）", "kind": "diverging",
        "colors": ["#D53E4F", "#F46D43", "#FDAE61", "#FEE08B",
                   "#E6F598", "#ABDDA4", "#66C2A5", "#3288BD"],
        "roles": {"高值": "#D53E4F", "中点": "#FEE08B", "低值": "#3288BD"},
    },
    "minimal6": {
        "label": "简约六色（珊瑚+雾绿+芥黄+钢蓝）", "kind": "qualitative",
        "colors": ["#F27468", "#F8C9C2", "#85B5B3", "#F4D26C", "#457B9D", "#F3AA63"],
        "roles": {"主强调": "#F27468", "浅底": "#F8C9C2", "分类A": "#85B5B3",
                  "分类B": "#F4D26C", "深色文字": "#457B9D", "辅助": "#F3AA63"},
    },
    "map7": {
        "label": "地图灵感（绿系+珊瑚+赭石）", "kind": "qualitative",
        "colors": ["#C9E0B2", "#6BB084", "#74DCCC", "#F9ADA4", "#F9BC4F", "#D7813C", "#F94D37"],
        "roles": {"浅底": "#C9E0B2", "绿": "#6BB084", "青": "#74DCCC", "粉": "#F9ADA4",
                  "黄": "#F9BC4F", "赭": "#D7813C", "强调红": "#F94D37"},
    },
    "medical6": {
        "label": "医学示意（青+明黄+珊瑚+灰绿+丁香紫）", "kind": "qualitative",
        "colors": ["#0098B6", "#F8D05B", "#F08072", "#F8B6B9", "#98A880", "#D090C5"],
        "roles": {"青": "#0098B6", "黄": "#F8D05B", "珊瑚": "#F08072",
                  "粉": "#F8B6B9", "灰绿": "#98A880", "紫": "#D090C5"},
    },
    "ocean8": {
        "label": "全球气候（云海白→冷区蓝）", "kind": "diverging",
        "colors": ["#F5F7F7", "#D0E8F8", "#90C0D8", "#708A98",
                   "#E2CCB2", "#E9AF99", "#D96E57", "#7195DB"],
        "roles": {"底色": "#F5F7F7", "冷": "#90C0D8", "中性": "#E2CCB2", "暖": "#D96E57"},
    },
    "material6": {
        "label": "材料科学（深蓝+雾蓝+蜜桃+深绯+孔雀青+珊瑚）", "kind": "qualitative",
        "colors": ["#187BBA", "#68A8CC", "#FCB796", "#C21F30", "#0C8B94", "#E7674A"],
        "roles": {"深蓝": "#187BBA", "雾蓝": "#68A8CC", "蜜桃": "#FCB796",
                  "深绯": "#C21F30", "孔雀青": "#0C8B94", "珊瑚": "#E7674A"},
    },
    "bluepink10": {
        "label": "蓝粉渐变（10 级，发散）", "kind": "diverging",
        "colors": ["#104FBC", "#376A9D", "#6089B2", "#AEC4D8", "#C5EAE4",
                   "#D7E0EC", "#F2DBDC", "#E5B5B4", "#DA9092", "#B02222"],
        "roles": {"深蓝端": "#104FBC", "中点": "#C5EAE4", "粉端": "#B02222"},
    },
    "medical2_6": {
        "label": "医学示意二（天蓝+蜜桃+米黄+浅绿+米灰+浅青）", "kind": "qualitative",
        "colors": ["#39B8E0", "#F0AF90", "#F8E8B3", "#9FD0A0", "#E8D0C0", "#C8E7E0"],
        "roles": {"天蓝": "#39B8E0", "蜜桃": "#F0AF90", "米黄": "#F8E8B3",
                  "浅绿": "#9FD0A0", "米灰": "#E8D0C0", "浅青": "#C8E7E0"},
    },
}

ALL_THEME_COLORS: list[str] = [c for t in THEMES.values() for c in t["colors"]]


def theme(name: str) -> dict:
    """取主题色卡；接受 key 或中文 label 片段。"""
    key = (name or "").strip()
    if key in THEMES:
        return THEMES[key]
    for t in THEMES.values():
        if key and (key in t["label"]):
            return t
    raise KeyError("未知色卡：%r（可用：%s）" % (name, "、".join(THEMES)))


def analyze(colors: list[str], min_cvd: float = 12.0, min_dl: float = 12.0) -> dict:
    """量一组颜色的可区分度：任意两两的红绿色盲距离与感知明度差。"""
    worst = (999.0, "", "")
    worst_dl = (999.0, "", "")
    red_green = 0
    for i in range(len(colors)):
        for j in range(i + 1, len(colors)):
            a, b = colors[i], colors[j]
            cvd = min(cvd_distance(a, b, "deuteranopia"),
                      cvd_distance(a, b, "protanopia"))
            dl = delta_l(a, b)
            if cvd < worst[0]:
                worst = (cvd, a, b)
            if dl < worst_dl[0]:
                worst_dl = (dl, a, b)
            if is_red_green_pair(a, b):
                red_green += 1
    if not colors:
        verdict = "空"
    elif worst[0] >= min_cvd and worst_dl[0] >= min_dl:
        verdict = "安全"
    elif worst[0] >= 8 and worst_dl[0] >= 8:
        verdict = "可用（需线型/标记冗余）"
    else:
        verdict = "不建议作分类色"
    return {
        "ok": verdict == "安全", "verdict": verdict,
        "min_cvd": worst[0], "cvd_pair": (worst[1], worst[2]),
        "min_dl": worst_dl[0], "dl_pair": (worst_dl[1], worst_dl[2]),
        "red_green_pairs": red_green,
    }


def safe_subset(name: str, n: int, min_cvd: float = 15.0,
                min_dl: float = 15.0) -> dict:
    """从主题色卡里挑 n 个"能当分类色用"的颜色（贪心最大化两两可区分度）。

    设计色卡常有同色系深浅，n 大于安全数量时会自动放宽阈值并在报告中说明。
    """
    raw = list(theme(name)["colors"])
    # 近白/近黑是背景或文字色，不能当分类填充色（条形填充会"消失"或糊成黑块）
    colors = [c for c in raw if 12.0 <= lstar(c) <= 93.0] or raw
    n = max(1, min(int(n), len(colors)))
    if n == 1:
        return {"colors": colors[:1], "ok": True, "note": "单系列"}

    def score(a: str, b: str) -> float:
        cvd = min(cvd_distance(a, b, "deuteranopia"),
                  cvd_distance(a, b, "protanopia"))
        return 0.6 * min(cvd / 100.0, 1.0) + 0.4 * min(delta_l(a, b) / 60.0, 1.0)

    # 种子：与其它颜色的距离之和最大者（避免选到渐变里的中间调）
    seed = max(colors, key=lambda c: sum(score(c, o) for o in colors if o != c))
    chosen = [seed]
    while len(chosen) < n:
        rest = [c for c in colors if c not in chosen]
        best = max(rest, key=lambda c: min(score(c, x) for x in chosen))
        chosen.append(best)

    chosen.sort(key=colors.index)          # 保持设计稿里的顺序
    report = analyze(chosen, min_cvd, min_dl)
    report["colors"] = chosen
    if not report["ok"]:
        report["note"] = ("该色卡只有 %d 个可安全区分的色阶，%d 类会接近；"
                          "已自动放宽，请配合线型/标记区分" % (len(chosen), n))
    return report
