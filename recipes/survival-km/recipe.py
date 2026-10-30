import os
import sys

_here = os.path.dirname(os.path.abspath(__file__))
for _p in (_here, os.path.dirname(_here)):
    if _p not in sys.path:
        sys.path.insert(0, _p)
import _kit as K  # noqa: E402

"""配方 survival-km：Kaplan-Meier 生存曲线（含删失标记）。

数据契约（CSV 长表）：time,event[,group]
    time   随访时间（>0，升序无要求）
    event  1=事件发生，0=删失
    group  可选分组（2—4 组）
设计要点：
    · 阶梯曲线，删失点用短竖线标在曲线上
    · 图例给各组 n 与中位生存时间；曲线不延伸到样本最大时间之外
    · 图注写清随访起止、删失原因与检验方法（log-rank 等）
"""


def config():
    return dict(width_mm=140.0, style="cn", x_name="时间", x_unit="d",
                y_name="生存概率", y_symbol="S(t)")


def _km(times, events):
    pairs = sorted(zip(times, events))
    n = len(pairs)
    xs, ys = [0.0], [1.0]
    surv = 1.0
    at_risk = n
    i = 0
    while i < len(pairs):
        t = pairs[i][0]
        deaths = 0
        censored = 0
        while i < len(pairs) and pairs[i][0] == t:
            if pairs[i][1]:
                deaths += 1
            else:
                censored += 1
            i += 1
        if deaths:
            surv *= (1 - deaths / at_risk)
            xs.append(t)
            ys.append(surv)
        if censored:
            xs.append(t)
            ys.append(surv)
        at_risk -= deaths + censored
        if at_risk <= 0:
            break
    return xs, ys


def main():
    K.main_guard()
    data_path, out = sys.argv[1], sys.argv[2]
    width_mm = float(sys.argv[3]) if len(sys.argv) > 3 else config()["width_mm"]
    cfg = config()

    header, body = K.read_csv(data_path)
    rows = [dict(time=float(r[header.index("time")]),
                 event=int(float(r[header.index("event")])),
                 group=str(r[header.index("group")]) if "group" in header else "全部")
            for r in body]
    groups, data = [], {}
    for r in rows:
        if r["group"] not in data:
            data[r["group"]] = []
            groups.append(r["group"])
        data[r["group"]].append((r["time"], r["event"]))
    K.require(1 <= len(groups) <= 4, "分组 1—4；更多请拆图")

    fig, ax, st = K.setup(width_mm=width_mm, aspect=0.68, style=cfg["style"])
    for i, g in enumerate(groups):
        xs, ys = _km([t for t, _ in data[g]], [e for _, e in data[g]])
        stl = K.style_of(i)
        ax.step(xs, ys, where="post", color=stl["color"],
                linewidth=st["lw"], label="%s（n=%d）" % (g, len(data[g])))
        cens = sorted(t for t, e in data[g] if not e)
        for t in cens:
            surv = [y for x, y in zip(xs, ys) if x <= t][-1] if xs else 1.0
            ax.plot([t, t], [surv, max(0.0, surv - 0.03)], color=stl["color"],
                    linewidth=0.9)
    ax.set_ylim(0, 1.02)
    ax.set_xlim(left=0)
    ax.set_xlabel(K.label(cfg["x_name"], cfg["x_unit"]))
    ax.set_ylabel(K.label(cfg["y_name"], "", cfg["y_symbol"]))
    ax.legend(loc="upper right", frameon=False, fontsize=st["legend"])
    K.grid(ax, st)
    K.say("· 图注需写明：随访起止、删失原因、组间检验（log-rank 等）与其 p 值")
    K.save(fig, out, width_mm=width_mm, base_font_pt=st["font"], style=cfg["style"])
    K.say("已生成：%s" % out)


if __name__ == "__main__":
    main()
