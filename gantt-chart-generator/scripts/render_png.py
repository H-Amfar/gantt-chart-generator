#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
render_png.py —— 把模型渲染为静态 PNG 横道图（用于 Word / PPT 报告）

特点：
  · 自适应画布尺寸（随任务数量与工期跨度伸缩）
  · 自动选择时间轴粒度（周 / 半月 / 月 / 季 / 年）
  · 里程碑菱形、依赖连线、关键路径、基线对比、今日线
  · 中文字体自动探测（Microsoft YaHei / SimHei / 苹方）
"""

from __future__ import annotations

import os
from datetime import date, datetime, timedelta

WEEKDAY_CN = ["周一", "周二", "周三", "周四", "周五", "周六", "周日"]

COLOR_CRITICAL = "#DC2626"
COLOR_SUMMARY = "#334155"
COLOR_MILESTONE = "#B45309"
COLOR_BASELINE = "#94A3B8"
COLOR_TODAY = "#EF4444"


def _setup_font():
    import matplotlib
    from matplotlib import font_manager

    wanted = ["Microsoft YaHei", "微软雅黑", "SimHei", "黑体",
              "PingFang SC", "Hiragino Sans GB", "Noto Sans CJK SC",
              "Source Han Sans SC", "WenQuanYi Micro Hei", "DejaVu Sans"]
    available = {f.name for f in font_manager.fontManager.ttflist}
    picked = [w for w in wanted if w in available]
    if not picked:
        picked = ["DejaVu Sans"]
    matplotlib.rcParams["font.sans-serif"] = picked
    matplotlib.rcParams["axes.unicode_minus"] = False
    return picked[0]


def _parse(s):
    if not s:
        return None
    return datetime.strptime(s, "%Y-%m-%d").date()


def render_png(model: dict, out_path: str, dpi: int = 200,
               show_deps: bool = True, show_baseline: bool = True,
               show_critical: bool = True, show_today: bool = True,
               width: float | None = None) -> str:
    import matplotlib
    matplotlib.use("Agg")
    import matplotlib.pyplot as plt
    import matplotlib.dates as mdates

    _setup_font()

    tasks = model["tasks"]
    proj = model["project"]
    p_start = _parse(proj["start"])
    p_end = _parse(proj["end"])
    span = (p_end - p_start).days + 1

    n = len(tasks)
    # 画布尺寸自适应
    fig_w = width or max(10.0, min(46.0, span / 14.0 + 5.0))
    fig_h = max(3.2, min(48.0, n * 0.36 + 2.0))
    fig, ax = plt.subplots(figsize=(fig_w, fig_h))

    # 时间轴粒度
    if span <= 45:
        major, minor, fmt = mdates.WeekdayLocator(byweekday=0), mdates.DayLocator(), "%m-%d"
    elif span <= 150:
        major, minor, fmt = mdates.MonthLocator(), mdates.WeekdayLocator(byweekday=0), "%Y-%m"
    elif span <= 760:
        major, minor, fmt = mdates.MonthLocator(interval=1), mdates.MonthLocator(), "%Y-%m"
    else:
        major, minor, fmt = mdates.MonthLocator(bymonth=[1, 4, 7, 10]), mdates.MonthLocator(), "%Y-%m"
    ax.xaxis.set_major_locator(major)
    ax.xaxis.set_major_formatter(mdates.DateFormatter(fmt))
    if minor is not None:
        ax.xaxis.set_minor_locator(minor)

    ys = list(range(n))

    # 周末底色 / 网格
    ax.grid(axis="x", which="major", color="#cbd5e1", linewidth=0.8, zorder=0)
    ax.grid(axis="x", which="minor", color="#eef2f7", linewidth=0.5, zorder=0)
    ax.set_axisbelow(True)

    by_id = {t["id"]: t for t in tasks}
    y_index = {t["id"]: i for i, t in enumerate(tasks)}

    # ---------- 任务条 ----------
    for i, t in enumerate(tasks):
        ts, te = _parse(t["start"]), _parse(t["end"])
        if ts is None or te is None:
            continue
        color = t.get("color") or "#3B82F6"
        if show_critical and t.get("critical") and t["type"] != "summary":
            color = COLOR_CRITICAL

        if t["type"] == "summary":
            ax.barh(i, (te - ts).days + 1, left=ts, height=0.28,
                    color=COLOR_SUMMARY, zorder=3, edgecolor="#1e293b", linewidth=0.5)
            for xx, sgn in ((ts, -1), (te, 1)):
                ax.plot([xx, xx], [i + 0.14, i + 0.30], color=COLOR_SUMMARY,
                        linewidth=1.1, zorder=3)
        elif t["type"] == "milestone":
            ax.scatter([ts], [i], marker="D", s=52, color=COLOR_MILESTONE,
                       zorder=6, edgecolor="#7c2d12", linewidth=0.6)
            ax.text(ts + timedelta(days=max(1, int(span * 0.006))), i,
                    t["name"], va="center", ha="left", fontsize=7.6,
                    color="#92400e", zorder=7)
        else:
            ax.barh(i, (te - ts).days + 1, left=ts, height=0.52,
                    color=color, zorder=3, edgecolor="white", linewidth=0.6)
            pct = max(0.0, min(100.0, t.get("progress", 0)))
            if pct > 0:
                ax.barh(i, ((te - ts).days + 1) * pct / 100.0, left=ts, height=0.52,
                        color="black", alpha=0.26, zorder=4)
            label = f"{pct:.0f}%" if (te - ts).days >= 10 else t["name"]
            inside = (te - ts).days >= 10
            ax.text(ts + timedelta(days=1) if inside else te + timedelta(days=1), i,
                    label, va="center", ha="left", fontsize=7.2,
                    color="white" if inside else "#475569",
                    fontweight="bold" if inside else "normal", zorder=7)

        # 基线
        if show_baseline and t.get("baseline"):
            bs, be = _parse(t["baseline"]["start"]), _parse(t["baseline"]["end"])
            if bs and be:
                ax.barh(i + 0.31, (be - bs).days + 1, left=bs, height=0.10,
                        color=COLOR_BASELINE, zorder=2)

    # ---------- 依赖线 ----------
    if show_deps:
        for t in tasks:
            for d in t.get("deps", []):
                p = by_id.get(d["id"])
                if not p or p["id"] not in y_index or t["id"] not in y_index:
                    continue
                ps, pe = _parse(p["start"]), _parse(p["end"])
                ts, te = _parse(t["start"]), _parse(t["end"])
                if None in (ps, pe, ts, te):
                    continue
                i_from, i_to = y_index[p["id"]], y_index[t["id"]]
                dt = d.get("type", "FS")
                if dt == "SS":
                    x1, x2 = ps, ts
                elif dt == "FF":
                    x1, x2 = pe, te
                elif dt == "SF":
                    x1, x2 = ps, te
                else:
                    x1, x2 = pe, ts
                crit = show_critical and t.get("critical") and p.get("critical")
                col = COLOR_CRITICAL if crit else "#64748B"
                ax.annotate("", xy=(x2, i_to), xytext=(x1, i_from),
                            arrowprops=dict(arrowstyle="-|>", color=col,
                                            linewidth=1.0 if not crit else 1.5,
                                            shrinkA=0, shrinkB=2,
                                            connectionstyle="angle,angleA=0,angleB=90,rad=2"),
                            zorder=5, alpha=0.85)

    # ---------- 今日线 ----------
    if show_today:
        today = date.today()
        if p_start <= today <= p_end:
            ax.axvline(today, color=COLOR_TODAY, linestyle="--", linewidth=1.4, zorder=8)
            ax.text(today, -1.15, "今天", color=COLOR_TODAY, fontsize=8,
                    ha="center", va="bottom", zorder=9, fontweight="bold")

    # ---------- 坐标轴 ----------
    labels = []
    for t in tasks:
        indent = "    " * int(t.get("level") or 0)
        wbs = (t.get("wbs") or "")
        prefix = (wbs + " ") if wbs else ""
        mark = "◆ " if t["type"] == "milestone" else ""
        labels.append(f"{indent}{prefix}{mark}{t['name']}")
    ax.set_yticks(ys)
    ax.set_yticklabels(labels, fontsize=8.4)
    for tick, t in zip(ax.get_yticklabels(), tasks):
        if t["type"] == "summary":
            tick.set_fontweight("bold")
            tick.set_color("#0F172A")
        elif t["type"] == "milestone":
            tick.set_color("#92400E")
        else:
            tick.set_color("#334155")

    ax.set_ylim(n - 0.5, -1.35)
    ax.set_xlim(p_start - timedelta(days=1), p_end + timedelta(days=1))
    ax.tick_params(axis="x", labelsize=8, labelrotation=0)
    ax.tick_params(axis="y", length=0)
    for side in ("top", "right", "left"):
        ax.spines[side].set_visible(False)
    ax.spines["bottom"].set_color("#cbd5e1")

    # ---------- 标题 ----------
    title = proj.get("name") or "项目进度计划"
    sub = proj.get("subtitle") or ""
    fig.suptitle(f"{title}\n{sub}".strip(), fontsize=14, fontweight="bold",
                 x=0.012, ha="left", color="#0F172A")
    fig.text(0.012, 0.012,
             f"计划周期 {proj['start']} ~ {proj['end']}　|　"
             f"任务 {proj['taskCount']} 个　里程碑 {proj['milestoneCount']} 个　"
             f"关键任务 {proj['criticalCount']} 个　|　生成于 {model['generatedAt']}",
             fontsize=7.5, color="#94A3B8", ha="left")

    fig.tight_layout(rect=(0, 0.028, 1, 0.955 if sub else 0.968))
    os.makedirs(os.path.dirname(os.path.abspath(out_path)) or ".", exist_ok=True)
    fig.savefig(out_path, dpi=dpi, facecolor="white", bbox_inches="tight")
    plt.close(fig)
    return out_path
