#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
build.py —— 横道图生成统一入口（CLI）

用法示例
--------
# 生成交互式 HTML（默认）
python build.py tasks.csv -o 进度计划.html

# 同时生成 PNG（用于 Word / PPT）
python build.py tasks.xlsx -o 进度计划.html --png 进度计划.png

# 指定项目名称与起始日期，按自然日排程（不跳周末）
python build.py tasks.csv -o out.html --name "某系统开发项目" --start 2026-03-01 --natural

# 只输出中间模型 JSON
python build.py tasks.csv --model model.json
"""

from __future__ import annotations

import argparse
import os
import sys

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

from gantt_core import load_tasks, build_model, model_to_json  # noqa: E402


def main(argv=None) -> int:
    ap = argparse.ArgumentParser(
        description="横道图（甘特图）生成器 —— 支持自动排程、关键路径、里程碑",
        formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("input", help="任务数据文件（.csv / .xlsx / .json）")
    ap.add_argument("-o", "--out", default=None, help="输出 HTML 路径")
    ap.add_argument("--png", default=None, help="输出 PNG 路径（需 matplotlib）")
    ap.add_argument("--model", default=None, help="同时输出中间模型 JSON 路径")
    ap.add_argument("--name", default=None, help="项目名称")
    ap.add_argument("--subtitle", default="", help="副标题")
    ap.add_argument("--start", default=None, help="项目起始日期 YYYY-MM-DD（缺省取最早任务）")
    ap.add_argument("--calendar", choices=["workday", "natural"], default="workday",
                    help="排程日历：workday 跳过周末（默认），natural 按自然日")
    ap.add_argument("--workdays", default="1,2,3,4,5",
                    help="工作日设置，1=周一 … 7=周日，默认 1,2,3,4,5")
    ap.add_argument("--holidays", default="",
                    help="额外节假日，YYYY-MM-DD 逗号分隔")
    ap.add_argument("--dpi", type=int, default=200, help="PNG 分辨率，默认 200")
    ap.add_argument("--no-deps", action="store_true", help="PNG 中不绘制依赖线")
    ap.add_argument("--no-baseline", action="store_true", help="PNG 中不绘制基线")
    ap.add_argument("--left-width", type=int, default=560, help="HTML 左侧任务表宽度(px)")
    args = ap.parse_args(argv)

    if not os.path.isfile(args.input):
        print(f"[错误] 找不到输入文件：{args.input}", file=sys.stderr)
        return 2

    base = os.path.splitext(os.path.basename(args.input))[0]
    out_html = args.out or f"{base}_甘特图.html"
    proj_name = args.name or f"{base} 进度计划"

    # 工作日历
    try:
        wd = [int(x.strip()) - 1 for x in args.workdays.split(",") if x.strip() != ""]
    except ValueError:
        print("[错误] --workdays 格式应为 1,2,3,4,5", file=sys.stderr)
        return 2

    holidays = []
    if args.holidays:
        from gantt_core import parse_date
        for h in args.holidays.split(","):
            d = parse_date(h.strip())
            if d:
                holidays.append(d)

    rows = load_tasks(args.input)
    if not rows:
        print("[错误] 输入文件没有可用的数据行", file=sys.stderr)
        return 2

    model = build_model(
        rows,
        project_start=args.start,
        workdays=wd,
        holidays=holidays,
        use_calendar=(args.calendar == "workday"),
        project_name=proj_name,
        subtitle=args.subtitle,
    )

    p = model["project"]
    print(f"[解析] 任务 {p['taskCount']} 个 · 里程碑 {p['milestoneCount']} 个 · "
          f"关键任务 {p['criticalCount']} 个")
    print(f"[周期] {p['start']} ~ {p['end']}（{p['totalDays']} 天）")

    # HTML
    from render_html import write_html
    write_html(model, out_html, left_width=args.left_width)
    print(f"[输出] HTML -> {os.path.abspath(out_html)}")

    # PNG
    if args.png:
        try:
            from render_png import render_png
            render_png(model, args.png, dpi=args.dpi,
                       show_deps=not args.no_deps,
                       show_baseline=not args.no_baseline)
            print(f"[输出] PNG  -> {os.path.abspath(args.png)}")
        except ImportError as e:
            print(f"[跳过] 未安装 matplotlib，无法生成 PNG：{e}", file=sys.stderr)

    # 模型
    if args.model:
        model_to_json(model, args.model)
        print(f"[输出] 模型 -> {os.path.abspath(args.model)}")

    return 0


if __name__ == "__main__":
    sys.exit(main())
