#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
render_html.py —— 把模型渲染为自包含的交互式 HTML 横道图

特性：
  · 日期自适应时间轴（自动/日/周/月/季 五种视图）
  · 关键节点（里程碑）菱形标记
  · 任务依赖箭头（FS/SS/FF/SF）
  · 关键路径高亮（CPM 计算）
  · 进度填充、基线对比、今日线
  · 拖拽调整日期，后续任务按依赖自动顺延
  · 导出 PNG / 打印
"""

from __future__ import annotations

import json
import os

TEMPLATE_NAME = "gantt_template.html"
DATA_PLACEHOLDER = '/*__GANTT_DATA__*/ {"project":{},"tasks":[]}'
DEFAULT_LEFT_W = 560


def template_path() -> str:
    here = os.path.dirname(os.path.abspath(__file__))
    return os.path.normpath(os.path.join(here, "..", "assets", TEMPLATE_NAME))


def render_html(model: dict, tmpl: str | None = None,
                left_width: int = DEFAULT_LEFT_W,
                title: str | None = None) -> str:
    """把模型数据注入 HTML 模板，返回完整 HTML 文本。"""
    path = tmpl or template_path()
    with open(path, "r", encoding="utf-8") as f:
        html = f.read()

    if DATA_PLACEHOLDER not in html:
        raise RuntimeError(f"模板缺少数据占位符：{DATA_PLACEHOLDER}")

    # 安全注入：转义 < 避免 </script> 截断
    payload = json.dumps(model, ensure_ascii=False).replace("<", "\\u003c")
    html = html.replace(DATA_PLACEHOLDER, payload)
    html = html.replace("__GANTT_TITLE__",
                        title or model.get("project", {}).get("name", "项目进度计划"))
    html = html.replace("__LEFT_W__", str(int(left_width)))
    return html


def write_html(model: dict, out_path: str, **kw) -> str:
    html = render_html(model, **kw)
    os.makedirs(os.path.dirname(os.path.abspath(out_path)) or ".", exist_ok=True)
    with open(out_path, "w", encoding="utf-8") as f:
        f.write(html)
    return out_path
