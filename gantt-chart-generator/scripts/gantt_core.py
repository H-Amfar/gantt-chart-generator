#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
gantt_core.py —— 横道图（甘特图）核心引擎

职责：
  1. 读取任务数据（CSV / XLSX / JSON）
  2. 归一化字段（兼容中英文列名、多种日期写法）
  3. 自动排程：按依赖关系（FS/SS/FF/SF）+ 滞后量 + 工作日历推导日期
  4. CPM 关键路径计算：ES / EF / LS / LF / 总浮动时间
  5. 输出统一的 JSON 模型，供 HTML 渲染与 PNG 渲染共用

本模块不产生任何可视化产物，只做「数据 -> 计算 -> 结构化模型」。

用法（被 build_html.py / build_png.py 调用）：
    from gantt_core import load_tasks, build_model
    model = build_model(load_tasks("tasks.csv"))
"""

from __future__ import annotations

import csv
import json
import math
import os
import re
from datetime import date, datetime, timedelta
from typing import Any, Dict, Iterable, List, Optional, Tuple

# --------------------------------------------------------------------------
# 常量
# --------------------------------------------------------------------------

DAY_MS = 86400

# 任务类型
TYPE_TASK = "task"
TYPE_SUMMARY = "summary"      # 摘要/阶段任务（自动汇总子任务）
TYPE_MILESTONE = "milestone"  # 里程碑（关键节点，零工期菱形）

# 依赖类型（与 Microsoft Project 一致）
DEP_TYPES = ("FS", "SS", "FF", "SF")
DEP_LABEL = {
    "FS": "完成-开始",
    "SS": "开始-开始",
    "FF": "完成-完成",
    "SF": "开始-完成",
}

# 列名别名映射（中文优先，兼容英文）
COLUMN_ALIASES: Dict[str, List[str]] = {
    "id":         ["id", "编号", "任务编号", "任务id", "wbs编号"],
    "wbs":        ["wbs", "wbs编码", "层级编码", "大纲号"],
    "name":       ["name", "任务名称", "任务名", "名称", "工作名称", "工序"],
    "level":      ["level", "层级", "级别", "缩进", "大纲级别"],
    "type":       ["type", "类型", "任务类型"],
    "start":      ["start", "开始日期", "开始时间", "计划开始", "起始日期", "开工日期"],
    "end":        ["end", "结束日期", "结束时间", "计划结束", "完成日期", "竣工日期", "截止日期"],
    "duration":   ["duration", "工期", "工期天数", "天数", "持续天数"],
    "progress":   ["progress", "进度", "完成度", "完成百分比", "完成率", "percent", "percentcomplete"],
    "assignee":   ["assignee", "负责人", "责任人", "执行人", "承担人", "owner"],
    "resources":  ["resources", "资源", "参与人", "人员", "团队"],
    "deps":       ["deps", "predecessors", "前置任务", "前置", "依赖", "紧前任务", "前序任务"],
    "deptype":    ["deptype", "依赖类型", "关系类型", "依赖关系"],
    "lag":        ["lag", "滞后", "滞后天数", "延隔", "延迟"],
    "color":      ["color", "颜色", "色彩"],
    "baseline_start": ["baseline_start", "baseline_start_date", "基线开始", "基准开始"],
    "baseline_end":   ["baseline_end", "baseline_end_date", "基线结束", "基准结束"],
    "notes":      ["notes", "备注", "说明", "描述", "note", "remark"],
    "milestone":  ["milestone", "里程碑", "是否里程碑"],
    "parent":     ["parent", "父任务", "上级任务", "所属阶段"],
    "priority":   ["priority", "优先级"],
}

# 默认调色板（高对比、印刷友好）
PALETTE = [
    "#3B82F6", "#10B981", "#F59E0B", "#8B5CF6", "#EC4899",
    "#06B6D4", "#EF4444", "#84CC16", "#F97316", "#6366F1",
    "#14B8A6", "#A855F7", "#EAB308", "#0EA5E9", "#22C55E",
]

# 状态色
COLOR_CRITICAL = "#DC2626"   # 关键路径
COLOR_DONE = "#059669"       # 已完成
COLOR_SUMMARY = "#334155"    # 摘要条
COLOR_MILESTONE = "#B45309"  # 里程碑
COLOR_BASELINE = "#94A3B8"   # 基线


# --------------------------------------------------------------------------
# 日期工具
# --------------------------------------------------------------------------

_DATE_PATTERNS = [
    "%Y-%m-%d", "%Y/%m/%d", "%Y.%m.%d", "%Y年%m月%d日",
    "%Y-%m-%d %H:%M:%S", "%Y/%m/%d %H:%M:%S",
    "%d/%m/%Y", "%m/%d/%Y", "%Y%m%d",
]


def parse_date(value: Any) -> Optional[date]:
    """解析各种写法的日期，失败返回 None。"""
    if value is None or value == "":
        return None
    if isinstance(value, datetime):
        return value.date()
    if isinstance(value, date):
        return value
    if isinstance(value, (int, float)):
        # Excel 序列号（1900 基准）
        if 20000 < value < 80000:
            return date(1899, 12, 30) + timedelta(days=int(value))
        return None
    text = str(value).strip()
    if not text:
        return None
    # 处理 "2026-01-01至2026-01-10" 这类区间，取前半段
    text = re.split(r"[至~到]|--", text)[0].strip()
    for fmt in _DATE_PATTERNS:
        try:
            return datetime.strptime(text, fmt).date()
        except ValueError:
            continue
    # 纯 "1月5日" 这类，补当前年
    m = re.fullmatch(r"(\d{1,2})月(\d{1,2})日?", text)
    if m:
        year = date.today().year
        return date(year, int(m.group(1)), int(m.group(2)))
    return None


def to_iso(d: Optional[date]) -> Optional[str]:
    return d.isoformat() if d else None


def is_workday(d: date, workdays: Iterable[int]) -> bool:
    """workdays 用 0=周一 ... 6=周日 表示。"""
    return d.weekday() in set(workdays)


def add_workdays(start: date, days: int, workdays: Iterable[int],
                 holidays: Iterable[date] = ()) -> date:
    """
    从 start 起，按工作日历推进 days 个工作日，返回结束日（含当日）。
    days 为工期；days<=1 时直接返回 start。
    """
    hol = set(holidays)
    wd = set(workdays)
    if days <= 1:
        return start
    cur = start
    counted = 1
    guard = 0
    while counted < days and guard < 20000:
        cur += timedelta(days=1)
        guard += 1
        if cur.weekday() in wd and cur not in hol:
            counted += 1
    return cur


def count_workdays(start: date, end: date, workdays: Iterable[int],
                   holidays: Iterable[date] = ()) -> int:
    """统计 start..end 含两端的有效工作日数。"""
    hol = set(holidays)
    wd = set(workdays)
    if end < start:
        return 0
    n = 0
    cur = start
    while cur <= end:
        if cur.weekday() in wd and cur not in hol:
            n += 1
        cur += timedelta(days=1)
    return n


# --------------------------------------------------------------------------
# 读取与归一化
# --------------------------------------------------------------------------

def _norm_key(key: str) -> Optional[str]:
    """把任意列名映射到标准字段名。"""
    k = str(key).strip().lower().replace(" ", "").replace("_", "")
    for field, aliases in COLUMN_ALIASES.items():
        for a in aliases:
            if k == a.lower().replace(" ", "").replace("_", ""):
                return field
    return None


def _norm_bool(v: Any) -> bool:
    if isinstance(v, bool):
        return v
    return str(v).strip().lower() in ("1", "true", "yes", "y", "是", "✓", "√", "里程碑")


def _norm_type(v: Any, is_milestone: bool = False) -> str:
    if is_milestone:
        return TYPE_MILESTONE
    t = str(v or "").strip().lower()
    if t in ("milestone", "里程碑", "节点", "关键节点", "m"):
        return TYPE_MILESTONE
    if t in ("summary", "摘要", "阶段", "汇总", "父任务", "s", "phase"):
        return TYPE_SUMMARY
    return TYPE_TASK


def _norm_int(v: Any) -> Optional[int]:
    """归一化工期为整数天数（支持 "5"、"5天"、"5d"），失败返回 None。"""
    if v is None or v == "":
        return None
    try:
        return int(round(float(str(v).replace("天", "").replace("d", "").strip())))
    except (TypeError, ValueError):
        return None


def _norm_progress(v: Any) -> float:
    """归一化进度到 0-100。"""
    if v is None or v == "":
        return 0.0
    if isinstance(v, str) and v.strip().endswith("%"):
        try:
            return max(0.0, min(100.0, float(v.strip().rstrip("%"))))
        except ValueError:
            return 0.0
    try:
        f = float(v)
    except (TypeError, ValueError):
        return 0.0
    if f <= 1.0 and f > 0:   # 0.85 之类的比例值
        f *= 100
    return max(0.0, min(100.0, f))


def _parse_deps(raw: Any) -> List[Dict[str, Any]]:
    """
    解析前置任务串。支持：
      "T1,T2"                 默认 FS
      "T1:SS:2"               指定类型与滞后
      "T1 FS 2" / "T1(FS,2)"
      "1.2.3"                 用 WBS 引用
    """
    if raw is None or str(raw).strip() == "":
        return []
    text = str(raw).replace("；", ";").replace("，", ",").replace("、", ",")
    parts = [p.strip() for p in re.split(r"[;,]", text) if p.strip()]
    out: List[Dict[str, Any]] = []
    for p in parts:
        dep_type = "FS"
        lag = 0
        tokens = re.split(r"[:\s\(\)\[\],，]+", p)
        tokens = [t for t in tokens if t]
        if not tokens:
            continue
        ref = tokens[0]
        for t in tokens[1:]:
            up = t.upper()
            if up in DEP_TYPES:
                dep_type = up
            else:
                m = re.fullmatch(r"[+-]?\d+(\.\d+)?[d天]?", t)
                if m:
                    lag = int(round(float(re.sub(r"[d天]", "", t))))
        out.append({"ref": ref, "type": dep_type, "lag": lag})
    return out


def read_csv(path: str) -> List[Dict[str, Any]]:
    with open(path, "r", encoding="utf-8-sig", newline="") as f:
        sample = f.read(16384)
        f.seek(0)
        delim = ","
        try:
            delim = csv.Sniffer().sniff(sample, delimiters=",;\t").delimiter
        except csv.Error:
            # 嗅探失败时按出现次数挑最可能的分隔符
            counts = {d: sample.count(d) for d in (",", ";", "\t")}
            best = max(counts, key=lambda k: counts[k])
            if counts[best] > 0:
                delim = best
        reader = csv.DictReader(f, delimiter=delim)
        return [dict(row) for row in reader]


def read_xlsx(path: str) -> List[Dict[str, Any]]:
    try:
        from openpyxl import load_workbook
    except ImportError as e:  # pragma: no cover
        raise RuntimeError("读取 xlsx 需要 openpyxl：pip install openpyxl") from e
    wb = load_workbook(path, data_only=True)
    ws = wb.active
    rows = list(ws.iter_rows(values_only=True))
    if not rows:
        return []
    header = [str(h).strip() if h is not None else f"col{i}" for i, h in enumerate(rows[0])]
    out = []
    for r in rows[1:]:
        if all(c is None or str(c).strip() == "" for c in r):
            continue
        out.append({header[i]: r[i] for i in range(min(len(header), len(r)))})
    return out


def load_tasks(path: str) -> List[Dict[str, Any]]:
    """从 CSV / XLSX / JSON 读取原始任务行。"""
    ext = os.path.splitext(path)[1].lower()
    if ext in (".xlsx", ".xlsm"):
        return read_xlsx(path)
    if ext == ".json":
        with open(path, "r", encoding="utf-8") as f:
            data = json.load(f)
        if isinstance(data, dict):
            for key in ("tasks", "Tasks", "任务", "items", "行"):
                if key in data:
                    return data[key]
            return [data]
        return data
    return read_csv(path)


def normalize_rows(rows: List[Dict[str, Any]]) -> List[Dict[str, Any]]:
    """把原始行映射为标准字段。"""
    tasks: List[Dict[str, Any]] = []
    for i, raw in enumerate(rows):
        item: Dict[str, Any] = {}
        for k, v in raw.items():
            field = _norm_key(k)
            if field and (field not in item or item.get(field) in (None, "")):
                item[field] = v
        if not item.get("name"):
            continue
        is_ms = _norm_bool(item.get("milestone", False))
        t: Dict[str, Any] = {
            "id": str(item.get("id") or "").strip(),
            "wbs": str(item.get("wbs") or "").strip(),
            "name": str(item["name"]).strip(),
            "level": item.get("level"),
            "type": _norm_type(item.get("type"), is_ms),
            "start": parse_date(item.get("start")),
            "end": parse_date(item.get("end")),
            "duration": _norm_int(item.get("duration")),
            "progress": _norm_progress(item.get("progress")),
            "assignee": str(item.get("assignee") or "").strip(),
            "resources": str(item.get("resources") or "").strip(),
            "deps_raw": _parse_deps(item.get("deps")),
            "deptype": str(item.get("deptype") or "").strip().upper(),
            "lag": item.get("lag"),
            "color": str(item.get("color") or "").strip(),
            "baseline_start": parse_date(item.get("baseline_start")),
            "baseline_end": parse_date(item.get("baseline_end")),
            "notes": str(item.get("notes") or "").strip(),
            "priority": str(item.get("priority") or "").strip(),
            "row": i,
        }
        # 编号兜底
        if not t["id"]:
            t["id"] = f"T{i + 1}"
        # 层级推导：优先 WBS 点号，其次显式 level
        if t["wbs"]:
            t["level"] = max(0, t["wbs"].count("."))
        elif t["level"] not in (None, ""):
            try:
                t["level"] = int(float(t["level"]))
            except (TypeError, ValueError):
                t["level"] = 0
        else:
            t["level"] = None  # 稍后按名称缩进或默认 0
        if t["type"] == TYPE_MILESTONE:
            t["duration"] = 0
        tasks.append(t)
    return tasks


# --------------------------------------------------------------------------
# 依赖解析与拓扑排序
# --------------------------------------------------------------------------

def resolve_deps(tasks: List[Dict[str, Any]]) -> None:
    """把 deps 里的 ref（id 或 WBS）解析为任务 id。"""
    by_id = {t["id"]: t for t in tasks}
    by_wbs = {t["wbs"]: t for t in tasks if t["wbs"]}
    by_name = {t["name"]: t for t in tasks}
    for t in tasks:
        deps: List[Dict[str, Any]] = []
        for d in t["deps_raw"]:
            ref = d["ref"]
            target = by_id.get(ref) or by_wbs.get(ref) or by_name.get(ref)
            if target is None:
                continue
            dep_type = d["type"]
            lag = d["lag"]
            # 若未在串内指定类型，用独立列兜底
            if dep_type == "FS" and t.get("deptype") in DEP_TYPES:
                dep_type = t["deptype"]
            if lag == 0 and t.get("lag") not in (None, ""):
                try:
                    lag = float(t["lag"])
                except (TypeError, ValueError):
                    lag = 0
            deps.append({"id": target["id"], "type": dep_type, "lag": lag})
        t["deps"] = deps


def topo_order(tasks: List[Dict[str, Any]]) -> List[Dict[str, Any]]:
    """拓扑排序；存在环时把剩余任务追加在末尾。"""
    ids = {t["id"] for t in tasks}
    indeg = {t["id"]: 0 for t in tasks}
    succ: Dict[str, List[str]] = {t["id"]: [] for t in tasks}
    for t in tasks:
        for d in t.get("deps", []):
            if d["id"] in ids and d["id"] != t["id"]:
                indeg[t["id"]] += 1
                succ[d["id"]].append(t["id"])
    queue = [t for t in tasks if indeg[t["id"]] == 0]
    out: List[Dict[str, Any]] = []
    seen = set()
    while queue:
        cur = queue.pop(0)
        if cur["id"] in seen:
            continue
        seen.add(cur["id"])
        out.append(cur)
        for s in succ[cur["id"]]:
            indeg[s] -= 1
            if indeg[s] <= 0:
                nxt = next((x for x in tasks if x["id"] == s), None)
                if nxt:
                    queue.append(nxt)
    for t in tasks:
        if t["id"] not in seen:
            out.append(t)
    return out


# --------------------------------------------------------------------------
# 排程
# --------------------------------------------------------------------------

def auto_schedule(tasks: List[Dict[str, Any]], project_start: date,
                  workdays: Iterable[int] = (0, 1, 2, 3, 4),
                  holidays: Iterable[date] = (),
                  use_calendar: bool = False) -> None:
    """
    自动排程：
      - 有明确开始日期的任务保持不动（人工锁定）
      - 无开始日期的任务，按前置依赖推导开始时间
      - 无工期时按 结束-开始 反推，或默认 1 天
    """
    by_id = {t["id"]: t for t in tasks}
    ordered = topo_order(tasks)
    for t in ordered:
        raw_dur = t.get("duration")
        dur: Optional[int] = None
        if raw_dur not in (None, ""):
            try:
                dur = int(round(float(str(raw_dur).replace("天", "").replace("d", ""))))
            except (TypeError, ValueError):
                dur = None
        if t["type"] == TYPE_MILESTONE:
            dur = 0
        # 补齐 start/end/duration
        if t["start"] and t["end"]:
            if dur is None:
                dur = (count_workdays(t["start"], t["end"], workdays, holidays)
                       if use_calendar else (t["end"] - t["start"]).days + 1)
            if t["type"] == TYPE_MILESTONE:
                t["end"] = t["start"]
                dur = 0
        elif t["start"] and dur is not None:
            if dur <= 0:
                t["end"] = t["start"]
            elif use_calendar:
                t["end"] = add_workdays(t["start"], dur, workdays, holidays)
            else:
                t["end"] = t["start"] + timedelta(days=dur - 1)
        elif t["end"] and dur is not None:
            if use_calendar:
                t["start"] = _back_workdays(t["end"], dur, workdays, holidays)
            else:
                t["start"] = t["end"] - timedelta(days=dur - 1)
        else:
            # 依赖驱动
            start_candidate = project_start
            if t.get("deps"):
                cands = []
                for d in t["deps"]:
                    pred = by_id.get(d["id"])
                    if not pred or not pred.get("start") or not pred.get("end"):
                        continue
                    lag = int(d.get("lag") or 0)
                    if d["type"] == "FS":
                        # 前驱完成后开始
                        base = pred["end"] + timedelta(days=1 + lag)
                    elif d["type"] == "SS":
                        # 与前驱同时开始
                        base = pred["start"] + timedelta(days=lag)
                    elif d["type"] == "FF":
                        # 与前驱同时完成
                        base = pred["end"] + timedelta(days=lag)
                        if dur is not None and dur > 0:
                            base = base - timedelta(days=dur - 1)
                    else:  # SF：前驱开始后才能完成
                        base = pred["start"] + timedelta(days=lag)
                        if dur is not None and dur > 0:
                            base = base - timedelta(days=dur - 1)
                    if use_calendar and not is_workday(base, workdays):
                        base = _next_workday(base, workdays, holidays)
                    cands.append(base)
                if cands:
                    start_candidate = max(cands)
            t["start"] = start_candidate
            if dur is None:
                dur = 0 if t["type"] == TYPE_MILESTONE else 1
            if dur <= 0:
                t["end"] = t["start"]
            elif use_calendar:
                t["end"] = add_workdays(t["start"], dur, workdays, holidays)
            else:
                t["end"] = t["start"] + timedelta(days=dur - 1)
        t["duration"] = dur if dur is not None else 1
        if t["end"] is None:
            t["end"] = t["start"]


def _next_workday(d: date, workdays: Iterable[int], holidays: Iterable[date]) -> date:
    hol = set(holidays)
    wd = set(workdays)
    guard = 0
    while (d.weekday() not in wd or d in hol) and guard < 400:
        d += timedelta(days=1)
        guard += 1
    return d


def _back_workdays(end: date, days: int, workdays: Iterable[int],
                   holidays: Iterable[date]) -> date:
    hol = set(holidays)
    wd = set(workdays)
    if days <= 1:
        return end
    cur = end
    counted = 1
    guard = 0
    while counted < days and guard < 20000:
        cur -= timedelta(days=1)
        guard += 1
        if cur.weekday() in wd and cur not in hol:
            counted += 1
    return cur


class CalendarIndex:
    """
    工作日序数索引。

    把日期映射为单调递增的整数序数（跳过非工作日），使 CPM 的前推/后推
    可以在「工作日」空间内做纯整数运算，从而正确处理跨周末、节假日的情形。

    自然日模式等价于「每周 7 天均为工作日、无节假日」。
    """

    def __init__(self, anchors: Iterable[date],
                 workdays: Iterable[int] = (0, 1, 2, 3, 4),
                 holidays: Iterable[date] = ()):
        self.wd = set(workdays)
        self.hol = set(holidays)
        ds = list(anchors)
        lo = (min(ds) if ds else date.today()) - timedelta(days=420)
        hi = (max(ds) if ds else date.today()) + timedelta(days=420)
        self.forward: Dict[date, int] = {}
        self.backward: Dict[int, date] = {}
        idx = 0
        cur = lo
        while cur <= hi:
            if cur.weekday() in self.wd and cur not in self.hol:
                self.forward[cur] = idx
                self.backward[idx] = cur
                idx += 1
            cur += timedelta(days=1)
        self._last = idx
        # 非工作日：映射到下一个工作日（找不到则用最后一个）
        cur = lo
        while cur <= hi:
            if cur not in self.forward:
                nxt = cur
                guard = 0
                while nxt not in self.forward and guard < 60:
                    nxt += timedelta(days=1)
                    guard += 1
                self.forward[cur] = self.forward.get(nxt, max(0, self._last - 1))
            cur += timedelta(days=1)

    def i(self, d: date) -> int:
        """日期 -> 序数。"""
        if d in self.forward:
            return self.forward[d]
        # 超出预生成范围时兜底
        lo = min(self.backward) if self.backward else 0
        hi = max(self.backward) if self.backward else 0
        if d < min(self.backward.values()):
            return lo - (min(self.backward.values()) - d).days
        return hi + (d - max(self.backward.values())).days

    def d(self, i: int) -> date:
        """序数 -> 日期（夹取到有效范围）。"""
        if i in self.backward:
            return self.backward[i]
        keys = self.backward
        if i < min(keys):
            base = keys[min(keys)]
            return base + timedelta(days=(min(keys) - i))
        base = keys[max(keys)]
        return base + timedelta(days=(i - max(keys)))

    def actual(self, i: int) -> int:
        """把任意序数夹取到有效工作日范围内。"""
        lo, hi = min(self.backward), max(self.backward)
        return max(lo, min(hi, i))


def summarize_parents(tasks: List[Dict[str, Any]],
                      workdays: Iterable[int] = (0, 1, 2, 3, 4),
                      use_calendar: bool = True) -> None:
    """
    摘要（父）任务自动按子任务范围汇总起止日期、工期与加权进度。
    依据任务出现顺序与层级缩进推断父子关系。
    """
    # 用栈按出现顺序推断包含关系
    stack: List[Dict[str, Any]] = []
    children: Dict[str, List[Dict[str, Any]]] = {}
    for t in tasks:
        lvl = t.get("level")
        if lvl is None:
            lvl = 0
        while stack and (stack[-1].get("_lvl") or 0) >= lvl:
            stack.pop()
        if stack:
            children.setdefault(stack[-1]["id"], []).append(t)
        t["_lvl"] = lvl
        stack.append(t)
    for t in tasks:
        kids = children.get(t["id"])
        if not kids:
            if t["type"] == TYPE_SUMMARY:
                t["type"] = TYPE_TASK
            continue
        t["type"] = TYPE_SUMMARY
        starts = [k["start"] for k in kids if k.get("start")]
        ends = [k["end"] for k in kids if k.get("end")]
        if starts and ends:
            t["start"] = min(starts)
            t["end"] = max(ends)
            if use_calendar:
                t["duration"] = max(1, count_workdays(t["start"], t["end"], workdays))
            else:
                t["duration"] = (t["end"] - t["start"]).days + 1
        progs = [k.get("progress", 0) for k in kids]
        if progs:
            weights = [max(1, int(k.get("duration") or 1)) for k in kids]
            t["progress"] = sum(p * w for p, w in zip(progs, weights)) / sum(weights)
        t["_has_children"] = True


# --------------------------------------------------------------------------
# CPM 关键路径
# --------------------------------------------------------------------------

def compute_cpm(tasks: List[Dict[str, Any]],
                cal: Optional["CalendarIndex"] = None) -> Dict[str, Any]:
    """
    前推 / 后推计算关键路径（Critical Path Method）。

    约束模型（与 Microsoft Project 一致）：
      FS  前驱完成后才能开始：  succ.ES >= pred.EF + lag + 1
      SS  与前驱同时开始：      succ.ES >= pred.ES + lag
      FF  与前驱同时完成：      succ.EF >= pred.EF + lag
      SF  前驱开始后才能完成：  succ.EF >= pred.ES + lag

    全部运算在 CalendarIndex 的「工作日序数」空间内进行，
    因此能正确处理跨周末与节假日的排期（使用自然日日历时退化为纯自然日算术）。

    总浮动时间（Total Float）为 0 的任务构成关键路径。

    返回：{'critical': [...], 'total_float': {...}, 'project_end': iso,
           'detail': {id: {es, ef, ls, lf, tf}}}
    """
    by_id = {t["id"]: t for t in tasks}
    ordered = topo_order(tasks)

    if cal is None:
        anchors = [x for t in tasks for x in (t.get("start"), t.get("end")) if x]
        cal = CalendarIndex(anchors)

    def dur_of(t: Dict[str, Any]) -> int:
        d = t.get("duration")
        if d is None:
            return 0 if t["type"] == TYPE_MILESTONE else 1
        return max(0, int(d))

    # -------- 正向：ES / EF（序数空间） --------
    for t in ordered:
        if t["type"] == TYPE_SUMMARY:
            continue
        es = cal.i(t["start"])
        td = dur_of(t)
        for d in t.get("deps", []):
            pred = by_id.get(d["id"])
            if not pred or pred["type"] == TYPE_SUMMARY or "ef" not in pred:
                continue
            lag = int(d.get("lag") or 0)
            if d["type"] == "FS":
                cand = pred["ef"] + 1 + lag
            elif d["type"] == "SS":
                cand = pred["es"] + lag
            elif d["type"] == "FF":
                cand = pred["ef"] + lag - max(0, td - 1)
            else:  # SF
                cand = pred["es"] + lag - max(0, td - 1)
            if cand > es:
                es = cand
        t["es"] = es
        t["ef"] = es + max(0, td - 1)

    leaves = [t for t in tasks if t["type"] != TYPE_SUMMARY]
    if not leaves:
        return {"critical": [], "total_float": {}, "project_end": None, "detail": {}}

    proj_end = max(t["ef"] for t in leaves)

    # 后继索引
    succ: Dict[str, List[str]] = {t["id"]: [] for t in tasks}
    for t in tasks:
        for d in t.get("deps", []):
            if d["id"] in succ and d["id"] != t["id"]:
                succ[d["id"]].append(t["id"])

    # -------- 反向：LF / LS（序数空间） --------
    for t in reversed(ordered):
        if t["type"] == TYPE_SUMMARY:
            continue
        nxt = [by_id[s] for s in succ.get(t["id"], [])
               if s in by_id and by_id[s]["type"] != TYPE_SUMMARY
               and by_id[s].get("lf") is not None]
        if not nxt:
            lf = proj_end
        else:
            cands: List[int] = []
            td = dur_of(t)
            for s in nxt:
                dep = next((d for d in s.get("deps", []) if d["id"] == t["id"]), None)
                if dep is None:
                    continue
                lag = int(dep.get("lag") or 0)
                if dep["type"] == "FS":
                    cands.append(s["ls"] - 1 - lag)
                elif dep["type"] == "SS":
                    cands.append(s["ls"] - lag + max(0, td - 1))
                elif dep["type"] == "FF":
                    cands.append(s["lf"] - lag)
                else:  # SF
                    cands.append(s["ls"] - lag)
            lf = min(cands) if cands else proj_end
        t["lf"] = lf
        t["ls"] = lf - max(0, dur_of(t) - 1)
        t["total_float"] = t["ls"] - t["es"]

    critical: List[str] = []
    total_float: Dict[str, int] = {}
    detail: Dict[str, Dict[str, Any]] = {}
    for t in leaves:
        tf = int(t.get("total_float", 0))
        total_float[t["id"]] = tf
        t["critical"] = (tf <= 0)
        if t["critical"]:
            critical.append(t["id"])
        detail[t["id"]] = {
            "es": to_iso(cal.d(t["es"])),
            "ef": to_iso(cal.d(t["ef"])),
            "ls": to_iso(cal.d(t["ls"])),
            "lf": to_iso(cal.d(t["lf"])),
            "totalFloat": tf,
        }
    return {
        "critical": critical,
        "total_float": total_float,
        "project_end": to_iso(cal.d(proj_end)),
        "detail": detail,
    }


# --------------------------------------------------------------------------
# 组装模型
# --------------------------------------------------------------------------

def build_model(rows: List[Dict[str, Any]],
                project_start: Optional[Any] = None,
                workdays: Iterable[int] = (0, 1, 2, 3, 4),
                holidays: Iterable[date] = (),
                use_calendar: bool = True,
                project_name: str = "项目进度计划",
                subtitle: str = "") -> Dict[str, Any]:
    """
    主入口：原始行 -> 完整模型。

    use_calendar=True（默认）：按工作日历排程与计算，自动跳过周末；
                              工期字段为「工作日天数」。
    use_calendar=False：按自然日排程与计算，工期为日历跨度天数。
    """
    tasks = normalize_rows(rows)
    if not tasks:
        raise ValueError("没有解析到任何任务，请检查输入文件的列名与内容。")

    resolve_deps(tasks)
    summarize_parents(tasks, workdays, use_calendar)

    # 项目起点
    starts = [t["start"] for t in tasks if t.get("start")]
    ps = parse_date(project_start) or (min(starts) if starts else date.today())

    # 依赖排程：已有明确日期的任务保持不动，只补齐工期/结束日；
    # 缺少日期的任务按前置依赖自动推导
    auto_schedule(tasks, ps, workdays, holidays, use_calendar)

    summarize_parents(tasks, workdays, use_calendar)   # 再汇总一次，吸收子任务更新后的日期

    # 关键路径计算（在工作日序数空间，正确处理跨周末/节假日）
    if use_calendar:
        cal_wd, cal_hol = tuple(workdays), tuple(holidays)
    else:
        cal_wd, cal_hol = (0, 1, 2, 3, 4, 5, 6), ()
    anchors = [x for t in tasks for x in (t.get("start"), t.get("end")) if x]
    cal = CalendarIndex(anchors, cal_wd, cal_hol)
    cpm = compute_cpm(tasks, cal)

    # 颜色分配
    color_idx = 0
    for t in tasks:
        if not t.get("color"):
            if t["type"] == TYPE_SUMMARY:
                t["color"] = COLOR_SUMMARY
            elif t["type"] == TYPE_MILESTONE:
                t["color"] = COLOR_MILESTONE
            elif t.get("critical"):
                t["color"] = COLOR_CRITICAL
            elif t.get("progress", 0) >= 100:
                t["color"] = COLOR_DONE
            else:
                t["color"] = PALETTE[color_idx % len(PALETTE)]
                color_idx += 1

    all_starts = [t["start"] for t in tasks if t.get("start")]
    all_ends = [t["end"] for t in tasks if t.get("end")]
    p_start = min(all_starts) if all_starts else ps
    p_end = max(all_ends) if all_ends else ps

    # 输出任务（去掉内部字段）
    out_tasks = []
    for t in tasks:
        out_tasks.append({
            "id": t["id"],
            "wbs": t.get("wbs", ""),
            "name": t["name"],
            "level": t.get("level") or 0,
            "type": t["type"],
            "start": to_iso(t.get("start")),
            "end": to_iso(t.get("end")),
            "duration": t.get("duration") or 0,
            "progress": round(t.get("progress", 0), 1),
            "assignee": t.get("assignee", ""),
            "resources": t.get("resources", ""),
            "deps": [{"id": d["id"], "type": d["type"], "lag": d["lag"]}
                     for d in t.get("deps", [])],
            "color": t.get("color"),
            "baseline": ({"start": to_iso(t["baseline_start"]), "end": to_iso(t["baseline_end"])}
                         if t.get("baseline_start") and t.get("baseline_end") else None),
            "notes": t.get("notes", ""),
            "priority": t.get("priority", ""),
            "critical": bool(t.get("critical", False)),
            "total_float": t.get("total_float"),
            "has_children": bool(t.get("_has_children", False)),
        })

    return {
        "project": {
            "name": project_name,
            "subtitle": subtitle,
            "start": to_iso(p_start),
            "end": to_iso(p_end),
            "totalDays": (p_end - p_start).days + 1,
            "taskCount": len([t for t in out_tasks if t["type"] != TYPE_SUMMARY]),
            "milestoneCount": len([t for t in out_tasks if t["type"] == TYPE_MILESTONE]),
            "criticalCount": len(cpm["critical"]),
            "workdays": [int(w) for w in workdays],
            "today": date.today().isoformat(),
        },
        "tasks": out_tasks,
        "criticalPath": cpm["critical"],
        "generatedAt": datetime.now().strftime("%Y-%m-%d %H:%M"),
    }


def model_to_json(model: Dict[str, Any], path: str) -> None:
    with open(path, "w", encoding="utf-8") as f:
        json.dump(model, f, ensure_ascii=False, indent=2)


if __name__ == "__main__":  # pragma: no cover
    import sys
    if len(sys.argv) < 2:
        print("用法: python gantt_core.py <任务文件.csv|xlsx|json> [输出.json]")
        sys.exit(1)
    m = build_model(load_tasks(sys.argv[1]))
    out = sys.argv[2] if len(sys.argv) > 2 else "gantt_model.json"
    model_to_json(m, out)
    print(f"OK -> {out}  ({m['project']['taskCount']} 个任务, "
          f"{m['project']['milestoneCount']} 个里程碑, "
          f"{m['project']['criticalCount']} 个关键任务)")
