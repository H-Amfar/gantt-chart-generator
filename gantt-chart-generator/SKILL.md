---
name: gantt-chart-generator
description: 生成横道图（甘特图）与项目进度计划。当用户需要制作/生成横道图、甘特图、项目进度表、施工进度计划、排期表、里程碑计划、工程进度网络图，或需要按任务依赖关系（FS/SS/FF/SF）自动排程、标注关键节点与里程碑、计算关键路径与浮动时间、对比基线计划、导出进度图表（HTML/PNG）时使用。支持从对话描述、CSV、Excel(.xlsx)、JSON 生成，功能对标 Microsoft Project。
---

# 横道图生成器（Gantt Chart Generator）

把项目任务清单渲染成专业横道图，支持日期自适应时间轴、关键节点（里程碑）、
任务依赖与关键路径分析。产出两种形态：

- **交互式 HTML**：单文件、离线可用、可缩放、双击编辑任务、拖拽改期、可添加/删除任务、可导出 PNG
- **静态 PNG**：用于插入 Word / PPT / 汇报材料

（中文别称：甘特图、条状图、进度条图、工程进度横道图）

---

## 一、何时使用本 Skill

命中以下任一情形即应使用：

- 用户给出或描述了一批带日期的任务，希望看到进度排布 → 直接生成横道图
- 用户需要「项目进度计划 / 施工进度计划 / 实施排期 / 里程碑计划 / 交付计划」
- 用户强调「关键节点」「关键里程碑」「节点时间」→ 用里程碑（milestone）类型
- 用户要求「按依赖自动排期」「前后置关系」「工期跟着前置走」→ 用依赖排程
- 用户询问「哪些任务不能延期」「哪条路径最关键」→ 输出关键路径与总浮动
- 用户提供 Excel / CSV 任务清单，要求可视化

---

## 二、快速开始

脚本位于本 skill 的 `scripts/` 目录。默认使用受管 Python 环境：

```bash
PY="C:/Users/yt947/.workbuddy/binaries/python/envs/default/Scripts/python.exe"
SKILL="<本 skill 目录>/scripts"

# 1) 只出交互式 HTML
"$PY" "$SKILL/build.py" tasks.csv -o 进度计划.html --name "项目名称"

# 2) 同时出 HTML + PNG（PNG 需 matplotlib）
"$PY" "$SKILL/build.py" tasks.xlsx -o 进度计划.html --png 进度计划.png --name "项目名称"

# 3) 只有工期和前置关系、没有日期的任务，交给脚本自动排程
"$PY" "$SKILL/build.py" tasks.csv -o out.html --start 2026-03-02 --name "厂区建设"
```

> 若无 matplotlib，`--png` 会被自动跳过并提示，HTML 仍正常产出。
> HTML 页面右上角自带「导出 PNG」按钮，可在浏览器里一键出图。

### 依赖安装（首次使用或换机器）

```bash
"C:/Users/yt947/.workbuddy/binaries/python/versions/3.13.12/python.exe" -m venv "C:/Users/yt947/.workbuddy/binaries/python/envs/default"
"C:/Users/yt947/.workbuddy/binaries/python/envs/default/Scripts/python.exe" -m pip install \
  -i https://mirrors.aliyun.com/pypi/simple/ --trusted-host mirrors.aliyun.com matplotlib openpyxl
```

---

## 三、输入数据格式

三种输入方式，字段名中英文皆可（如 `开始日期` / `start` 等价）。

### 3.1 CSV / Excel 列定义

| 列名 | 别名 | 必填 | 说明 |
|---|---|---|---|
| `WBS` | 编号、大纲号 | 否 | 形如 `1.2.3`，自动据此推导层级与缩进 |
| `任务名称` | name、工序 | **是** | 任务/阶段/节点名称 |
| `类型` | type | 否 | `任务` / `摘要`（阶段）/ `里程碑`（关键节点） |
| `开始日期` | start、计划开始 | 条件 | 与「工期」至少给一个 |
| `结束日期` | end、完成日期 | 否 | 给了开始+工期时可省略 |
| `工期` | duration、天数 | 否 | 默认按**工作日**计（自动跳过周末），`--calendar natural` 改自然日 |
| `进度` | progress、完成率 | 否 | `0-100` 或 `0-1` 或 `85%` |
| `负责人` | assignee、责任人 | 否 | |
| `前置任务` | 依赖、紧前任务 | 否 | 见下方依赖语法 |
| `基线开始` / `基线结束` | baseline | 否 | 画对比灰条，用于「计划 vs 实际」 |
| `备注` | notes、说明 | 否 | 鼠标悬浮时显示 |

示例文件：`scripts/sample_project.csv`（完整日期）、`scripts/sample_auto.csv`（依赖自动排程）。

### 3.2 依赖语法（关键）

```
1.2                    默认 FS（完成-开始）
1.2:SS:3               开始-开始，滞后 3 天
1.2:FS:2               完成-开始，滞后 2 天
1.2;1.5                多个前置，逗号或分号分隔
```

四种关系语义（与 Microsoft Project 一致）：

| 代号 | 含义 | 约束 |
|---|---|---|
| `FS` | 完成-开始（最常用） | 后续开始 ≥ 前置结束 + 1 + 滞后 |
| `SS` | 开始-开始（搭接施工） | 后续开始 ≥ 前置开始 + 滞后 |
| `FF` | 完成-完成 | 后续结束 ≥ 前置结束 + 滞后 |
| `SF` | 开始-完成（少用） | 后续结束 ≥ 前置开始 + 滞后 |

引用方式：写 WBS 编号（`1.2`）或任务编号（`T3`）或任务名称均可。

详见 `references/01-数据格式说明.md`。

---

## 四、命令行参数

| 参数 | 说明 |
|---|---|
| `input` | 任务文件（.csv / .xlsx / .json） |
| `-o, --out` | 输出 HTML 路径 |
| `--png` | 输出 PNG 路径（需 matplotlib） |
| `--model` | 同时导出中间模型 JSON |
| `--name` / `--subtitle` | 项目名称与副标题 |
| `--start` | 项目起始日期（缺省取最早任务日期） |
| `--calendar` | `workday`（默认，跳周末）/ `natural`（自然日） |
| `--workdays` | 工作日设置，默认 `1,2,3,4,5`（1=周一） |
| `--holidays` | 额外节假日，如 `2026-10-01,2026-10-02` |
| `--dpi` | PNG 分辨率，默认 200 |
| `--no-deps` / `--no-baseline` | PNG 中隐藏依赖线 / 基线 |
| `--left-width` | HTML 左侧任务表宽度，默认 560px |

---

## 五、自动排程与关键路径

**自动排程**：任务缺少「开始日期」时，脚本按前置依赖 + 滞后量 + 工作日历推导；
已有明确日期的任务视为人工锁定，不被改动。

**关键路径（CPM）**：脚本自动做前推/后推计算，得出：

- `ES/EF`：最早开始/完成
- `LS/LF`：最晚开始/完成
- `总浮动时间`（Total Float）：不影响项目总工期的最大可延迟天数
- 浮动为 0 的任务串成**关键路径**，HTML 中红色高亮，PNG 中红色描边

**摘要任务**：其起止日期与工期由子任务自动汇总，进度按工期加权平均。

---

## 六、交互式 HTML 的用法

生成后直接用浏览器打开，支持：

| 操作 | 效果 |
|---|---|
| **双击任务条 / 左侧行 / 右侧行空白处** | **打开任务编辑窗口**：改名称、类型（普通/里程碑）、开始/结束/工期、进度、负责人、备注、前置依赖（FS/SS/FF/SF + 滞后天数），保存后自动按依赖顺延 |
| 工具栏「＋ 添加任务」 | 新建任务并设置前置依赖 |
| 编辑窗口「删除任务」 | 删除叶子任务并自动清理指向它的依赖 |
| 视图「自动/日/周/月/季」 | 时间轴粒度自适应切换 |
| 缩放 `+` `−` / 适应宽度 | 调整时间轴密度 |
| 定位今天 | 滚动到今日线位置 |
| 拖动任务条 | 左右移动调整日期 |
| 拖动任务条右端 | 改变工期（自动重算工期天数） |
| 释放鼠标后 | 后续任务按依赖关系自动顺延 |
| 「按依赖重排」 | 全量按依赖重新计算所有任务 |
| 悬浮任务条 | 查看工期、进度、负责人、前置、总浮动、基线 |
| 开关项 | 关键路径 / 基线 / 依赖线 / 周末底色 |
| 「? 操作说明」 | 查看全部交互操作指引 |
| 导出 PNG / 打印 | 输出图片或 A3 横版打印 |

> 注意：页面内编辑保存在当前浏览器会话中（刷新后恢复源数据）。需要固化修改时，把调整后的日期告诉我，由我更新 CSV/Excel 后重新生成；或直接在编辑窗口核对日期后修改源文件。

---

## 七、参考文档

- `references/01-数据格式说明.md` —— 完整字段表、依赖语法、日期写法、Excel 模板
- `references/02-Project功能对照.md` —— 与 Microsoft Project 的功能对应与差异
- `references/03-使用示例.md` —— 若干典型场景的完整命令与数据样例

## 八、脚本清单

| 文件 | 作用 |
|---|---|
| `scripts/gantt_core.py` | 数据解析、归一化、自动排程、CPM 关键路径 |
| `scripts/render_html.py` | 模型 → 交互式 HTML |
| `scripts/render_png.py` | 模型 → 静态 PNG（matplotlib） |
| `scripts/build.py` | 命令行统一入口 |
| `assets/gantt_template.html` | 前端渲染引擎模板（数据由脚本注入） |

也可单独调用核心模块：

```bash
"$PY" "$SKILL/gantt_core.py" tasks.csv model.json   # 只生成模型 JSON
```
