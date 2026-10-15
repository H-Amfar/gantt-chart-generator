# 📊 横道图生成器（Gantt Chart Generator）

一个 WorkBuddy 技能（Skill），把项目任务清单一键渲染成**专业横道图（甘特图）**。
支持日期自适应时间轴、关键节点（里程碑）、任务依赖自动排程、关键路径分析，
对标 Microsoft Project 的核心排程能力。

![icon](_icon.png)

---

## ✨ 核心功能

| 能力 | 说明 |
|---|---|
| **日期自适应** | 时间轴按任务起止自动计算，日 / 周 / 月 / 季自动切换 |
| **关键节点** | 里程碑以菱形标注，可作为依赖锚点 |
| **依赖排程** | FS / SS / FF / SF 四种关系 + 滞后量，按工作日历跳过周末、支持自定义节假日 |
| **关键路径** | 标准 CPM 前推 / 后推，自动算总浮动时间并红色高亮 |
| **交互编辑** | 生成的 HTML 里双击任务条即可改名称 / 日期 / 工期 / 进度 / 负责人 / 依赖，并能增删任务 |
| **拖拽改期** | 拖动任务条改日期、拖右端改工期，松手后后继任务自动顺延 |
| **双形态输出** | 交互式 HTML（离线可用、可缩放、可导出 PNG）+ 静态 PNG（插入 Word / PPT） |
| **多源输入** | 对话描述、CSV、Excel(.xlsx)、JSON 均可，中英文列名通吃 |

---

## 🚀 安装到 WorkBuddy

1. 下载本仓库（ZIP 或 `git clone`）。
2. 把整个文件夹放到 WorkBuddy 的技能目录：
   - 用户级：`C:/Users/<你的用户名>/.workbuddy/skills/gantt-chart-generator/`
   - 或项目级：`<你的项目>/.workbuddy/skills/gantt-chart-generator/`
3. 重启 WorkBuddy（或刷新技能列表）。
4. 在对话里直接说「帮我做一个项目进度横道图……」即可触发。

---

## 🛠 命令行用法

依赖受管 Python 环境（已装 `matplotlib`、`openpyxl`）：

```bash
PY="C:/Users/<用户名>/.workbuddy/binaries/python/envs/default/Scripts/python.exe"
SKILL="<本仓库>/scripts"

# 仅出交互式 HTML
"$PY" "$SKILL/build.py" tasks.csv -o 进度计划.html --name "项目名称"

# 同时导出 PNG
"$PY" "$SKILL/build.py" tasks.csv -o 进度计划.html --png 进度计划.png \
     --name "项目名称" --dpi 200

# 只给工期 + 前后置关系，由引擎自动排程
"$PY" "$SKILL/build.py" tasks.csv -o 进度计划.html --name "项目名称" --start 2026-03-02
```

### 数据格式（CSV 表头示例，中英文列名都支持）

| 任务名称 | 开始 | 结束 | 工期 | 进度% | 前置任务 | 类型 |
|---|---|---|---|---|---|---|
| 图纸设计 | 2026-03-02 | 2026-03-13 | | 100 | | task |
| 土建施工 | | | 30 | 0 | 图纸设计 | task |
| 主体封顶 | | | | 0 | 土建施工 | milestone |

- 留空「开始 / 结束」→ 引擎按依赖自动排程；
- `类型` 填 `milestone` 即关键节点（菱形）；
- `前置任务` 填前驱的 ID 或名称，可用 `FS+2` 表示滞后 2 天（FS / SS / FF / SF 均可）。

更完整的字段说明见 [`references/01-数据格式说明.md`](references/01-数据格式说明.md)。

---

## 📁 目录结构

```
gantt-chart-generator/
├── SKILL.md                  # 技能清单（触发词、用法、对照表）
├── README.md                 # 本文件
├── _icon.png                 # 技能图标
├── assets/
│   └── gantt_template.html   # 交互式渲染引擎（自包含、离线可用）
├── scripts/
│   ├── gantt_core.py         # 数据解析 / 依赖排程 / CPM 关键路径
│   ├── build.py              # 命令行入口
│   ├── render_html.py        # HTML 渲染
│   ├── render_png.py         # PNG 静态图渲染
│   ├── sample_project.csv    # 示例：完整日期项目
│   └── sample_auto.csv       # 示例：依赖自动排程
└── references/
    ├── 01-数据格式说明.md
    ├── 02-Project功能对照.md
    └── 03-使用示例.md
```

---

## 📄 许可

MIT —— 可自由使用、修改、再分发。
