# Fill Transport Detail Sheet

将滴滴/哈啰出行行程单自动填入交通明细表 Excel 模板。

## 功能

- 解析滴滴出行行程报销单 PDF（快车/特惠快车/专车/优享/出租车/顺风车/惊喜特价）
- 解析哈啰出行粘贴表格（tab 或空格分隔）
- 自动识别表头行、保留模板样式、按时间排序
- 支持高速费行、日期对应工作内容映射
- 金额校验，行数校验

## 文件说明

```
.
├── SKILL.md                          # 技能说明书（AI 触发用）
├── README.md                         # 本文件
├── package.json                      # 依赖声明
├── .gitignore
├── agents/openai.yaml                # Claude Code agent 配置
├── 交通明细表.xlsx                   # 空白模板
└── scripts/
    └── fill_transport_sheet.js       # 核心填充脚本
```

---

## 在 Claude Code 中使用

### 安装

```bash
# 克隆到用户级技能目录（所有项目可用）
git clone git@github.com:gaoyuaneos-byte/fill-transport-detail-sheet-codex-skills.git \
  ~/.claude/skills/fill-transport-detail-sheet

# 安装依赖
cd ~/.claude/skills/fill-transport-detail-sheet/scripts
npm install
```

也可以放到单个项目的 `.claude/skills/` 目录下。

### 使用

把行程 PDF 和模板 `交通明细表.xlsx` 放在工作目录下，直接对话：

> "把当前目录的滴滴行程单填进交通明细表"

Claude Code 会自动识别技能并执行。也可以直接跑脚本：

```powershell
node ~/.claude/skills/fill-transport-detail-sheet/scripts/fill_transport_sheet.js `
  --source "交通明细表.xlsx" `
  --target "交通明细表_已填.xlsx" `
  --project-no "PAEEXXXX" `
  --project-name "XX项目" `
  --clear-existing
```

---

## 在 OpenClaw 中使用

### 安装

```bash
# 克隆到 OpenClaw 技能目录
git clone git@github.com:gaoyuaneos-byte/fill-transport-detail-sheet-codex-skills.git \
  ~/.openclaw/skills/fill-transport-detail-sheet

# 安装依赖
cd ~/.openclaw/skills/fill-transport-detail-sheet/scripts
npm install

# 重启 gateway 使技能生效
openclaw gateway restart
```

### 使用

把行程 PDF 和模板放在工作目录下，直接对话：

> "把滴滴行程单填进交通明细表"

或直接跑脚本：

```powershell
node ~/.openclaw/skills/fill-transport-detail-sheet/scripts/fill_transport_sheet.js `
  --source "交通明细表.xlsx" `
  --target "交通明细表_已填.xlsx" `
  --project-no "PAEEXXXX" `
  --project-name "XX项目" `
  --clear-existing
```

> **注意**：仓库中的 `agents/openai.yaml` 是 Claude Code 专用的 agent 配置文件，OpenClaw 会自动忽略，不影响使用。

---

## CLI 参数

| 参数 | 说明 | 默认值 |
|------|------|--------|
| `--source` | 模板文件路径 | `交通明细表.xlsx` |
| `--target` | 输出文件路径 | `交通明细表_已填.xlsx` |
| `--project-no` | 项目号（**必填**） | — |
| `--project-name` | 项目名称（**必填**） | — |
| `--pdf-dir` | 行程 PDF 目录 | 当前目录 |
| `--source-format` | 强制解析器：`didi` / `hellobike` / `auto` | `auto` |
| `--trips-json` | 手动行程 JSON 文件 | — |
| `--work-json` | 日期→工作内容映射 JSON | — |
| `--clear-existing` | 清除旧数据行后填充 | `false` |
| `--blank-highway-notes` | 高速费行不填工作内容 | `false` |

## 行程 JSON 格式

手动提供行程数据时使用：

```json
[
  {
    "dateTime": "2026-01-21 10:27",
    "city": "XX市",
    "start": "XX区|XX地名",
    "end": "XX县|XX站-进站口",
    "amount": 39.10,
    "feeType": "行程费"
  }
]
```

高速费行使用 `"feeType": "高速费"`。

## 依赖

```bash
npm install pdf-parse xlsx
```
