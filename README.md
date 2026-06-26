# qwen-question-gpt-audit

> 用 qwen-plus 批量重出面试题与参考答案，再由 GPT 逐字依据审题提示词审核，只输出通过的条目。

这是一个面向「面试题质量把控」的工作流技能（Skill）。它把**生成**和**审核**两件事拆开：`qwen-plus` 只负责按行级提示词生成内容，GPT 代理负责编排、逐字审核、迭代重出，确保最终落表的每一道题都符合既定的审题规则。

## 它解决什么问题

面试题库在迭代时常遇到两类麻烦：

- 直接让大模型「自己生成自己把关」，知识性错误（如把失效状态答成正常状态）、编造的固定阈值、中英文混杂等问题很难被稳定拦下。
- 审核标准容易在多轮传话中被改写、缩写或扩展，导致每个人/每次审得不一样。

本技能用一条明确的流水线来应对：

```
Excel 行（提示词 + 变量）
        │
        ▼
  qwen-plus 生成        ← 只生成，绝不自评
        │
        ▼
  GPT 逐字审核          ← 唯一标准：prompts/审题提示词.txt
        │
   ┌────┴────┐
 通过        未通过 → 追加针对性约束 → 重新生成 → 重新审核
   │
   ▼
 最终 Excel（审核状态=通过）
```

> [!IMPORTANT]
> 审核标准以 `prompts/审题提示词.txt` 为**唯一来源**，必须逐字套用，不得改写、缩写或扩展。脚本只能做量化预筛，绝不替代 GPT 审核，也绝不自动判「通过」。

## 仓库结构

```
qwen-question-gpt-audit/
├── SKILL.md                    # 技能说明：用途、编排规则、输入形态、输出格式
├── prompts/
│   └── 审题提示词.txt          # 审核标准（唯一来源，逐字套用）
└── scripts/
    ├── fill_and_call.py        # 填变量 + 调用 qwen-plus（只生成）
    └── write_final_xlsx.py     # 把已通过条目落成最终 Excel（只落表）
```

## 快速开始

### 前置条件

- Python 3.9+
- `openpyxl`（读写 Excel）
- 一个可用的千问 `qwen-plus` API key

```bash
pip install openpyxl
export DASHSCOPE_API_KEY=你的千问key
```

> [!WARNING]
> 不要把 API key 写死进脚本或输出文件。优先用环境变量 `DASHSCOPE_API_KEY`；若 key 曾在对话中粘贴，建议任务后轮换。

### 三步走

**1. 生成** —— 用 `fill_and_call.py` 读取 Excel、填充提示词变量、调用 qwen-plus，产出待审 JSON：

```bash
# 出题 + 出答案（自动识别含「出题」「出答案」两个 sheet 的工作簿）
python3 scripts/fill_and_call.py 需要重新出题.xlsx -o out.json

# 仅重出答案（题干保持不变）；行内提示词为空时回退到模板文件
python3 scripts/fill_and_call.py 重新出答案.xlsx -o out.json \
  --answer-prompt-file 出答案提示词.txt
```

**2. 审核** —— GPT 逐字依据 `prompts/审题提示词.txt`（先把 `{{岗位名称}}` 替换为实际岗位）审核 `out.json` 的每一条。未通过的，按失败原因写一段针对性约束，再带约束重出：

```bash
python3 scripts/fill_and_call.py 需要重新出题.xlsx -o out.json \
  --extra-q-file q_约束.txt --extra-a-file a_约束.txt
```

**3. 落表** —— 当所有条目都被审核判为「肯定无」（通过）后，用 `write_final_xlsx.py` 写出最终 Excel：

```bash
python3 scripts/write_final_xlsx.py audited.json -o 最终_GPT审核通过.xlsx
```

## 输入形态

工作流自动适配三种常见结构：

| 结构 | 触发条件 | 行为 |
| --- | --- | --- |
| `question` | 工作簿含 `出题` + `出答案` 两个 sheet | 用 `出题` 提示词出题，再用 `出答案` 提示词出答案 |
| `answers` | 单 sheet 且含 `题目`/`问题` 列 | 题干不变，只重出三层参考答案 |
| `questions` | 单 sheet 且无题目列 | 按每行 `提示词` 各出一题 |

`fill_and_call.py` 默认 `--mode auto` 自动判断，也可显式指定。

## 脚本参考

### `scripts/fill_and_call.py`

只生成，不审核、不判定通过。逐字用行内字段替换提示词中的 `{{...}}` 占位符后调用 qwen-plus，对瞬时网络错误自动重试。

| 参数 | 说明 |
| --- | --- |
| `excel` | 输入 Excel 路径（位置参数） |
| `-o, --out` | 输出 JSON 路径（供 GPT 审核） |
| `--mode` | `auto`（默认）/ `question` / `answers` / `questions` |
| `--answer-prompt-file` | `answers` 结构中行内 `提示词` 为空时使用的出答案模板 |
| `--extra-q-file` | 出题阶段追加的针对性约束文件（迭代重出用） |
| `--extra-a-file` | 出答案阶段追加的针对性约束文件（迭代重出用） |

### `scripts/write_final_xlsx.py`

只把**已通过审核**的条目序列化落表，不做审核。会校验 `题目` 与三层答案均非空，将 `审核状态` 置为 `通过`，并按固定列序输出。

| 参数 | 说明 |
| --- | --- |
| `json` | 输入 JSON：数组、`{"items":[...]}` 或 `{"出题":{...},"出答案#2":{...}}` 键值字典 |
| `-o, --out` | 输出 Excel 路径 |
| `--sheet` | 输出 sheet 名（默认 `已确认题目答案`） |
| `--passed-only` | 只写被标记为通过/肯定无的条目 |
| `--keep-prompt` | 保留条目中的 `提示词` 列（默认清空） |

> [!NOTE]
> 若任一条目缺少必填字段，`write_final_xlsx.py` 会中止并以非零码退出，避免残缺条目混入终稿。

## 提示词变量映射

| 模板变量 | Excel 字段 |
| --- | --- |
| `{{industry}}` | `行业` |
| `{{job}}` | `岗位` / `岗位名称` |
| `{{scene}}` | `招聘场景` / `场景` |
| `{{category}}` | `技能分类` |
| `{{skill}}` | `技能` |
| `{{knowledge}}` | `知识点` |
| `{{level}}` | `难度` |
| `{{JD}}` | `JD` |
| `{{Q}}` | 当前或新生成的题目 |
| `{{num}}` | 通常为 `1` |

替换前会把写错的 `{{{industry}}` 归一化为 `{{industry}}`。

## 审核标准说明

`prompts/审题提示词.txt` 定义了完整的审核规则，包括：

- **题干错误** `A1–A4`：知识性错误、场景偏门、偏离知识点、含不必要英文或无法口头作答的要求。
- **答案错误** `B1–B6`：知识性/逻辑错误、方案过时偏门、跑题、第三层异常不合理、中英混杂堆叠、岗位口径不符。
- **重复** `C1–C2`：题干完全相同，或不同表述但作答要点高度雷同。
- **风险标注**：`肯定无` / `疑似无` / `疑似有` / `肯定有`——除「肯定无」外均判定删除。

GPT 审核时只替换 `{{岗位名称}}`，其余规则一字不改地套用。

## 输出格式

最终 Excel 采用固定列序：

`招聘场景 · 行业 · 岗位 · JD · 技能分类 · 技能 · 知识点 · 难度 · 题目 · 审核状态 · 参考答案-第一层 · 参考答案-第二层 · 参考答案-第三层 · 提示词`

所有最终行 `审核状态=通过`，`提示词` 列默认留空。
