---
name: qwen-question-gpt-audit
description: 以 qwen-plus 按 Excel 行级提示词与变量生成面试题或参考答案，再由当前 GPT 代理逐字依据审题提示词审核并迭代，直到所有生成项通过。当用户要求调用千问/qwen-plus 重新出题、重出答案，或处理带提示词与变量的 Excel 文件、并由 GPT 负责审核时使用本技能。
disable-model-invocation: true
---

# Qwen 出题 + GPT 审核

## 用途

用本技能从 Excel 文件重新生成面试题和/或参考答案：

- `qwen-plus` 只负责按行级提示词与变量生成内容。
- 当前 GPT 代理负责编排与审核。
- GPT 审核必须逐字依据 `prompts/审题提示词.txt`，绝不改写、缩写或扩展审核规则。
- 不要把整个工作簿交给 qwen-plus 自主处理。
- 对未通过的条目，向 qwen 提示词追加针对性约束后重出，再重新审核。
- 最终输出只能包含通过 GPT 审核的条目。

## 必需的用户输入

开始前需取得以下三项：

1. `qwen-plus` 使用的千问 API key。
2. 输入文件路径：要处理的 Excel 文件。
3. 输出文件路径：最终 Excel 文件路径。

若有任一缺失，先询问再开始。

安全处理：

- 优先从环境变量 `DASHSCOPE_API_KEY` 读取 API key。
- 若用户明确授权直接使用其提供的 key，也不要把它写进长期保存的脚本或输出文件。
- 不要把 key 写入任何生成的日志、JSON、Excel 或最终汇总中。
- 若 key 是在对话中粘贴的，建议任务结束后轮换该 key。

## 支持的输入形态

本工作流支持以下常见 sheet/文件：

### 1. 需要重新出题并出答案

行内包含如下变量：

- `问题`（或原始未通过的题目）
- `技能分类`
- `技能`
- `知识点`
- `难度`
- `行业`
- `岗位`
- `招聘场景`
- `JD`
- `提示词`

若工作簿含 `出题`、`出答案` 两个独立 sheet：

- 用 `出题` sheet 的提示词生成新题目。
- 用 `出答案` sheet 的提示词为新生成的题目生成答案。

### 2. 仅需要重出答案

行内包含：

- `题目` 或 `问题`
- `技能分类`
- `技能`
- `知识点`
- `难度`
- `行业`
- `岗位`
- `招聘场景`
- `JD`
- `提示词`

保持原题干不变，只重新生成 `参考答案-第一层`、`参考答案-第二层`、`参考答案-第三层`。

## 编排规则

1. 先在本地读取 Excel，检查 sheet 与列结构。
2. 逐行处理。「填变量 + 调用 qwen-plus」这一步可使用 `scripts/fill_and_call.py`（它只生成，绝不审核）。
3. 对每一行，先替换提示词中的变量再调用 qwen-plus。
4. 只把当前行的提示词与变量发给 qwen-plus。
5. 要求 qwen-plus 返回严格 JSON。
6. 将 qwen 输出解析为结构化字段。
7. GPT 严格逐字依据 `prompts/审题提示词.txt` 审核该行（见「GPT 审核标准」）。不要改写、缩写或扩展这些规则。
8. 若 GPT 审核未通过：
   - 记录具体失败原因。
   - 仅针对该条目向 qwen 提示词追加针对性约束。
   - 重新生成并重新审核。
9. 只有当每一行都通过 GPT 审核才停止；否则报告被卡住的条目及原因。

## Qwen API 调用方式

你既可以为每个任务现写一次性脚本，也可以使用内置通用脚本 `scripts/fill_and_call.py`（推荐用于「填变量 + 调用」这一步）。

### 内置脚本：scripts/fill_and_call.py

该脚本只负责生成（绝不审核、也不标记「通过」）。它读取 Excel，逐字用行内字段填充 `{{...}}` 占位符，调用 qwen-plus，并写出一个供 GPT 审核的 JSON。

它会自动识别三种结构：

- `question`：工作簿同时含 `出题` 和 `出答案` 两个 sheet → 用 `出题` 提示词出题，再用 `出答案` 提示词出答案。
- `answers`：单个含 `题目`/`问题` 的 sheet → 保持题干不变，只重出三层答案（优先用行内 `提示词`，为空时回退到 `--answer-prompt-file`）。
- `questions`：不含题目列的单 sheet → 按每行的 `提示词` 各出一题。

用法：

```bash
export DASHSCOPE_API_KEY=...   # 切勿把 key 写死在代码里

# 出题+出答案（自动识别）
python3 scripts/fill_and_call.py 需要重新出题.xlsx -o out.json

# 仅重出答案（题干不变）；行内提示词为空时用 fallback 模板
python3 scripts/fill_and_call.py 重新出答案.xlsx -o out.json \
  --answer-prompt-file 出答案提示词.txt

# 审核失败迭代：把针对性约束写进文件，按阶段追加
python3 scripts/fill_and_call.py 需要重新出题.xlsx -o out.json \
  --extra-q-file q_constraints.txt --extra-a-file a_constraints.txt
```

参数：`--mode auto|question|answers|questions`、`--answer-prompt-file`、`--extra-q-file`、`--extra-a-file`、`-o/--out`。

脚本写出 `out.json` 后，GPT 必须逐字依据 `prompts/审题提示词.txt` 审核每一条；脚本输出绝不自动视为通过。

### 原始 API 调用模板

若你自己写脚本，使用 DashScope 的 OpenAI 兼容接口：

```python
import json
from urllib import request

API_URL = "https://dashscope.aliyuncs.com/compatible-mode/v1/chat/completions"

payload = {
    "model": "qwen-plus",
    "messages": [
        {"role": "system", "content": "你只输出严格JSON，不输出任何解释。"},
        {"role": "user", "content": prompt},
    ],
    "temperature": 0.2,
    "response_format": {"type": "json_object"},
}

req = request.Request(
    API_URL,
    data=json.dumps(payload, ensure_ascii=False).encode("utf-8"),
    headers={
        "Authorization": f"Bearer {api_key}",
        "Content-Type": "application/json",
    },
    method="POST",
)
```

对网络瞬时错误（如 `RemoteDisconnected`、超时、5xx）要加重试。不要把网络失败当作审核失败。

## 提示词变量映射

使用以下通用映射：

| 模板变量 | Excel 字段 |
| --- | --- |
| `{{industry}}` | `行业` |
| `{{job}}` | `岗位` 或 `岗位名称` |
| `{{scene}}` | `招聘场景` 或 `场景` |
| `{{category}}` | `技能分类` |
| `{{skill}}` | `技能` |
| `{{knowledge}}` | `知识点` |
| `{{level}}` | `难度` |
| `{{JD}}` | `JD` |
| `{{Q}}` | 当前 `题目` 或生成出的题目 |
| `{{num}}` | 通常为 `1` |

替换前先把写错的模板占位符 `{{{industry}}` 归一化为 `{{industry}}`。

## qwen 输出格式要求

出题时：

```json
{
  "questions": [
    {
      "question": "题目题干",
      "level": "难度",
      "time": 120
    }
  ]
}
```

出答案时：

```json
{
  "questions": [
    {
      "question": "题目题干",
      "answer1": "参考答案-第一层",
      "answer2": "参考答案-第二层",
      "answer3": "参考答案-第三层"
    }
  ]
}
```

## GPT 审核标准

GPT 必须严格依据本技能内 `prompts/审题提示词.txt` 这份审题提示词审核每一个生成项。该文件是唯一标准来源。

硬性规则：

- 逐字使用这份审题提示词。绝不改写、转述、缩写、总结、重排或扩展其中的规则。
- 不得用你自己的清单替代它，也不得新增任何通过/不通过的判定标准。唯一的审核标准就是该文件里写明的内容。
- 审核前只能把 `{{岗位名称}}` 占位符替换为实际岗位名，其余一字不改。
- 完整套用其中每一条规则，包括错误类型代码（A1–A4、B1–B6、C1–C2）、三层答案定义、难度定义，以及风险标注规则（肯定无 / 疑似无 / 疑似有 / 肯定有，其中除「肯定无」外均需删除）。
- 若该文件缺失或无法读取，应停下来向用户索取，而不要凭记忆审核。

审核前，先完整读取 `prompts/审题提示词.txt` 并原样套用。

## 迭代指引

当某条目未通过审核时，编写针对性约束。示例：

- `不要写"抽真空阶段阀门保持关闭却形成抽气通路"；必须说明关闭破空阀后按联锁打开抽气通路阀。`
- `不要编造固定阈值；改用厂内报警线、机台基线、工艺窗口或历史趋势。`
- `不要使用 Wafer/RR/SPC/pH 等英文裸词或符号；改成中文口语表达。`
- `化学品泄漏必须优先人员撤离、区域隔离、停供泄压、EHS/设施响应，不要以产线连续性优先。`
- `保留题干重出答案时不得改题干。`

每次重出后都要重新依据 `prompts/审题提示词.txt` 逐字审核。本地脚本只能预筛量化类问题（长度、零散英文、泄漏的元指令），绝不替代或覆盖这份审题提示词。不要仅凭本地脚本就判定通过。

## 最终输出格式

将最终 Excel 写到用户提供的输出路径。若用户提供了参考题目表格式，则与其保持一致。推荐列：

- `招聘场景`
- `行业`
- `岗位`
- `JD`
- `技能分类`
- `技能`
- `知识点`
- `难度`
- `题目`
- `审核状态`
- `参考答案-第一层`
- `参考答案-第二层`
- `参考答案-第三层`
- `提示词`

所有最终行的 `审核状态` 设为 `通过`。除非用户明确要求保留，否则 `提示词` 列留空。

### 内置脚本：scripts/write_final_xlsx.py

当 GPT 已逐字依据 `prompts/审题提示词.txt` 审核完毕、并只保留通过项（肯定无）后，用该脚本写出最终 Excel。它不做审核，只负责把已通过的条目序列化落表。

输入 JSON 可以是数组 `[{...}]`、对象 `{"items": [...]}`，或形如 `{"出题": {...}, "出答案#2": {...}}` 的键值字典。它会校验 `题目` 与三层答案均非空，将 `审核状态` 设为 `通过`，清空 `提示词`（除非加 `--keep-prompt`），并按上面的固定列序输出。

```bash
# 默认：输入即终稿，全部写出
python3 scripts/write_final_xlsx.py audited.json -o 最终_GPT审核通过.xlsx

# 只写被标记通过/肯定无的条目
python3 scripts/write_final_xlsx.py audited.json -o out.xlsx --passed-only
```

参数：`-o/--out`、`--sheet`、`--passed-only`、`--keep-prompt`。若任一条目缺少必填字段，脚本会中止并以非零码退出。

## 汇报

最终回复应包含：

- 最终输出文件路径。
- 重新生成的题目数量。
- 重新生成的答案数量。
- 最终通过的行数。
- 如有被卡住的条目，一并列出。

回复中不要包含 API key。
