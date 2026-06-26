---
name: qwen-question-gpt-audit
description: Orchestrates qwen-plus question or answer generation from Excel rows, then uses the current GPT agent to audit and iterate until all generated interview questions pass. Use when the user asks to call 千问/qwen-plus to regenerate interview questions, regenerate answers, or process Excel files with prompts and variables while GPT handles review.
disable-model-invocation: true
---

# Qwen Question Generation + GPT Audit

## Purpose

Use this skill to regenerate interview questions and/or reference answers from Excel files:

- `qwen-plus` only generates content from row-level prompts and variables.
- The current GPT agent is the orchestrator and auditor.
- GPT audit MUST follow `prompts/审题提示词.txt` verbatim. Never rewrite, abbreviate, or extend the audit rules.
- Do not hand an entire workbook to qwen-plus for autonomous processing.
- Iterate failed items by adding targeted constraints to the qwen prompt, then re-audit.
- Final output must contain only items that pass GPT audit.

## Required user inputs

Before running, obtain all three:

1. 千问 API key for `qwen-plus`.
2. Input file path(s): Excel files to operate on.
3. Output file path: final Excel file path.

If any is missing, ask for it before starting.

Security handling:

- Prefer reading the API key from `DASHSCOPE_API_KEY`.
- If the user explicitly authorizes direct use of a provided key, avoid writing it into persistent scripts or output files.
- Do not include the key in generated logs, JSON, Excel, or final summaries.
- If the key was pasted in chat, recommend rotating it after the task.

## Expected input patterns

The workflow supports these common sheets/files:

### 1. Need to regenerate questions and answers

Rows contain variables such as:

- `问题` or original failed question
- `技能分类`
- `技能`
- `知识点`
- `难度`
- `行业`
- `岗位`
- `招聘场景`
- `JD`
- `提示词`

If the workbook contains separate sheets such as `出题` and `出答案`:

- Use `出题` sheet prompt to generate the new question.
- Use `出答案` sheet prompt to generate answers for the newly generated question.

### 2. Need to regenerate answers only

Rows contain:

- `题目` or `问题`
- `技能分类`
- `技能`
- `知识点`
- `难度`
- `行业`
- `岗位`
- `招聘场景`
- `JD`
- `提示词`

Keep the original question text unchanged; regenerate only `参考答案-第一层`、`参考答案-第二层`、`参考答案-第三层`.

## Orchestration rules

1. Read Excel locally and inspect sheets/columns first.
2. Process rows one by one.
3. For each row, replace prompt variables before calling qwen-plus.
4. Send qwen-plus only the current row prompt and variables.
5. Ask qwen-plus to return strict JSON.
6. Parse qwen output into structured fields.
7. GPT audits the generated row strictly using `prompts/审题提示词.txt` verbatim (see "GPT audit rubric"). Do not rewrite, abbreviate, or extend those rules.
8. If GPT audit fails:
   - Record specific failure reasons.
   - Add targeted constraints to the qwen prompt for that item only.
   - Regenerate and re-audit.
9. Stop only when every final row passes GPT audit, or report the blocked items and reasons.

## Qwen API call pattern

Use DashScope OpenAI-compatible endpoint:

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

Add retry for transient network errors such as `RemoteDisconnected`, timeouts, and 5xx responses. Do not treat network failures as audit failures.

## Prompt variable mapping

Use common mappings:

| Template variable | Excel field |
| --- | --- |
| `{{industry}}` | `行业` |
| `{{job}}` | `岗位` or `岗位名称` |
| `{{scene}}` | `招聘场景` or `场景` |
| `{{category}}` | `技能分类` |
| `{{skill}}` | `技能` |
| `{{knowledge}}` | `知识点` |
| `{{level}}` | `难度` |
| `{{JD}}` | `JD` |
| `{{Q}}` | current `题目` or generated question |
| `{{num}}` | usually `1` |

Normalize malformed template token `{{{industry}}` to `{{industry}}` before replacement.

## Required qwen output formats

For question generation:

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

For answer generation:

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

## GPT audit rubric

GPT MUST audit every generated item strictly according to the canonical audit prompt stored at `prompts/审题提示词.txt` in this skill. This file is the single source of truth.

Hard rules:

- Use the canonical 审题提示词 verbatim. Do NOT rewrite, paraphrase, abbreviate, summarize, reorder, or extend its rules.
- Do NOT substitute your own checklist or invent additional pass/fail criteria. The only audit criteria are the ones written in that file.
- Replace only the `{{岗位名称}}` placeholder with the actual job title before auditing; change nothing else.
- Apply every rule exactly as written, including the error-type codes (A1–A4, B1–B6, C1–C2), the three-layer answer definitions, the difficulty definitions, and the risk-labeling rule (肯定无 / 疑似无 / 疑似有 / 肯定有，其中除“肯定无”外均需删除).
- If the file is missing or unreadable, stop and ask the user for it rather than auditing from memory.

Before auditing, read `prompts/审题提示词.txt` in full and apply it as-is.

## Iteration guidance

When an item fails audit, create targeted constraints. Examples:

- `不要写“抽真空阶段阀门保持关闭却形成抽气通路”；必须说明关闭破空阀后按联锁打开抽气通路阀。`
- `不要编造固定阈值；改用厂内报警线、机台基线、工艺窗口或历史趋势。`
- `不要使用 Wafer/RR/SPC/pH 等英文裸词或符号；改成中文口语表达。`
- `化学品泄漏必须优先人员撤离、区域隔离、停供泄压、EHS/设施响应，不要以产线连续性优先。`
- `保留题干重出答案时不得改题干。`

Re-audit after every regeneration using `prompts/审题提示词.txt` verbatim. Local scripts may pre-filter only quantitative issues (length, stray tokens, leaked meta-instructions); they never replace or override the canonical audit prompt. Do not mark as passed based only on local scripts.

## Final output format

Write a final Excel file to the user-provided output path. Match the reference question table format when provided. Preferred columns:

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

Set `审核状态` to `通过` for all final rows. Keep `提示词` empty unless the user explicitly wants it retained.

## Reporting

Final response should include:

- Final output file path.
- Number of regenerated questions.
- Number of regenerated answers.
- Number of final passed rows.
- Any blocked items, if present.

Do not include the API key in the response.
