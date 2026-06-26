#!/usr/bin/env python3
"""
write_final_xlsx.py — 把「已通过 GPT 审核」的条目落成最终 Excel。

重要边界：
- 本脚本不做审核、不判定通过与否。审核必须由 GPT agent 逐字按 prompts/审题提示词.txt 完成。
- 仅当某条目已被 GPT 审核判为「肯定无」（保留）后，才应进入输入 JSON。
- 输入 JSON 通常是 GPT 审核后整理的最终条目数组（或带 passed 标记的对象）。

输入 JSON 支持两种形态：
1) 数组：[{...}, {...}]，默认每条都视为已通过；
2) 对象：{"items": [...]}，可选 "passed_only": true 时只写 item.get("通过"/"passed") 为真的行。

输出列固定为参考题目表格式：
招聘场景 | 行业 | 岗位 | JD | 技能分类 | 技能 | 知识点 | 难度 | 题目 |
审核状态 | 参考答案-第一层 | 参考答案-第二层 | 参考答案-第三层 | 提示词
- 审核状态 统一写「通过」。
- 提示词 默认留空，除非 --keep-prompt 且条目里带「提示词」。
"""
from __future__ import annotations

import argparse
import json
import sys

import openpyxl

OUTPUT_COLUMNS = [
    "招聘场景", "行业", "岗位", "JD", "技能分类", "技能", "知识点", "难度", "题目", "审核状态",
    "参考答案-第一层", "参考答案-第二层", "参考答案-第三层", "提示词",
]

REQUIRED_NONEMPTY = ["题目", "参考答案-第一层", "参考答案-第二层", "参考答案-第三层"]


def load_items(path: str, passed_only: bool) -> list[dict]:
    with open(path, encoding="utf-8") as f:
        data = json.load(f)
    if isinstance(data, dict):
        items = data.get("items")
        if items is None:
            # 兼容 {"出题": {...}, "出答案#2": {...}} 这类键值字典
            items = [v for v in data.values() if isinstance(v, dict)]
        if data.get("passed_only"):
            passed_only = True
    elif isinstance(data, list):
        items = data
    else:
        raise ValueError("输入 JSON 必须是数组或对象")

    if passed_only:
        def is_passed(it: dict) -> bool:
            for k in ("通过", "passed", "审核通过"):
                if k in it:
                    return bool(it[k])
            status = str(it.get("审核状态", "")).strip()
            return status in ("通过", "肯定无")
        items = [it for it in items if is_passed(it)]
    return items


def cell(item: dict, key: str) -> str:
    # 兼容若干别名
    aliases = {
        "招聘场景": ["招聘场景", "场景"],
        "岗位": ["岗位", "岗位名称"],
        "题目": ["题目", "问题"],
    }
    for k in aliases.get(key, [key]):
        if k in item and item[k] is not None:
            return str(item[k])
    return ""


def main() -> None:
    ap = argparse.ArgumentParser(description="把已通过 GPT 审核的条目写成最终 Excel")
    ap.add_argument("json", help="输入 JSON（GPT 审核后整理的最终条目）")
    ap.add_argument("-o", "--out", required=True, help="输出 Excel 路径")
    ap.add_argument("--sheet", default="已确认题目答案", help="输出 sheet 名")
    ap.add_argument("--passed-only", action="store_true",
                    help="只写被标记通过/肯定无的条目（默认全部写出，假定已是终稿）")
    ap.add_argument("--keep-prompt", action="store_true",
                    help="保留条目里的『提示词』列内容（默认留空）")
    args = ap.parse_args()

    items = load_items(args.json, args.passed_only)
    if not items:
        print("没有可写出的条目（检查输入 JSON 或 --passed-only 过滤）", file=sys.stderr)
        sys.exit(1)

    problems = []
    for i, it in enumerate(items, 1):
        for req in REQUIRED_NONEMPTY:
            if not cell(it, req).strip():
                problems.append(f"第{i}条缺少必填字段：{req}")
    if problems:
        print("发现不完整条目，已中止（请先补全或重新生成审核）：", file=sys.stderr)
        for p in problems:
            print("  - " + p, file=sys.stderr)
        sys.exit(2)

    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = args.sheet
    ws.append(OUTPUT_COLUMNS)
    for it in items:
        row = []
        for col in OUTPUT_COLUMNS:
            if col == "审核状态":
                row.append("通过")
            elif col == "提示词":
                row.append(cell(it, "提示词") if args.keep_prompt else "")
            else:
                row.append(cell(it, col))
        ws.append(row)
    wb.save(args.out)
    print(f"written_rows={len(items)}  out={args.out}  sheet={args.sheet}")


if __name__ == "__main__":
    main()
