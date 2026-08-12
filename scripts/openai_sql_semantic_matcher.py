import json
import os
from dataclasses import dataclass
from typing import Dict, List, Sequence

import requests

from semantic_sql_compare import SqlRecord, format_sql_lines


OPENAI_API_URL = "https://api.openai.com/v1/responses"
DEFAULT_MODEL = os.getenv("OPENAI_MODEL", "gpt-5.6")


@dataclass
class LlmJudgeResult:
    judgement: str
    confidence: str
    semantic_score: float
    reasoning: str
    same_business: bool
    join_change: str
    where_change: str
    group_by_change: str
    subquery_change: str
    base_line_refs: str
    candidate_line_refs: str
    common_tables: List[str]
    key_differences: List[str]


def require_api_key() -> str:
    api_key = os.getenv("OPENAI_API_KEY", "").strip()
    if not api_key:
        raise RuntimeError("未检测到 OPENAI_API_KEY，无法调用大模型。")
    return api_key


def numbered_sql(sql: str, max_lines: int = 220) -> str:
    lines = format_sql_lines(sql)
    lines = lines[:max_lines]
    return "\n".join(f"{idx:03d} | {line}" for idx, line in enumerate(lines, start=1))


def build_features_text(record: SqlRecord) -> str:
    f = record.features
    return (
        f"核心表: {', '.join(f.tables[:20])}\n"
        f"JOIN数量: {f.join_count}\n"
        f"WHERE条件数: {len(f.where_conditions)}\n"
        f"GROUP BY字段: {', '.join(f.group_by[:10])}\n"
        f"子查询数量: {f.subquery_count}\n"
        f"UNION数量: {f.union_count}\n"
        f"函数: {', '.join(f.functions[:20])}\n"
        f"SELECT字段样本: {', '.join(f.select_targets[:20])}"
    )


def schema() -> Dict[str, object]:
    return {
        "type": "object",
        "additionalProperties": False,
        "properties": {
            "judgement": {
                "type": "string",
                "enum": ["同一业务SQL", "可能同一业务SQL", "非同一业务SQL"],
            },
            "confidence": {
                "type": "string",
                "enum": ["极高", "高", "中", "低"],
            },
            "semantic_score": {"type": "number"},
            "same_business": {"type": "boolean"},
            "reasoning": {"type": "string"},
            "join_change": {"type": "string"},
            "where_change": {"type": "string"},
            "group_by_change": {"type": "string"},
            "subquery_change": {"type": "string"},
            "base_line_refs": {"type": "string"},
            "candidate_line_refs": {"type": "string"},
            "common_tables": {
                "type": "array",
                "items": {"type": "string"},
            },
            "key_differences": {
                "type": "array",
                "items": {"type": "string"},
            },
        },
        "required": [
            "judgement",
            "confidence",
            "semantic_score",
            "same_business",
            "reasoning",
            "join_change",
            "where_change",
            "group_by_change",
            "subquery_change",
            "base_line_refs",
            "candidate_line_refs",
            "common_tables",
            "key_differences",
        ],
    }


def prompt_for_pair(base: SqlRecord, candidate: SqlRecord, base_label: str, candidate_label: str) -> str:
    return f"""
你是资深 SQL 语义分析专家。你的任务不是看字符串像不像，而是判断两条 SQL 是否属于同一业务 SQL。

判断标准：
1. 重点看查询目标、核心表、JOIN 关系、WHERE 主干、GROUP BY 口径、子查询/CTE/UNION 的业务含义是否一致。
2. 如果只是改写、分页变化、条件下推、历史分表、字段增减、参数字面值差异，通常仍可判为同一业务SQL或可能同一业务SQL。
3. 如果查询目标、统计口径、核心过滤逻辑明显不同，就判为非同一业务SQL。
4. 必须引用我给出的带行号 SQL，输出 base_line_refs 和 candidate_line_refs，标出你主要依据的行号。
5. reasoning 必须用中文，写清楚为什么这么判。

{base_label} 元信息：
任务编号: {base.meta.get("任务编号", "")}
应用场景: {base.meta.get("应用场景", "")}
代码位置: {base.meta.get("代码位置", "")}
SQL结构特征:
{build_features_text(base)}

{candidate_label} 元信息：
指纹: {candidate.meta.get("指纹", "")}
服务: {candidate.meta.get("服务", "")}
SQL结构特征:
{build_features_text(candidate)}

{base_label} SQL（带行号）:
{numbered_sql(base.sql)}

{candidate_label} SQL（带行号）:
{numbered_sql(candidate.sql)}
""".strip()


def parse_response_json(resp_json: Dict[str, object]) -> Dict[str, object]:
    text = ""
    for item in resp_json.get("output", []):
        if item.get("type") != "message":
            continue
        for content in item.get("content", []):
            if content.get("type") == "output_text":
                text += content.get("text", "")
    if not text.strip():
        raise RuntimeError("大模型返回为空，未获取到结构化结果。")
    return json.loads(text)


def judge_pair_with_openai(base: SqlRecord, candidate: SqlRecord, base_label: str, candidate_label: str) -> LlmJudgeResult:
    api_key = require_api_key()
    payload = {
        "model": DEFAULT_MODEL,
        "store": False,
        "input": [
            {
                "role": "system",
                "content": [
                    {
                        "type": "input_text",
                        "text": "你只输出符合 JSON Schema 的结果，不要输出额外说明。",
                    }
                ],
            },
            {
                "role": "user",
                "content": [{"type": "input_text", "text": prompt_for_pair(base, candidate, base_label, candidate_label)}],
            },
        ],
        "text": {
            "format": {
                "type": "json_schema",
                "name": "sql_semantic_judgement",
                "schema": schema(),
                "strict": True,
            }
        },
    }
    response = requests.post(
        OPENAI_API_URL,
        headers={
            "Authorization": f"Bearer {api_key}",
            "Content-Type": "application/json",
        },
        json=payload,
        timeout=180,
    )
    response.raise_for_status()
    data = parse_response_json(response.json())
    return LlmJudgeResult(
        judgement=data["judgement"],
        confidence=data["confidence"],
        semantic_score=float(data["semantic_score"]),
        reasoning=data["reasoning"],
        same_business=bool(data["same_business"]),
        join_change=data["join_change"],
        where_change=data["where_change"],
        group_by_change=data["group_by_change"],
        subquery_change=data["subquery_change"],
        base_line_refs=data["base_line_refs"],
        candidate_line_refs=data["candidate_line_refs"],
        common_tables=list(data["common_tables"]),
        key_differences=list(data["key_differences"]),
    )
