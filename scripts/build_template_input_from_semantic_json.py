import json
from pathlib import Path


ROOT = Path(__file__).resolve().parent.parent
INPUT_PATH = ROOT / "outputs" / "semantic_sql_compare_results.json"
OUTPUT_PATH = ROOT / "outputs" / "ai_match_template_input.json"


def main() -> None:
    raw = json.loads(INPUT_PATH.read_text(encoding="utf-8"))

    summary_rows = []
    for row in raw.get("summary", []):
        if "统计项" in row and "数量" in row:
            summary_rows.append({"统计项": row["统计项"], "数量": row["数量"]})

    detail_rows = []
    for row in raw.get("details", []):
        detail_rows.append(
            {
                "0709序号": row.get("0709序号", ""),
                "0709原表行号": row.get("0709原表行号", ""),
                "任务编号": row.get("任务编号", ""),
                "应用场景": row.get("应用场景", ""),
                "代码位置": row.get("代码位置", ""),
                "原修复情况": row.get("修复情况", ""),
                "0709 SQL": row.get("0709 SQL", ""),
                "7月10日判断": row.get("7月10日判断", ""),
                "7月10日置信度": row.get("7月10日置信度", ""),
                "7月10日匹配序号": row.get("7月10日匹配序号", ""),
                "7月10日原表行号": row.get("7月10日原表行号", ""),
                "7月10日指纹": row.get("7月10日指纹", ""),
                "7月10日服务": row.get("7月10日服务", ""),
                "7月10日候选SQL": row.get("7月10日候选SQL", ""),
                "7月10日详细分析": row.get("7月10日详细分析", ""),
                "7月13日判断": row.get("7月13日判断", ""),
                "7月13日置信度": row.get("7月13日置信度", ""),
                "7月13日匹配序号": row.get("7月13日匹配序号", ""),
                "7月13日原表行号": row.get("7月13日原表行号", ""),
                "7月13日指纹": row.get("7月13日指纹", ""),
                "7月13日服务": row.get("7月13日服务", ""),
                "7月13日候选SQL": row.get("7月13日候选SQL", ""),
                "7月13日详细分析": row.get("7月13日详细分析", ""),
                "综合结论": row.get("综合结论", ""),
            }
        )

    output = {"summary_rows": summary_rows, "detail_rows": detail_rows}
    OUTPUT_PATH.parent.mkdir(parents=True, exist_ok=True)
    OUTPUT_PATH.write_text(json.dumps(output, ensure_ascii=False, indent=2), encoding="utf-8")
    print(OUTPUT_PATH)


if __name__ == "__main__":
    main()
