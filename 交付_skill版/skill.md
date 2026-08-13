---
name: sql-semantic-file-to-file
description: 用于两个 Excel 文件之间的 SQL 语义比对。将主文件中的每条 SQL 与对比文件中的候选 SQL 做语义判断，输出是否为同一业务 SQL、依据行号、结构变化说明和最终 Excel。适用于 .xlsx，也适用于 HTML 导出的 .xls。
---

# 文件对文件 SQL 语义比对 Skill 版

这是 `skill版` 交付物，适合直接放到支持 skill 的平台里使用。

这一版的职责是：
- 用本地 Python 脚本读取两个文件
- 做 SQL 清洗、结构特征提取、候选召回
- 产出待智能体判断的 `prepare.json`
- 接收智能体返回的 `review_results.json`
- 生成最终 Excel

Python 脚本本身不调用任何大模型。

## 输入

必填：
- 主文件路径
- 对比文件路径

可选：
- 主文件 SQL 列名
- 对比文件 SQL 列名
- 主文件主键列名
- 候选召回数量 `top_k`

默认优先识别的 SQL 列名：
- `SQL语句`
- `SQL`
- `sql`
- `任务详情`
- `慢SQL`

## 适用文件类型

- 标准 `.xlsx`
- HTML 导出的 `.xls`

## 工作流程

### 1. Prepare

执行：

```bash
python compare_sql_files.py \
  --mode prepare \
  --base 主文件.xlsx \
  --target 对比文件.xls \
  --base-sql-column "SQL语句" \
  --target-sql-column "SQL语句" \
  --base-id-column "任务编号" \
  --top-k 3 \
  --output result/prepare.json
```

这一阶段负责：
- 读取文件
- 兼容 HTML `.xls`
- 提取 SQL 结构特征
- 候选召回
- 生成待智能体处理的数据

### 2. 智能体语义判断

智能体逐对判断 `prepare.json` 里的候选 SQL，对每个 `pair_id` 输出：

- `judgement`
- `confidence`
- `semantic_score`
- `same_business`
- `reasoning`
- `join_change`
- `where_change`
- `group_by_change`
- `subquery_change`
- `base_line_refs`
- `target_line_refs`
- `common_tables`
- `key_differences`

### 3. Finalize

执行：

```bash
python compare_sql_files.py \
  --mode finalize \
  --prepared result/prepare.json \
  --review-results result/review_results.json \
  --output result/final_result.xlsx
```

输出：
- `final_result.xlsx`
- 同名 JSON 底稿

## 判定原则

必须按业务语义判断，不按字符串相似度判断。

重点分析：
- 查询目标
- 核心表
- JOIN
- WHERE / 过滤条件
- GROUP BY / 聚合口径
- 子查询 / CTE / UNION

## 最终结果

Excel 至少包含：
- `汇总`
- `逐条语义比对`
- `候选明细`

## 推荐窗口输入模板

```text
请用 sql-semantic-file-to-file skill 比较这两个文件：
主文件：/path/to/base.xlsx
对比文件：/path/to/target.xls
主文件 SQL 列：SQL语句
对比文件 SQL 列：SQL语句
主文件主键列：任务编号
输出到：/path/to/final_result.xlsx
```
