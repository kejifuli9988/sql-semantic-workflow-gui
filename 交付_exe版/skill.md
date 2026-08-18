---
name: sql-semantic-file-to-file
description: 用于配合 SQL 语义比对 exe/GUI 工具工作。exe 负责生成 prepare、拆分批次、合并 review 结果并输出最终 Excel；智能体只负责根据 prepare_part 做 SQL 语义判断并输出 review_results_part。
---

# 文件对文件 SQL 语义比对 EXE 协同版

这是 `exe版` 配套 skill。

这一版不是让智能体自己重读 Excel、自己写额外脚本、自己生成最终 Excel，而是和本地 GUI / exe 分工协作：

- exe / GUI 负责：
  - 读取主文件和对比文件
  - 自动识别列名
  - 生成 `prepare.json`
  - 拆分 `prepare_part_x.json`
  - 合并 `review_results_part_x.json`
  - 生成最终 Excel
- 智能体负责：
  - 读取 `prepare_part_x.json`
  - 对其中每个 `pair_id` 做 SQL 语义判断
  - 输出对应的 `review_results_part_x.json`

## 智能体只做什么

只做语义评审，不做这些事：
- 不直接读取 Excel
- 不自己做候选召回
- 不自己重新生成 prepare
- 不自己写额外 Python 脚本
- 不自己生成最终 Excel

## 智能体输入

输入来自 exe 生成的 `prepare_part_x.json`。

每个 part 里已经包含：
- 主 SQL
- 候选 SQL
- 结构特征
- 带行号 SQL
- `expected_result_count`
- `expected_pair_ids`
- 需要返回的字段说明

## 智能体输出

每个 `prepare_part_x.json` 对应输出一个：
- `review_results_part_x.json`

每条结果必须包含：
- `pair_id`
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

## 判定原则

必须按业务语义判断，不按字符串相似度判断。

重点看：
- 查询目标
- 核心表
- JOIN
- WHERE / 过滤条件
- GROUP BY / 聚合口径
- 子查询 / CTE / UNION

## 推荐窗口输入模板

```text
请按 prepare_part 分批处理 SQL 语义比对任务。
不要重新读取 Excel，不要自己生成脚本，不要改 pair_id。
每个 prepare_part 都要先读取 expected_result_count 和 expected_pair_ids。
输出结果条数必须等于 expected_result_count。
输出的 pair_id 必须且只能来自 expected_pair_ids，禁止新增、遗漏、重复、改写 pair_id。
请直接读取结果目录下的 prepare_part 文件，并为每个 prepare_part 生成对应的 review_results_part JSON。
完成后返回：已生成了哪些 review_results_part 文件。
```
