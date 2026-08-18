#!/usr/bin/env python3
import json
import os
import re
import subprocess
import sys
import threading
import traceback
from datetime import datetime
from pathlib import Path
from typing import Dict, List

import tkinter as tk
from tkinter import filedialog, messagebox, ttk

import compare_sql_files as workflow


SETTINGS_FILE = Path(__file__).with_name("sql_semantic_workflow_gui_settings.json")
DEFAULT_SKILL_MD = """---
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
"""
DEFAULT_ROLE_MD = """# 智能体角色定义

## 角色名称

SQL 语义评审专家

## 角色定位

你在 `exe版` 工作流中，只负责 SQL 语义评审这一段。

本地 exe / GUI 已经完成了：
- 读取 Excel
- 识别列名
- 候选召回
- prepare 拆分

你不需要再重复做这些步骤。

## 你的职责

1. 读取 `prepare_part_x.json`
2. 先读取其中的 `expected_result_count` 和 `expected_pair_ids`
3. 逐条处理其中的 `pair_id`
4. 输出结果条数必须等于 `expected_result_count`
5. 输出的 `pair_id` 必须且只能来自 `expected_pair_ids`
6. 判断主 SQL 和候选 SQL 是否属于同一业务 SQL
7. 分析 JOIN、WHERE、GROUP BY、子查询、CTE、UNION、过滤条件变化
8. 标注依据行号
9. 输出 `review_results_part_x.json`

## 你不要做的事

- 不要重新读取 Excel
- 不要自己候选召回
- 不要自己生成 prepare
- 不要自己写额外脚本
- 不要改写 `pair_id`
- 不要新增未出现在 `expected_pair_ids` 里的结果
- 不要漏掉 `expected_pair_ids` 里的结果
- 不要重复 `pair_id`
- 不要自己生成最终 Excel

## 判断原则

按业务语义判断，不按字符串相似度判断。

重点分析：
- 查询目标是否一致
- 核心表是否一致
- JOIN 是否表达同一业务逻辑
- WHERE 主干是否一致
- GROUP BY / 聚合口径是否一致
- 子查询、CTE、UNION 是否只是写法变化

## 输出字段

必须输出：
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

## 输出要求

- 一个 `prepare_part` 对应一个 `review_results_part`
- 输出结果条数必须等于 `expected_result_count`
- 输出的 `pair_id` 必须与 `expected_pair_ids` 完全一致
- JSON 必须完整
- 不要漏结果
- 不要重复 `pair_id`
- 用中文写理由
- 行号必须可追溯
"""


def load_settings() -> dict:
    if SETTINGS_FILE.exists():
        try:
            return json.loads(SETTINGS_FILE.read_text(encoding="utf-8"))
        except Exception:
            return {}
    return {}


def save_settings(data: dict) -> None:
    SETTINGS_FILE.write_text(json.dumps(data, ensure_ascii=False, indent=2), encoding="utf-8")


class SqlSemanticWorkflowApp:
    def __init__(self, root: tk.Tk) -> None:
        self.root = root
        self.root.title("SQL 语义比对工作流")
        self.settings = load_settings()
        self.page_bg = "#F5F7FA"
        self.card_bg = "#FFFFFF"
        self.border_color = "#E5E7EB"
        self.primary = "#2563EB"
        self.primary_hover = "#1D4ED8"
        self.primary_soft = "#EFF6FF"
        self.text_primary = "#111827"
        self.text_secondary = "#6B7280"
        self.code_bg = "#F8FAFC"
        self.success = "#16A34A"
        self.error = "#DC2626"
        self.running = "#2563EB"
        self.root.configure(bg=self.page_bg)
        self._configure_styles()
        self._set_default_window()

        self.base_file = tk.StringVar(value=self.settings.get("base_file", ""))
        self.target_file = tk.StringVar(value=self.settings.get("target_file", ""))
        self.base_sql_column = tk.StringVar(value=self.settings.get("base_sql_column", ""))
        self.target_sql_column = tk.StringVar(value=self.settings.get("target_sql_column", ""))
        self.base_display_columns_summary = tk.StringVar(value="未选择显示列")
        self.target_display_columns_summary = tk.StringVar(value="未选择显示列")
        self.top_k = tk.StringVar(value=str(self.settings.get("top_k", 3)))
        self.batch_size = tk.StringVar(value=str(self.settings.get("batch_size", 20)))
        self.result_dir = tk.StringVar(value=self.settings.get("result_dir", str(Path.cwd() / "result")))
        self.base_columns: List[str] = []
        self.target_columns: List[str] = []
        self.base_display_column_vars: Dict[str, tk.BooleanVar] = {}
        self.target_display_column_vars: Dict[str, tk.BooleanVar] = {}
        self.saved_base_display_columns = list(self.settings.get("base_display_columns", []))
        self.saved_target_display_columns = list(self.settings.get("target_display_columns", []))
        self.active_display_popup: tk.Toplevel | None = None
        self.active_display_popup_kind: str | None = None
        self.display_popup_click_bind_id: str | None = None

        self.review_result_files: List[Path] = []
        self.is_busy = False
        self.action_buttons: List[ttk.Button] = []
        self.secondary_buttons: List[ttk.Button] = []
        self.detail_tab_var = tk.StringVar(value="command")
        self.status_heading_var = tk.StringVar(value="待开始")
        self.status_message_var = tk.StringVar(value="请先选择文件并配置参数。")
        self.status_var = self.status_message_var
        self.status_color = self.text_secondary

        self._build_ui()
        self._load_saved_column_options()
        self.refresh_prompt_preview()
        self.refresh_command_preview()

    def _configure_styles(self) -> None:
        style = ttk.Style()
        try:
            style.theme_use("clam")
        except tk.TclError:
            pass
        default_font = ("Microsoft YaHei UI", 10)
        text_font = ("Microsoft YaHei UI", 10)
        title_font = ("Microsoft YaHei UI", 16, "bold")
        subtitle_font = ("Microsoft YaHei UI", 9)
        section_font = ("Microsoft YaHei UI", 11, "bold")
        code_font = ("Consolas", 10)

        style.configure(".", font=default_font)
        style.configure("Page.TFrame", background=self.page_bg)
        style.configure("CardInner.TFrame", background=self.card_bg)
        style.configure("CardBody.TFrame", background=self.card_bg)
        style.configure("HeaderTitle.TLabel", background=self.page_bg, foreground=self.text_primary, font=title_font)
        style.configure("HeaderSubtitle.TLabel", background=self.page_bg, foreground=self.text_secondary, font=subtitle_font)
        style.configure("SectionTitle.TLabel", background=self.card_bg, foreground=self.text_primary, font=section_font)
        style.configure("FieldLabel.TLabel", background=self.card_bg, foreground=self.text_primary, font=text_font)
        style.configure("Muted.TLabel", background=self.card_bg, foreground=self.text_secondary, font=subtitle_font)
        style.configure("StatusMain.TLabel", background=self.card_bg, foreground=self.text_primary, font=("Microsoft YaHei UI", 11, "bold"))
        style.configure("StatusFile.TLabel", background=self.card_bg, foreground=self.text_primary, font=("Microsoft YaHei UI", 10, "bold"))
        style.configure("Primary.TButton", foreground="#FFFFFF", background=self.primary, borderwidth=0, padding=(12, 7))
        style.map("Primary.TButton", background=[("active", self.primary_hover), ("pressed", self.primary_hover)])
        style.configure("Secondary.TButton", foreground=self.text_primary, background=self.card_bg, bordercolor=self.border_color, relief="solid", padding=(10, 6))
        style.map("Secondary.TButton", background=[("active", self.primary_soft)], bordercolor=[("active", self.primary)])
        style.configure("Accent.TButton", foreground=self.primary, background=self.card_bg, bordercolor=self.primary, relief="solid", padding=(10, 6))
        style.map("Accent.TButton", background=[("active", self.primary_soft)], bordercolor=[("active", self.primary_hover)], foreground=[("active", self.primary_hover)])
        style.configure("FilePick.TButton", foreground=self.text_primary, background=self.card_bg, bordercolor=self.border_color, relief="solid", padding=(8, 3))
        style.map("FilePick.TButton", background=[("active", self.primary_soft)], bordercolor=[("active", self.primary)])
        style.configure("TEntry", padding=5, fieldbackground="#FFFFFF", bordercolor=self.border_color)
        style.configure("TCombobox", padding=3, fieldbackground="#FFFFFF", bordercolor=self.border_color)
        style.map(
            "TCombobox",
            fieldbackground=[("readonly", "#FFFFFF"), ("!disabled", "#FFFFFF")],
            background=[("readonly", "#FFFFFF"), ("!disabled", "#FFFFFF")],
            foreground=[("readonly", self.text_primary), ("!disabled", self.text_primary)],
        )
        style.configure(
            "White.TMenubutton",
            background="#FFFFFF",
            foreground=self.text_primary,
            bordercolor=self.border_color,
            relief="solid",
            padding=(8, 5),
        )
        style.map(
            "White.TMenubutton",
            background=[("active", "#FFFFFF"), ("!disabled", "#FFFFFF")],
            foreground=[("active", self.text_primary), ("!disabled", self.text_primary)],
            bordercolor=[("active", self.primary), ("!disabled", self.border_color)],
        )
        style.configure(
            "White.TButton",
            background="#FFFFFF",
            foreground=self.text_primary,
            bordercolor=self.border_color,
            relief="solid",
            padding=(8, 5),
            anchor="w",
        )
        style.map(
            "White.TButton",
            background=[("active", "#FFFFFF"), ("!disabled", "#FFFFFF")],
            foreground=[("active", self.text_primary), ("!disabled", self.text_primary)],
            bordercolor=[("active", self.primary), ("!disabled", self.border_color)],
        )
        style.configure(
            "Vertical.TScrollbar",
            background="#D6DEEA",
            troughcolor="#F3F6FA",
            bordercolor="#F3F6FA",
            arrowcolor=self.text_secondary,
            lightcolor="#F3F6FA",
            darkcolor="#F3F6FA",
            gripcount=0,
        )
        style.map(
            "Vertical.TScrollbar",
            background=[("active", "#AFC2F5"), ("pressed", self.primary)],
            arrowcolor=[("active", self.primary), ("pressed", "#FFFFFF")],
        )
        style.configure(
            "Horizontal.TScrollbar",
            background="#D6DEEA",
            troughcolor="#F3F6FA",
            bordercolor="#F3F6FA",
            arrowcolor=self.text_secondary,
            lightcolor="#F3F6FA",
            darkcolor="#F3F6FA",
            gripcount=0,
        )
        style.map(
            "Horizontal.TScrollbar",
            background=[("active", "#AFC2F5"), ("pressed", self.primary)],
            arrowcolor=[("active", self.primary), ("pressed", "#FFFFFF")],
        )
        style.configure("DetailTab.TButton", foreground=self.text_secondary, background=self.card_bg, bordercolor=self.border_color, relief="solid", padding=(10, 4))
        style.map("DetailTab.TButton", background=[("active", self.card_bg)], foreground=[("active", self.text_primary)], bordercolor=[("active", self.primary)])
        style.configure("DetailTabActive.TButton", foreground=self.primary, background="#FFFFFF", bordercolor=self.primary, relief="solid", padding=(10, 4))
        style.map("DetailTabActive.TButton", background=[("active", "#FFFFFF")], foreground=[("active", self.primary)], bordercolor=[("active", self.primary)])
        self.code_font = code_font

    def _set_default_window(self) -> None:
        width = 720
        height = 790
        self.root.minsize(680, 720)
        self.root.geometry(f"{width}x{height}")
        self.root.update_idletasks()
        screen_w = self.root.winfo_screenwidth()
        screen_h = self.root.winfo_screenheight()
        x = max(0, int((screen_w - width) / 2))
        y = max(0, int((screen_h - height) / 2) - 20)
        self.root.geometry(f"{width}x{height}+{x}+{y}")

    def _build_ui(self) -> None:
        self.page = ttk.Frame(self.root, style="Page.TFrame", padding=(14, 10, 14, 10))
        self.page.pack(fill=tk.BOTH, expand=True)
        self.page.columnconfigure(0, weight=1)

        self._build_header(self.page)

        form_card, form_body = self._create_card(self.page, "参数配置", "⚙")
        form_card.pack(fill=tk.X, pady=(0, 10))
        self._build_parameter_section(form_body)

        flow_card, flow_body = self._create_card(self.page, "比对流程", "⇄")
        flow_card.pack(fill=tk.X, pady=(0, 10))
        self._build_flow_section(flow_body)

        quick_card, quick_body = self._create_card(self.page, "快捷操作", "⌘")
        quick_card.pack(fill=tk.X, pady=(0, 10))
        self._build_quick_actions(quick_body)

        bottom = ttk.Frame(self.page, style="Page.TFrame")
        bottom.pack(fill=tk.BOTH, expand=True)
        bottom.columnconfigure(0, minsize=190)
        bottom.columnconfigure(0, weight=2)
        bottom.columnconfigure(1, weight=18)
        bottom.rowconfigure(0, minsize=248)

        status_card, status_body = self._create_card(bottom, "运行状态", "●")
        status_card.configure(width=190, height=248)
        status_card.grid(row=0, column=0, sticky="nsew", padx=(0, 6))
        status_card.grid_propagate(False)
        status_card.pack_propagate(False)
        self._build_status_section(status_body)

        detail_card, detail_body = self._create_card(bottom, "执行详情", "☰")
        detail_card.configure(height=248)
        detail_card.grid(row=0, column=1, sticky="nsew")
        detail_card.grid_propagate(False)
        detail_card.pack_propagate(False)
        self._build_detail_section(detail_body)

    def _build_header(self, parent) -> None:
        header = ttk.Frame(parent, style="Page.TFrame")
        header.pack(fill=tk.X, pady=(0, 8))

        badge = tk.Canvas(header, width=40, height=40, bg=self.page_bg, highlightthickness=0)
        badge.pack(side=tk.LEFT, padx=(0, 10))
        badge.create_oval(2, 2, 38, 38, fill=self.primary_soft, outline="")
        badge.create_text(20, 20, text="SQL", fill=self.primary, font=("Microsoft YaHei UI", 11, "bold"))

        text_wrap = ttk.Frame(header, style="Page.TFrame")
        text_wrap.pack(side=tk.LEFT, fill=tk.X, expand=True)
        ttk.Label(text_wrap, text="SQL 语义比对工作流", style="HeaderTitle.TLabel").pack(anchor="w")
        ttk.Label(
            text_wrap,
            text="用于批量完成 SQL 候选召回、AI 语义分析与结果汇总",
            style="HeaderSubtitle.TLabel",
        ).pack(anchor="w", pady=(1, 0))

    def _create_card(self, parent, title: str, icon: str = ""):
        card = tk.Frame(parent, bg=self.card_bg, highlightbackground=self.border_color, highlightthickness=1, bd=0)
        title_row = tk.Frame(card, bg=self.card_bg)
        title_row.pack(fill=tk.X, padx=12, pady=(7, 4))
        tk.Frame(title_row, bg=self.primary, width=4, height=16).pack(side=tk.LEFT, padx=(0, 8))
        if icon:
            tk.Label(
                title_row,
                text=icon,
                bg=self.card_bg,
                fg=self.primary,
                font=("Microsoft YaHei UI", 10, "bold"),
            ).pack(side=tk.LEFT, padx=(0, 6))
        title_label = ttk.Label(title_row, text=title, style="SectionTitle.TLabel")
        title_label.pack(side=tk.LEFT)
        body = ttk.Frame(card, style="CardBody.TFrame", padding=(12, 1, 12, 8))
        body.pack(fill=tk.BOTH, expand=True)
        return card, body

    def _build_parameter_section(self, body) -> None:
        form = ttk.Frame(body, style="CardBody.TFrame")
        form.pack(fill=tk.X)
        form.columnconfigure(0, minsize=72)
        form.columnconfigure(1, weight=1)
        form.columnconfigure(2, minsize=72)
        form.columnconfigure(3, weight=1)
        form.columnconfigure(4, minsize=72)

        self._row_file(form, 0, "主文件", self.base_file, self.pick_base_file)
        self._row_file(form, 1, "对比文件", self.target_file, self.pick_target_file)
        self.base_sql_column_box, self.target_sql_column_box = self._row_dual_select(
            form, 2, "主文件 SQL 列", self.base_sql_column, "对比文件 SQL 列", self.target_sql_column
        )
        self._row_dual_multi_select(form, 3, "主文件显示列", "对比文件显示列")
        self.base_sql_column_box.bind("<<ComboboxSelected>>", self._on_base_sql_column_changed)
        self.target_sql_column_box.bind("<<ComboboxSelected>>", self._on_target_sql_column_changed)
        self._row_dual_text(form, 4, "候选召回 top_k", self.top_k, "每批条数", self.batch_size)
        self._row_file(form, 5, "结果目录", self.result_dir, self.pick_result_dir, directory=True)

    def _build_flow_section(self, body) -> None:
        flow = ttk.Frame(body, style="CardBody.TFrame")
        flow.pack()
        prepare_button = ttk.Button(flow, text="① 生成预处理文件", style="Primary.TButton", command=self.generate_prepare_and_split, takefocus=False)
        prepare_button.grid(row=0, column=0)
        ttk.Label(flow, text="→", style="SectionTitle.TLabel").grid(row=0, column=1, padx=8)
        prompt_button = ttk.Button(flow, text="② 复制 AI 提示词", style="Primary.TButton", command=self.refresh_and_copy_prompt, takefocus=False)
        prompt_button.grid(row=0, column=2)
        ttk.Label(flow, text="→", style="SectionTitle.TLabel").grid(row=0, column=3, padx=8)
        merge_button = ttk.Button(flow, text="③ 合并生成 Excel", style="Primary.TButton", command=self.merge_and_finalize, takefocus=False)
        merge_button.grid(row=0, column=4)
        self.action_buttons.extend([prepare_button, prompt_button, merge_button])

    def _build_quick_actions(self, body) -> None:
        row = ttk.Frame(body, style="CardBody.TFrame")
        row.pack()
        export_skill_button = ttk.Button(row, text="📄 导出 Skill", style="Secondary.TButton", command=self.export_skill_md, takefocus=False)
        export_skill_button.grid(row=0, column=0)
        copy_role_button = ttk.Button(row, text="🧠 复制智能体角色文案", style="Secondary.TButton", command=self.copy_role_md, takefocus=False)
        copy_role_button.grid(row=0, column=1, padx=(8, 0))
        open_dir_button = ttk.Button(row, text="📁 打开结果目录", style="Secondary.TButton", command=self.open_result_dir, takefocus=False)
        open_dir_button.grid(row=0, column=2, padx=(8, 0))
        open_excel_button = ttk.Button(row, text="📗 打开最终 Excel", style="Secondary.TButton", command=self.open_final_excel, takefocus=False)
        open_excel_button.grid(row=0, column=3, padx=(8, 0))
        self.secondary_buttons.extend([export_skill_button, copy_role_button, open_dir_button, open_excel_button])

    def _build_status_section(self, body) -> None:
        self.status_badge = tk.Label(body, text="●", bg=self.card_bg, fg=self.status_color, font=("Microsoft YaHei UI", 14, "bold"))
        self.status_badge.pack(anchor="w")
        self.status_heading_label = ttk.Label(body, textvariable=self.status_heading_var, style="StatusMain.TLabel")
        self.status_heading_label.pack(anchor="w", pady=(1, 4))
        ttk.Label(body, textvariable=self.status_message_var, style="Muted.TLabel", wraplength=150, justify=tk.LEFT).pack(anchor="w", pady=(0, 4))

    def _build_detail_section(self, body) -> None:
        detail_height = 206
        detail_container = tk.Frame(body, bg=self.card_bg, height=detail_height, bd=0, highlightthickness=0)
        detail_container.pack(fill=tk.BOTH, expand=True)
        detail_container.pack_propagate(False)
        detail_container.grid_propagate(False)

        tab_row = ttk.Frame(detail_container, style="CardBody.TFrame")
        tab_row.pack(fill=tk.X, pady=(0, 4))
        self.command_tab_button = ttk.Button(
            tab_row,
            text="命令预览",
            style="DetailTabActive.TButton",
            command=lambda: self._switch_detail_tab("command"),
            takefocus=False,
        )
        self.command_tab_button.pack(side=tk.LEFT)
        self.prompt_tab_button = ttk.Button(
            tab_row,
            text="AI 提示词",
            style="DetailTab.TButton",
            command=lambda: self._switch_detail_tab("prompt"),
            takefocus=False,
        )
        self.prompt_tab_button.pack(side=tk.LEFT, padx=(6, 0))

        content_height = detail_height - 34
        self.detail_content = tk.Frame(detail_container, bg=self.card_bg, height=content_height, bd=0, highlightthickness=0)
        self.detail_content.pack(fill=tk.BOTH, expand=True)
        self.detail_content.pack_propagate(False)
        self.detail_content.grid_propagate(False)

        self.command_tab_frame = ttk.Frame(self.detail_content, style="CardBody.TFrame")
        self.prompt_tab_frame = ttk.Frame(self.detail_content, style="CardBody.TFrame")
        for tab in (self.command_tab_frame, self.prompt_tab_frame):
            tab.configure(height=content_height)
            tab.pack_propagate(False)
            tab.grid_propagate(False)

        self.command_text = self._build_text_panel(self.command_tab_frame, wrap=tk.NONE)
        self.prompt_text = self._build_text_panel(self.prompt_tab_frame, wrap=tk.WORD)
        self._switch_detail_tab("command")

    def _switch_detail_tab(self, tab_name: str) -> None:
        self.detail_tab_var.set(tab_name)
        self.command_tab_button.configure(style="DetailTabActive.TButton" if tab_name == "command" else "DetailTab.TButton")
        self.prompt_tab_button.configure(style="DetailTabActive.TButton" if tab_name == "prompt" else "DetailTab.TButton")
        self.command_tab_frame.pack_forget()
        self.prompt_tab_frame.pack_forget()
        if tab_name == "command":
            self.command_tab_frame.pack(fill=tk.BOTH, expand=True)
        else:
            self.prompt_tab_frame.pack(fill=tk.BOTH, expand=True)

    def _build_text_panel(self, parent, wrap: str):
        container = tk.Frame(parent, bg=self.code_bg, highlightbackground=self.border_color, highlightthickness=1, bd=0)
        container.pack(fill=tk.BOTH, expand=True)
        text = tk.Text(
            container,
            wrap=wrap,
            font=self.code_font,
            bg=self.code_bg,
            fg=self.text_primary,
            bd=0,
            relief=tk.FLAT,
            padx=9,
            pady=7,
            insertbackground=self.text_primary,
        )
        y_scroll = ttk.Scrollbar(container, orient=tk.VERTICAL, command=text.yview)
        text.configure(yscrollcommand=y_scroll.set)
        text.grid(row=0, column=0, sticky="nsew")
        y_scroll.grid(row=0, column=1, sticky="ns")
        container.rowconfigure(0, weight=1)
        container.columnconfigure(0, weight=1)
        if wrap == tk.NONE:
            x_scroll = ttk.Scrollbar(container, orient=tk.HORIZONTAL, command=text.xview)
            text.configure(xscrollcommand=x_scroll.set)
            x_scroll.grid(row=1, column=0, sticky="ew")
        else:
            spacer = tk.Frame(container, bg=self.code_bg, height=16, bd=0, highlightthickness=0)
            spacer.grid(row=1, column=0, sticky="ew")
            spacer.grid_propagate(False)
        self._bind_mousewheel_to_scroll_target(container, text, text if wrap == tk.NONE else None)
        return text

    def _bind_mousewheel_to_scroll_target(self, widget, scroll_target, horizontal_target=None) -> None:
        def on_mousewheel(event):
            if event.delta:
                direction = -1 if event.delta > 0 else 1
                scroll_target.yview_scroll(direction, "units")
            elif getattr(event, "num", None) == 4:
                scroll_target.yview_scroll(-1, "units")
            elif getattr(event, "num", None) == 5:
                scroll_target.yview_scroll(1, "units")
            return "break"

        def on_shift_mousewheel(event):
            if horizontal_target is None:
                return "break"
            if event.delta:
                direction = -1 if event.delta > 0 else 1
                horizontal_target.xview_scroll(direction, "units")
            elif getattr(event, "num", None) == 4:
                horizontal_target.xview_scroll(-1, "units")
            elif getattr(event, "num", None) == 5:
                horizontal_target.xview_scroll(1, "units")
            return "break"

        for target in (widget, scroll_target):
            target.bind("<MouseWheel>", on_mousewheel, add="+")
            target.bind("<Button-4>", on_mousewheel, add="+")
            target.bind("<Button-5>", on_mousewheel, add="+")
            target.bind("<Shift-MouseWheel>", on_shift_mousewheel, add="+")
            target.bind("<Shift-Button-4>", on_shift_mousewheel, add="+")
            target.bind("<Shift-Button-5>", on_shift_mousewheel, add="+")

    def _toggle_display_popup(self, kind: str) -> None:
        if self.active_display_popup_kind == kind and self.active_display_popup is not None:
            self._close_display_popup()
            return
        self._close_display_popup()

        if kind == "base":
            button = self.base_display_menu_button
            options = list(self.base_display_column_vars.items())
            title = "主文件显示列"
        else:
            button = self.target_display_menu_button
            options = list(self.target_display_column_vars.items())
            title = "对比文件显示列"

        if not options:
            return

        popup = tk.Toplevel(self.root)
        popup.withdraw()
        popup.overrideredirect(True)
        popup.configure(bg=self.border_color)

        outer = tk.Frame(popup, bg=self.border_color, bd=0, highlightthickness=0)
        outer.pack(fill=tk.BOTH, expand=True)
        inner = tk.Frame(outer, bg=self.card_bg, bd=0, highlightthickness=0)
        inner.pack(fill=tk.BOTH, expand=True, padx=1, pady=1)

        title_label = tk.Label(
            inner,
            text=title,
            bg=self.card_bg,
            fg=self.text_secondary,
            font=("Microsoft YaHei UI", 9),
            anchor="w",
        )
        title_label.pack(fill=tk.X, padx=10, pady=(8, 4))

        canvas = tk.Canvas(inner, bg=self.card_bg, highlightthickness=0, bd=0)
        scrollbar = ttk.Scrollbar(inner, orient=tk.VERTICAL, command=canvas.yview)
        list_frame = tk.Frame(canvas, bg=self.card_bg, bd=0, highlightthickness=0)
        list_frame.bind(
            "<Configure>",
            lambda e: canvas.configure(scrollregion=canvas.bbox("all")),
        )
        canvas.create_window((0, 0), window=list_frame, anchor="nw")
        canvas.configure(yscrollcommand=scrollbar.set)
        canvas.pack(side=tk.LEFT, fill=tk.BOTH, expand=True, padx=(6, 0), pady=(0, 6))
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y, padx=(0, 6), pady=(0, 6))
        self._bind_mousewheel_to_scroll_target(canvas, canvas)
        self._bind_mousewheel_to_scroll_target(list_frame, canvas)

        for column, var in options:
            check = tk.Checkbutton(
                list_frame,
                text=column,
                variable=var,
                command=self._toggle_base_display_column if kind == "base" else self._toggle_target_display_column,
                anchor="w",
                justify=tk.LEFT,
                bg=self.card_bg,
                activebackground=self.primary_soft,
                fg=self.text_primary,
                selectcolor="#FFFFFF",
                highlightthickness=0,
                bd=0,
                padx=6,
                pady=3,
                takefocus=0,
            )
            check.pack(fill=tk.X, anchor="w")

        self.root.update_idletasks()
        button.update_idletasks()
        popup_width = max(button.winfo_width(), 180)
        popup_height = min(240, max(110, 32 + len(options) * 28))
        x = button.winfo_rootx()
        y = button.winfo_rooty() + button.winfo_height() + 2
        root_left = self.root.winfo_rootx()
        root_right = root_left + self.root.winfo_width()
        popup_width = min(popup_width, max(180, root_right - x - 12))
        popup.geometry(f"{popup_width}x{popup_height}+{x}+{y}")
        canvas.configure(width=popup_width - 22, height=popup_height - 34)

        self.active_display_popup = popup
        self.active_display_popup_kind = kind
        popup.deiconify()

        self.display_popup_click_bind_id = self.root.bind("<Button-1>", self._handle_global_click_for_popup, add="+")

    def _close_display_popup(self) -> None:
        if self.display_popup_click_bind_id:
            self.root.unbind("<Button-1>", self.display_popup_click_bind_id)
            self.display_popup_click_bind_id = None
        if self.active_display_popup is not None:
            try:
                self.active_display_popup.destroy()
            except tk.TclError:
                pass
        self.active_display_popup = None
        self.active_display_popup_kind = None

    def _widget_is_descendant_of(self, widget, ancestor) -> bool:
        current = widget
        while current is not None:
            if current == ancestor:
                return True
            try:
                parent_name = current.winfo_parent()
                if not parent_name:
                    return False
                current = current.nametowidget(parent_name)
            except Exception:
                return False
        return False

    def _handle_global_click_for_popup(self, event) -> None:
        if self.active_display_popup is None:
            return
        button = self.base_display_menu_button if self.active_display_popup_kind == "base" else self.target_display_menu_button
        widget = event.widget
        if self._widget_is_descendant_of(widget, self.active_display_popup):
            return
        if self._widget_is_descendant_of(widget, button):
            return
        self._close_display_popup()

    def _row_file(self, parent, row: int, label: str, var: tk.StringVar, cmd, directory: bool = False, save: bool = False) -> None:
        ttk.Label(parent, text=label, style="FieldLabel.TLabel").grid(row=row, column=0, sticky="w", padx=(0, 8), pady=3)
        ttk.Entry(parent, textvariable=var).grid(row=row, column=1, columnspan=3, sticky="ew", pady=3)
        ttk.Button(parent, text="选择", width=5, style="FilePick.TButton", command=cmd, takefocus=False).grid(row=row, column=4, padx=(6, 0), pady=3, sticky="ew")

    def _row_dual_text(
        self,
        parent,
        row: int,
        left_label: str,
        left_var: tk.StringVar,
        right_label: str,
        right_var: tk.StringVar,
    ) -> None:
        ttk.Label(parent, text=left_label, style="FieldLabel.TLabel").grid(row=row, column=0, sticky="w", padx=(0, 8), pady=3)
        ttk.Entry(parent, textvariable=left_var, width=18).grid(row=row, column=1, sticky="ew", pady=3)
        ttk.Label(parent, text=right_label, style="FieldLabel.TLabel").grid(row=row, column=2, sticky="w", padx=(8, 8), pady=3)
        ttk.Entry(parent, textvariable=right_var, width=18).grid(row=row, column=3, columnspan=2, sticky="ew", pady=3)

    def _row_dual_multi_select(
        self,
        parent,
        row: int,
        left_label: str,
        right_label: str,
    ) -> None:
        ttk.Label(parent, text=left_label, style="FieldLabel.TLabel").grid(row=row, column=0, sticky="w", padx=(0, 8), pady=3)
        left_wrapper = ttk.Frame(parent)
        left_wrapper.grid(row=row, column=1, sticky="ew", pady=3)
        left_wrapper.columnconfigure(0, weight=1)
        self.base_display_menu_button = ttk.Button(
            left_wrapper,
            textvariable=self.base_display_columns_summary,
            style="White.TButton",
            command=lambda: self._toggle_display_popup("base"),
            takefocus=False,
        )
        self.base_display_menu_button.grid(row=0, column=0, sticky="ew")

        ttk.Label(parent, text=right_label, style="FieldLabel.TLabel").grid(row=row, column=2, sticky="w", padx=(8, 8), pady=3)
        right_wrapper = ttk.Frame(parent)
        right_wrapper.grid(row=row, column=3, columnspan=2, sticky="ew", pady=3)
        right_wrapper.columnconfigure(0, weight=1)
        self.target_display_menu_button = ttk.Button(
            right_wrapper,
            textvariable=self.target_display_columns_summary,
            style="White.TButton",
            command=lambda: self._toggle_display_popup("target"),
            takefocus=False,
        )
        self.target_display_menu_button.grid(row=0, column=0, sticky="ew")

    def _row_dual_select(
        self,
        parent,
        row: int,
        left_label: str,
        left_var: tk.StringVar,
        right_label: str,
        right_var: tk.StringVar,
    ) -> tuple[ttk.Combobox, ttk.Combobox]:
        ttk.Label(parent, text=left_label, style="FieldLabel.TLabel").grid(row=row, column=0, sticky="w", padx=(0, 8), pady=3)
        left_combo = ttk.Combobox(parent, textvariable=left_var, width=22, state="readonly", takefocus=False)
        left_combo.grid(row=row, column=1, sticky="ew", pady=3)
        ttk.Label(parent, text=right_label, style="FieldLabel.TLabel").grid(row=row, column=2, sticky="w", padx=(8, 8), pady=3)
        right_combo = ttk.Combobox(parent, textvariable=right_var, width=22, state="readonly", takefocus=False)
        right_combo.grid(row=row, column=3, columnspan=2, sticky="ew", pady=3)
        return left_combo, right_combo

    def _load_saved_column_options(self) -> None:
        try:
            if Path(self.base_file.get().strip()).exists():
                self.load_base_columns()
            if Path(self.target_file.get().strip()).exists():
                self.load_target_columns()
        except Exception:
            pass

    def pick_base_file(self) -> None:
        path = filedialog.askopenfilename(title="选择主文件", filetypes=[("Excel Files", "*.xlsx *.xls *.html *.htm"), ("All Files", "*.*")])
        if path:
            self.base_file.set(path)
            self.load_base_columns()
            self.refresh_all_previews()

    def pick_target_file(self) -> None:
        path = filedialog.askopenfilename(title="选择对比文件", filetypes=[("Excel Files", "*.xlsx *.xls *.html *.htm"), ("All Files", "*.*")])
        if path:
            self.target_file.set(path)
            self.load_target_columns()
            self.refresh_all_previews()

    def pick_result_dir(self) -> None:
        path = filedialog.askdirectory(title="选择结果目录")
        if path:
            self.result_dir.set(path)
            self.refresh_all_previews()

    def _safe_filename_part(self, text: str) -> str:
        value = re.sub(r"[\\\\/:*?\"<>|]+", "_", text.strip())
        value = re.sub(r"\s+", "_", value)
        value = value.strip("._")
        return value or "未命名"

    def _quote(self, text: str) -> str:
        return f'"{text}"'

    def _final_excel_path(self) -> Path:
        result_dir = Path(self.result_dir.get().strip() or ".")
        base_name = self._safe_filename_part(Path(self.base_file.get().strip()).stem) if self.base_file.get().strip() else "主文件"
        target_name = self._safe_filename_part(Path(self.target_file.get().strip()).stem) if self.target_file.get().strip() else "对比文件"
        return result_dir / f"对比结果_{base_name}vs{target_name}.xlsx"

    def _prepare_command(self) -> str:
        base_sql_column = self.base_sql_column.get().strip()
        target_sql_column = self.target_sql_column.get().strip()
        cmd = [
            "python",
            "compare_sql_files.py",
            "--mode", "prepare",
            "--base", self._quote(self.base_file.get().strip()),
            "--target", self._quote(self.target_file.get().strip()),
        ]
        if base_sql_column:
            cmd.extend(["--base-sql-column", self._quote(base_sql_column)])
        if target_sql_column:
            cmd.extend(["--target-sql-column", self._quote(target_sql_column)])
        base_display_columns = self.get_selected_base_display_columns()
        target_display_columns = self.get_selected_target_display_columns()
        if base_display_columns:
            cmd.extend(["--base-display-columns", self._quote(",".join(base_display_columns))])
        if target_display_columns:
            cmd.extend(["--target-display-columns", self._quote(",".join(target_display_columns))])
        cmd.extend([
            "--top-k", self.top_k.get().strip() or "3",
            "--output", self._quote(str(Path(self.result_dir.get().strip() or ".") / "prepare.json")),
        ])
        return " ".join(cmd)

    def _split_command(self) -> str:
        result_dir = Path(self.result_dir.get().strip() or ".")
        return " ".join([
            "python",
            "compare_sql_files.py",
            "--mode", "split-prepare",
            "--prepared", self._quote(str(result_dir / "prepare.json")),
            "--batch-size", self.batch_size.get().strip() or "20",
            "--output-dir", self._quote(str(result_dir / "prepare_parts")),
        ])

    def _merge_command(self) -> str:
        result_dir = Path(self.result_dir.get().strip() or ".")
        return (
            'python compare_sql_files.py --mode merge-review-results '
            '--review-results-files "result/review_results_part_1.json" "result/review_results_part_2.json" '
            f'--output {self._quote(str(result_dir / "review_results.json"))}'
        )

    def _finalize_command(self) -> str:
        result_dir = Path(self.result_dir.get().strip() or ".")
        return " ".join([
            "python",
            "compare_sql_files.py",
            "--mode", "finalize",
            "--prepared", self._quote(str(result_dir / "prepare.json")),
            "--review-results", self._quote(str(result_dir / "review_results.json")),
            "--output", self._quote(str(self._final_excel_path())),
        ])

    def refresh_command_preview(self) -> None:
        self.command_text.delete("1.0", tk.END)
        text = (
            "Prepare 命令：\n"
            f"{self._prepare_command()}\n\n"
            "拆分命令：\n"
            f"{self._split_command()}\n\n"
            "合并命令：\n"
            f"{self._merge_command()}\n\n"
            "Finalize 命令：\n"
            f"{self._finalize_command()}\n"
        )
        self.command_text.insert(tk.END, text)

    def refresh_all_previews(self) -> None:
        self.refresh_prompt_preview()
        self.refresh_command_preview()

    def _set_status_card(self, state: str, heading: str, message: str, file_path: Path | None = None) -> None:
        color_map = {
            "idle": self.text_secondary,
            "running": self.running,
            "success": self.success,
            "error": self.error,
        }
        self.status_color = color_map.get(state, self.text_secondary)
        self.status_badge.configure(fg=self.status_color)
        self.status_heading_var.set(heading)
        self.status_message_var.set(message)

    def _set_status(self, message: str) -> None:
        self.status_var.set(message)
        self.root.update_idletasks()

    def _set_busy(self, busy: bool, message: str | None = None) -> None:
        self.is_busy = busy
        state = tk.DISABLED if busy else tk.NORMAL
        for button in self.action_buttons:
            button.config(state=state)
        if message is not None:
            self._set_status(message)
            self._set_status_card("running" if busy else "idle", "执行中" if busy else "待命", message)

    def _run_in_background(self, worker, start_message: str, success_message: str, success_detail: str) -> None:
        if self.is_busy:
            return

        self._set_busy(True, start_message)

        def job() -> None:
            try:
                detail = worker()
                self.root.after(0, lambda: self._background_success(success_message, success_detail.format(detail=detail)))
            except Exception as exc:
                error_text = f"{exc}\n\n{traceback.format_exc()}"
                self.root.after(0, lambda: self._background_error(error_text))

        threading.Thread(target=job, daemon=True).start()

    def _background_success(self, status_message: str, dialog_message: str) -> None:
        self._set_busy(False, status_message)
        self._set_status_card("success", "已完成", status_message)
        self.refresh_all_previews()
        messagebox.showinfo("完成", dialog_message)

    def _background_error(self, error_text: str) -> None:
        self._set_busy(False, "执行失败")
        self._set_status_card("error", "执行失败", "请检查输入文件、参数配置或结果文件完整性。")
        messagebox.showerror("错误", error_text)

    def _copy_text(self, text: str, success_message: str) -> None:
        self.root.clipboard_clear()
        self.root.clipboard_append(text)
        self.root.update()
        self._set_status(success_message)
        messagebox.showinfo("完成", success_message)

    def _read_packaged_text(self, filename: str, fallback: str) -> str:
        candidate = Path(__file__).with_name(filename)
        try:
            if candidate.exists():
                return candidate.read_text(encoding="utf-8")
        except Exception:
            pass
        return fallback

    def export_skill_md(self) -> None:
        suggested_path = Path(self.result_dir.get().strip() or ".") / "skill.md"
        output_path = filedialog.asksaveasfilename(
            title="导出 skill.md",
            initialfile=suggested_path.name,
            initialdir=str(suggested_path.parent),
            defaultextension=".md",
            filetypes=[("Markdown Files", "*.md"), ("All Files", "*.*")],
        )
        if not output_path:
            return
        content = self._read_packaged_text("skill.md", DEFAULT_SKILL_MD)
        Path(output_path).write_text(content, encoding="utf-8")
        self._set_status(f"已导出 skill.md：{output_path}")
        self._set_status_card("success", "已完成", "skill.md 已导出。", Path(output_path))
        messagebox.showinfo("完成", f"已导出 skill.md：\n{output_path}")

    def copy_role_md(self) -> None:
        content = self._read_packaged_text("智能体角色定义.md", DEFAULT_ROLE_MD).strip()
        self._copy_text(content, "智能体角色文案已复制")

    def _open_path(self, path: Path) -> None:
        if not path.exists():
            raise RuntimeError(f"路径不存在: {path}")
        if sys.platform.startswith("darwin"):
            subprocess.Popen(["open", str(path)])
        elif os.name == "nt":
            os.startfile(str(path))
        else:
            subprocess.Popen(["xdg-open", str(path)])

    def open_result_dir(self) -> None:
        try:
            self._open_path(Path(self.result_dir.get().strip()))
        except Exception as exc:
            messagebox.showerror("错误", f"{exc}\n\n{traceback.format_exc()}")

    def open_final_excel(self) -> None:
        try:
            self._open_path(self._final_excel_path())
        except Exception as exc:
            messagebox.showerror("错误", f"{exc}\n\n{traceback.format_exc()}")

    def _set_combobox_values(self, combo: ttk.Combobox, values: List[str]) -> None:
        combo["values"] = values

    def _auto_pick_sql_column(self, columns: List[str]) -> str:
        try:
            return workflow.pick_sql_column(columns, None)
        except Exception:
            return columns[0] if columns else ""

    def get_selected_base_display_columns(self) -> List[str]:
        return [
            column
            for column, var in self.base_display_column_vars.items()
            if var.get()
        ]

    def get_selected_target_display_columns(self) -> List[str]:
        return [
            column
            for column, var in self.target_display_column_vars.items()
            if var.get()
        ]

    def _update_base_display_columns_summary(self) -> None:
        selected = self.get_selected_base_display_columns()
        if not selected:
            self.base_display_columns_summary.set("未选择显示列")
        else:
            self.base_display_columns_summary.set(f"已选择 {len(selected)} 列")

    def _update_target_display_columns_summary(self) -> None:
        selected = self.get_selected_target_display_columns()
        if not selected:
            self.target_display_columns_summary.set("未选择显示列")
        else:
            self.target_display_columns_summary.set(f"已选择 {len(selected)} 列")

    def _toggle_base_display_column(self) -> None:
        self._update_base_display_columns_summary()
        self.refresh_all_previews()

    def _toggle_target_display_column(self) -> None:
        self._update_target_display_columns_summary()
        self.refresh_all_previews()

    def _on_base_sql_column_changed(self, _event=None) -> None:
        self._refresh_base_display_column_options()
        self.refresh_all_previews()

    def _on_target_sql_column_changed(self, _event=None) -> None:
        self._refresh_target_display_column_options()
        self.refresh_all_previews()

    def _refresh_base_display_column_options(self) -> None:
        selected_before = set(self.get_selected_base_display_columns()) | set(self.saved_base_display_columns)
        current_sql = self.base_sql_column.get().strip()
        selectable_columns = [
            column
            for column in self.base_columns
            if column and column != current_sql
        ]
        self.base_display_column_vars = {}
        for column in selectable_columns:
            var = tk.BooleanVar(value=column in selected_before)
            self.base_display_column_vars[column] = var
        self.saved_base_display_columns = self.get_selected_base_display_columns()
        self._update_base_display_columns_summary()
        if not selectable_columns:
            self.base_display_columns_summary.set("无可选显示列")
        if self.active_display_popup_kind == "base":
            self._close_display_popup()
            self.root.after(0, lambda: self._toggle_display_popup("base"))

    def _refresh_target_display_column_options(self) -> None:
        selected_before = set(self.get_selected_target_display_columns()) | set(self.saved_target_display_columns)
        current_sql = self.target_sql_column.get().strip()
        selectable_columns = [
            column
            for column in self.target_columns
            if column and column != current_sql
        ]
        self.target_display_column_vars = {}
        for column in selectable_columns:
            var = tk.BooleanVar(value=column in selected_before)
            self.target_display_column_vars[column] = var
        self.saved_target_display_columns = self.get_selected_target_display_columns()
        self._update_target_display_columns_summary()
        if not selectable_columns:
            self.target_display_columns_summary.set("无可选显示列")
        if self.active_display_popup_kind == "target":
            self._close_display_popup()
            self.root.after(0, lambda: self._toggle_display_popup("target"))

    def load_base_columns(self) -> None:
        path = Path(self.base_file.get().strip())
        if not path.exists():
            return
        df = workflow.load_table(path)
        self.base_columns = [str(col).strip() for col in df.columns if str(col).strip()]
        self._set_combobox_values(self.base_sql_column_box, self.base_columns)
        if self.base_sql_column.get().strip() not in self.base_columns:
            self.base_sql_column.set(self._auto_pick_sql_column(self.base_columns))
        self._refresh_base_display_column_options()

    def load_target_columns(self) -> None:
        path = Path(self.target_file.get().strip())
        if not path.exists():
            return
        df = workflow.load_table(path)
        self.target_columns = [str(col).strip() for col in df.columns if str(col).strip()]
        self._set_combobox_values(self.target_sql_column_box, self.target_columns)
        if self.target_sql_column.get().strip() not in self.target_columns:
            self.target_sql_column.set(self._auto_pick_sql_column(self.target_columns))
        self._refresh_target_display_column_options()

    def pick_review_result_files(self) -> None:
        paths = filedialog.askopenfilenames(title="选择 review_results_part 文件", filetypes=[("JSON Files", "*.json"), ("All Files", "*.*")])
        if paths:
            self.review_result_files = [Path(path) for path in paths]
            self._render_review_result_files()

    def auto_detect_review_result_files(self) -> None:
        result_dir = Path(self.result_dir.get().strip())
        files = sorted(result_dir.glob("review_results_part_*.json"))
        if not files:
            files = sorted((result_dir / "prepare_parts").glob("review_results_part_*.json"))
        self.review_result_files = files
        self._render_review_result_files()
        if files:
            self._set_status(f"已自动发现 {len(files)} 个 review_results_part 文件")
        else:
            self._set_status("未发现 review_results_part 文件")

    def _render_review_result_files(self) -> None:
        pass

    def _extract_part_no(self, path: Path) -> int | None:
        match = re.search(r"review_results_part_(\d+)\.json$", path.name)
        if not match:
            match = re.search(r"prepare_part_(\d+)\.json$", path.name)
        if not match:
            return None
        return int(match.group(1))

    def _validate_review_parts(self, result_dir: Path) -> None:
        prepare_parts_dir = result_dir / "prepare_parts"
        prepare_parts = sorted(prepare_parts_dir.glob("prepare_part_*.json"))
        if not prepare_parts:
            raise RuntimeError("未找到 prepare_part 文件，请先执行第 1 步生成 Prepare 并拆分")

        expected_parts = sorted(
            part_no for part_no in
            (self._extract_part_no(path) for path in prepare_parts)
            if part_no is not None
        )
        found_parts = sorted(
            part_no for part_no in
            (self._extract_part_no(path) for path in self.review_result_files)
            if part_no is not None
        )

        missing_parts = [part_no for part_no in expected_parts if part_no not in found_parts]
        extra_parts = [part_no for part_no in found_parts if part_no not in expected_parts]

        if missing_parts:
            raise RuntimeError(f"缺少 review_results_part 文件：第 {', '.join(map(str, missing_parts))} 批")
        if extra_parts:
            raise RuntimeError(f"发现多余的 review_results_part 文件：第 {', '.join(map(str, extra_parts))} 批")

    def _persist_settings(self) -> None:
        self.saved_base_display_columns = self.get_selected_base_display_columns()
        self.saved_target_display_columns = self.get_selected_target_display_columns()
        save_settings(
            {
                "base_file": self.base_file.get().strip(),
                "target_file": self.target_file.get().strip(),
                "base_sql_column": self.base_sql_column.get().strip(),
                "target_sql_column": self.target_sql_column.get().strip(),
                "base_display_columns": self.saved_base_display_columns,
                "target_display_columns": self.saved_target_display_columns,
                "top_k": int(self.top_k.get().strip() or "3"),
                "batch_size": int(self.batch_size.get().strip() or "20"),
                "result_dir": self.result_dir.get().strip(),
            }
        )

    def _build_prepare_payload(self) -> dict:
        base_path = Path(self.base_file.get().strip())
        target_path = Path(self.target_file.get().strip())
        selected_base_display_columns = self.get_selected_base_display_columns()
        selected_target_display_columns = self.get_selected_target_display_columns()
        base_records, base_sql_col = workflow.load_records(
            base_path,
            self.base_sql_column.get().strip() or None,
            selected_base_display_columns,
        )
        target_records, target_sql_col = workflow.load_records(
            target_path,
            self.target_sql_column.get().strip() or None,
            selected_target_display_columns,
        )
        return workflow.build_prepare_payload(
            base_path,
            target_path,
            base_records,
            target_records,
            base_sql_col,
            target_sql_col,
            selected_base_display_columns,
            selected_target_display_columns,
            int(self.top_k.get().strip() or "3"),
        )

    def generate_prepare_and_split(self) -> None:
        def worker() -> str:
            self._persist_settings()
            result_dir = Path(self.result_dir.get().strip())
            prepare_path = result_dir / "prepare.json"
            parts_dir = result_dir / "prepare_parts"

            self.root.after(0, lambda: self._set_status("正在读取主文件和对比文件..."))
            payload = self._build_prepare_payload()
            self.root.after(0, lambda: self._set_status("正在写出 prepare.json..."))
            workflow.write_json(prepare_path, payload)
            self.root.after(0, lambda: self._set_status("正在拆分 prepare_part 文件..."))
            parts = workflow.split_prepare_payload(payload, int(self.batch_size.get().strip() or "20"))
            parts_dir.mkdir(parents=True, exist_ok=True)
            for part in parts:
                workflow.write_json(parts_dir / f"prepare_part_{int(part['part_no'])}.json", part)
            return f"review_tasks={len(payload['review_tasks'])}，拆分为 {len(parts)} 批"

        self._run_in_background(
            worker,
            "正在生成 Prepare，请稍候...",
            "已生成 prepare.json",
            "已生成 prepare.json，{detail}。",
        )

    def merge_and_finalize(self) -> None:
        try:
            self._persist_settings()
            result_dir = Path(self.result_dir.get().strip())
            prepare_path = result_dir / "prepare.json"
            review_results_path = result_dir / "review_results.json"
            final_excel_path = self._final_excel_path()

            if not self.review_result_files:
                self._set_status("正在自动扫描 review_results_part 文件...")
                self.auto_detect_review_result_files()
            if not self.review_result_files:
                raise RuntimeError("未选择 review_results_part 文件")

            selected_names = "；".join(path.name for path in self.review_result_files)
            self._set_status(f"已识别 review_results_part 文件：{selected_names}")
            self._set_status("正在检查 review_results_part 是否齐全...")
            self._validate_review_parts(result_dir)
            self._set_status("正在合并 review_results_part 文件...")
            merged = workflow.merge_review_result_files(self.review_result_files)
            workflow.write_json(review_results_path, {"review_results": merged})

            self._set_status("正在生成最终 Excel...")
            prepared = workflow.load_json(prepare_path)
            payload = workflow.build_finalize_payload(prepared, merged)
            json_output_path = final_excel_path.with_suffix(".json")
            workflow.write_json(json_output_path, payload)
            workflow.export_excel(payload, final_excel_path)

            self._set_status(f"已生成 {final_excel_path}")
            self._set_status_card("success", "已完成", "最终 Excel 已生成。", final_excel_path)
            messagebox.showinfo("完成", f"已生成最终 Excel：\n{final_excel_path}")
        except Exception as exc:
            self._set_status("合并或导出失败")
            self._set_status_card("error", "执行失败", "合并 review 结果或导出 Excel 时发生错误。")
            messagebox.showerror("错误", f"{exc}\n\n{traceback.format_exc()}")

    def refresh_prompt_preview(self) -> None:
        self.prompt_text.delete("1.0", tk.END)
        result_dir = Path(self.result_dir.get().strip() or ".")
        parts_dir = result_dir / "prepare_parts"
        parts = sorted(parts_dir.glob("prepare_part_*.json"))

        if not parts:
            self.prompt_text.insert(
                tk.END,
                "先点击“生成 Prepare 并拆分”，软件会按实际批次数自动生成 prepare_part_x.json，之后这里会出现每一批给智能体的提示词。\n",
            )
            return
        input_lines = "\n".join(f"- {path}" for path in parts)
        output_lines = "\n".join(
            f"- {result_dir / f'review_results_part_{idx}.json'}" for idx in range(1, len(parts) + 1)
        )
        prompt = (
            f"请按顺序一次性处理全部 {len(parts)} 个批次的 SQL 语义比对任务。\n\n"
            "输入文件：\n"
            f"{input_lines}\n\n"
            "输出文件：\n"
            f"{output_lines}\n\n"
            "处理要求：\n"
            "1. 按顺序逐个处理，每次只处理一个 prepare_part。\n"
            "2. 先完整处理 prepare_part_1.json，再处理 prepare_part_2.json，再处理 prepare_part_3.json，以此类推。\n"
            "3. 每个 prepare_part 单独生成对应的 review_results_part_x.json。\n"
            "4. 不要重新读取原始 Excel。\n"
            "5. 不要重新做 prepare。\n"
            "6. 不要新写任何 Python 脚本。\n"
            "7. 不要生成 generate_results.py 或任何辅助脚本。\n"
            "8. 每个 prepare_part 都要先读取 expected_result_count 和 expected_pair_ids。\n"
            "9. 每个 pair_id 只输出一条结果。\n"
            "10. 每处理完一个 prepare_part，先完整输出该 part 的 review_results_part_x.json，再继续下一个。\n"
            "11. 每个 prepare_part 的输出结果条数必须等于 expected_result_count。\n"
            "12. 输出的 pair_id 必须且只能来自当前 prepare_part 的 expected_pair_ids。\n"
            "13. 禁止新增、遗漏、重复、改写 pair_id。\n"
            "14. 如果单个 part 输出过长，请继续补完这个 part，不要跳到下一个，更不要改成写脚本。\n"
            "15. 不要把多个 part 的结果混在一个 JSON 里。\n"
            "16. 每个 review_results_part_x.json 都必须是完整合法 JSON，不允许截断。\n"
            "17. 全部 part 完成后，再告诉我“所有 review_results_part 已完成”。\n\n"
            "字段必须包含：\n"
            "pair_id, judgement, confidence, semantic_score, same_business, reasoning, "
            "join_change, where_change, group_by_change, subquery_change, base_line_refs, "
            "target_line_refs, common_tables, key_differences\n\n"
            "如果一次回复放不下，请继续输出，不要中断任务，不要改成写脚本。\n"
        )
        self.prompt_text.insert(tk.END, prompt)

    def refresh_and_copy_prompt(self) -> None:
        self.refresh_prompt_preview()
        self.copy_prompt()

    def copy_prompt(self) -> None:
        prompt = self.prompt_text.get("1.0", tk.END).strip()
        if not prompt:
            messagebox.showwarning("提示", "当前没有可复制的提示词。")
            return
        self._copy_text(prompt, "提示词已复制，可直接粘贴到智能体窗口。")


def main() -> None:
    root = tk.Tk()
    app = SqlSemanticWorkflowApp(root)
    root.mainloop()


if __name__ == "__main__":
    main()
