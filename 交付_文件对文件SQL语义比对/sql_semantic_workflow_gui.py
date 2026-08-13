#!/usr/bin/env python3
import json
import os
import re
import subprocess
import sys
import threading
import traceback
from pathlib import Path
from typing import List

import tkinter as tk
from tkinter import filedialog, messagebox, ttk

import compare_sql_files as workflow


SETTINGS_FILE = Path(__file__).with_name("sql_semantic_workflow_gui_settings.json")


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
        self.root.geometry("600x800")
        self.settings = load_settings()

        self.base_file = tk.StringVar(value=self.settings.get("base_file", ""))
        self.target_file = tk.StringVar(value=self.settings.get("target_file", ""))
        self.base_sql_column = tk.StringVar(value=self.settings.get("base_sql_column", ""))
        self.target_sql_column = tk.StringVar(value=self.settings.get("target_sql_column", ""))
        self.base_id_column = tk.StringVar(value=self.settings.get("base_id_column", "任务编号"))
        self.top_k = tk.StringVar(value=str(self.settings.get("top_k", 3)))
        self.batch_size = tk.StringVar(value=str(self.settings.get("batch_size", 20)))
        self.result_dir = tk.StringVar(value=self.settings.get("result_dir", str(Path.cwd() / "result")))
        self.base_columns: List[str] = []
        self.target_columns: List[str] = []

        self.review_result_files: List[Path] = []
        self.is_busy = False
        self.action_buttons: List[ttk.Button] = []
        self.secondary_buttons: List[ttk.Button] = []

        self._build_ui()
        self._load_saved_column_options()
        self.refresh_prompt_preview()

    def _build_ui(self) -> None:
        frame = ttk.Frame(self.root, padding=12)
        frame.pack(fill=tk.BOTH, expand=True)

        form = ttk.LabelFrame(frame, text="输入参数", padding=10)
        form.pack(fill=tk.X)

        self._row_file(form, 0, "主文件", self.base_file, self.pick_base_file)
        self._row_file(form, 1, "对比文件", self.target_file, self.pick_target_file)
        self.base_sql_column_box = self._row_select(form, 2, "主文件 SQL 列", self.base_sql_column)
        self.target_sql_column_box = self._row_select(form, 3, "对比文件 SQL 列", self.target_sql_column)
        self.base_id_column_box = self._row_select(form, 4, "主文件主键列", self.base_id_column)
        self._row_text(form, 5, "候选召回 top_k", self.top_k)
        self._row_text(form, 6, "每批条数", self.batch_size)
        self._row_file(form, 7, "结果目录", self.result_dir, self.pick_result_dir, directory=True)

        actions = ttk.LabelFrame(frame, text="操作", padding=10)
        actions.pack(fill=tk.X, pady=(12, 0))

        prepare_button = ttk.Button(actions, text="1. 生成 Prepare 并拆分", command=self.generate_prepare_and_split)
        prepare_button.grid(row=0, column=0, padx=6, pady=6, sticky="w")
        prompt_button = ttk.Button(actions, text="2. 刷新并复制提示词", command=self.refresh_and_copy_prompt)
        prompt_button.grid(row=0, column=1, padx=6, pady=6, sticky="w")
        merge_button = ttk.Button(actions, text="3. 合并 review_results_part 生成 Excel", command=self.merge_and_finalize)
        merge_button.grid(row=0, column=2, padx=6, pady=6, sticky="w")
        self.action_buttons.extend([prepare_button, prompt_button, merge_button])
        second_row = ttk.Frame(actions)
        second_row.grid(row=1, column=0, columnspan=3, sticky="w", padx=6, pady=(6, 0))
        open_dir_button = ttk.Button(second_row, text="打开结果目录", command=self.open_result_dir)
        open_dir_button.pack(side=tk.LEFT, padx=(0, 8))
        open_excel_button = ttk.Button(second_row, text="打开最终 Excel", command=self.open_final_excel)
        open_excel_button.pack(side=tk.LEFT, padx=(0, 8))
        copy_prepare_button = ttk.Button(second_row, text="复制 Prepare 命令", command=self.copy_prepare_command)
        copy_prepare_button.pack(side=tk.LEFT, padx=(0, 8))
        copy_finalize_button = ttk.Button(second_row, text="复制 Finalize 命令", command=self.copy_finalize_command)
        copy_finalize_button.pack(side=tk.LEFT)
        self.secondary_buttons.extend([open_dir_button, open_excel_button, copy_prepare_button, copy_finalize_button])

        info = ttk.LabelFrame(frame, text="状态", padding=10)
        info.pack(fill=tk.X, pady=(12, 0))
        self.status_var = tk.StringVar(value="待开始")
        ttk.Label(info, textvariable=self.status_var).pack(anchor="w")
        self.progress = ttk.Progressbar(info, mode="indeterminate")
        self.progress.pack(fill=tk.X, pady=(8, 0))

        command_frame = ttk.LabelFrame(frame, text="命令预览", padding=10)
        command_frame.pack(fill=tk.X, pady=(12, 0))
        self.command_text = tk.Text(command_frame, height=8, wrap=tk.WORD)
        self.command_text.pack(fill=tk.X)

        prompt_frame = ttk.LabelFrame(frame, text="给智能体的提示词", padding=10)
        prompt_frame.pack(fill=tk.BOTH, expand=True, pady=(12, 0))
        self.prompt_text = tk.Text(prompt_frame, wrap=tk.WORD)
        self.prompt_text.pack(fill=tk.BOTH, expand=True)

    def _row_file(self, parent: ttk.LabelFrame, row: int, label: str, var: tk.StringVar, cmd, directory: bool = False, save: bool = False) -> None:
        ttk.Label(parent, text=label).grid(row=row, column=0, sticky="w", padx=(0, 8), pady=4)
        ttk.Entry(parent, textvariable=var, width=100).grid(row=row, column=1, sticky="ew", pady=4)
        ttk.Button(parent, text="选择", command=cmd).grid(row=row, column=2, padx=(8, 0), pady=4)
        parent.columnconfigure(1, weight=1)

    def _row_text(self, parent: ttk.LabelFrame, row: int, label: str, var: tk.StringVar) -> None:
        ttk.Label(parent, text=label).grid(row=row, column=0, sticky="w", padx=(0, 8), pady=4)
        ttk.Entry(parent, textvariable=var, width=30).grid(row=row, column=1, sticky="w", pady=4)

    def _row_select(self, parent: ttk.LabelFrame, row: int, label: str, var: tk.StringVar) -> ttk.Combobox:
        ttk.Label(parent, text=label).grid(row=row, column=0, sticky="w", padx=(0, 8), pady=4)
        combo = ttk.Combobox(parent, textvariable=var, width=40, state="readonly")
        combo.grid(row=row, column=1, sticky="w", pady=4)
        return combo

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

    def _quote(self, text: str) -> str:
        return f'"{text}"'

    def _final_excel_path(self) -> Path:
        return Path(self.result_dir.get().strip() or ".") / "final_result.xlsx"

    def _prepare_command(self) -> str:
        base_sql_column = self.base_sql_column.get().strip()
        target_sql_column = self.target_sql_column.get().strip()
        base_id_column = self.base_id_column.get().strip()
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
        if base_id_column:
            cmd.extend(["--base-id-column", self._quote(base_id_column)])
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

    def _set_status(self, message: str) -> None:
        self.status_var.set(message)
        self.root.update_idletasks()

    def _set_busy(self, busy: bool, message: str | None = None) -> None:
        self.is_busy = busy
        state = tk.DISABLED if busy else tk.NORMAL
        for button in self.action_buttons:
            button.config(state=state)
        if busy:
            self.progress.start(10)
        else:
            self.progress.stop()
        if message is not None:
            self._set_status(message)

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
        self.refresh_all_previews()
        messagebox.showinfo("完成", dialog_message)

    def _background_error(self, error_text: str) -> None:
        self._set_busy(False, "执行失败")
        messagebox.showerror("错误", error_text)

    def _copy_text(self, text: str, success_message: str) -> None:
        self.root.clipboard_clear()
        self.root.clipboard_append(text)
        self.root.update()
        self.status_var.set(success_message)
        messagebox.showinfo("完成", success_message)

    def copy_prepare_command(self) -> None:
        self._copy_text(self._prepare_command(), "Prepare 命令已复制")

    def copy_finalize_command(self) -> None:
        self._copy_text(self._finalize_command(), "Finalize 命令已复制")

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

    def _auto_pick_id_column(self, columns: List[str]) -> str:
        preferred = ["任务编号", "技术治理单号", "治理单号", "需求ID", "需求编号", "ID", "id"]
        for name in preferred:
            if name in columns:
                return name
        return columns[0] if columns else ""

    def load_base_columns(self) -> None:
        path = Path(self.base_file.get().strip())
        if not path.exists():
            return
        df = workflow.load_table(path)
        self.base_columns = [str(col).strip() for col in df.columns if str(col).strip()]
        self._set_combobox_values(self.base_sql_column_box, self.base_columns)
        self._set_combobox_values(self.base_id_column_box, [""] + self.base_columns)
        if self.base_sql_column.get().strip() not in self.base_columns:
            self.base_sql_column.set(self._auto_pick_sql_column(self.base_columns))
        if self.base_id_column.get().strip() not in self.base_columns:
            self.base_id_column.set(self._auto_pick_id_column(self.base_columns))

    def load_target_columns(self) -> None:
        path = Path(self.target_file.get().strip())
        if not path.exists():
            return
        df = workflow.load_table(path)
        self.target_columns = [str(col).strip() for col in df.columns if str(col).strip()]
        self._set_combobox_values(self.target_sql_column_box, self.target_columns)
        if self.target_sql_column.get().strip() not in self.target_columns:
            self.target_sql_column.set(self._auto_pick_sql_column(self.target_columns))

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
            self.status_var.set(f"已自动发现 {len(files)} 个 review_results_part 文件")
        else:
            self.status_var.set("未发现 review_results_part 文件")

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
        save_settings(
            {
                "base_file": self.base_file.get().strip(),
                "target_file": self.target_file.get().strip(),
                "base_sql_column": self.base_sql_column.get().strip(),
                "target_sql_column": self.target_sql_column.get().strip(),
                "base_id_column": self.base_id_column.get().strip(),
                "top_k": int(self.top_k.get().strip() or "3"),
                "batch_size": int(self.batch_size.get().strip() or "20"),
                "result_dir": self.result_dir.get().strip(),
            }
        )

    def _build_prepare_payload(self) -> dict:
        base_path = Path(self.base_file.get().strip())
        target_path = Path(self.target_file.get().strip())
        base_records, base_sql_col = workflow.load_records(base_path, self.base_sql_column.get().strip() or None, self.base_id_column.get().strip() or None)
        target_records, target_sql_col = workflow.load_records(target_path, self.target_sql_column.get().strip() or None, None)
        return workflow.build_prepare_payload(
            base_path,
            target_path,
            base_records,
            target_records,
            base_sql_col,
            target_sql_col,
            self.base_id_column.get().strip() or None,
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
            messagebox.showinfo("完成", f"已生成最终 Excel：\n{final_excel_path}")
        except Exception as exc:
            self._set_status("合并或导出失败")
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
            "8. 每个 pair_id 只输出一条结果。\n"
            "9. 每处理完一个 prepare_part，先完整输出该 part 的 review_results_part_x.json，再继续下一个。\n"
            "10. 如果单个 part 输出过长，请继续补完这个 part，不要跳到下一个，更不要改成写脚本。\n"
            "11. 不要把多个 part 的结果混在一个 JSON 里。\n"
            "12. 只允许输出当前 prepare_part 文件中的 pair_id，禁止输出其他 part 的 pair_id。\n"
            "13. 每个 review_results_part_x.json 都必须是完整合法 JSON，不允许截断。\n"
            "14. 全部 part 完成后，再告诉我“所有 review_results_part 已完成”。\n\n"
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
