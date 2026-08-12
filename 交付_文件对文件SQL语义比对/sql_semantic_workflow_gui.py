#!/usr/bin/env python3
import json
import os
import subprocess
import sys
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
        self.root.geometry("1180x920")
        self.settings = load_settings()

        self.base_file = tk.StringVar(value=self.settings.get("base_file", ""))
        self.target_file = tk.StringVar(value=self.settings.get("target_file", ""))
        self.base_sql_column = tk.StringVar(value=self.settings.get("base_sql_column", ""))
        self.target_sql_column = tk.StringVar(value=self.settings.get("target_sql_column", ""))
        self.base_id_column = tk.StringVar(value=self.settings.get("base_id_column", "任务编号"))
        self.top_k = tk.StringVar(value=str(self.settings.get("top_k", 3)))
        self.batch_size = tk.StringVar(value=str(self.settings.get("batch_size", 20)))
        self.result_dir = tk.StringVar(value=self.settings.get("result_dir", str(Path.cwd() / "result")))
        self.final_excel = tk.StringVar(value=self.settings.get("final_excel", str(Path.cwd() / "result" / "result.xlsx")))

        self.review_result_files: List[Path] = []

        self._build_ui()
        self.refresh_prompt_preview()

    def _build_ui(self) -> None:
        frame = ttk.Frame(self.root, padding=12)
        frame.pack(fill=tk.BOTH, expand=True)

        form = ttk.LabelFrame(frame, text="输入参数", padding=10)
        form.pack(fill=tk.X)

        self._row_file(form, 0, "主文件", self.base_file, self.pick_base_file)
        self._row_file(form, 1, "对比文件", self.target_file, self.pick_target_file)
        self._row_text(form, 2, "主文件 SQL 列", self.base_sql_column)
        self._row_text(form, 3, "对比文件 SQL 列", self.target_sql_column)
        self._row_text(form, 4, "主文件主键列", self.base_id_column)
        self._row_text(form, 5, "候选召回 top_k", self.top_k)
        self._row_text(form, 6, "每批条数", self.batch_size)
        self._row_file(form, 7, "结果目录", self.result_dir, self.pick_result_dir, directory=True)
        self._row_file(form, 8, "最终 Excel", self.final_excel, self.pick_final_excel, save=True)

        actions = ttk.LabelFrame(frame, text="操作", padding=10)
        actions.pack(fill=tk.X, pady=(12, 0))

        ttk.Button(actions, text="查看主文件列名", command=self.preview_base_columns).grid(row=0, column=0, padx=6, pady=6, sticky="w")
        ttk.Button(actions, text="查看对比文件列名", command=self.preview_target_columns).grid(row=0, column=1, padx=6, pady=6, sticky="w")
        ttk.Button(actions, text="1. 生成 Prepare 并拆分", command=self.generate_prepare_and_split).grid(row=0, column=2, padx=6, pady=6, sticky="w")
        ttk.Button(actions, text="2. 选择 review_results_part 文件", command=self.pick_review_result_files).grid(row=0, column=3, padx=6, pady=6, sticky="w")
        ttk.Button(actions, text="3. 自动发现 review_results_part 文件", command=self.auto_detect_review_result_files).grid(row=0, column=4, padx=6, pady=6, sticky="w")
        ttk.Button(actions, text="4. 合并并生成 Excel", command=self.merge_and_finalize).grid(row=0, column=5, padx=6, pady=6, sticky="w")
        ttk.Button(actions, text="刷新提示词", command=self.refresh_prompt_preview).grid(row=0, column=6, padx=6, pady=6, sticky="w")
        ttk.Button(actions, text="复制提示词", command=self.copy_prompt).grid(row=0, column=7, padx=6, pady=6, sticky="w")
        ttk.Button(actions, text="打开结果目录", command=self.open_result_dir).grid(row=1, column=0, padx=6, pady=6, sticky="w")
        ttk.Button(actions, text="打开最终 Excel", command=self.open_final_excel).grid(row=1, column=1, padx=6, pady=6, sticky="w")
        ttk.Button(actions, text="复制 Prepare 命令", command=self.copy_prepare_command).grid(row=1, column=2, padx=6, pady=6, sticky="w")
        ttk.Button(actions, text="复制 Finalize 命令", command=self.copy_finalize_command).grid(row=1, column=3, padx=6, pady=6, sticky="w")

        info = ttk.LabelFrame(frame, text="状态", padding=10)
        info.pack(fill=tk.X, pady=(12, 0))
        self.status_var = tk.StringVar(value="待开始")
        ttk.Label(info, textvariable=self.status_var).pack(anchor="w")

        columns_frame = ttk.LabelFrame(frame, text="列名预览", padding=10)
        columns_frame.pack(fill=tk.X, pady=(12, 0))
        self.columns_text = tk.Text(columns_frame, height=6, wrap=tk.WORD)
        self.columns_text.pack(fill=tk.X)

        selected = ttk.LabelFrame(frame, text="当前已选择的 review_results_part 文件", padding=10)
        selected.pack(fill=tk.X, pady=(12, 0))
        self.review_files_text = tk.Text(selected, height=5, wrap=tk.WORD)
        self.review_files_text.pack(fill=tk.X)

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

    def pick_base_file(self) -> None:
        path = filedialog.askopenfilename(title="选择主文件", filetypes=[("Excel Files", "*.xlsx *.xls *.html *.htm"), ("All Files", "*.*")])
        if path:
            self.base_file.set(path)
            self.refresh_all_previews()

    def pick_target_file(self) -> None:
        path = filedialog.askopenfilename(title="选择对比文件", filetypes=[("Excel Files", "*.xlsx *.xls *.html *.htm"), ("All Files", "*.*")])
        if path:
            self.target_file.set(path)
            self.refresh_all_previews()

    def pick_result_dir(self) -> None:
        path = filedialog.askdirectory(title="选择结果目录")
        if path:
            self.result_dir.set(path)
            self.refresh_all_previews()

    def pick_final_excel(self) -> None:
        path = filedialog.asksaveasfilename(title="选择最终 Excel 路径", defaultextension=".xlsx", filetypes=[("Excel Files", "*.xlsx")])
        if path:
            self.final_excel.set(path)
            self.refresh_all_previews()

    def _quote(self, text: str) -> str:
        return f'"{text}"'

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
            "--output", self._quote(self.final_excel.get().strip()),
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
            self._open_path(Path(self.final_excel.get().strip()))
        except Exception as exc:
            messagebox.showerror("错误", f"{exc}\n\n{traceback.format_exc()}")

    def _preview_columns(self, file_path: str, title: str) -> None:
        path = Path(file_path.strip())
        if not path.exists():
            raise RuntimeError(f"文件不存在: {path}")
        df = workflow.load_table(path)
        columns = [str(col).strip() for col in df.columns]
        self.columns_text.delete("1.0", tk.END)
        self.columns_text.insert(tk.END, f"{title}：{path}\n")
        self.columns_text.insert(tk.END, f"共 {len(columns)} 列\n")
        self.columns_text.insert(tk.END, "\n".join(f"- {name}" for name in columns) + "\n")
        self.status_var.set(f"已读取{title}列名，共 {len(columns)} 列")

    def preview_base_columns(self) -> None:
        try:
            self._preview_columns(self.base_file.get(), "主文件列名")
        except Exception as exc:
            messagebox.showerror("错误", f"{exc}\n\n{traceback.format_exc()}")

    def preview_target_columns(self) -> None:
        try:
            self._preview_columns(self.target_file.get(), "对比文件列名")
        except Exception as exc:
            messagebox.showerror("错误", f"{exc}\n\n{traceback.format_exc()}")

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
        self.review_files_text.delete("1.0", tk.END)
        if not self.review_result_files:
            self.review_files_text.insert(tk.END, "未选择文件\n")
            return
        for path in self.review_result_files:
            self.review_files_text.insert(tk.END, str(path) + "\n")

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
                "final_excel": self.final_excel.get().strip(),
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
        try:
            self._persist_settings()
            result_dir = Path(self.result_dir.get().strip())
            prepare_path = result_dir / "prepare.json"
            parts_dir = result_dir / "prepare_parts"

            payload = self._build_prepare_payload()
            workflow.write_json(prepare_path, payload)
            parts = workflow.split_prepare_payload(payload, int(self.batch_size.get().strip() or "20"))
            parts_dir.mkdir(parents=True, exist_ok=True)
            for part in parts:
                workflow.write_json(parts_dir / f"prepare_part_{int(part['part_no'])}.json", part)

            self.status_var.set(f"已生成 prepare.json，review_tasks={len(payload['review_tasks'])}，拆分为 {len(parts)} 批")
            self.refresh_all_previews()
            messagebox.showinfo("完成", f"已生成 prepare.json，并拆分为 {len(parts)} 批。")
        except Exception as exc:
            self.status_var.set("生成 prepare 失败")
            messagebox.showerror("错误", f"{exc}\n\n{traceback.format_exc()}")

    def merge_and_finalize(self) -> None:
        try:
            self._persist_settings()
            result_dir = Path(self.result_dir.get().strip())
            prepare_path = result_dir / "prepare.json"
            review_results_path = result_dir / "review_results.json"
            final_excel_path = Path(self.final_excel.get().strip())

            if not self.review_result_files:
                self.auto_detect_review_result_files()
            if not self.review_result_files:
                raise RuntimeError("未选择 review_results_part 文件")

            merged = workflow.merge_review_result_files(self.review_result_files)
            workflow.write_json(review_results_path, {"review_results": merged})

            prepared = workflow.load_json(prepare_path)
            payload = workflow.build_finalize_payload(prepared, merged)
            json_output_path = final_excel_path.with_suffix(".json")
            workflow.write_json(json_output_path, payload)
            workflow.export_excel(payload, final_excel_path)

            self.status_var.set(f"已生成 {final_excel_path}")
            messagebox.showinfo("完成", f"已生成最终 Excel：\n{final_excel_path}")
        except Exception as exc:
            self.status_var.set("合并或导出失败")
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
