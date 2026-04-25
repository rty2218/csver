#!/usr/bin/env python3
"""
Small GUI for batch converting CSV files.

Supports selecting one folder or selecting multiple CSV files, then converting
to XLSX, fixed-width TXT, Markdown, or all three formats.
"""

from __future__ import annotations

import shlex
import threading
from pathlib import Path
from typing import Optional

try:
    import tkinter as tk
    from tkinter import filedialog, messagebox, scrolledtext, ttk

    TK_AVAILABLE = True
    TK_IMPORT_ERROR: Optional[Exception] = None
except Exception as exc:
    tk = None
    filedialog = messagebox = scrolledtext = ttk = None
    TK_AVAILABLE = False
    TK_IMPORT_ERROR = exc

from csv_batch_convert import (
    collect_csv_files,
    convert_one_csv,
    delimiter_label,
    resolve_formats,
)


SCRIPT_DIR = Path(__file__).resolve().parent
RESULT_FOLDER_NAME = "转换结果"
BG_COLOR = "#f5f7fb"
PANEL_COLOR = "#ffffff"
TEXT_COLOR = "#1f2937"
MUTED_COLOR = "#4b5563"
PRIMARY_COLOR = "#2563eb"
PRIMARY_ACTIVE_COLOR = "#1d4ed8"
BUTTON_COLOR = "#e5e7eb"
BUTTON_ACTIVE_COLOR = "#d1d5db"


class CsvBatchConvertApp:
    def __init__(self, root: tk.Tk) -> None:
        self.root = root
        self.root.title("CSV 批量转换工具")
        self.root.geometry("760x560")
        self.root.minsize(680, 500)

        self.input_mode = tk.StringVar(value="")
        self.selected_paths: list[Path] = []
        self.output_dir = tk.StringVar(value="选择文件或文件夹后自动生成")
        self.output_per_file = False
        self.format_name = tk.StringVar(value="all")
        self.recursive = tk.BooleanVar(value=False)
        self.infer_types = tk.BooleanVar(value=True)

        self.build_ui()

    def build_ui(self) -> None:
        root = self.root
        root.configure(bg=BG_COLOR)
        root.columnconfigure(0, weight=1)
        root.rowconfigure(5, weight=1)

        title = tk.Label(root, text="CSV 批量转换工具", font=("", 20, "bold"), bg=BG_COLOR, fg=TEXT_COLOR)
        title.grid(row=0, column=0, sticky="w", padx=18, pady=(16, 8))

        source_frame = tk.LabelFrame(
            root,
            text="1. 选择要上传/转换的 CSV",
            bg=PANEL_COLOR,
            fg=TEXT_COLOR,
            padx=10,
            pady=8,
            font=("", 13, "bold"),
        )
        source_frame.grid(row=1, column=0, sticky="ew", padx=18, pady=8)
        source_frame.columnconfigure(2, weight=1)

        self.make_button(source_frame, "选择 CSV 文件（可多选）", self.choose_files).grid(
            row=0, column=0, padx=10, pady=10, sticky="w"
        )
        self.make_button(source_frame, "选择文件夹", self.choose_folder).grid(
            row=0, column=1, padx=10, pady=10, sticky="w"
        )
        tk.Checkbutton(
            source_frame,
            text="文件夹包含子文件夹",
            variable=self.recursive,
            bg=PANEL_COLOR,
            fg=TEXT_COLOR,
            activebackground=PANEL_COLOR,
            activeforeground=TEXT_COLOR,
            selectcolor=PANEL_COLOR,
        ).grid(
            row=0, column=2, padx=10, pady=10, sticky="w"
        )
        self.source_label = tk.Label(
            source_frame,
            text="还没有选择文件或文件夹",
            bg=PANEL_COLOR,
            fg=MUTED_COLOR,
            anchor="w",
        )
        self.source_label.grid(row=1, column=0, columnspan=3, padx=10, pady=(0, 10), sticky="ew")

        format_frame = tk.LabelFrame(
            root,
            text="2. 选择转换类型",
            bg=PANEL_COLOR,
            fg=TEXT_COLOR,
            padx=10,
            pady=8,
            font=("", 13, "bold"),
        )
        format_frame.grid(row=2, column=0, sticky="ew", padx=18, pady=8)
        for col in range(4):
            format_frame.columnconfigure(col, weight=1)

        options = [
            ("只转 XLSX", "xlsx"),
            ("只转 TXT 表格", "txt"),
            ("只转 Markdown 表格", "md"),
            ("全部转换", "all"),
        ]
        for col, (label, value) in enumerate(options):
            tk.Radiobutton(
                format_frame,
                text=label,
                value=value,
                variable=self.format_name,
                bg=PANEL_COLOR,
                fg=TEXT_COLOR,
                activebackground=PANEL_COLOR,
                activeforeground=TEXT_COLOR,
                selectcolor=PANEL_COLOR,
            ).grid(
                row=0, column=col, padx=10, pady=10, sticky="w"
            )

        output_frame = tk.LabelFrame(
            root,
            text="3. 输出位置",
            bg=PANEL_COLOR,
            fg=TEXT_COLOR,
            padx=10,
            pady=8,
            font=("", 13, "bold"),
        )
        output_frame.grid(row=3, column=0, sticky="ew", padx=18, pady=8)
        output_frame.columnconfigure(0, weight=1)
        tk.Label(
            output_frame,
            text="自动输出到源目录里的“转换结果”文件夹",
            bg=PANEL_COLOR,
            fg=TEXT_COLOR,
            anchor="w",
        ).grid(
            row=0, column=0, padx=10, pady=(10, 2), sticky="w"
        )
        self.output_label = tk.Label(
            output_frame,
            textvariable=self.output_dir,
            bg=PANEL_COLOR,
            fg=MUTED_COLOR,
            anchor="w",
        )
        self.output_label.grid(row=1, column=0, padx=10, pady=(0, 10), sticky="ew")
        tk.Checkbutton(
            output_frame,
            text="XLSX 自动识别数字",
            variable=self.infer_types,
            bg=PANEL_COLOR,
            fg=TEXT_COLOR,
            activebackground=PANEL_COLOR,
            activeforeground=TEXT_COLOR,
            selectcolor=PANEL_COLOR,
        ).grid(
            row=2, column=0, padx=10, pady=(0, 10), sticky="w"
        )

        action_frame = tk.Frame(root, bg=BG_COLOR)
        action_frame.grid(row=4, column=0, sticky="ew", padx=18, pady=8)
        action_frame.columnconfigure(0, weight=1)
        self.start_button = self.make_button(
            action_frame,
            "开始批量转换",
            self.start_conversion,
            bg=PRIMARY_COLOR,
            fg="#ffffff",
            active_bg=PRIMARY_ACTIVE_COLOR,
            active_fg="#ffffff",
            padx=18,
            pady=8,
        )
        self.start_button.grid(row=0, column=1, sticky="e")

        log_frame = tk.LabelFrame(
            root,
            text="转换日志",
            bg=PANEL_COLOR,
            fg=TEXT_COLOR,
            padx=10,
            pady=8,
            font=("", 13, "bold"),
        )
        log_frame.grid(row=5, column=0, sticky="nsew", padx=18, pady=(8, 16))
        log_frame.columnconfigure(0, weight=1)
        log_frame.rowconfigure(0, weight=1)
        self.log = scrolledtext.ScrolledText(
            log_frame,
            height=12,
            wrap="word",
            bg="#ffffff",
            fg=TEXT_COLOR,
            insertbackground=TEXT_COLOR,
            relief="solid",
            borderwidth=1,
        )
        self.log.grid(row=0, column=0, sticky="nsew", padx=10, pady=10)

    def make_button(
        self,
        parent,
        text: str,
        command,
        bg: str = BUTTON_COLOR,
        fg: str = TEXT_COLOR,
        active_bg: str = BUTTON_ACTIVE_COLOR,
        active_fg: str = TEXT_COLOR,
        padx: int = 12,
        pady: int = 6,
    ):
        return tk.Button(
            parent,
            text=text,
            command=command,
            bg=bg,
            fg=fg,
            activebackground=active_bg,
            activeforeground=active_fg,
            relief="raised",
            borderwidth=1,
            padx=padx,
            pady=pady,
            highlightthickness=0,
        )

    def choose_files(self) -> None:
        files = filedialog.askopenfilenames(
            title="选择一个或多个 CSV 文件",
            filetypes=[("CSV files", "*.csv"), ("All files", "*.*")],
        )
        if not files:
            return
        self.input_mode.set("files")
        self.selected_paths = [Path(file).resolve() for file in files]
        self.source_label.configure(text=f"已选择 {len(self.selected_paths)} 个 CSV 文件")
        self.set_default_output_for_files(self.selected_paths)

    def choose_folder(self) -> None:
        folder = filedialog.askdirectory(title="选择包含 CSV 的文件夹")
        if not folder:
            return
        folder_path = Path(folder).resolve()
        self.input_mode.set("folder")
        self.selected_paths = [folder_path]
        self.source_label.configure(text=f"已选择文件夹：{folder_path}")
        self.output_per_file = False
        self.output_dir.set(str(folder_path / RESULT_FOLDER_NAME))

    def set_default_output_for_files(self, files: list[Path]) -> None:
        parents = {file.parent for file in files}
        if len(parents) == 1:
            self.output_per_file = False
            self.output_dir.set(str(next(iter(parents)) / RESULT_FOLDER_NAME))
        else:
            self.output_per_file = True
            self.output_dir.set(f"多个目录：每个 CSV 所在目录 / {RESULT_FOLDER_NAME}")

    def start_conversion(self) -> None:
        if not self.selected_paths:
            messagebox.showwarning("请先选择", "请先选择 CSV 文件，或者选择一个包含 CSV 的文件夹。")
            return

        csv_files = self.get_csv_files()
        if not csv_files:
            messagebox.showwarning("没有 CSV", "没有找到可转换的 CSV 文件。")
            return

        output_root = None if self.output_per_file else Path(self.output_dir.get()).expanduser()
        format_name = self.format_name.get()
        infer_types = self.infer_types.get()

        self.start_button.configure(state="disabled")
        self.log.delete("1.0", tk.END)
        thread = threading.Thread(
            target=self.convert_in_background,
            args=(csv_files, output_root, format_name, infer_types),
            daemon=True,
        )
        thread.start()

    def convert_in_background(
        self,
        csv_files: list[Path],
        output_root: Optional[Path],
        format_name: str,
        infer_types: bool,
    ) -> None:
        try:
            if output_root is not None:
                output_root = output_root.resolve()
                output_root.mkdir(parents=True, exist_ok=True)
            formats = resolve_formats(format_name)
            used_outputs: set[Path] = set()

            self.log_line(f"找到 {len(csv_files)} 个 CSV 文件。")
            if output_root is None:
                self.log_line(f"输出文件夹：每个 CSV 所在目录 / {RESULT_FOLDER_NAME}")
            else:
                self.log_line(f"输出文件夹：{output_root}")
            self.log_line("")

            converted = 0
            failed = 0
            for csv_path in csv_files:
                try:
                    base_dir = output_root if output_root is not None else csv_path.parent / RESULT_FOLDER_NAME
                    base_dir.mkdir(parents=True, exist_ok=True)
                    outputs, encoding, delimiter = convert_one_csv(
                        csv_path,
                        base_dir=base_dir,
                        used_outputs=used_outputs,
                        formats=formats,
                        encoding="auto",
                        delimiter="auto",
                        infer_types=infer_types,
                        no_header=False,
                        max_col_width=60,
                    )
                    converted += 1
                    self.log_line(f"成功：{csv_path.name}  encoding={encoding}, delimiter={delimiter_label(delimiter)}")
                    for output_path in outputs:
                        self.log_line(f"  -> {output_path}")
                except Exception as exc:
                    failed += 1
                    self.log_line(f"失败：{csv_path}：{exc}")
                self.log_line("")

            ok = converted > 0
            message = f"完成：成功 {converted} 个，失败 {failed} 个。"
            self.log_line(message)
            self.finish_conversion(ok, message)
        except Exception as exc:
            self.log_line(f"出错：{exc}")
            self.finish_conversion(False, str(exc))

    def get_csv_files(self) -> list[Path]:
        if self.input_mode.get() == "files":
            return [path for path in self.selected_paths if path.is_file() and path.suffix.lower() == ".csv"]
        if self.input_mode.get() == "folder":
            return collect_csv_files([str(self.selected_paths[0])], self.recursive.get())
        return []

    def log_line(self, text: str) -> None:
        self.root.after(0, self._append_log_line, text)

    def _append_log_line(self, text: str) -> None:
        self.log.insert(tk.END, text + "\n")
        self.log.see(tk.END)

    def finish_conversion(self, ok: bool, message: str) -> None:
        self.root.after(0, self._finish_conversion, ok, message)

    def _finish_conversion(self, ok: bool, message: str) -> None:
        self.start_button.configure(state="normal")
        if ok:
            messagebox.showinfo("转换完成", message)
        else:
            messagebox.showwarning("转换未完成", message)


def terminal_main(error: Optional[Exception]) -> int:
    print("CSV 批量转换工具")
    print("图形窗口不可用，已切换到终端提问模式。")
    if error is not None:
        print(f"原因：{error}")
    print()

    source_choice = ask_choice("请选择上传方式：", [("1", "选择文件夹"), ("2", "选择 CSV 文件（可多个）")])
    csv_files: list[Path]
    output_root: Optional[Path]

    if source_choice == "1":
        folder = Path(strip_outer_quotes(input("请输入或拖入文件夹路径：").strip())).expanduser().resolve()
        recursive = ask_yes_no("是否包含子文件夹？y/N：")
        csv_files = collect_csv_files([str(folder)], recursive)
        output_root = folder / RESULT_FOLDER_NAME
    else:
        raw = input("请输入或拖入 CSV 文件路径；多个文件可用空格或逗号分开：").strip()
        csv_files = [path for path in parse_input_paths(raw) if path.is_file() and path.suffix.lower() == ".csv"]
        parents = {path.parent for path in csv_files}
        output_root = next(iter(parents)) / RESULT_FOLDER_NAME if len(parents) == 1 else None

    if not csv_files:
        print("没有找到 CSV 文件。")
        input("按 Enter 退出。")
        return 1

    format_choice = ask_choice(
        "请选择转换类型：",
        [("1", "只转 XLSX"), ("2", "只转 TXT 表格"), ("3", "只转 Markdown 表格"), ("4", "全部转换")],
    )
    format_name = {"1": "xlsx", "2": "txt", "3": "md", "4": "all"}[format_choice]
    formats = resolve_formats(format_name)
    used_outputs: set[Path] = set()

    print()
    print(f"找到 {len(csv_files)} 个 CSV 文件。")
    if output_root is None:
        print(f"输出文件夹：每个 CSV 所在目录 / {RESULT_FOLDER_NAME}")
    else:
        print(f"输出文件夹：{output_root}")
    print()

    converted = 0
    failed = 0
    for csv_path in csv_files:
        try:
            base_dir = output_root if output_root is not None else csv_path.parent / RESULT_FOLDER_NAME
            outputs, encoding, delimiter = convert_one_csv(
                csv_path,
                base_dir=base_dir,
                used_outputs=used_outputs,
                formats=formats,
                encoding="auto",
                delimiter="auto",
                infer_types=True,
                no_header=False,
                max_col_width=60,
            )
            converted += 1
            print(f"成功：{csv_path.name}  encoding={encoding}, delimiter={delimiter_label(delimiter)}")
            for output_path in outputs:
                print(f"  -> {output_path}")
        except Exception as exc:
            failed += 1
            print(f"失败：{csv_path}：{exc}")
        print()

    print(f"完成：成功 {converted} 个，失败 {failed} 个。")
    input("按 Enter 退出。")
    return 0 if converted else 1


def ask_choice(prompt: str, options: list[tuple[str, str]]) -> str:
    while True:
        print(prompt)
        for value, label in options:
            print(f"  {value}. {label}")
        answer = input("请输入选项编号：").strip()
        if answer in {value for value, _ in options}:
            return answer
        print("选项无效，请重新输入。")
        print()


def ask_yes_no(prompt: str) -> bool:
    answer = input(prompt).strip().lower()
    return answer in {"y", "yes", "是", "1"}


def parse_input_paths(raw: str) -> list[Path]:
    if "," in raw:
        parts = [strip_outer_quotes(part.strip()) for part in raw.split(",") if part.strip()]
    else:
        parts = [strip_outer_quotes(part) for part in shlex.split(raw)]
    return [Path(part).expanduser().resolve() for part in parts]


def strip_outer_quotes(value: str) -> str:
    if len(value) >= 2 and value[0] == value[-1] and value[0] in {"'", '"'}:
        return value[1:-1]
    return value


def main() -> int:
    if not TK_AVAILABLE:
        return terminal_main(TK_IMPORT_ERROR)

    root = tk.Tk()
    app = CsvBatchConvertApp(root)
    root.mainloop()
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
