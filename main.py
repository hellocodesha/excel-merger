"""
Excel多表合并工具
- 选择文件夹，自动读取所有 .xlsx/.xls 文件并合并
- 支持子文件夹扫描、来源列、多种合并模式
- tkinter GUI，中文界面
"""

import os
import sys
import platform
import subprocess
import threading
import tkinter as tk
from tkinter import ttk, filedialog, messagebox
from pathlib import Path
from datetime import datetime

import pandas as pd
from openpyxl.styles import Font, Alignment, Border, Side, PatternFill
from openpyxl.utils import get_column_letter


# ── 颜色 / 字体常量 ──────────────────────────────────────────────
BG       = "#F7F8FA"
CARD_BG  = "#FFFFFF"
PRIMARY  = "#4A90D9"
TEXT     = "#333333"
SUB_TEXT = "#888888"
FONT     = ("Microsoft YaHei UI", 10)
FONT_B   = ("Microsoft YaHei UI", 10, "bold")
FONT_S   = ("Microsoft YaHei UI", 9)
FONT_T   = ("Microsoft YaHei UI", 14, "bold")


class ExcelMergerApp:
    """主窗口."""

    def __init__(self, root: tk.Tk):
        self.root = root
        self.root.title("Excel多表合并工具")
        self.root.geometry("620x700")
        self.root.resizable(False, True)
        self.root.configure(bg=BG)

        # ── 变量 ──
        self.folder_path = tk.StringVar()
        self.include_subfolders = tk.BooleanVar(value=False)
        self.add_source_col = tk.BooleanVar(value=True)
        # 合并模式: "single" = 合并到一个sheet, "separate" = 每个文件一个sheet
        self.merge_mode = tk.StringVar(value="single")
        # sheet范围: "first" = 只读第一个sheet, "all" = 读所有sheet
        self.sheet_scope = tk.StringVar(value="first")

        self._build_ui()

    # ── UI 构建 ───────────────────────────────────────────────────
    def _build_ui(self):
        # 标题
        title = tk.Label(
            self.root, text="Excel 多表合并工具", font=FONT_T,
            bg=BG, fg=PRIMARY,
        )
        title.pack(pady=(18, 6))

        subtitle = tk.Label(
            self.root, text="选择文件夹，一键合并所有 Excel 文件", font=FONT_S,
            bg=BG, fg=SUB_TEXT,
        )
        subtitle.pack(pady=(0, 12))

        # ── 文件夹选择卡片 ──
        card1 = self._card(self.root)
        card1.pack(fill="x", padx=24, pady=(0, 8))

        row = tk.Frame(card1, bg=CARD_BG)
        row.pack(fill="x", padx=12, pady=10)

        tk.Label(row, text="文件夹路径：", font=FONT_B, bg=CARD_BG, fg=TEXT).pack(
            side="left"
        )
        entry = tk.Entry(
            row, textvariable=self.folder_path, font=FONT, width=34,
            relief="solid", bd=1,
        )
        entry.pack(side="left", padx=(4, 8))

        btn_browse = tk.Button(
            row, text="浏 览", font=FONT_S, bg=PRIMARY, fg="white",
            activebackground="#3A7BC8", activeforeground="white",
            relief="flat", padx=12, pady=2, cursor="hand2",
            command=self._browse_folder,
        )
        btn_browse.pack(side="left")

        # ── 选项卡片 ──
        card2 = self._card(self.root)
        card2.pack(fill="x", padx=24, pady=(0, 8))

        opt_title = tk.Label(
            card2, text="合并选项", font=FONT_B, bg=CARD_BG, fg=TEXT,
        )
        opt_title.pack(anchor="w", padx=14, pady=(10, 4))

        sep = ttk.Separator(card2, orient="horizontal")
        sep.pack(fill="x", padx=14)

        opts = tk.Frame(card2, bg=CARD_BG)
        opts.pack(fill="x", padx=14, pady=8)

        # 第一行复选框
        chk_row1 = tk.Frame(opts, bg=CARD_BG)
        chk_row1.pack(fill="x", pady=2)

        tk.Checkbutton(
            chk_row1, text="包含子文件夹", variable=self.include_subfolders,
            font=FONT, bg=CARD_BG, fg=TEXT, activebackground=CARD_BG,
            selectcolor=CARD_BG,
        ).pack(side="left", padx=(0, 24))

        tk.Checkbutton(
            chk_row1, text="添加「来源文件名」列", variable=self.add_source_col,
            font=FONT, bg=CARD_BG, fg=TEXT, activebackground=CARD_BG,
            selectcolor=CARD_BG,
        ).pack(side="left")

        # 合并模式
        mode_frame = tk.LabelFrame(
            opts, text="合并模式", font=FONT_S, bg=CARD_BG, fg=SUB_TEXT,
            bd=1, relief="groove",
        )
        mode_frame.pack(fill="x", pady=(8, 4))

        tk.Radiobutton(
            mode_frame, text="合并到同一个 Sheet", variable=self.merge_mode,
            value="single", font=FONT, bg=CARD_BG, fg=TEXT,
            activebackground=CARD_BG, selectcolor=CARD_BG,
        ).pack(side="left", padx=(8, 20), pady=4)

        tk.Radiobutton(
            mode_frame, text="每个文件单独一个 Sheet", variable=self.merge_mode,
            value="separate", font=FONT, bg=CARD_BG, fg=TEXT,
            activebackground=CARD_BG, selectcolor=CARD_BG,
        ).pack(side="left", padx=(0, 8), pady=4)

        # Sheet 范围
        scope_frame = tk.LabelFrame(
            opts, text="读取范围", font=FONT_S, bg=CARD_BG, fg=SUB_TEXT,
            bd=1, relief="groove",
        )
        scope_frame.pack(fill="x", pady=(4, 4))

        tk.Radiobutton(
            scope_frame, text="只读取第一个 Sheet", variable=self.sheet_scope,
            value="first", font=FONT, bg=CARD_BG, fg=TEXT,
            activebackground=CARD_BG, selectcolor=CARD_BG,
        ).pack(side="left", padx=(8, 20), pady=4)

        tk.Radiobutton(
            scope_frame, text="读取所有 Sheet", variable=self.sheet_scope,
            value="all", font=FONT, bg=CARD_BG, fg=TEXT,
            activebackground=CARD_BG, selectcolor=CARD_BG,
        ).pack(side="left", padx=(0, 8), pady=4)

        # ── 进度条 + 状态 ──
        card3 = self._card(self.root)
        card3.pack(fill="x", padx=24, pady=(0, 8))

        prog_frame = tk.Frame(card3, bg=CARD_BG)
        prog_frame.pack(fill="x", padx=14, pady=10)

        self.progress = ttk.Progressbar(
            prog_frame, orient="horizontal", length=540, mode="determinate",
        )
        self.progress.pack(fill="x")

        self.status_var = tk.StringVar(value="就绪")
        status_lbl = tk.Label(
            prog_frame, textvariable=self.status_var, font=FONT_S,
            bg=CARD_BG, fg=SUB_TEXT, anchor="w",
        )
        status_lbl.pack(fill="x", pady=(4, 0))

        # ── 日志区域 ──
        card4 = self._card(self.root)
        card4.pack(fill="x", padx=24, pady=(0, 8))

        self.log_text = tk.Text(
            card4, height=5, font=("Consolas", 9), bg="#FAFAFA",
            fg=TEXT, relief="flat", wrap="word", state="disabled",
        )
        self.log_text.pack(fill="x", padx=10, pady=6)

        # ── 执行按钮（醒目大按钮） ──
        btn_frame = tk.Frame(self.root, bg=BG)
        btn_frame.pack(fill="x", padx=24, pady=(8, 16))

        btn_run = tk.Button(
            btn_frame, text="开 始 合 并", font=("Microsoft YaHei UI", 13, "bold"),
            bg=PRIMARY, fg="white", activebackground="#3A7BC8",
            activeforeground="white", relief="flat",
            padx=60, pady=10, cursor="hand2",
            command=self._start_merge,
        )
        btn_run.pack(expand=True)

    # ── 辅助 UI ──────────────────────────────────────────────────
    @staticmethod
    def _card(parent) -> tk.Frame:
        """带圆角感的白色卡片容器."""
        f = tk.Frame(parent, bg=CARD_BG, bd=1, relief="solid", highlightthickness=0)
        return f

    def _log(self, msg: str):
        self.log_text.configure(state="normal")
        self.log_text.insert("end", msg + "\n")
        self.log_text.see("end")
        self.log_text.configure(state="disabled")

    def _clear_log(self):
        self.log_text.configure(state="normal")
        self.log_text.delete("1.0", "end")
        self.log_text.configure(state="disabled")

    # ── 文件夹浏览 ──
    def _browse_folder(self):
        path = filedialog.askdirectory(title="选择包含 Excel 文件的文件夹")
        if path:
            self.folder_path.set(path)

    # ── 启动合并（线程） ─────────────────────────────────────────
    def _start_merge(self):
        folder = self.folder_path.get().strip()
        if not folder or not os.path.isdir(folder):
            messagebox.showwarning("提示", "请先选择一个有效的文件夹路径。")
            return

        self._clear_log()
        self.progress["value"] = 0
        self.status_var.set("扫描文件中…")

        thread = threading.Thread(target=self._merge_worker, args=(folder,), daemon=True)
        thread.start()

    # ── 合并核心逻辑 ─────────────────────────────────────────────
    def _merge_worker(self, folder: str):
        try:
            # 1. 扫描文件
            files = self._scan_files(folder)
            if not files:
                self.root.after(0, lambda: messagebox.showinfo("提示", "所选文件夹内未找到 .xlsx 或 .xls 文件。"))
                self.root.after(0, lambda: self.status_var.set("就绪"))
                return

            total = len(files)
            self.root.after(0, lambda: self._log(f"共找到 {total} 个 Excel 文件"))

            merge_mode = self.merge_mode.get()
            sheet_scope = self.sheet_scope.get()
            add_source = self.add_source_col.get()

            all_frames: list[pd.DataFrame] = []  # 用于 single 模式
            file_frames: dict[str, pd.DataFrame] = {}  # 用于 separate 模式
            skipped: list[str] = []

            for idx, fpath in enumerate(files, 1):
                fname = os.path.basename(fpath)
                self.root.after(0, lambda f=fname, i=idx: self.status_var.set(f"处理中 ({i}/{total})：{f}"))

                try:
                    dfs = self._read_excel(fpath, sheet_scope)
                except Exception as e:
                    skipped.append(fname)
                    self.root.after(0, lambda f=fname, err=e: self._log(f"[跳过] {f}：{err}"))
                    self._update_progress(idx, total)
                    continue

                if not dfs:
                    skipped.append(fname)
                    self.root.after(0, lambda f=fname: self._log(f"[跳过] {f}：文件无有效数据"))
                    self._update_progress(idx, total)
                    continue

                for sheet_name, df in dfs:
                    if add_source:
                        df.insert(0, "来源文件名", fname)
                        if sheet_scope == "all" and len(dfs) > 1:
                            df.insert(1, "来源Sheet", sheet_name)

                    if merge_mode == "single":
                        all_frames.append(df)
                    else:
                        # separate: 用 文件名 或 文件名_Sheet名 作为sheet名
                        base = Path(fname).stem
                        key = base if len(dfs) == 1 else f"{base}_{sheet_name}"
                        # Excel sheet名最长31字符
                        key = key[:31]
                        # 处理重名
                        orig_key = key
                        counter = 1
                        while key in file_frames:
                            suffix = f"_{counter}"
                            key = orig_key[: 31 - len(suffix)] + suffix
                            counter += 1
                        file_frames[key] = df

                self.root.after(0, lambda f=fname: self._log(f"[完成] {f}"))
                self._update_progress(idx, total)

            # 2. 写出结果
            if merge_mode == "single" and not all_frames:
                self.root.after(0, lambda: messagebox.showinfo("提示", "没有可合并的数据。"))
                self.root.after(0, lambda: self.status_var.set("就绪"))
                return
            if merge_mode == "separate" and not file_frames:
                self.root.after(0, lambda: messagebox.showinfo("提示", "没有可合并的数据。"))
                self.root.after(0, lambda: self.status_var.set("就绪"))
                return

            timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
            out_name = f"合并结果_{timestamp}.xlsx"
            out_path = os.path.join(folder, out_name)

            self.root.after(0, lambda: self.status_var.set("正在写入文件…"))

            if merge_mode == "single":
                merged = pd.concat(all_frames, ignore_index=True)
                with pd.ExcelWriter(out_path, engine="openpyxl") as writer:
                    merged.to_excel(writer, sheet_name="合并结果", index=False)
                    self._format_sheet(writer.sheets["合并结果"])
            else:
                with pd.ExcelWriter(out_path, engine="openpyxl") as writer:
                    for sname, df in file_frames.items():
                        df.to_excel(writer, sheet_name=sname, index=False)
                        self._format_sheet(writer.sheets[sname])

            # 3. 完成
            summary = f"合并完成！输出文件：{out_name}"
            if skipped:
                summary += f"\n\n以下 {len(skipped)} 个文件被跳过：\n" + "\n".join(skipped)

            self.root.after(0, lambda: self.status_var.set("完成"))
            self.root.after(0, lambda: self._log(f"输出文件：{out_path}"))
            self.root.after(0, lambda s=summary: messagebox.showinfo("完成", s))
            self.root.after(0, lambda: self._open_folder(os.path.dirname(out_path)))

        except Exception as e:
            self.root.after(0, lambda err=e: messagebox.showerror("错误", f"合并过程中发生错误：\n{err}"))
            self.root.after(0, lambda: self.status_var.set("出错"))

    # ── 扫描文件 ──
    def _scan_files(self, folder: str) -> list[str]:
        result = []
        if self.include_subfolders.get():
            for root_dir, _, filenames in os.walk(folder):
                for fn in filenames:
                    if fn.lower().endswith((".xlsx", ".xls")) and not fn.startswith("~$"):
                        result.append(os.path.join(root_dir, fn))
        else:
            for fn in os.listdir(folder):
                if fn.lower().endswith((".xlsx", ".xls")) and not fn.startswith("~$"):
                    full = os.path.join(folder, fn)
                    if os.path.isfile(full):
                        result.append(full)
        result.sort()
        return result

    # ── 读取单个 Excel ──
    def _read_excel(self, fpath: str, scope: str) -> list[tuple[str, pd.DataFrame]]:
        results = []
        if scope == "first":
            df = pd.read_excel(fpath, sheet_name=0, engine=self._engine(fpath))
            if not df.empty:
                xl = pd.ExcelFile(fpath, engine=self._engine(fpath))
                sheet_name = xl.sheet_names[0]
                xl.close()
                results.append((sheet_name, df))
        else:
            sheets = pd.read_excel(fpath, sheet_name=None, engine=self._engine(fpath))
            for name, df in sheets.items():
                if not df.empty:
                    results.append((name, df))
        return results

    @staticmethod
    def _engine(fpath: str) -> str:
        return "xlrd" if fpath.lower().endswith(".xls") else "openpyxl"

    # ── 格式化 Sheet ──
    @staticmethod
    def _format_sheet(ws):
        """给 sheet 添加表头样式、边框、自动列宽、冻结首行."""
        thin_border = Border(
            left=Side(style="thin"),
            right=Side(style="thin"),
            top=Side(style="thin"),
            bottom=Side(style="thin"),
        )
        header_font = Font(name="Microsoft YaHei UI", size=11, bold=True)
        header_fill = PatternFill(start_color="4A90D9", end_color="4A90D9", fill_type="solid")
        header_align = Alignment(horizontal="center", vertical="center", wrap_text=True)
        cell_font = Font(name="Microsoft YaHei UI", size=10)
        cell_align = Alignment(vertical="center", wrap_text=False)

        # 遍历所有单元格
        for row_idx, row in enumerate(ws.iter_rows(min_row=1, max_row=ws.max_row, max_col=ws.max_column), 1):
            for cell in row:
                cell.border = thin_border
                if row_idx == 1:
                    cell.font = Font(name="Microsoft YaHei UI", size=11, bold=True, color="FFFFFF")
                    cell.fill = header_fill
                    cell.alignment = header_align
                else:
                    cell.font = cell_font
                    cell.alignment = cell_align

        # 自动列宽
        for col_idx in range(1, ws.max_column + 1):
            max_len = 0
            col_letter = get_column_letter(col_idx)
            for row in ws.iter_rows(min_row=1, max_row=min(ws.max_row, 200), min_col=col_idx, max_col=col_idx):
                cell = row[0]
                if cell.value is not None:
                    # 中文字符按2个宽度算
                    val = str(cell.value)
                    length = sum(2 if ord(c) > 127 else 1 for c in val)
                    max_len = max(max_len, length)
            # 限制在 8~50 之间
            ws.column_dimensions[col_letter].width = min(max(max_len + 3, 8), 50)

        # 冻结首行（滚动时表头不动）
        ws.freeze_panes = "A2"

        # 行高
        ws.row_dimensions[1].height = 28
        for r in range(2, ws.max_row + 1):
            ws.row_dimensions[r].height = 22

    # ── 进度条更新 ──
    def _update_progress(self, current: int, total: int):
        pct = current / total * 100
        self.root.after(0, lambda: self.progress.configure(value=pct))

    # ── 打开文件夹 ──
    @staticmethod
    def _open_folder(path: str):
        system = platform.system()
        if system == "Windows":
            os.startfile(path)
        elif system == "Darwin":
            subprocess.Popen(["open", path])
        else:
            subprocess.Popen(["xdg-open", path])


def main():
    # Windows DPI 适配（必须在创建 Tk 之前）
    try:
        if platform.system() == "Windows":
            from ctypes import windll
            windll.shcore.SetProcessDpiAwareness(1)
    except Exception:
        pass

    root = tk.Tk()
    ExcelMergerApp(root)
    root.mainloop()


if __name__ == "__main__":
    main()
