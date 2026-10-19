#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
债券询价双表合并工具 - Windows GUI
两个拖拽区（品种一/品种二）+ 合并按钮 + 结果另存
"""

import os
import sys
import traceback
import tkinter as tk
from tkinter import ttk, filedialog, messagebox
from pathlib import Path
import threading

# ---------- 导入合并处理逻辑 ----------
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
try:
    from bond_merge_processor import merge_two_tables, write_output_excel
except ImportError:
    merge_two_tables = None
    write_output_excel = None


# ---------- Windows 拖拽支持 ----------
if sys.platform == "win32":
    import ctypes
    from ctypes import wintypes

    user32 = ctypes.windll.user32
    shell32 = ctypes.windll.shell32

    WM_DROPFILES = 0x0233
    GWL_WNDPROC = -4
    GWL_EXSTYLE = -20
    WS_EX_ACCEPTFILES = 0x00000010
    DragQueryFile = shell32.DragQueryFileW
    DragFinish = shell32.DragFinish


# ---------- GUI ----------
class BondMergeGUI:
    def __init__(self, root):
        self.root = root
        self.root.title("债券询价双表合并工具")
        self.root.geometry("620x540")
        self.root.resizable(True, True)
        self.root.minsize(540, 480)

        # 文件路径
        self.file1_path = tk.StringVar()
        self.file2_path = tk.StringVar()
        self.file1_label_text = tk.StringVar(value="拖拽或点击选择 品种一 文件")
        self.file2_label_text = tk.StringVar(value="拖拽或点击选择 品种二 文件")
        self.status_var = tk.StringVar(value="就绪 — 请选择两个文件后点击「开始合并」")

        # 配色
        self.bg = "#f0f4f8"
        self.accent = "#2563eb"
        self.accent_hover = "#1d4ed8"
        self.success = "#059669"
        self.error = "#dc2626"
        self.text_color = "#1e293b"
        self.sub_color = "#64748b"
        self.card_bg = "#ffffff"
        self.drop_bg = "#f8fafc"
        self.drop_border = "#cbd5e1"
        self.drop_border_active = "#2563eb"
        self.drop_filled_bg = "#eff6ff"
        self.drop_filled_border = "#3b82f6"

        self.root.configure(bg=self.bg)

        # 拖拽区引用（_build_drop_zone 填充）
        self.drop1_frame = None
        self.drop2_frame = None

        self.build_ui()

        # 启用拖拽
        self.root.after(100, self._enable_dnd)

    def _enable_dnd(self):
        try:
            hwnd = self.root.winfo_id()
            if hwnd and sys.platform == "win32":
                self._setup_wndproc(hwnd)
        except Exception:
            pass

    def _setup_wndproc(self, hwnd):
        if sys.platform != "win32":
            return
        self._current_drop_zone = None
        self._hwnd = hwnd

        WNDPROC = ctypes.WINFUNCTYPE(
            ctypes.c_int64, ctypes.c_int64, ctypes.c_uint,
            ctypes.c_int64, ctypes.c_int64
        )
        self._original_wndproc = user32.SetWindowLongW(
            hwnd, GWL_WNDPROC, WNDPROC(self._wnd_proc)
        )

    def _wnd_proc(self, hwnd, msg, wparam, lparam):
        if msg == WM_DROPFILES:
            count = DragQueryFile(wparam, 0xFFFFFFFF, None, 0)
            files = []
            for i in range(count):
                buf = ctypes.create_unicode_buffer(260)
                DragQueryFile(wparam, i, buf, 260)
                files.append(buf.value)
            DragFinish(wparam)
            if files:
                # 检测鼠标位置确定落在哪个拖拽区
                pt = ctypes.c_long()
                user32.GetCursorPos(ctypes.byref(pt))
                self.root.after(0, self._on_drop_by_position, files, pt.value)
            return 0
        return user32.CallWindowProcW(self._original_wndproc, hwnd, msg, wparam, lparam)

    def _on_drop_by_position(self, files, screen_pt):
        """根据屏幕坐标判断拖到哪个区域"""
        x = screen_pt & 0xFFFF
        y = (screen_pt >> 16) & 0xFFFF

        # 获取两个 drop frame 的屏幕位置
        try:
            fx1 = self.drop1_frame.winfo_rootx()
            fy1 = self.drop1_frame.winfo_rooty()
            fw1 = self.drop1_frame.winfo_width()
            fh1 = self.drop1_frame.winfo_height()

            fx2 = self.drop2_frame.winfo_rootx()
            fy2 = self.drop2_frame.winfo_rooty()
            fw2 = self.drop2_frame.winfo_width()
            fh2 = self.drop2_frame.winfo_height()

            if fx1 <= x <= fx1 + fw1 and fy1 <= y <= fy1 + fh1:
                self._set_file1(files[0])
            elif fx2 <= x <= fx2 + fw2 and fy2 <= y <= fy2 + fh2:
                self._set_file2(files[0])
            else:
                # 默认放进品种一
                if not self.file1_path.get():
                    self._set_file1(files[0])
                else:
                    self._set_file2(files[0])
        except Exception:
            if not self.file1_path.get():
                self._set_file1(files[0])
            else:
                self._set_file2(files[0])

    def build_ui(self):
        main_frame = tk.Frame(self.root, bg=self.bg, padx=28, pady=24)
        main_frame.pack(fill=tk.BOTH, expand=True)

        # 标题
        title = tk.Label(
            main_frame,
            text="债券询价双表合并工具",
            font=("Microsoft YaHei UI", 17, "bold"),
            fg=self.text_color,
            bg=self.bg,
        )
        title.pack(pady=(0, 4))

        subtitle = tk.Label(
            main_frame,
            text="分别拖入品种一和品种二的询价导出表（.xls / .xlsx），一键合并",
            font=("Microsoft YaHei UI", 9),
            fg=self.sub_color,
            bg=self.bg,
        )
        subtitle.pack(pady=(0, 22))

        # 两个拖拽区（左右并排）
        drop_row = tk.Frame(main_frame, bg=self.bg)
        drop_row.pack(fill=tk.BOTH, expand=True, pady=(0, 20))

        # 品种一 (左)
        left_col = tk.Frame(drop_row, bg=self.bg)
        left_col.pack(side=tk.LEFT, fill=tk.BOTH, expand=True, padx=(0, 10))

        self.drop1_frame = self._build_drop_zone(left_col, "品种一",
                              self.file1_label_text, self._set_file1)

        # 品种二 (右)
        right_col = tk.Frame(drop_row, bg=self.bg)
        right_col.pack(side=tk.RIGHT, fill=tk.BOTH, expand=True, padx=(10, 0))

        self.drop2_frame = self._build_drop_zone(right_col, "品种二",
                              self.file2_label_text, self._set_file2)

        # 合并按钮
        btn_frame = tk.Frame(main_frame, bg=self.bg)
        btn_frame.pack(fill=tk.X, pady=(0, 14))

        self.merge_btn = tk.Button(
            btn_frame,
            text="开始合并",
            font=("Microsoft YaHei UI", 12, "bold"),
            bg=self.accent,
            fg="#ffffff",
            activebackground=self.accent_hover,
            activeforeground="#ffffff",
            relief="flat",
            bd=0,
            padx=28,
            pady=8,
            cursor="hand2",
            command=self.start_merge,
        )
        self.merge_btn.pack()

        # 进度条
        self.progress = ttk.Progressbar(
            main_frame, mode="indeterminate", length=400
        )

        # 状态栏
        status_frame = tk.Frame(main_frame, bg="#e2e8f0", padx=12, pady=8)
        status_frame.pack(fill=tk.X)

        self.status_label = tk.Label(
            status_frame,
            textvariable=self.status_var,
            font=("Microsoft YaHei UI", 9),
            fg=self.text_color,
            bg="#e2e8f0",
            anchor="w",
            justify="left",
        )
        self.status_label.pack(fill=tk.X)

    def _build_drop_zone(self, parent, title_text, label_var, setter):
        """构建单个拖拽区，返回 drop_frame"""
        lbl = tk.Label(
            parent,
            text=title_text,
            font=("Microsoft YaHei UI", 11, "bold"),
            fg=self.text_color,
            bg=self.bg,
        )
        lbl.pack(anchor="w", pady=(0, 6))

        drop_frame = tk.Frame(
            parent,
            bg=self.drop_bg,
            highlightbackground=self.drop_border,
            highlightthickness=1,
            relief="flat",
            bd=0,
            cursor="hand2",
        )
        drop_frame.pack(fill=tk.BOTH, expand=True)

        # 绑定点击
        drop_frame.bind("<Button-1>", lambda e, s=setter: self._browse_file(s))
        # 悬停效果
        drop_frame.bind("<Enter>", lambda e, df=drop_frame: df.configure(
            highlightbackground=self.drop_border_active))
        drop_frame.bind("<Leave>", lambda e, df=drop_frame, v=label_var: df.configure(
            highlightbackground=self.drop_filled_border if v.get() != f"拖拽或点击选择 {title_text} 文件"
            else self.drop_border))

        drop_icon = tk.Label(
            drop_frame,
            text="📂",
            font=("Microsoft YaHei UI", 30),
            bg=self.drop_bg,
            fg=self.accent,
            cursor="hand2",
        )
        drop_icon.pack(pady=(24, 6))
        drop_icon.bind("<Button-1>", lambda e, s=setter: self._browse_file(s))

        file_label = tk.Label(
            drop_frame,
            textvariable=label_var,
            font=("Microsoft YaHei UI", 9),
            bg=self.drop_bg,
            fg=self.sub_color,
            wraplength=240,
            justify="center",
            cursor="hand2",
        )
        file_label.pack(pady=(0, 24))
        file_label.bind("<Button-1>", lambda e, s=setter: self._browse_file(s))

        # 存引用
        drop_frame._file_label = file_label
        drop_frame._drop_icon = drop_icon
        drop_frame._label_var = label_var

        return drop_frame

    def _browse_file(self, setter_func):
        filepath = filedialog.askopenfilename(
            title="选择询价导出表",
            filetypes=[
                ("Excel 文件", "*.xls *.xlsx"),
                ("所有文件", "*.*"),
            ],
        )
        if filepath:
            setter_func(filepath)

    def _set_file1(self, filepath):
        if not filepath.lower().endswith(('.xls', '.xlsx')):
            messagebox.showwarning("文件格式", "请选择 .xls 或 .xlsx 文件")
            return
        self.file1_path.set(filepath)
        fname = os.path.basename(filepath)
        self.file1_label_text.set(f"✅ {fname}")
        self._update_drop_style(self.drop1_frame, True)
        self._update_status()

    def _set_file2(self, filepath):
        if not filepath.lower().endswith(('.xls', '.xlsx')):
            messagebox.showwarning("文件格式", "请选择 .xls 或 .xlsx 文件")
            return
        self.file2_path.set(filepath)
        fname = os.path.basename(filepath)
        self.file2_label_text.set(f"✅ {fname}")
        self._update_drop_style(self.drop2_frame, True)
        self._update_status()

    def _update_drop_style(self, drop_frame, filled):
        if filled:
            drop_frame.configure(
                bg=self.drop_filled_bg,
                highlightbackground=self.drop_filled_border,
            )
            if hasattr(drop_frame, '_drop_icon'):
                drop_frame._drop_icon.configure(bg=self.drop_filled_bg)
            if hasattr(drop_frame, '_file_label'):
                drop_frame._file_label.configure(bg=self.drop_filled_bg)
        else:
            drop_frame.configure(
                bg=self.drop_bg,
                highlightbackground=self.drop_border,
            )
            if hasattr(drop_frame, '_drop_icon'):
                drop_frame._drop_icon.configure(bg=self.drop_bg)
            if hasattr(drop_frame, '_file_label'):
                drop_frame._file_label.configure(bg=self.drop_bg)

    def _update_status(self):
        if self.file1_path.get() and self.file2_path.get():
            self.status_var.set("就绪 — 两个文件已选择，点击「开始合并」")
            self.merge_btn.configure(bg=self.success, activebackground="#047857")
        elif self.file1_path.get() or self.file2_path.get():
            self.status_var.set("请选择另一个文件")
        else:
            self.status_var.set("就绪 — 请选择两个文件后点击「开始合并」")
            self.merge_btn.configure(bg=self.accent, activebackground=self.accent_hover)

    def start_merge(self):
        if not self.file1_path.get() or not self.file2_path.get():
            messagebox.showwarning("文件缺失", "请先选择品种一和品种二两个文件")
            return

        # 选择保存位置
        output_path = filedialog.asksaveasfilename(
            title="保存合并结果",
            defaultextension=".xlsx",
            filetypes=[
                ("Excel 文件", "*.xlsx"),
                ("所有文件", "*.*"),
            ],
            initialfile="询价明细表（）-合并结果.xlsx",
        )
        if not output_path:
            return  # 用户取消

        self.set_processing(True)
        self.status_var.set("正在合并…")

        thread = threading.Thread(
            target=self._do_merge,
            args=(self.file1_path.get(), self.file2_path.get(), output_path),
            daemon=True,
        )
        thread.start()

    def _do_merge(self, path1, path2, output_path):
        try:
            if merge_two_tables is None or write_output_excel is None:
                raise RuntimeError("合并处理模块未加载，请检查 bond_merge_processor.py 是否存在")

            # 执行合并
            df = merge_two_tables(path1, path2)

            if df.empty:
                self.root.after(0, self._on_error, "两张表均无可处理的数据", "")
                return

            # 生成 Excel
            write_output_excel(df, output_path,
                               title="债券询价合并表",
                               info_text="债券信息")

            # 统计
            merged_count = int(df["_is_merged"].sum()) if "_is_merged" in df else 0
            variety1_only = int((df["备注"] == "品种一").sum())
            variety2_only = int((df["备注"] == "品种二").sum())
            variety_both = int((df["备注"] == "品种一/品种二").sum())
            total = len(df)

            summary = (
                f"总行数: {total}\n"
                f"品种一独有: {variety1_only}\n"
                f"品种二独有: {variety2_only}\n"
                f"合并行: {variety_both}\n"
                f"合并单元格数: {merged_count}"
            )

            self.root.after(0, self._on_success, output_path, summary)
        except Exception as e:
            err_msg = traceback.format_exc()
            self.root.after(0, self._on_error, str(e), err_msg)

    def _on_success(self, output_path, summary):
        self.set_processing(False)
        self.status_var.set(f"✅ 合并完成 → {output_path}")
        self.merge_btn.configure(bg=self.accent, activebackground=self.accent_hover)
        messagebox.showinfo(
            "合并完成",
            f"文件已保存至：\n{output_path}\n\n{summary}",
        )

    def _on_error(self, err_str, detail):
        self.set_processing(False)
        self.status_var.set(f"❌ 合并失败：{err_str[:80]}")
        self.merge_btn.configure(bg=self.error, activebackground="#b91c1c")
        self.root.after(500, lambda: self.merge_btn.configure(
            bg=self.accent, activebackground=self.accent_hover))
        messagebox.showerror(
            "合并失败",
            f"错误：{err_str}\n\n详细信息：\n{detail[-500:]}" if detail else f"错误：{err_str}",
        )

    def set_processing(self, processing):
        if processing:
            self.merge_btn.configure(state="disabled", text="合并中…")
            self.progress.pack(fill=tk.X, pady=(0, 12))
            self.progress.start(10)
        else:
            self.merge_btn.configure(state="normal", text="开始合并")
            self.progress.stop()
            self.progress.pack_forget()


# ---------- 入口 ----------
def main():
    root = tk.Tk()

    # 如果有命令行参数（拖拽到 exe 上），作为品种一
    app = BondMergeGUI(root)
    if len(sys.argv) > 1:
        filepath = sys.argv[1]
        if os.path.isfile(filepath) and filepath.lower().endswith(('.xls', '.xlsx')):
            app._set_file1(filepath)

    root.mainloop()


if __name__ == "__main__":
    main()
