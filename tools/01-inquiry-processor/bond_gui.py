#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
债券询价明细表处理工具 - Windows GUI
双击运行，拖拽或选择 Excel 文件即可自动处理。
"""

import os
import sys
import traceback
import tkinter as tk
from tkinter import ttk, filedialog, messagebox
from pathlib import Path
import threading

# ---------- 导入处理逻辑 ----------
try:
    sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
    from bond_inquiry_processor import (
        process_file,
        save_outputs,
    )
except ImportError:
    try:
        from bond_inquiry_processor import process_file, save_outputs
    except ImportError:
        process_file = None
        save_outputs = None


# ---------- Windows 拖拽支持 (WM_DROPFILES) ----------
if sys.platform == "win32":
    import ctypes
    from ctypes import wintypes

    # 定义所需 Windows API
    user32 = ctypes.windll.user32
    shell32 = ctypes.windll.shell32
    kernel32 = ctypes.windll.kernel32

    WM_DROPFILES = 0x0233
    GWL_WNDPROC = -4
    GWL_EXSTYLE = -20
    WS_EX_ACCEPTFILES = 0x00000010
    DragQueryFile = shell32.DragQueryFileW
    DragFinish = shell32.DragFinish

    DropCallback = None


def enable_drag_drop(hwnd, callback):
    """为窗口启用文件拖拽支持（仅 Windows）"""
    global DropCallback
    if sys.platform != "win32":
        return
    DropCallback = callback
    # 设置窗口为接受拖拽文件
    user32.SetWindowLongW(hwnd, GWL_EXSTYLE,
                          user32.GetWindowLongW(hwnd, GWL_EXSTYLE) | WS_EX_ACCEPTFILES)


# ---------- GUI ----------
class BondProcessorGUI:
    def __init__(self, root):
        self.root = root
        self.root.title("债券询价明细表处理工具")
        self.root.geometry("520x380")
        self.root.resizable(True, True)
        self.root.minsize(420, 320)

        # 配色
        self.bg = "#f0f4f8"
        self.accent = "#2563eb"
        self.accent_hover = "#1d4ed8"
        self.success = "#059669"
        self.error = "#dc2626"
        self.text_color = "#1e293b"
        self.sub_color = "#64748b"
        self.card_bg = "#ffffff"

        self.root.configure(bg=self.bg)

        # 设置图标（如有）
        try:
            if sys.platform == "win32":
                self.root.iconbitmap(default="")
        except Exception:
            pass

        self.build_ui()

        # 启用拖拽
        self.root.after(100, self._enable_dnd)

    def _enable_dnd(self):
        try:
            hwnd = self.root.winfo_id()
            if hwnd:
                enable_drag_drop(hwnd, self.on_file_drop)
        except Exception:
            pass

    def build_ui(self):
        # 主容器
        main_frame = tk.Frame(self.root, bg=self.bg, padx=24, pady=24)
        main_frame.pack(fill=tk.BOTH, expand=True)

        # 标题
        title = tk.Label(
            main_frame,
            text="📊 债券询价明细表处理工具",
            font=("Microsoft YaHei UI", 16, "bold"),
            fg=self.text_color,
            bg=self.bg,
        )
        title.pack(pady=(0, 6))

        subtitle = tk.Label(
            main_frame,
            text="支持 .xls / .xlsx 询价导出表，自动分类整理并生成标准格式",
            font=("Microsoft YaHei UI", 9),
            fg=self.sub_color,
            bg=self.bg,
        )
        subtitle.pack(pady=(0, 20))

        # 拖拽区域
        self.drop_frame = tk.Frame(
            main_frame,
            bg=self.card_bg,
            highlightbackground="#cbd5e1",
            highlightthickness=1,
            relief="flat",
            bd=0,
        )
        self.drop_frame.pack(fill=tk.BOTH, expand=True, pady=(0, 16))

        # 拖拽图标和文字
        self.drop_icon = tk.Label(
            self.drop_frame,
            text="📁",
            font=("Microsoft YaHei UI", 36),
            bg=self.card_bg,
            fg=self.accent,
        )
        self.drop_icon.pack(pady=(30, 8))

        self.drop_label = tk.Label(
            self.drop_frame,
            text="拖拽 Excel 文件到此处",
            font=("Microsoft YaHei UI", 13, "bold"),
            fg=self.text_color,
            bg=self.card_bg,
        )
        self.drop_label.pack()

        self.drop_hint = tk.Label(
            self.drop_frame,
            text="或点击下方按钮选择文件",
            font=("Microsoft YaHei UI", 9),
            fg=self.sub_color,
            bg=self.card_bg,
        )
        self.drop_hint.pack(pady=(4, 0))

        # 选择文件按钮
        self.btn_select = tk.Button(
            self.drop_frame,
            text="选择 Excel 文件",
            font=("Microsoft YaHei UI", 11),
            bg=self.accent,
            fg="white",
            activebackground=self.accent_hover,
            activeforeground="white",
            relief="flat",
            bd=0,
            padx=24,
            pady=8,
            cursor="hand2",
            command=self.select_file,
        )
        self.btn_select.pack(pady=(16, 30))

        # 状态栏
        self.status_var = tk.StringVar(value="等待选择文件…")
        self.status_bar = tk.Label(
            main_frame,
            textvariable=self.status_var,
            font=("Microsoft YaHei UI", 9),
            fg=self.sub_color,
            bg=self.bg,
            anchor="w",
        )
        self.status_bar.pack(fill=tk.X)

        # 底部进度条
        self.progress = ttk.Progressbar(
            main_frame,
            mode="indeterminate",
            length=200,
        )

    def on_file_drop(self, files):
        """拖拽文件回调"""
        if files:
            filepath = files[0]
            self.process_file(filepath)

    def select_file(self):
        """点击选择文件"""
        filepath = filedialog.askopenfilename(
            title="选择询价导出表",
            filetypes=[
                ("Excel 文件", "*.xls *.xlsx"),
                ("所有文件", "*.*"),
            ],
        )
        if filepath:
            self.process_file(filepath)

    def process_file(self, filepath):
        """处理文件（在后台线程运行）"""
        if not filepath.lower().endswith(('.xls', '.xlsx')):
            messagebox.showwarning("文件格式", "请选择 .xls 或 .xlsx 文件")
            return

        # 更新 UI
        self.set_processing(True)
        self.status_var.set(f"处理中：{os.path.basename(filepath)}…")

        # 后台处理
        thread = threading.Thread(
            target=self._do_process,
            args=(filepath,),
            daemon=True,
        )
        thread.start()

    def _do_process(self, filepath):
        """实际处理逻辑"""
        try:
            from datetime import datetime
            input_dir = os.path.dirname(os.path.abspath(filepath))
            output_filename = f"输出询价明细表（）-{datetime.now().strftime('%Y%m%d')}.xlsx"
            output_path = os.path.join(input_dir, output_filename)

            # 查找同目录下的模板文件
            template_path = os.path.join(input_dir, "询价明细表模板.xlsx")
            if not os.path.exists(template_path):
                alt_template = os.path.join(os.path.dirname(os.path.abspath(__file__)), "询价明细表模板.xlsx")
                if os.path.exists(alt_template):
                    template_path = alt_template
                else:
                    template_path = None

            if process_file is not None and save_outputs is not None:
                # 使用原有处理器
                df = process_file(filepath)
                save_outputs(df, output_path, input_path=filepath, template_path=template_path)
            else:
                # 内联调用
                self._run_inline_processor(filepath, output_path, template_path)

            self.root.after(0, self._on_success, output_path)
        except Exception as e:
            err_msg = traceback.format_exc()
            self.root.after(0, self._on_error, str(e), err_msg)

    def _run_inline_processor(self, input_path, output_path, template_path):
        """内联处理（当无法导入 bond_inquiry_processor 时）"""
        # 这里会由 PyInstaller 打包时自动包含 bond_inquiry_processor
        # 如果到这里还没导入，说明打包有问题
        raise RuntimeError("处理模块未加载，请检查打包配置")

    def _on_success(self, output_path):
        """处理成功回调"""
        self.set_processing(False)
        self.status_var.set(f"✅ 完成 → 已保存至：{output_path}")
        self.root.after(500, lambda: messagebox.showinfo(
            "处理完成",
            f"文件已保存至：\n{output_path}",
        ))

    def _on_error(self, err_str, detail):
        """处理失败回调"""
        self.set_processing(False)
        self.status_var.set(f"❌ 处理失败：{err_str[:80]}")
        self.root.after(500, lambda: messagebox.showerror(
            "处理失败",
            f"错误：{err_str}\n\n详细信息：\n{detail[-500:]}",
        ))

    def set_processing(self, processing):
        """切换 UI 处理状态"""
        if processing:
            self.btn_select.config(state="disabled", text="处理中…")
            self.progress.pack(fill=tk.X, pady=(8, 0))
            self.progress.start(10)
        else:
            self.btn_select.config(state="normal", text="选择 Excel 文件")
            self.progress.stop()
            self.progress.pack_forget()


# ---------- CLI 无窗口处理（拖拽文件到 EXE） ----------
def _cli_process(filepath: str) -> None:
    """命令行模式：直接处理文件，不弹 GUI 窗口，完成后输出结果路径"""
    from datetime import datetime

    if not os.path.isfile(filepath):
        _show_error("文件不存在", f"找不到文件：\n{filepath}")
        return

    if not filepath.lower().endswith(('.xls', '.xlsx')):
        _show_error("不支持的文件格式", "请拖拽 .xls 或 .xlsx 文件到 EXE 图标上")
        return

    input_dir = os.path.dirname(os.path.abspath(filepath))
    output_filename = f"输出询价明细表（）-{datetime.now().strftime('%Y%m%d')}.xlsx"
    output_path = os.path.join(input_dir, output_filename)

    # 查找模板文件
    template_path = os.path.join(input_dir, "询价明细表模板.xlsx")
    if not os.path.exists(template_path):
        alt_template = os.path.join(os.path.dirname(os.path.abspath(__file__)), "询价明细表模板.xlsx")
        if os.path.exists(alt_template):
            template_path = alt_template
        else:
            template_path = None

    try:
        print(f"正在处理：{filepath}")
        if process_file is None or save_outputs is None:
            raise RuntimeError("处理模块未加载，请检查打包配置（bond_inquiry_processor.py 未包含在 EXE 中）")
        df = process_file(filepath)
        save_outputs(df, output_path, input_path=filepath, template_path=template_path)
        print(f"处理完成 → {output_path}")
        _show_info("处理完成", f"文件已保存至：\n{output_path}")
    except Exception as e:
        err_detail = traceback.format_exc()
        print(f"处理失败：{e}\n{err_detail}")
        _show_error("处理失败", f"{e}\n\n{err_detail[-300:]}")


def _show_info(title: str, message: str) -> None:
    """弹出信息提示框（仅 Windows）"""
    if sys.platform == "win32":
        ctypes.windll.user32.MessageBoxW(0, message, title, 0x40)
    else:
        print(f"[{title}] {message}")


def _show_error(title: str, message: str) -> None:
    """弹出错误提示框（仅 Windows）"""
    if sys.platform == "win32":
        ctypes.windll.user32.MessageBoxW(0, message, title, 0x10)
    else:
        print(f"[{title}] {message}", file=sys.stderr)


# ---------- 入口 ----------
def main():
    # 拖拽文件到 EXE → sys.argv[1] 为文件路径 → 无窗口直接处理
    if len(sys.argv) > 1:
        filepath = sys.argv[1]
        _cli_process(filepath)
        return

    # 双击 EXE → 启动 GUI
    root = tk.Tk()

    app = BondProcessorGUI(root)

    # Windows 消息循环（处理 WM_DROPFILES）—— GUI 内的拖拽
    if sys.platform == "win32":
        original_wndproc = None

        def wnd_proc(hwnd, msg, wparam, lparam):
            if msg == WM_DROPFILES:
                count = DragQueryFile(wparam, 0xFFFFFFFF, None, 0)
                files = []
                for i in range(count):
                    buf = ctypes.create_unicode_buffer(260)
                    DragQueryFile(wparam, i, buf, 260)
                    files.append(buf.value)
                DragFinish(wparam)
                if files and app:
                    root.after(0, app.on_file_drop, files)
                return 0
            return user32.CallWindowProcW(original_wndproc, hwnd, msg, wparam, lparam)

        def setup_wndproc():
            nonlocal original_wndproc
            hwnd = root.winfo_id()
            if hwnd:
                WNDPROC = ctypes.WINFUNCTYPE(ctypes.c_int64, ctypes.c_int64, ctypes.c_uint, ctypes.c_int64, ctypes.c_int64)
                original_wndproc = user32.SetWindowLongW(hwnd, GWL_WNDPROC, WNDPROC(wnd_proc))

        root.after(200, setup_wndproc)

    root.mainloop()


if __name__ == "__main__":
    main()
