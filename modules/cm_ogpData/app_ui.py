#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
@Project : tools
@File    : app_ui.py
@Author  : Shawn
@Date    : 2026/1/8 9:45
@Info    : 主程序界面，主要是可视化界面及实现
"""

import tkinter as tk
from tkinter import filedialog, messagebox, scrolledtext
import os
from datetime import datetime
from tkinter import ttk
import threading

# 导入数据处理模块
from ogp_processor import OGPProcessor
from three_d_processor import ThreeDProcessor

# 检查pandas是否可用
PANDAS_AVAILABLE = False
try:
    import pandas as pd
    PANDAS_AVAILABLE = True
except ImportError:
    pass

class DataSorterApp:
    def __init__(self, root):
        self.root = root
        self.root.title("测量数据汇总工具")
        self.root.geometry("900x850")

        # 设置默认值
        self.header_lines = tk.IntVar(value=6)
        self.file_paths = []
        self.output_dir = tk.StringVar(value=os.path.expanduser("~"))
        self.processing = False
        
        # 三次元相关变量
        self.three_d_file_paths = []
        self.three_d_output_dir = tk.StringVar(value=os.path.expanduser("~"))
        self.three_d_processing = False
        self.three_d_create_subfolder = tk.BooleanVar(value=True)

        # 创建数据处理器实例
        self.ogp_processor = OGPProcessor()
        self.three_d_processor = ThreeDProcessor()

        # 创建UI组件
        self.create_widgets()

    def create_widgets(self):
        # 创建Notebook（Tab页容器）
        self.notebook = ttk.Notebook(self.root)
        self.notebook.pack(fill="both", expand=True, padx=10, pady=5)
        
        # 创建第一个Tab：OGP数据处理（原有功能）
        self.ogp_frame = tk.Frame(self.notebook)
        self.notebook.add(self.ogp_frame, text="OGP")
        self.create_ogp_widgets()
        
        # 创建第二个Tab：三次元数据处理
        self.three_d_frame = tk.Frame(self.notebook)
        self.notebook.add(self.three_d_frame, text="三次元")
        self.create_three_d_widgets()
        
        # 状态栏
        self.status_bar = tk.Label(self.root, text="就绪", bd=1, relief=tk.SUNKEN, anchor=tk.W)
        self.status_bar.pack(side=tk.BOTTOM, fill=tk.X)

    def create_ogp_widgets(self):
        """创建OGP数据处理Tab的界面"""
        # 创建主框架，使用PanedWindow布局以支持左右调整宽度
        main_frame = tk.PanedWindow(self.ogp_frame, orient=tk.HORIZONTAL, sashrelief=tk.RAISED, sashwidth=5)
        main_frame.pack(fill="both", expand=True, padx=10, pady=5)

        # === 左侧区域：文件选择和参数设置 ===
        left_frame = tk.Frame(main_frame)
        main_frame.add(left_frame, width=400)

        # 处理模式选择
        self.process_mode = tk.StringVar(value="summary")  # 默认汇总模式

        # 文件选择区域
        file_frame = tk.LabelFrame(left_frame, text="文件选择", padx=8, pady=8)
        file_frame.pack(fill="x", pady=5)

        # 文件列表显示
        tk.Label(file_frame, text="已选择的文件:", font=("Arial", 9, "bold")).pack(anchor="w")

        # Treeview框架
        tree_frame = tk.Frame(file_frame)
        tree_frame.pack(fill="both", expand=True, pady=5)

        # 滚动条
        scrollbar = tk.Scrollbar(tree_frame)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)

        # Treeview
        self.file_tree = ttk.Treeview(tree_frame, columns=("序号", "文件名"),
                                      show="headings", yscrollcommand=scrollbar.set, height=6)

        # 设置列
        self.file_tree.heading("序号", text="序号")
        self.file_tree.heading("文件名", text="文件名")

        self.file_tree.column("序号", width=50, anchor="center")
        self.file_tree.column("文件名", width=250)

        self.file_tree.pack(side=tk.LEFT, fill="both", expand=True)
        scrollbar.config(command=self.file_tree.yview)

        # 右键菜单
        self.file_tree.bind("<Button-3>", self.show_file_context_menu)

        # 文件操作按钮（水平排列）
        file_buttons_frame = tk.Frame(file_frame)
        file_buttons_frame.pack(fill="x", pady=5)

        tk.Button(file_buttons_frame, text="添加文件", command=self.browse_files,
                  width=12).pack(side=tk.LEFT, padx=2)
        tk.Button(file_buttons_frame, text="移除选中", command=self.remove_selected_files,
                  width=12).pack(side=tk.LEFT, padx=2)
        tk.Button(file_buttons_frame, text="清空列表", command=self.clear_file_list,
                  width=12).pack(side=tk.LEFT, padx=2)

        # 文件统计信息
        self.file_count_label = tk.Label(file_frame, text="已选择 0 个文件", fg="blue", font=("Arial", 9))
        self.file_count_label.pack(anchor="w")

        # 输出文件夹选择区域
        output_frame = tk.LabelFrame(left_frame, text="输出设置", padx=8, pady=8)
        output_frame.pack(fill="x", pady=5)

        # 输出文件夹输入框和按钮
        output_input_frame = tk.Frame(output_frame)
        output_input_frame.pack(fill="x", pady=2)

        tk.Label(output_input_frame, text="文件夹:", width=8).pack(side=tk.LEFT)
        tk.Entry(output_input_frame, textvariable=self.output_dir, width=30).pack(side=tk.LEFT, padx=2, fill="x",
                                                                                  expand=True)
        tk.Button(output_input_frame, text="选择", command=self.browse_output_dir, width=8).pack(side=tk.LEFT, padx=2)

        # 处理选项
        options_frame = tk.Frame(output_frame)
        options_frame.pack(fill="x", pady=5)

        self.create_subfolder = tk.BooleanVar(value=True)
        tk.Checkbutton(options_frame, text="创建子文件夹", variable=self.create_subfolder,
                       font=("Arial", 9)).pack(anchor="w")

        

        # 参数设置区域
        param_frame = tk.LabelFrame(left_frame, text="处理参数", padx=8, pady=8)
        param_frame.pack(fill="x", pady=5)

        param_input_frame = tk.Frame(param_frame)
        param_input_frame.pack(fill="x", pady=5)

        tk.Label(param_input_frame, text="表头行数:", width=10).pack(side=tk.LEFT)
        tk.Entry(param_input_frame, textvariable=self.header_lines, width=10).pack(side=tk.LEFT, padx=5)

        # 处理按钮区域（在左侧框架底部）
        button_frame = tk.Frame(left_frame)
        button_frame.pack(fill="x", pady=10)

        # 开始处理按钮
        self.start_button = tk.Button(button_frame, text="开始批量处理", command=self.start_batch_processing,
                                      bg="#4CAF50", fg="white", font=("Arial", 11, "bold"),
                                      height=2, state=tk.DISABLED)
        self.start_button.pack(fill="x", pady=2)

        # 停止处理按钮
        self.stop_button = tk.Button(button_frame, text="停止处理", command=self.stop_processing,
                                     bg="#FF9800", fg="white", height=1, state=tk.DISABLED)
        self.stop_button.pack(fill="x", pady=2)

        # === 右侧区域：结果显示 ===
        right_frame = tk.Frame(main_frame)
        main_frame.add(right_frame)

        # 结果显示区域
        result_frame = tk.LabelFrame(right_frame, text="处理结果", padx=8, pady=8)
        result_frame.pack(fill="both", expand=True)

        # 使用ScrolledText显示结果
        self.result_text_area = scrolledtext.ScrolledText(result_frame, wrap=tk.WORD, font=("Courier New", 9))
        self.result_text_area.pack(fill="both", expand=True)

    def browse_files_generic(self, title, filetypes, file_paths_var, tree_var, count_label_var, output_dir_var, start_button_var, pandas_check=False):
        """通用文件选择方法（OGP和三次元兼容）"""
        if pandas_check and not PANDAS_AVAILABLE:
            messagebox.showerror("错误", "pandas库未安装，无法处理Excel文件。\n请运行: pip install pandas openpyxl")
            return

        filenames = filedialog.askopenfilenames(
            title=title,
            filetypes=filetypes
        )

        if filenames:
            file_paths = getattr(self, file_paths_var)
            tree = getattr(self, tree_var)
            count_label = getattr(self, count_label_var)
            output_dir = getattr(self, output_dir_var)
            start_button = getattr(self, start_button_var)

            for filename in filenames:
                if filename not in file_paths:
                    file_paths.append(filename)
                    # 添加到Treeview
                    index = len(file_paths)
                    tree.insert("", "end", values=(
                        index,
                        os.path.basename(filename)
                    ))

            # 更新统计信息
            self.update_file_count_generic(file_paths_var, count_label_var, start_button_var, pandas_check)

            # 自动设置输出文件夹为第一个文件所在目录
            if file_paths and (not output_dir.get() or output_dir.get() == os.path.expanduser("~")):
                input_dir = os.path.dirname(file_paths[0])
                output_dir.set(input_dir)

            self.status_bar.config(text=f"已添加 {len(filenames)} 个文件")

    def browse_output_dir_generic(self, title, output_dir_var):
        """通用输出文件夹选择方法（OGP和三次元兼容）"""
        output_dir = getattr(self, output_dir_var)
        dirname = filedialog.askdirectory(
            title=title,
            initialdir=output_dir.get()
        )
        if dirname:
            output_dir.set(dirname)
            self.status_bar.config(text=f"输出文件夹: {dirname}")

    def clear_file_list_generic(self, file_paths_var, tree_var, count_label_var, start_button_var, pandas_check=False):
        """通用清空文件列表方法（OGP和三次元兼容）"""
        file_paths = getattr(self, file_paths_var)
        if file_paths:
            if messagebox.askyesno("确认", "确定要清空所有文件吗？"):
                file_paths.clear()
                # 清空Treeview
                tree = getattr(self, tree_var)
                for item in tree.get_children():
                    tree.delete(item)
                self.update_file_count_generic(file_paths_var, count_label_var, start_button_var, pandas_check)
                self.status_bar.config(text="已清空文件列表")

    def remove_selected_files_generic(self, tree_var, file_paths_var, count_label_var, start_button_var, pandas_check=False):
        """通用移除选中文件方法（OGP和三次元兼容）"""
        tree = getattr(self, tree_var)
        selected_items = tree.selection()
        if not selected_items:
            messagebox.showinfo("提示", "请先选择要移除的文件")
            return

        # 确认移除
        if messagebox.askyesno("确认", f"确定要移除选中的 {len(selected_items)} 个文件吗？"):
            file_paths = getattr(self, file_paths_var)
            for item in selected_items:
                values = tree.item(item, "values")
                if values:
                    # 从文件路径列表中移除
                    filename = values[1]
                    # 需要找到完整路径
                    for file_path in file_paths[:]:
                        if os.path.basename(file_path) == filename:
                            file_paths.remove(file_path)
                            break
                # 从Treeview中移除
                tree.delete(item)

            # 重新排序
            self.reorder_file_list_generic(tree_var, file_paths_var)
            self.update_file_count_generic(file_paths_var, count_label_var, start_button_var, pandas_check)
            self.status_bar.config(text=f"已移除 {len(selected_items)} 个文件")

    def reorder_file_list_generic(self, tree_var, file_paths_var):
        """通用重新排序文件列表方法（OGP和三次元兼容）"""
        tree = getattr(self, tree_var)
        file_paths = getattr(self, file_paths_var)
        # 清空Treeview
        for item in tree.get_children():
            tree.delete(item)

        # 重新添加
        for i, filename in enumerate(file_paths, 1):
            tree.insert("", "end", values=(
                i,
                os.path.basename(filename)
            ))

    def update_file_count_generic(self, file_paths_var, count_label_var, start_button_var, pandas_check=False):
        """通用更新文件计数方法（OGP和三次元兼容）"""
        file_paths = getattr(self, file_paths_var)
        count = len(file_paths)
        count_label = getattr(self, count_label_var)
        count_label.config(text=f"已选择 {count} 个文件")

        # 根据文件数量启用/禁用处理按钮
        start_button = getattr(self, start_button_var)
        if count > 0 and (not pandas_check or PANDAS_AVAILABLE):
            start_button.config(state=tk.NORMAL)
        else:
            start_button.config(state=tk.DISABLED)

    def show_file_context_menu_generic(self, event, tree_var, file_paths_var, remove_method):
        """通用显示文件右键菜单方法（OGP和三次元兼容）"""
        tree = getattr(self, tree_var)
        # 选择右键点击的项目
        item = tree.identify_row(event.y)
        if item:
            tree.selection_set(item)

            menu = tk.Menu(self.root, tearoff=0)
            menu.add_command(label="打开文件", command=lambda: self.open_selected_file_generic(tree_var, file_paths_var))
            menu.add_command(label="打开所在文件夹", command=lambda: self.open_file_folder_generic(tree_var, file_paths_var))
            menu.add_separator()
            menu.add_command(label="从列表中移除", command=remove_method)
            menu.tk_popup(event.x_root, event.y_root)

    def open_selected_file_generic(self, tree_var, file_paths_var):
        """通用打开选中文件方法（OGP和三次元兼容）"""
        tree = getattr(self, tree_var)
        file_paths = getattr(self, file_paths_var)
        selected_items = tree.selection()
        if selected_items:
            values = tree.item(selected_items[0], "values")
            if values:
                filename = values[1]
                # 找到完整路径
                for file_path in file_paths:
                    if os.path.basename(file_path) == filename:
                        try:
                            if os.name == 'nt':
                                os.startfile(file_path)
                            elif os.name == 'posix':
                                import subprocess
                                subprocess.call(
                                    ['open', file_path] if os.sys.platform == 'darwin' else ['xdg-open', file_path])
                        except Exception as e:
                            messagebox.showwarning("打开失败", f"无法打开文件:\n{str(e)}")
                        break

    def open_file_folder_generic(self, tree_var, file_paths_var):
        """通用打开文件所在文件夹方法（OGP和三次元兼容）"""
        tree = getattr(self, tree_var)
        file_paths = getattr(self, file_paths_var)
        selected_items = tree.selection()
        if selected_items:
            values = tree.item(selected_items[0], "values")
            if values:
                filename = values[1]
                # 找到完整路径
                for file_path in file_paths:
                    if os.path.basename(file_path) == filename:
                        folder_path = os.path.dirname(file_path)
                        try:
                            if os.name == 'nt':
                                os.startfile(folder_path)
                            elif os.name == 'posix':
                                import subprocess
                                subprocess.call(
                                    ['open', folder_path] if os.sys.platform == 'darwin' else ['xdg-open', folder_path])
                        except Exception as e:
                            messagebox.showwarning("打开失败", f"无法打开文件夹:\n{str(e)}")
                        break

    def browse_files(self):
        """浏览并选择多个OGP文件"""
        self.browse_files_generic(
            title="选择OGP数据文件",
            filetypes=[("Excel files", "*.xls *.xlsx"), ("All files", "*.*")],
            file_paths_var="file_paths",
            tree_var="file_tree",
            count_label_var="file_count_label",
            output_dir_var="output_dir",
            start_button_var="start_button",
            pandas_check=True
        )

    def browse_output_dir(self):
        """浏览并选择OGP输出文件夹"""
        self.browse_output_dir_generic(
            title="选择输出文件夹",
            output_dir_var="output_dir"
        )

    def clear_file_list(self):
        """清空OGP文件列表"""
        self.clear_file_list_generic(
            file_paths_var="file_paths",
            tree_var="file_tree",
            count_label_var="file_count_label",
            start_button_var="start_button",
            pandas_check=True
        )

    def remove_selected_files(self):
        """移除选中的OGP文件"""
        self.remove_selected_files_generic(
            tree_var="file_tree",
            file_paths_var="file_paths",
            count_label_var="file_count_label",
            start_button_var="start_button",
            pandas_check=True
        )

    def reorder_file_list(self):
        """重新排序OGP文件列表"""
        self.reorder_file_list_generic(
            tree_var="file_tree",
            file_paths_var="file_paths"
        )

    def update_file_count(self):
        """更新OGP文件计数"""
        self.update_file_count_generic(
            file_paths_var="file_paths",
            count_label_var="file_count_label",
            start_button_var="start_button",
            pandas_check=True
        )

    def show_file_context_menu(self, event):
        """显示OGP文件右键菜单"""
        self.show_file_context_menu_generic(
            event=event,
            tree_var="file_tree",
            file_paths_var="file_paths",
            remove_method=self.remove_selected_files
        )

    def open_selected_file(self):
        """打开选中的OGP文件"""
        self.open_selected_file_generic(
            tree_var="file_tree",
            file_paths_var="file_paths"
        )

    def open_file_folder(self):
        """打开OGP文件所在文件夹"""
        self.open_file_folder_generic(
            tree_var="file_tree",
            file_paths_var="file_paths"
        )

    def start_batch_processing(self):
        """开始批量处理OGP文件（OGP专用）"""
        if not self.file_paths:
            messagebox.showerror("错误", "请先选择要处理的文件！")
            return

        # 检查输出文件夹
        if not self.output_dir.get():
            messagebox.showerror("错误", "请选择输出文件夹！")
            return

        # 创建输出文件夹（如果需要）
        try:
            os.makedirs(self.output_dir.get(), exist_ok=True)
        except Exception as e:
            messagebox.showerror("错误", f"无法创建输出文件夹:\n{str(e)}")
            return

        # 禁用处理按钮，启用停止按钮
        self.toggle_buttons(processing=True)
        self.processing = True

        # 在新线程中处理
        thread = threading.Thread(target=self.process_files_thread, daemon=True)
        thread.start()

    def stop_processing(self):
        """停止OGP处理（OGP专用）"""
        self.processing = False
        self.status_bar.config(text="正在停止处理...")

    def toggle_buttons(self, processing):
        """切换OGP按钮状态（OGP专用）"""
        if processing:
            self.start_button.config(state=tk.DISABLED)
            self.stop_button.config(state=tk.NORMAL)
        else:
            self.start_button.config(state=tk.NORMAL if self.file_paths else tk.DISABLED)
            self.stop_button.config(state=tk.DISABLED)

    def process_files_thread(self):
        """处理OGP文件的线程函数（OGP专用）"""
        total_files = len(self.file_paths)
        success_count = 0
        fail_count = 0

        # 清空结果区域
        self.result_text_area.delete(1.0, tk.END)

        # 显示开始信息
        self.show_result_header(total_files)

        # 根据处理模式选择不同的处理方式
        mode_text = "排序" if self.process_mode.get() == "sort" else "汇总"

        if self.process_mode.get() == "summary" and total_files > 1:
            # 汇总模式且多个文件，使用合并方法
            self.status_bar.config(text="正在合并多个文件到一个汇总文件...")

            # 调用合并方法
            success, message, output_file = self.ogp_processor.merge_ogp_files(
                self.file_paths, self.output_dir.get(), self.processing, 
                self.create_subfolder.get()
            )

            if success:
                success_count = 1
                self.append_result(f"✓ {message}\n")
            else:
                fail_count = 1
                self.append_result(f"✗ {message}\n")
        else:
            # 排序模式或单个文件，使用原有的处理方式
            for i, file_path in enumerate(self.file_paths, 1):
                if not self.processing:
                    self.append_result("\n\n处理已停止！\n")
                    break

                self.status_bar.config(text=f"正在处理: {os.path.basename(file_path)}")

                # 处理单个文件
                success, message = self.ogp_processor.process_single_file(
                    file_path, i, self.process_mode.get(), "format2",
                    self.header_lines.get(), self.output_dir.get(),
                    self.create_subfolder.get(), False
                )

                if success:
                    success_count += 1
                    self.append_result(f"✓ {message}\n")
                else:
                    fail_count += 1
                    self.append_result(f"✗ {message}\n")

        # 处理完成
        self.processing = False
        self.toggle_buttons(processing=False)

        # 显示总结
        self.show_result_summary(success_count, fail_count)

        self.status_bar.config(text=f"批量处理完成: 成功 {success_count} 个，失败 {fail_count} 个")

    def show_result_header(self, total_files):
        """显示OGP结果头部信息（OGP专用）"""
        mode_text = "排序" if self.process_mode.get() == "sort" else "汇总"

        result_info = f"{'=' * 70}\n"
        result_info += f"批量{mode_text}开始 - {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}\n"
        result_info += f"{'=' * 70}\n\n"

        result_info += f"文件总数: {total_files}\n"
        result_info += f"输出文件夹: {self.output_dir.get()}\n"
        result_info += f"表头行数: {self.header_lines.get()}\n"
        result_info += f"处理模式: {mode_text}\n\n"

        result_info += f"{'=' * 70}\n"
        result_info += "处理结果:\n"
        result_info += f"{'=' * 70}\n\n"

        self.result_text_area.insert(1.0, result_info)

    def append_result(self, message):
        """在OGP结果区域追加信息（OGP专用）"""
        self.result_text_area.insert(tk.END, message)
        self.result_text_area.see(tk.END)
        self.root.update()

    def show_result_summary(self, success_count, fail_count):
        """显示OGP处理总结（OGP专用）"""
        mode_text = "排序" if self.process_mode.get() == "sort" else "汇总"

        summary = f"\n{'=' * 70}\n"
        summary += f"批量{mode_text}完成 - {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}\n"
        summary += f"{'=' * 70}\n\n"

        summary += f"✓ 成功处理: {success_count} 个文件\n"
        summary += f"✗ 处理失败: {fail_count} 个文件\n"
        summary += f"共计: {success_count + fail_count} 个文件\n\n"

        if fail_count == 0:
            summary += f"🎉 所有文件{mode_text}成功！\n"
        else:
            summary += f"⚠️  有 {fail_count} 个文件{mode_text}失败\n"

        summary += f"{'=' * 70}\n"

        self.append_result(summary)

    def create_three_d_widgets(self):
        """创建三次元数据处理Tab的界面"""
        # 创建主框架，使用PanedWindow布局以支持左右调整宽度
        main_frame = tk.PanedWindow(self.three_d_frame, orient=tk.HORIZONTAL, sashrelief=tk.RAISED, sashwidth=5)
        main_frame.pack(fill="both", expand=True, padx=10, pady=5)

        # === 左侧区域：文件选择和参数设置 ===
        left_frame = tk.Frame(main_frame)
        main_frame.add(left_frame, width=400)

        # 文件选择区域
        file_frame = tk.LabelFrame(left_frame, text="Excel文件选择", padx=8, pady=8)
        file_frame.pack(fill="x", pady=5)

        # 文件列表显示
        tk.Label(file_frame, text="已选择的Excel文件:", font=("Arial", 9, "bold")).pack(anchor="w")

        # Treeview框架
        tree_frame = tk.Frame(file_frame)
        tree_frame.pack(fill="both", expand=True, pady=5)

        # 滚动条
        scrollbar_3d = tk.Scrollbar(tree_frame)
        scrollbar_3d.pack(side=tk.RIGHT, fill=tk.Y)

        # Treeview
        self.three_d_file_tree = ttk.Treeview(tree_frame, columns=("序号", "文件名"),
                                              show="headings", yscrollcommand=scrollbar_3d.set, height=8)

        # 设置列
        self.three_d_file_tree.heading("序号", text="序号")
        self.three_d_file_tree.heading("文件名", text="文件名")

        self.three_d_file_tree.column("序号", width=50, anchor="center")
        self.three_d_file_tree.column("文件名", width=250)

        self.three_d_file_tree.pack(side=tk.LEFT, fill="both", expand=True)
        scrollbar_3d.config(command=self.three_d_file_tree.yview)

        # 右键菜单
        self.three_d_file_tree.bind("<Button-3>", self.show_three_d_file_context_menu)

        # 文件操作按钮（水平排列）
        file_buttons_frame = tk.Frame(file_frame)
        file_buttons_frame.pack(fill="x", pady=5)

        tk.Button(file_buttons_frame, text="添加Excel文件", command=self.browse_three_d_files,
                  width=12).pack(side=tk.LEFT, padx=2)
        tk.Button(file_buttons_frame, text="移除选中", command=self.remove_selected_three_d_files,
                  width=12).pack(side=tk.LEFT, padx=2)
        tk.Button(file_buttons_frame, text="清空列表", command=self.clear_three_d_file_list,
                  width=12).pack(side=tk.LEFT, padx=2)

        # 文件统计信息
        self.three_d_file_count_label = tk.Label(file_frame, text="已选择 0 个文件", fg="blue", font=("Arial", 9))
        self.three_d_file_count_label.pack(anchor="w")

        # 规则栏
        rule_frame = tk.LabelFrame(left_frame, text="规则", padx=8, pady=8)
        rule_frame.pack(fill="x", pady=5)

        self.three_d_layout_mode = tk.StringVar(value="horizontal")  # 默认横排
        tk.Radiobutton(rule_frame, text="横排", variable=self.three_d_layout_mode,
                       value="horizontal", font=("Arial", 9)).pack(anchor="w", pady=2)
        tk.Radiobutton(rule_frame, text="竖排", variable=self.three_d_layout_mode,
                       value="vertical", font=("Arial", 9)).pack(anchor="w", pady=2)

        # 输出文件夹选择区域
        output_frame = tk.LabelFrame(left_frame, text="输出设置", padx=8, pady=8)
        output_frame.pack(fill="x", pady=5)

        # 输出文件夹输入框和按钮
        output_input_frame = tk.Frame(output_frame)
        output_input_frame.pack(fill="x", pady=2)

        tk.Label(output_input_frame, text="文件夹:", width=8).pack(side=tk.LEFT)
        tk.Entry(output_input_frame, textvariable=self.three_d_output_dir, width=30).pack(side=tk.LEFT, padx=2, fill="x",
                                                                                          expand=True)
        tk.Button(output_input_frame, text="选择", command=self.browse_three_d_output_dir, width=8).pack(side=tk.LEFT, padx=2)

        # 处理选项
        options_frame = tk.Frame(output_frame)
        options_frame.pack(fill="x", pady=5)

        tk.Checkbutton(options_frame, text="创建子文件夹", variable=self.three_d_create_subfolder,
                       font=("Arial", 9)).pack(anchor="w")

        # 处理按钮区域
        button_frame = tk.Frame(left_frame)
        button_frame.pack(fill="x", pady=10)

        # 开始处理按钮
        self.three_d_start_button = tk.Button(button_frame, text="开始汇总", command=self.start_three_d_processing,
                                              bg="#4CAF50", fg="white", font=("Arial", 11, "bold"),
                                              height=2, state=tk.DISABLED)
        self.three_d_start_button.pack(fill="x", pady=2)

        # 停止处理按钮
        self.three_d_stop_button = tk.Button(button_frame, text="停止处理", command=self.stop_three_d_processing,
                                             bg="#FF9800", fg="white", height=1, state=tk.DISABLED)
        self.three_d_stop_button.pack(fill="x", pady=2)

        # === 右侧区域：结果显示 ===
        right_frame = tk.Frame(main_frame)
        main_frame.add(right_frame)

        # 结果显示区域
        result_frame = tk.LabelFrame(right_frame, text="处理结果", padx=8, pady=8)
        result_frame.pack(fill="both", expand=True)

        # 使用ScrolledText显示结果
        self.three_d_result_text_area = scrolledtext.ScrolledText(result_frame, wrap=tk.WORD, font=("Courier New", 9))
        self.three_d_result_text_area.pack(fill="both", expand=True)

        # 检查pandas是否可用
        if not PANDAS_AVAILABLE:
            warning_text = "警告: pandas库未安装，无法处理Excel文件。\n请运行: pip install pandas openpyxl\n"
            self.three_d_result_text_area.insert(1.0, warning_text)
            self.three_d_start_button.config(state=tk.DISABLED)

    def browse_three_d_files(self):
        """浏览并选择多个Excel文件"""
        self.browse_files_generic(
            title="选择Excel文件",
            filetypes=[("Excel files", "*.xlsx *.xls"), ("All files", "*.*")],
            file_paths_var="three_d_file_paths",
            tree_var="three_d_file_tree",
            count_label_var="three_d_file_count_label",
            output_dir_var="three_d_output_dir",
            start_button_var="three_d_start_button",
            pandas_check=True
        )

    def browse_three_d_output_dir(self):
        """浏览并选择输出文件夹"""
        self.browse_output_dir_generic(
            title="选择输出文件夹",
            output_dir_var="three_d_output_dir"
        )

    def clear_three_d_file_list(self):
        """清空三次元文件列表（三次元专用）"""
        self.clear_file_list_generic(
            file_paths_var="three_d_file_paths",
            tree_var="three_d_file_tree",
            count_label_var="three_d_file_count_label",
            start_button_var="three_d_start_button",
            pandas_check=True
        )

    def remove_selected_three_d_files(self):
        """移除选中的三次元文件（三次元专用）"""
        self.remove_selected_files_generic(
            tree_var="three_d_file_tree",
            file_paths_var="three_d_file_paths",
            count_label_var="three_d_file_count_label",
            start_button_var="three_d_start_button",
            pandas_check=True
        )

    def reorder_three_d_file_list(self):
        """重新排序三次元文件列表（三次元专用）"""
        self.reorder_file_list_generic(
            tree_var="three_d_file_tree",
            file_paths_var="three_d_file_paths"
        )

    def update_three_d_file_count(self):
        """更新三次元文件计数（三次元专用）"""
        self.update_file_count_generic(
            file_paths_var="three_d_file_paths",
            count_label_var="three_d_file_count_label",
            start_button_var="three_d_start_button",
            pandas_check=True
        )

    def show_three_d_file_context_menu(self, event):
        """显示三次元文件右键菜单（三次元专用）"""
        self.show_file_context_menu_generic(
            event=event,
            tree_var="three_d_file_tree",
            file_paths_var="three_d_file_paths",
            remove_method=self.remove_selected_three_d_files
        )

    def open_selected_three_d_file(self):
        """打开选中的三次元文件（三次元专用）"""
        self.open_selected_file_generic(
            tree_var="three_d_file_tree",
            file_paths_var="three_d_file_paths"
        )

    def open_three_d_file_folder(self):
        """打开三次元文件所在文件夹（三次元专用）"""
        self.open_file_folder_generic(
            tree_var="three_d_file_tree",
            file_paths_var="three_d_file_paths"
        )

    def start_three_d_processing(self):
        """开始三次元数据处理"""
        if not PANDAS_AVAILABLE:
            messagebox.showerror("错误", "pandas库未安装，无法处理Excel文件。\n请运行: pip install pandas openpyxl")
            return

        if not self.three_d_file_paths:
            messagebox.showerror("错误", "请先选择要处理的Excel文件！")
            return

        # 检查输出文件夹
        if not self.three_d_output_dir.get():
            messagebox.showerror("错误", "请选择输出文件夹！")
            return

        # 创建输出文件夹（如果需要）
        try:
            os.makedirs(self.three_d_output_dir.get(), exist_ok=True)
        except Exception as e:
            messagebox.showerror("错误", f"无法创建输出文件夹:\n{str(e)}")
            return

        # 禁用处理按钮，启用停止按钮
        self.toggle_three_d_buttons(processing=True)
        self.three_d_processing = True

        # 在新线程中处理
        thread = threading.Thread(target=self.process_three_d_files_thread, daemon=True)
        thread.start()

    def stop_three_d_processing(self):
        """停止处理"""
        self.three_d_processing = False
        self.status_bar.config(text="正在停止处理...")

    def toggle_three_d_buttons(self, processing):
        """切换三次元按钮状态（三次元专用）"""
        if processing:
            self.three_d_start_button.config(state=tk.DISABLED)
            self.three_d_stop_button.config(state=tk.NORMAL)
        else:
            self.three_d_start_button.config(state=tk.NORMAL if self.three_d_file_paths and PANDAS_AVAILABLE else tk.DISABLED)
            self.three_d_stop_button.config(state=tk.DISABLED)

    def process_three_d_files_thread(self):
        """处理三次元文件的线程函数（三次元专用）"""
        try:
            # 清空结果区域
            self.three_d_result_text_area.delete(1.0, tk.END)

            # 显示开始信息
            self.show_three_d_result_header(len(self.three_d_file_paths))

            # 汇总所有文件的数据
            success, message, output_file = self.three_d_processor.merge_three_d_files(
                self.three_d_file_paths, self.three_d_output_dir.get(), self.three_d_processing, 
                self.three_d_create_subfolder.get(), self.three_d_layout_mode.get()
            )

            if success:
                self.append_three_d_result(f"✓ {message}\n")
                self.status_bar.config(text=f"汇总完成: {os.path.basename(output_file)}")
            else:
                self.append_three_d_result(f"✗ {message}\n")
                self.status_bar.config(text="汇总失败")

        except Exception as e:
            self.append_three_d_result(f"✗ 处理失败: {str(e)}\n")
        finally:
            # 处理完成
            self.three_d_processing = False
            self.toggle_three_d_buttons(processing=False)

    def show_three_d_result_header(self, total_files):
        """显示三次元结果头部信息（三次元专用）"""
        result_info = f"{'=' * 70}\n"
        result_info += f"三次元数据汇总开始 - {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}\n"
        result_info += f"{'=' * 70}\n\n"

        result_info += f"文件总数: {total_files}\n"
        result_info += f"输出文件夹: {self.three_d_output_dir.get()}\n\n"
        result_info += f"{'=' * 70}\n"
        result_info += "处理结果:\n"
        result_info += f"{'=' * 70}\n\n"

        self.three_d_result_text_area.insert(1.0, result_info)

    def append_three_d_result(self, message):
        """在三次元结果区域追加信息（三次元专用）"""
        self.three_d_result_text_area.insert(tk.END, message)
        self.three_d_result_text_area.see(tk.END)
        self.root.update()
